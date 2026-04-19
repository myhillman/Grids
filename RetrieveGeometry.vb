Imports System.Net.Http
Imports System.Text
Imports System.Text.Json
Imports Microsoft.Data.Sqlite
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Union

Module RetrieveGeometry
    Public Sub InitializeOSMCache()

        Using conn As New SqliteConnection(DXCC_DATA)
            conn.Open()
            ' 1. Open DB
            ' 2. Ensure table exists
            Using cmd As New SqliteCommand("
            CREATE TABLE IF NOT EXISTS OsmGeometryCache (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                OsmId INTEGER NOT NULL,
                OsmType TEXT NOT NULL,
                Geometry TEXT NOT NULL,
                RetrievedUtc TEXT NOT NULL
            );", conn)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub
    Public Class OsmGeometryRecord
        Public Property OsmId As Long
        Public Property OsmType As String   ' "relation" or "way"
        Public Property Geometry As String  ' WKT or GeoJSON
        Public Property RetrievedUtc As DateTime
    End Class

    Public Sub StoreInCache(osmId As Long, osmType As String, geometry As String)
        Using conn As New SqliteConnection(DXCC_DATA)
            conn.Open()
            Using cmd As New SqliteCommand("
        INSERT INTO OsmGeometryCache (OsmId, OsmType, Geometry, RetrievedUtc)
        VALUES (@id, @type, @geom, @utc)", conn)

                cmd.Parameters.AddWithValue("@id", osmId)
                cmd.Parameters.AddWithValue("@type", osmType)
                cmd.Parameters.AddWithValue("@geom", geometry)
                cmd.Parameters.AddWithValue("@utc", DateTime.UtcNow.ToString("o"))

                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub


    Public Class OverpassClient
    Private ReadOnly _http As HttpClient = New HttpClient With {
        .Timeout = TimeSpan.FromSeconds(120)
    }
        Private Shared ReadOnly throttleLock As New Object()
        Private Shared lastRequest As DateTime = DateTime.MinValue
        Private Const MinSpacingMs As Integer = 1000   ' 1 second

        Public Shared Async Function ThrottleAsync() As Task
            Dim delay As Integer = 0
            Dim now = DateTime.UtcNow

            ' Calculate required delay
            SyncLock throttleLock
                Dim nextAllowed = lastRequest.AddMilliseconds(MinSpacingMs)

                If nextAllowed > now Then
                    delay = CInt((nextAllowed - now).TotalMilliseconds)
                End If
            End SyncLock

            ' Wait outside the lock
            If delay > 0 Then
                Await Task.Delay(delay)
            End If

            ' Update lastRequest AFTER waiting
            SyncLock throttleLock
                lastRequest = DateTime.UtcNow
            End SyncLock
        End Function

        Private Shared ReadOnly Endpoints As String() = {
    "https://overpass-api.de/api/interpreter",
    "https://lz4.overpass-api.de/api/interpreter",
    "https://overpass.kumi.systems/api/interpreter",
    "https://overpass.nchc.org.tw/api/interpreter"
}

        Public Async Function FetchAsync(osmId As Long, osmType As String) As Task(Of String)
            Dim query As String

            If osmType = "relation" Then
                query = $"[out:json][timeout:120];relation({osmId});(._;>>;);out body;"
            ElseIf osmType = "way" Then
                query = $"[out:json][timeout:120];way({osmId});(._;>>;);out body;"
            Else
                Throw New ArgumentException("osmType must be 'relation' or 'way'")
            End If

            For attempt = 1 To 5
                For Each ep In Endpoints
                    Try
                        Using client As New HttpClient()
                            client.Timeout = TimeSpan.FromSeconds(120)

                            Dim content = New StringContent(query, Encoding.UTF8, "application/x-www-form-urlencoded")
                            Debug.WriteLine($"Attemp {attempt} on endpoint {ep} with query: {query}")
                            Dim response = Await client.PostAsync(ep, content)
                            Debug.WriteLine($"Response code {response.ReasonPhrase}")
                            If response.IsSuccessStatusCode Then
                                Dim result = Await response.Content.ReadAsStringAsync()
                                Debug.WriteLine($"Returned {result.Count} bytes")
                                Return result
                            End If

                        End Using

                    Catch ex As TaskCanceledException
                        ' timeout → retry
                    Catch ex As HttpRequestException
                        ' 504, 429, etc → retry
                    End Try
                Next

                Await Task.Delay(1000 * attempt)
            Next

            Throw New Exception($"Overpass failed after multiple attempts for {osmType} {osmId}")
        End Function

    End Class

    Public Class OsmGeometryBuilder

        Private ReadOnly _geomFactory As New GeometryFactory()
        Private ReadOnly _writer As New WKTWriter()

        Public Function BuildGeometry(rawOsmJson As String) As String
            ' Parse JSON
            Dim doc = JsonDocument.Parse(rawOsmJson)
            Dim elements = doc.RootElement.GetProperty("elements")

            Dim nodes As New Dictionary(Of Long, NetTopologySuite.Geometries.Coordinate)
            Dim ways As New Dictionary(Of Long, List(Of NetTopologySuite.Geometries.Coordinate))
            Dim outerRings As New List(Of NetTopologySuite.Geometries.LinearRing)
            Dim innerRings As New List(Of NetTopologySuite.Geometries.LinearRing)

            ' -------------------------------
            ' PASS 1: Collect nodes
            ' -------------------------------
            For Each el In elements.EnumerateArray()
                If el.GetProperty("type").GetString() = "node" Then
                    Dim id = el.GetProperty("id").GetInt64()
                    Dim lat = el.GetProperty("lat").GetDouble()
                    Dim lon = el.GetProperty("lon").GetDouble()
                    nodes(id) = New NetTopologySuite.Geometries.Coordinate(lon, lat)
                End If
            Next

            ' -------------------------------
            ' PASS 2: Collect ways
            ' -------------------------------
            For Each el In elements.EnumerateArray()
                If el.GetProperty("type").GetString() = "way" Then
                    Dim id = el.GetProperty("id").GetInt64()
                    Dim coords As New List(Of NetTopologySuite.Geometries.Coordinate)
                    Dim missing As Boolean = False

                    For Each nd In el.GetProperty("nodes").EnumerateArray()
                        Dim nid = nd.GetInt64()
                        If nodes.ContainsKey(nid) Then
                            coords.Add(nodes(nid))
                        Else
                            missing = True
                        End If
                    Next

                    If Not missing AndAlso coords.Count >= 2 Then
                        ways(id) = coords
                    End If
                End If
            Next

            ' -------------------------------
            ' PASS 3: Detect multipolygon roles
            ' -------------------------------
            Dim hasRoles As Boolean = False

            For Each el In elements.EnumerateArray()
                If el.GetProperty("type").GetString() = "relation" Then
                    For Each mem In el.GetProperty("members").EnumerateArray()
                        If mem.GetProperty("type").GetString() = "way" Then
                            Dim wid = mem.GetProperty("ref").GetInt64()
                            Dim role = mem.GetProperty("role").GetString()

                            ' Only "outer" and "inner" count as roles
                            If role = "outer" OrElse role = "inner" Then
                                hasRoles = True
                            End If

                            If ways.ContainsKey(wid) AndAlso (role = "outer" OrElse role = "inner") Then
                                Dim coords = ways(wid).ToArray()

                                ' Ensure closed
                                If Not coords.First().Equals2D(coords.Last()) Then
                                    coords = coords.Concat({coords.First()}).ToArray()
                                End If

                                ' Must be valid ring
                                If coords.Length >= 4 AndAlso
                                   coords.Select(Function(c) (c.X, c.Y)).Distinct().Count() >= 3 Then

                                    Dim ring = _geomFactory.CreateLinearRing(coords)

                                    If role = "outer" Then
                                        outerRings.Add(ring)
                                    ElseIf role = "inner" Then
                                        innerRings.Add(ring)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next

            ' ---------------------------------------------------------
            ' CASE 1: Proper multipolygon with roles
            ' ---------------------------------------------------------
            If hasRoles AndAlso outerRings.Count > 0 Then
                Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)

                For Each outer In outerRings
                    Dim outerPoly As NetTopologySuite.Geometries.Polygon = _geomFactory.CreatePolygon(outer)
                    Dim holes As New List(Of NetTopologySuite.Geometries.LinearRing)

                    For Each inner In innerRings
                        Dim innerPoly As NetTopologySuite.Geometries.Polygon = _geomFactory.CreatePolygon(inner)
                        If outerPoly.Covers(innerPoly) Then
                            holes.Add(inner)
                        End If
                    Next

                    polys.Add(_geomFactory.CreatePolygon(outer, holes.ToArray()))
                Next

                Dim mp = _geomFactory.CreateMultiPolygon(polys.ToArray())
                Return _writer.Write(mp)
            End If

            ' ---------------------------------------------------------
            ' CASE 2: Coastline-style relation → stitch ways
            ' ---------------------------------------------------------

            ' Convert all ways to LineStrings
            Dim lineStrings As New List(Of NetTopologySuite.Geometries.LineString)
            For Each kvp In ways
                Dim coords = kvp.Value.ToArray()
                lineStrings.Add(_geomFactory.CreateLineString(coords))
            Next

            ' Merge connected lines
            Dim merger As New NetTopologySuite.Operation.Linemerge.LineMerger()
            merger.Add(lineStrings)
            Dim merged = merger.GetMergedLineStrings().Cast(Of NetTopologySuite.Geometries.LineString)().ToList()

            ' Polygonize merged lines
            Dim polygonizer As New NetTopologySuite.Operation.Polygonize.Polygonizer()
            polygonizer.Add(merged)

            Dim polys2 = polygonizer.GetPolygons().Cast(Of NetTopologySuite.Geometries.Polygon)().ToList()

            If polys2.Count = 0 Then
                Throw New Exception("Polygonizer produced no polygons.")
            End If

            ' Union all polygons
            Dim unioned = NetTopologySuite.Operation.Union.UnaryUnionOp.Union(polys2)

            Return _writer.Write(unioned)
        End Function


    End Class

    Public Class OsmGeometryCacheService
        Public Property CacheHits As Integer = 0
        Public Property CacheMisses As Integer = 0

        Private ReadOnly _conn As SqliteConnection
        Private ReadOnly _overpass As OverpassClient = New OverpassClient()
        Private ReadOnly _builder As OsmGeometryBuilder = New OsmGeometryBuilder()

        Public Sub New(conn As SqliteConnection)
            _conn = conn
        End Sub

        Private Function TryLoadFromCache(osmId As Long, osmType As String) As String
            Using cmd As New SqliteCommand("
            SELECT Geometry 
            FROM OsmGeometryCache 
            WHERE OsmId = @id AND OsmType = @type
            LIMIT 1", _conn)

                cmd.Parameters.AddWithValue("@id", osmId)
                cmd.Parameters.AddWithValue("@type", osmType)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    CacheHits += 1
                    Return CStr(result)
                End If
            End Using
            CacheMisses += 1
            Return Nothing
        End Function
        Private Sub StoreInCache(osmId As Long, osmType As String, geometry As String)
            Using cmd As New SqliteCommand("
            INSERT INTO OsmGeometryCache (OsmId, OsmType, Geometry, RetrievedUtc)
            VALUES (@id, @type, @geom, @utc)", _conn)

                cmd.Parameters.AddWithValue("@id", osmId)
                cmd.Parameters.AddWithValue("@type", osmType)
                cmd.Parameters.AddWithValue("@geom", geometry)
                cmd.Parameters.AddWithValue("@utc", DateTime.UtcNow.ToString("o"))

                cmd.ExecuteNonQuery()
            End Using
        End Sub

        Public Async Function GetGeometryAsync(osmId As Long, osmType As String) As Task(Of String)
            Dim cached = TryLoadFromCache(osmId, osmType)
            If cached IsNot Nothing Then Return cached

            Dim raw = Await _overpass.FetchAsync(osmId, osmType)
            Dim geom = _builder.BuildGeometry(raw)

            StoreInCache(osmId, osmType, geom)

            Return geom
        End Function
    End Class

    Public Class DxccPart
        Public Property OsmId As Long
        Public Property OsmType As String   ' "relation" or "way"
    End Class

    Public Class DxccDefinition
        Public Property Name As String
        Public Property Parts As List(Of DxccPart)
    End Class

    Public Class DxccGeometryService
        Private ReadOnly _cache As OsmGeometryCacheService
        Private ReadOnly _reader As WKTReader
        Private ReadOnly _writer As WKTWriter

        Public Sub New(cache As OsmGeometryCacheService)
            _cache = cache
            _reader = New WKTReader()
            _writer = New WKTWriter()
        End Sub
        
    End Class
    Public Function ParseExpression(expr As String) As List(Of (Op As String, OsmId As Long, OsmType As String))
        Dim tokens As New List(Of (String, Long, String))

        Dim cleaned = expr.Replace(" ", "")
        Dim i As Integer = 0

        ' First token must be an ID (relation or way)
        Dim first = ParseId(cleaned, i)
        tokens.Add(("+", first.OsmId, first.OsmType))

        While i < cleaned.Length
            Dim op = cleaned(i)
            i += 1

            Dim nextId = ParseId(cleaned, i)
            tokens.Add((op, nextId.OsmId, nextId.OsmType))
        End While

        Return tokens
    End Function

    Private Function ParseId(expr As String, ByRef i As Integer) As (OsmId As Long, OsmType As String)
        Dim osmType As String = "relation"

        If expr(i) = "W"c Or expr(i) = "w"c Then
            osmType = "way"
            i += 1
        End If

        Dim start = i
        While i < expr.Length AndAlso Char.IsDigit(expr(i))
            i += 1
        End While

        Dim id = Long.Parse(expr.Substring(start, i - start))
        Return (id, osmType)
    End Function
    Public Async Function EvaluateExpressionAsync(expr As String,
                                              cache As OsmGeometryCacheService) _
                                              As Task(Of NetTopologySuite.Geometries.Geometry)

        Dim tokens = ParseExpression(expr)

        Dim reader As New WKTReader()

        ' Load and clean first geometry
        Dim first = tokens(0)
        Dim currentWkt = Await cache.GetGeometryAsync(first.OsmId, first.OsmType)
        Dim currentGeom = reader.Read(currentWkt).Buffer(0)   ' CLEAN

        ' Apply remaining operations
        For i = 1 To tokens.Count - 1
            Dim t = tokens(i)

            Dim nextWkt = Await cache.GetGeometryAsync(t.OsmId, t.OsmType)
            Dim nextGeom = reader.Read(nextWkt).Buffer(0)     ' CLEAN

            If t.Op = "+"c Then
                Dim gc = currentGeom.Factory.CreateGeometryCollection({currentGeom, nextGeom})

                ' UnaryUnionOp is in NetTopologySuite.Operation.Union
                Dim uu = New UnaryUnionOp(CType(gc, NetTopologySuite.Geometries.Geometry))
                currentGeom = uu.Union().Buffer(0)

            ElseIf t.Op = "-"c Then
                currentGeom = currentGeom.Difference(nextGeom).Buffer(0)
            End If

        Next

        ' Final clean
        currentGeom = currentGeom.Buffer(0)
        Debug.WriteLine(currentGeom.AsText())

        Return currentGeom
    End Function

    Public Function NtsToEsri(nts As NetTopologySuite.Geometries.Geometry) As Esri.ArcGISRuntime.Geometry.Geometry
        Dim esriJson = NtsToEsriJson(nts)
        Return Esri.ArcGISRuntime.Geometry.Geometry.FromJson(esriJson)
    End Function
    Private Function NtsToEsriJson(geom As NetTopologySuite.Geometries.Geometry) As String
        If TypeOf geom Is NetTopologySuite.Geometries.Polygon Then
            Return PolygonToEsriJson(CType(geom, NetTopologySuite.Geometries.Polygon))

        ElseIf TypeOf geom Is NetTopologySuite.Geometries.MultiPolygon Then
            Return MultiPolygonToEsriJson(CType(geom, NetTopologySuite.Geometries.MultiPolygon))

        Else
            Throw New NotSupportedException("Only Polygon and MultiPolygon supported.")
        End If
    End Function

    Private Function PolygonToEsriJson(poly As NetTopologySuite.Geometries.Polygon) As String
        Dim rings As New List(Of List(Of Double()))

        rings.Add(CoordsToList(poly.ExteriorRing.Coordinates))

        For i = 0 To poly.NumInteriorRings - 1
            rings.Add(CoordsToList(poly.GetInteriorRingN(i).Coordinates))
        Next

        Dim obj = New With {
        .rings = rings,
        .spatialReference = New With {.wkid = 4326}
    }

        Return JsonSerializer.Serialize(obj)
    End Function

    Private Function MultiPolygonToEsriJson(mp As NetTopologySuite.Geometries.MultiPolygon) As String
        Dim rings As New List(Of List(Of Double()))

        For i = 0 To mp.NumGeometries - 1
            Dim poly As NetTopologySuite.Geometries.Polygon =
            CType(mp.GetGeometryN(i), NetTopologySuite.Geometries.Polygon)

            ' Exterior ring
            rings.Add(CoordsToList(poly.ExteriorRing.Coordinates))

            ' Interior rings (holes)
            For h = 0 To poly.NumInteriorRings - 1
                rings.Add(CoordsToList(poly.GetInteriorRingN(h).Coordinates))
            Next
        Next

        Dim obj = New With {
        .rings = rings,
        .spatialReference = New With {.wkid = 4326}
    }

        Return JsonSerializer.Serialize(obj)
    End Function

    Private Function CoordsToList(coords As NetTopologySuite.Geometries.Coordinate()) As List(Of Double())
        Dim list As New List(Of Double())
        For Each c In coords
            list.Add({c.X, c.Y})
        Next
        Return list
    End Function

End Module
