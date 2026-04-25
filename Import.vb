Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Security.Policy
Imports System.Text.Json.Nodes
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Xml.Linq
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports HtmlAgilityPack
Imports Microsoft.VisualBasic.FileIO
Imports NetTopologySuite
Imports NetTopologySuite.Features
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Linemerge
Imports NetTopologySuite.Operation.Union
Imports NetTopologySuite.Precision
Imports NetTopologySuite.Simplify


Module Import
    ' Module contains all import functions
    Sub ImportISO3166()
        ' Import ISO3166-1 country codes
        Dim sql As SqliteCommand, updated As Integer = 0
        Using connect As New SqliteConnection(DXCC_DATA),
              ISO3166 As New TextFieldParser($"{Application.StartupPath}\ISO3166.csv")
            connect.Open()
            sql = connect.CreateCommand
            With ISO3166
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(",")
            End With
            Dim lines As Integer = 1
            Try
                sql.CommandText = $"DELETE FROM `ISO31661`"    ' remove existing data
                sql.ExecuteNonQuery()

                While Not ISO3166.EndOfData
                    If lines > 1 Then
                        Dim currentRow = ISO3166.ReadFields
                        sql.CommandText = $"INSERT INTO `ISO31661`(`Entity`,`Code`) VALUES('{SQLescape(currentRow(0))}','{currentRow(1)}')"
                        updated += sql.ExecuteNonQuery
                    End If
                    lines += 1
                End While
            Catch ex As MalformedLineException
                MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
            AppendText(Form1.TextBox1, $"{updated} ISO6133 codes imported{vbCrLf}")
        End Using
    End Sub

    Sub ImportEUASBorder()

        Dim factory As GeometryFactory = NtsGeometryServices.Instance.CreateGeometryFactory(4326)
        Dim segments As New List(Of LineString)

        ' ------------------------------------------------------------
        ' LOAD KML
        ' ------------------------------------------------------------
        Dim doc = XDocument.Load(Path.Combine(Application.StartupPath, "BorderEUAS.kml"))
        Dim ns = doc.Root.Name.Namespace
        Dim linestrings = doc.Descendants(ns + "LineString")

        AppendText(Form1.TextBox1, $"{linestrings.Count} linestrings loaded{vbCrLf}")

        ' ------------------------------------------------------------
        ' PARSE EACH <coordinates> BLOCK
        ' ------------------------------------------------------------
        For Each coordinates In linestrings.Descendants(ns + "coordinates")

            Dim raw = Regex.Replace(coordinates.Value, "[^0-9\-\., ]", "").Trim()
            Dim coordPairs = raw.Split(" "c, StringSplitOptions.RemoveEmptyEntries)

            Dim pts As New List(Of Coordinate)

            For Each pair In coordPairs
                Dim parts = pair.Split(","c)
                If parts.Length >= 2 Then
                    Dim lon = Math.Round(Double.Parse(parts(0)), 5)
                    Dim lat = Math.Round(Double.Parse(parts(1)), 5)
                    pts.Add(New Coordinate(lon, lat))
                End If
            Next

            If pts.Count >= 2 Then
                segments.Add(factory.CreateLineString(pts.ToArray()))
            End If
        Next

        ' ------------------------------------------------------------
        ' SNAP TO 1e-5 GRID
        ' ------------------------------------------------------------
        Dim pm As New PrecisionModel(100000.0)
        Dim reducer As New GeometryPrecisionReducer(pm)

        Dim snapped As New List(Of LineString)
        For Each ls In segments
            snapped.Add(CType(reducer.Reduce(ls), LineString))
        Next

        ' ------------------------------------------------------------
        ' COLLECT ALL ENDPOINTS
        ' ------------------------------------------------------------
        Dim endpoints As New List(Of Coordinate)

        For Each ls In snapped
            endpoints.Add(ls.StartPoint.Coordinate)
            endpoints.Add(ls.EndPoint.Coordinate)
        Next

        endpoints = endpoints.Distinct(New CoordinateEqualityComparer()).ToList()

        ' ------------------------------------------------------------
        ' BUILD COMPLETE GRAPH (Euclidean distances)
        ' ------------------------------------------------------------
        Dim n = endpoints.Count
        Dim edges As New List(Of Tuple(Of Integer, Integer, Double))

        For i = 0 To n - 1
            For j = i + 1 To n - 1
                Dim dx = endpoints(i).X - endpoints(j).X
                Dim dy = endpoints(i).Y - endpoints(j).Y
                Dim d2 = dx * dx + dy * dy
                edges.Add(Tuple.Create(i, j, d2))
            Next
        Next

        ' ------------------------------------------------------------
        ' KRUSKAL MST
        ' ------------------------------------------------------------
        edges = edges.OrderBy(Function(e) e.Item3).ToList()

        Dim parent(n - 1) As Integer
        For i = 0 To n - 1
            parent(i) = i
        Next

        Dim Find As Func(Of Integer, Integer) =
        Function(x)
            While parent(x) <> x
                parent(x) = parent(parent(x))
                x = parent(x)
            End While
            Return x
        End Function

        Dim Union As Action(Of Integer, Integer) =
        Sub(a, b)
            parent(Find(a)) = Find(b)
        End Sub

        Dim adj As New Dictionary(Of Integer, List(Of Integer))
        For i = 0 To n - 1
            adj(i) = New List(Of Integer)
        Next

        For Each e In edges
            Dim a = e.Item1
            Dim b = e.Item2
            If Find(a) <> Find(b) Then
                Union(a, b)
                adj(a).Add(b)
                adj(b).Add(a)
            End If
        Next

        ' ------------------------------------------------------------
        ' FIND LONGEST PATH IN MST (two BFS passes)
        ' ------------------------------------------------------------
        Dim BFS As Func(Of Integer, Tuple(Of Integer, Dictionary(Of Integer, Integer))) =
        Function(start)
            Dim q As New Queue(Of Integer)
            Dim dist As New Dictionary(Of Integer, Integer)
            Dim parentNode As New Dictionary(Of Integer, Integer)

            q.Enqueue(start)
            dist(start) = 0
            parentNode(start) = -1

            While q.Count > 0
                Dim u = q.Dequeue()
                For Each v In adj(u)
                    If Not dist.ContainsKey(v) Then
                        dist(v) = dist(u) + 1
                        parentNode(v) = u
                        q.Enqueue(v)
                    End If
                Next
            End While

            Dim far = dist.OrderByDescending(Function(kv) kv.Value).First().Key
            Return Tuple.Create(far, parentNode)
        End Function

        Dim first = BFS(0).Item1
        Dim secondResult = BFS(first)
        Dim last = secondResult.Item1
        Dim parentMap = secondResult.Item2

        ' ------------------------------------------------------------
        ' RECONSTRUCT LONGEST PATH
        ' ------------------------------------------------------------
        Dim chain As New List(Of Coordinate)
        Dim cur = last

        While cur <> -1
            chain.Add(endpoints(cur))
            cur = parentMap(cur)
        End While

        chain.Reverse()

        ' ------------------------------------------------------------
        ' BUILD LINESTRING
        ' ------------------------------------------------------------
        Dim fullBorder As LineString = factory.CreateLineString(chain.ToArray())

        ' ------------------------------------------------------------
        ' SIMPLIFY TO 1 KM
        ' ------------------------------------------------------------
        Dim tolDeg As Double = 1000 / 111000.0
        Dim simplified As NetTopologySuite.Geometries.Geometry = DouglasPeuckerSimplifier.Simplify(fullBorder, tolDeg)

        AppendText(Form1.TextBox1, $" Border simplified to {simplified.NumPoints} points{vbCrLf}")

        ' ------------------------------------------------------------
        ' WRITE GEOJSON
        ' ------------------------------------------------------------
        Dim writer As New GeoJsonWriter()
        Dim json = writer.Write(simplified)

        File.WriteAllText(Path.Combine(Application.StartupPath, "BorderEUAS.json"), json)

        AppendText(Form1.TextBox1, $"Done{vbCrLf}")

    End Sub

    ' ------------------------------------------------------------
    ' SUPPORT: Coordinate comparer
    ' ------------------------------------------------------------
    Public Class CoordinateEqualityComparer
        Implements IEqualityComparer(Of Coordinate)

        Public Overloads Function Equals(a As Coordinate, b As Coordinate) As Boolean _
        Implements IEqualityComparer(Of Coordinate).Equals
            Return a.X = b.X AndAlso a.Y = b.Y
        End Function

        Public Overloads Function GetHashCode(c As Coordinate) As Integer _
        Implements IEqualityComparer(Of Coordinate).GetHashCode
            Return c.X.GetHashCode() Xor c.Y.GetHashCode()
        End Function
    End Class

    Private Function SimplifyLinework(
                                      geom As NetTopologySuite.Geometries.Geometry,
                                      tolerance As Double
                                      ) As NetTopologySuite.Geometries.Geometry

        Dim simplifier As New NetTopologySuite.Simplify.DouglasPeuckerSimplifier(geom)
        simplifier.DistanceTolerance = tolerance
        Return simplifier.GetResultGeometry()
    End Function

    Public Class zonedata
        Public Property Id As String
        Public Property Polyline As String
        Public Property Mask As String
        Public Property Color As String
        Public Property Lat As String
        Public Property Lon As String
        Public Property Description As String

        Public Sub New(data As List(Of String))
            Id = data(0)
            Polyline = data(1)
            Mask = data(2)
            Color = data(3)
            Lat = data(4)
            Lon = data(5)
            Description = data(6)
        End Sub
    End Class
    Public Async Sub ImportCQITUZones()

        ' Import CQ and ITU zones by reverse engineering some old KML files by IV3TMM
        Dim ZoneCheck As New Dictionary(Of String, String) From {
        {"CQ", "CQ.KML"},
        {"ITU", "ITU.KML"}
    }
        For Each ZoneType In ZoneCheck
            DecodeZone(ZoneType.Key, ZoneType.Value)
        Next
    End Sub

    Sub DecodeZone(ZoneType As String, FileName As String)

        Dim doc As XDocument = XDocument.Load(FileName)
        Dim ns As XNamespace = "http://www.opengis.net/kml/2.2"

        Dim linepattern As String, areapattern As String, AreaFolder As String
        Select Case ZoneType
            Case "CQ" : linepattern = "^Linea-(\d+)$" : areapattern = "^WAZ\s(\d+)$" : AreaFolder = "WAZ Area"
            Case "ITU" : linepattern = "^XXX-(\d+)$" : areapattern = "^ITU\s(\d+)$" : AreaFolder = "ITU Area"
            Case Else
                Throw New Exception($"Unrecognised zone type {ZoneType}")
        End Select
        ' ---------------------------------------------------------
        ' 1. Find the <Folder> with <name>Boundary</name> to extract zone lines
        ' ---------------------------------------------------------
        Dim boundariesFolder As XElement = doc.Descendants(ns + "Folder").FirstOrDefault(Function(f) (f.Element(ns + "name")?.Value.Trim() = "Boundary"))
        If boundariesFolder Is Nothing Then Throw New Exception($"No boundary folder found")

        Dim areasFolder As XElement =
                doc.Descendants(ns + "Folder").FirstOrDefault(Function(f) (f.Element(ns + "name")?.Value.Trim() = AreaFolder))
        If areasFolder Is Nothing Then Throw New Exception($"No areas folder found")

        ' ---------------------------------------------------------
        ' 2. Extract all <Placemark> nodes inside that folder
        ' ---------------------------------------------------------
        Dim Placemarks As IEnumerable(Of XElement) = boundariesFolder.Descendants(ns + "Placemark")
        If Placemarks Is Nothing Then Throw New Exception($"No Placemarks found in boundary folder")

        Dim AreaPlacemarks As IEnumerable(Of XElement) = areasFolder.Descendants(ns + "Placemark")
        If AreaPlacemarks Is Nothing Then Throw New Exception($"No Placemarks found in Area folder")

        ' ---------------------------------------------------------
        ' 3. Extract all <coordinates> nodes from each placemark
        ' ---------------------------------------------------------
        Dim coords As New List(Of Coordinate)

        Dim line = 1, number As Integer, gf As New GeometryFactory()
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            ' remove existing lines for this zone type
            Dim deleteCmd = connect.CreateCommand()
            deleteCmd.CommandText = $"DELETE FROM ZoneLines WHERE Type = '{ZoneType}'"
            deleteCmd.ExecuteNonQuery()
            Dim delete1Cmd = connect.CreateCommand()
            delete1Cmd.CommandText = $"DELETE FROM ZoneLabels WHERE Type = '{ZoneType}'"
            delete1Cmd.ExecuteNonQuery()

            Dim linescmd = connect.CreateCommand()
            linescmd.CommandText = "INSERT INTO ZoneLines (Type, zone, geometry) VALUES ($type, $zone, $geom);"

            Dim labelscmd = connect.CreateCommand()
            labelscmd.CommandText = "INSERT INTO ZoneLabels (Type, zone, geometry) VALUES ($type, $zone, $geom);"
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = Placemarks.Count
            End With
            For Each pm In Placemarks
                ' get the name
                Dim nameElement As XElement = pm.Element(ns + "name")

                Dim m = Regex.Match(nameElement.Value, linepattern)
                If m.Success Then
                    number = Integer.Parse(m.Groups(1).Value)
                End If
                ' get the coordinates
                Dim coordsElement As XElement = pm.Descendants(ns + "coordinates").FirstOrDefault()
                Dim coordinates = Split(coordsElement.Value.Trim, " ")
                coords.Clear()
                For Each grp In coordinates
                    Dim values = Split(grp, ",")
                    Dim coord As New Coordinate(values(0), values(1))
                    coords.Add(coord)
                Next
                ' Create the NTS LineString
                Dim rawLine As LineString = gf.CreateLineString(coords.ToArray())

                ' Split at anti-meridian
                Dim splitGeom As NetTopologySuite.Geometries.Geometry = SplitAntiMeridian(rawLine, gf)

                ' Densify each segment
                Dim densified As NetTopologySuite.Geometries.Geometry = DensifyGeometry(splitGeom, 5.0)
                ' Handle single or multiple LineStrings
                If TypeOf densified Is LineString Then

                    ' Single segment
                    Dim geoJson = FromNTSToGeoJson(densified)
                    linescmd.Parameters.Clear()
                    linescmd.Parameters.AddWithValue("$type", ZoneType)
                    linescmd.Parameters.AddWithValue("$zone", number)
                    linescmd.Parameters.AddWithValue("$geom", geoJson)
                    linescmd.ExecuteNonQuery()

                ElseIf TypeOf densified Is GeometryCollection Then

                    ' Multiple segments
                    Dim gc As GeometryCollection = CType(densified, GeometryCollection)

                    For i = 0 To gc.NumGeometries - 1
                        Dim seg As NetTopologySuite.Geometries.Geometry = gc.GetGeometryN(i)

                        If TypeOf seg Is LineString Then
                            Dim geoJson = FromNTSToGeoJson(seg)
                            linescmd.Parameters.Clear()
                            linescmd.Parameters.AddWithValue("$type", ZoneType)
                            linescmd.Parameters.AddWithValue("$zone", number)
                            linescmd.Parameters.AddWithValue("$geom", geoJson)
                            linescmd.ExecuteNonQuery()
                        End If
                    Next

                Else
                    Throw New Exception("Unexpected geometry type after AM split + densify.")
                End If
                line += 1
                Form1.ProgressBar1.Value += 1
            Next
            ' Extract zone labels
            line = 1
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = AreaPlacemarks.Count
            End With
            For Each area In AreaPlacemarks
                Dim nameElement As XElement = area.Element(ns + "name")
                Dim m = Regex.Match(nameElement.Value, areapattern)
                If m.Success Then
                    number = Integer.Parse(m.Groups(1).Value)
                End If
                Dim coordsElement As XElement = area.Descendants(ns + "coordinates").FirstOrDefault()
                Dim point = coordsElement.Value.Trim()
                Dim coord = Split(point, ",")
                Dim labelpoint = New Coordinate(coord(0), coord(1))
                Dim pt As Point = gf.CreatePoint(labelpoint)
                Dim GeoJson = FromNTSToGeoJson(pt)
                labelscmd.Parameters.Clear()
                labelscmd.Parameters.AddWithValue("$type", ZoneType)
                labelscmd.Parameters.AddWithValue("$zone", number)
                labelscmd.Parameters.AddWithValue("$geom", GeoJson)
                labelscmd.ExecuteNonQuery()
                line += 1
                Form1.ProgressBar1.Value += 1
            Next
            AppendText(Form1.TextBox1, $"Imported {line - 1} {ZoneType} lines{vbCrLf}")
        End Using
    End Sub
    ''' <summary>
    ''' Splits a LineString into multiple LineStrings such that no segment crosses
    ''' the anti-meridian (±180°). Longitudes are normalized to [-180, 180].
    ''' </summary>
    Private Function SplitAntiMeridian(ls As LineString, gf As GeometryFactory) As NetTopologySuite.Geometries.Geometry

        Dim parts As New List(Of LineString)
        Dim current As New List(Of Coordinate)

        Dim prev As Coordinate = Nothing

        For i = 0 To ls.NumPoints - 1

            Dim c As Coordinate = ls.GetCoordinateN(i)
            Dim lon As Double = NormalizeLon(c.X)
            Dim lat As Double = c.Y
            Dim curr As New Coordinate(lon, lat)

            If prev IsNot Nothing Then

                Dim prevLon = prev.X
                Dim currLon = curr.X

                ' Check if segment crosses AM
                If Math.Abs(prevLon - currLon) > 180 Then

                    ' Compute intersection latitude at AM
                    Dim targetLon As Double = If(prevLon > 0, 180, -180)
                    Dim t As Double = (targetLon - prevLon) / (currLon - prevLon)
                    Dim interLat As Double = prev.Y + t * (curr.Y - prev.Y)

                    ' Add intersection point to current segment
                    current.Add(New Coordinate(targetLon, interLat))

                    ' Close current segment
                    parts.Add(gf.CreateLineString(current.ToArray()))

                    ' Start new segment
                    current = New List(Of Coordinate)
                    current.Add(New Coordinate(If(targetLon = 180, -180, 180), interLat))
                End If
            End If

            current.Add(curr)
            prev = curr
        Next

        ' Add final segment
        If current.Count >= 2 Then
            parts.Add(gf.CreateLineString(current.ToArray()))
        End If

        If parts.Count = 1 Then
            Return parts(0)
        End If

        Return gf.CreateGeometryCollection(parts.Cast(Of NetTopologySuite.Geometries.Geometry).ToArray())
    End Function

    ''' <summary>
    ''' Normalizes longitude to [-180, 180].
    ''' </summary>
    Private Function NormalizeLon(lon As Double) As Double
        While lon > 180
            lon -= 360
        End While
        While lon < -180
            lon += 360
        End While
        Return lon
    End Function

    ''' <summary>
    ''' Imports Natural Earth timezone polygons from a GeoJSON file, groups all features
    ''' by their timezone name, and produces a single merged <c>MultiPolygon</c> for each
    ''' timezone.  
    ''' <para>
    ''' Each timezone in the source dataset may contain many individual polygons. These
    ''' are collected, unioned, simplified, and orientation-corrected before being
    ''' written to the <c>Timezones</c> table as one consolidated geometry.
    ''' </para>
    ''' <para>
    ''' The <c>places</c> field is generated by combining the distinct <c>places</c>
    ''' attributes from all polygons belonging to the same timezone. The <c>color</c>
    ''' field is taken from the first feature in each group.
    ''' </para>
    ''' </summary>
    ''' <remarks>
    ''' This method clears the <c>Timezones</c> table before inserting new rows.
    ''' Geometry is stored as GeoJSON.  
    ''' The operation is asynchronous and performs all database writes using a single
    ''' open SQLite connection.
    ''' </remarks>
    ''' <returns>
    ''' A task representing the asynchronous import operation.
    ''' </returns>
    ''' <exception cref="IOException">
    ''' Thrown if the GeoJSON file cannot be read.
    ''' </exception>
    ''' <exception cref="JsonException">
    ''' Thrown if the GeoJSON content is invalid or cannot be parsed.
    ''' </exception>
    ''' <exception cref="SqliteException">
    ''' Thrown if the database insert operations fail.
    ''' </exception>

    Async Function ImportTimeZones() As Task

        Dim geoJsonPath As String = "D:\GIS Data\Natural Earth\Timezones\ne_10m_time_zones.json"

        ' ---------------------------------------------------------
        ' 1. Load GeoJSON
        ' ---------------------------------------------------------
        Dim reader As New GeoJsonReader()
        Dim fc As NetTopologySuite.Features.FeatureCollection

        Using sr As New StreamReader(geoJsonPath)
            Dim json As String = Await sr.ReadToEndAsync()
            fc = reader.Read(Of NetTopologySuite.Features.FeatureCollection)(json)
        End Using

        ' ---------------------------------------------------------
        ' 2. Group features by timezone name
        ' ---------------------------------------------------------
        Dim groups =
        fc.GroupBy(Function(f) f.Attributes("name").ToString()) _
          .OrderBy(Function(g) Convert.ToSingle(g.Key)) ' numeric sort

        Dim gf As New GeometryFactory()
        Dim writer As New GeoJsonWriter()

        Using conn As New SqliteConnection(DXCC_DATA)
            Await conn.OpenAsync()

            ' Clear table
            Using delCmd = conn.CreateCommand()
                delCmd.CommandText = "DELETE FROM Timezones;"
                delCmd.ExecuteNonQuery()
            End Using

            ' Prepare insert
            Using cmd = conn.CreateCommand()
                cmd.CommandText =
                "INSERT INTO Timezones (name, places, color, geometry) VALUES ($name, $places, $color, $geometry);"

                cmd.Parameters.Add("$name", SqliteType.Text)
                cmd.Parameters.Add("$places", SqliteType.Text)
                cmd.Parameters.Add("$color", SqliteType.Text)
                cmd.Parameters.Add("$geometry", SqliteType.Text)

                With Form1.ProgressBar1
                    .Minimum = 0
                    .Maximum = groups.Count()
                    .Value = 0
                End With

                ' ---------------------------------------------------------
                ' 3. Process each timezone group
                ' ---------------------------------------------------------
                For Each g In groups

                    Dim tzName As String = g.Key

                    ' Merge places
                    Dim allPlaces As String =
                    String.Join(", ",
                        g.Select(Function(f) f.Attributes("places").ToString()) _
                         .Distinct())

                    ' Pick color (all polygons for a zone share the same)
                    Dim color As String = g.First().Attributes("map_color6").ToString()

                    ' Collect all geometries
                    Dim geoms As New List(Of NetTopologySuite.Geometries.Geometry)
                    For Each f In g
                        Dim ntsGeom As NetTopologySuite.Geometries.Geometry = f.Geometry
                        If ntsGeom IsNot Nothing AndAlso Not ntsGeom.IsEmpty Then
                            geoms.Add(ntsGeom)
                        End If
                    Next

                    ' Build MultiPolygon
                    Dim merged As NetTopologySuite.Geometries.Geometry = gf.BuildGeometry(geoms).Union()

                    ' Simplify
                    Dim simplified = NetTopologySuite.Simplify.TopologyPreservingSimplifier.Simplify(merged, 0.1)
                    If simplified Is Nothing OrElse simplified.IsEmpty Then simplified = merged

                    ' Convert to GeoJSON
                    Dim geoJson As String = writer.Write(simplified)

                    ' Insert
                    cmd.Parameters("$name").Value = tzName
                    cmd.Parameters("$places").Value = allPlaces
                    cmd.Parameters("$color").Value = color
                    cmd.Parameters("$geometry").Value = geoJson

                    cmd.ExecuteNonQuery()

                    Form1.ProgressBar1.Value += 1
                Next
            End Using
        End Using

        AppendText(Form1.TextBox1, $"Imported {groups.Count()} merged timezones{vbCrLf}")

    End Function
    ''' <summary>
    ''' Ensures that polygon and multipolygon rings follow the correct KML/GeoJSON
    ''' orientation rules.  
    ''' <para>
    ''' Exterior rings are forced to counter‑clockwise orientation, while interior
    ''' rings (holes) are forced to clockwise orientation. This guarantees consistent
    ''' winding order for downstream consumers and prevents rendering issues in
    ''' mapping engines.
    ''' </para>
    ''' </summary>
    ''' <param name="geom">
    ''' The input geometry whose ring orientation should be normalized.
    ''' </param>
    ''' <returns>
    ''' A geometry with corrected ring orientation. The original geometry is returned
    ''' unchanged if no orientation adjustments are required.
    ''' </returns>

    Private Function FixOrientation(
                                    geom As NetTopologySuite.Geometries.Geometry
                                    ) As NetTopologySuite.Geometries.Geometry

        If TypeOf geom Is NetTopologySuite.Geometries.Polygon Then
            Return FixPolygonOrientation(DirectCast(geom, NetTopologySuite.Geometries.Polygon))
        End If

        If TypeOf geom Is NetTopologySuite.Geometries.MultiPolygon Then
            Dim mp = DirectCast(geom, NetTopologySuite.Geometries.MultiPolygon)
            Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)

            For i = 0 To mp.NumGeometries - 1
                polys.Add(FixPolygonOrientation(DirectCast(mp.GetGeometryN(i),
                                                           NetTopologySuite.Geometries.Polygon)))
            Next

            Return geom.Factory.CreateMultiPolygon(polys.ToArray())
        End If

        Return geom
    End Function

    Private Function FixPolygonOrientation(
                                           poly As NetTopologySuite.Geometries.Polygon
                                           ) As NetTopologySuite.Geometries.Polygon

        Dim gf = poly.Factory

        ' Fix shell (must be CCW)
        Dim shell = poly.Shell
        Dim shellCoords = shell.Coordinates

        If Not NetTopologySuite.Algorithm.Orientation.IsCCW(shellCoords) Then
            shell = DirectCast(shell.Reverse(), NetTopologySuite.Geometries.LinearRing)
        End If

        ' Fix holes (must be CW)
        Dim holes(poly.NumInteriorRings - 1) As NetTopologySuite.Geometries.LinearRing

        For i = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(i)
            Dim holeCoords = hole.Coordinates

            If NetTopologySuite.Algorithm.Orientation.IsCCW(holeCoords) Then
                hole = DirectCast(hole.Reverse(), NetTopologySuite.Geometries.LinearRing)
            End If

            holes(i) = hole
        Next

        Return gf.CreatePolygon(shell, holes)
    End Function

    Public Async Function ImportIARURegions() As Task
        Dim GeoJson As String
        Dim factory As GeometryFactory = NetTopologySuite.NtsGeometryServices.Instance.CreateGeometryFactory(4326)
        Dim reader As New GeoJsonReader()

        ' Load shapefile using ArcGIS Runtime
        Dim table = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\IARU\IARU Region Lines_50m_No_Edges_Simplified.shp")
        Dim qp As New QueryParameters With {
        .OutSpatialReference = SpatialReferences.Wgs84,
        .ReturnGeometry = True
    }

        Dim features = Await table.QueryFeaturesAsync(qp)
        With Form1.ProgressBar1
            .Minimum = 0
            .Maximum = features.Count
            .Value = 0
        End With

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            Using cmd = connect.CreateCommand()
                cmd.CommandText = "DELETE FROM IARU"
                cmd.ExecuteNonQuery()
            End Using

            Using cmd = connect.CreateCommand()
                cmd.CommandText = "INSERT INTO IARU (line, geometry) VALUES (@id, @geom)"

                Dim pId = cmd.CreateParameter()
                pId.ParameterName = "@id"
                cmd.Parameters.Add(pId)

                Dim pGeom = cmd.CreateParameter()
                pGeom.ParameterName = "@geom"
                cmd.Parameters.Add(pGeom)

                Dim id As Integer = 0

                For Each f In features

                    GeoJson = FromArcGisToGeoJson(f.Geometry)   ' convert ArcGIS geometry to GeoJSON

                    ' Convert GeoJSON → NTS
                    Dim ntsGeom As NetTopologySuite.Geometries.Geometry = FromGeoJsonToNTS(GeoJson)

                    ' Simplify linework (reduce points)
                    Dim simplified As NetTopologySuite.Geometries.Geometry = DouglasPeuckerSimplifier.Simplify(ntsGeom, 0.01)

                    ' Densify each segment
                    Dim densified As NetTopologySuite.Geometries.Geometry = DensifyGeometry(simplified, 5.0)

                    If TypeOf densified Is LineString Then

                        ' Single segment
                        GeoJson = FromNTSToGeoJson(densified)
                        id += 1
                        pId.Value = id
                        pGeom.Value = geoJson
                        cmd.ExecuteNonQuery()

                    ElseIf TypeOf densified Is GeometryCollection Then

                        ' Multiple segments
                        Dim gc As GeometryCollection = CType(densified, GeometryCollection)

                        For i = 0 To gc.NumGeometries - 1
                            Dim seg As NetTopologySuite.Geometries.Geometry = gc.GetGeometryN(i)

                            If TypeOf seg Is LineString Then
                                GeoJson = FromNTSToGeoJson(seg)
                                id += 1
                                pId.Value = id
                                pGeom.Value = geoJson
                                cmd.ExecuteNonQuery()
                            End If
                        Next

                    Else
                        Throw New Exception("Unexpected geometry type after AM split + densify.")
                    End If

                    Form1.ProgressBar1.Value += 1
                Next
            End Using
        End Using
        AppendText(Form1.TextBox1, $"Imported {features.Count} IARU regions{vbCrLf}")
    End Function

    Public Async Function ImportAntarctica() As Task
        ' DONT NEED THIS ANYMORE - the data in NE is OK
        AppendText(Form1.TextBox1, "Obsolete" & vbCrLf)
    End Function


    Public Async Function ImportAntarcticBases() As Task

        Dim responseString As String = ""

        ' ---------------------------------------------------------
        ' 1. Download HTML
        ' ---------------------------------------------------------
        Using httpClient As New HttpClient()
            httpClient.Timeout = TimeSpan.FromMinutes(10)

            Dim url = "https://www.coolantarctica.com/Community/antarctic_bases.php"

            Try
                Dim httpResult = Await httpClient.GetAsync(url)
                httpResult.EnsureSuccessStatusCode()
                responseString = Await httpResult.Content.ReadAsStringAsync()
            Catch ex As HttpRequestException
                MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
                Return
            End Try
        End Using

        ' ---------------------------------------------------------
        ' 2. Parse HTML
        ' ---------------------------------------------------------
        Dim replacements As New Dictionary(Of String, String) From {
        {"&deg;", "°"},
        {"&#39;", "'"},
        {"&quot;", """"}
    }

        Dim allowedChars As String = "(\t|\n|&nbsp;)"
        Dim htmldoc As New HtmlDocument()
        htmldoc.LoadHtml(responseString)

        Dim table = htmldoc.DocumentNode.SelectSingleNode("//table")
        Dim rows = table.SelectNodes("tr")

        ' ---------------------------------------------------------
        ' 3. SQL setup
        ' ---------------------------------------------------------
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            Dim sql = connect.CreateCommand()
            sql.CommandText = "DELETE FROM Antarctic"
            sql.ExecuteNonQuery()

            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = rows.Count - 1
            End With

            ' ---------------------------------------------------------
            ' 4. Process each row
            ' ---------------------------------------------------------
            For row = 0 To rows.Count - 2       ' ignore totals row

                Form1.ProgressBar1.Value += 1

                Dim tds = rows(row).SelectNodes("td")
                If tds Is Nothing OrElse tds.Count < 5 Then Continue For

                Dim bgcolor = tds(0).Attributes("bgcolor")
                Dim open As String = If(bgcolor Is Nothing, "Summer only", "Year round")

                Dim name = CleanText(tds(0).InnerText, allowedChars)
                Dim nation = CleanText(tds(1).InnerText, allowedChars)

                Dim coordinates = ExtractLatLon(tds(2))
                Dim lat = coordinates.Item1
                Dim lon = coordinates.Item2

                Dim situation = CleanText(tds(3).InnerText, allowedChars)
                Dim altitude = CleanText(tds(4).InnerText, allowedChars)

                sql.CommandText =
                "INSERT INTO Antarctic (open, name, nation, coordinates, situation, altitude) " &
                "VALUES (@open, @name, @nation, @coords, @situation, @altitude)"

                sql.Parameters.Clear()
                sql.Parameters.AddWithValue("@open", open)
                sql.Parameters.AddWithValue("@name", name)
                sql.Parameters.AddWithValue("@nation", nation)
                sql.Parameters.AddWithValue("@coords", $"{{""type"":""Point"",""coordinates"":[{lon:F6},{lat:F6}]}}")
                sql.Parameters.AddWithValue("@situation", situation)
                sql.Parameters.AddWithValue("@altitude", altitude)

                sql.ExecuteNonQuery()
            Next
        End Using

        AppendText(Form1.TextBox1, $"Done{vbCrLf}")

    End Function
    Private Function CleanText(s As String, allowed As String) As String
        s = HttpUtility.HtmlDecode(s)
        s = Regex.Replace(s, allowed, "")
        Return s.Trim()
    End Function
    Private Function ExtractLatLon(td As HtmlNode) As (Double, Double)
        ' 1. Get raw HTML (keeps <br>)
        Dim html = td.InnerHtml

        ' 2. Split at <br> (handles <br>, <br/>, <br />)
        Dim parts = Regex.Split(html, "<br\s*/?>", RegexOptions.IgnoreCase)
        If parts.Length < 2 Then Return (0, 0)

        ' 3. Strip HTML tags from each part
        Dim latStr = Regex.Replace(parts(0), "</?[^>]+>", "").Trim()
        Dim lonStr = Regex.Replace(parts(1), "</?[^>]+>", "").Trim()

        ' 4. Decode HTML entities
        latStr = HttpUtility.HtmlDecode(latStr)
        lonStr = HttpUtility.HtmlDecode(lonStr)

        ' 5. Parse each DMS value
        Dim lat = ParseOne(latStr)
        Dim lon = ParseOne(lonStr)

        Return (lat, lon)
    End Function
    Private Function ParseOne(s As String) As Double
        ' Match degrees, minutes (integer or decimal), seconds (integer or decimal), hemisphere
        Dim m = Regex.Match(s, "(\d+)[°\s]+(\d+(?:\.\d+)?)?['\s]*?(\d+(?:\.\d+)?)?[""]?\s*([NSEW])", RegexOptions.IgnoreCase)
        If Not m.Success Then Return 0

        Dim deg = Double.Parse(m.Groups(1).Value)
        Dim min = If(m.Groups(2).Success, Double.Parse(m.Groups(2).Value), 0)
        Dim sec = If(m.Groups(3).Success, Double.Parse(m.Groups(3).Value), 0)
        Dim hemi = m.Groups(4).Value.ToUpper()

        ' Convert to decimal degrees
        Dim value = deg + (min / 60.0) + (sec / 3600.0)

        If hemi = "S" Or hemi = "W" Then value = -value

        Return value
    End Function

    Async Function ImportIOTAGroups() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        'Dim url = $"https://www.iota-world.org/rest/get/iota/groups?api_key={IOTA_API_KEY}"
        Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=groups.json"
        Using httpClient As New System.Net.Http.HttpClient()
            httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
            Try
                Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
                httpResult.EnsureSuccessStatusCode()
                responseString = Await httpResult.Content.ReadAsStringAsync()
            Catch ex As HttpRequestException
                MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
            End Try
        End Using

        ' Extract groups data
        Dim response = JsonNode.Parse(responseString).AsArray
        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = response.Count
        End With
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "BEGIN TRANSACTION"
            sql.ExecuteNonQuery()                             ' delete existing data
            sql.CommandText = "DELETE FROM IOTA_Groups"
            sql.ExecuteNonQuery()                             ' delete existing data
            ' use prepared statement for speed
            sql.CommandText = $"INSERT INTO IOTA_Groups 
        (`refno`,`name`,`dxcc_num`,latitude_max,latitude_min,longitude_max,longitude_min,grp_region,whitelist,comment) VALUES (@refno,@name,@dxcc_num,@latitude_max,@latitude_min,@longitude_max,@longitude_min,@grp_region,@whitelist,@comment)
"
            sql.Prepare()
            For Each group In response
                ' process each group
                Dim refno = SafeStr(group.Item("refno"))
                Form1.ProgressBar1.Value += 1
                With sql.Parameters
                    .Clear()
                    ' Sometimes they have things backwards so fix it if necessary
                    Dim LatMin As Double = SafeDbl(group.Item("latitude_min"))
                    Dim LatMax As Double = SafeDbl(group.Item("latitude_max"))
                    Dim LonMin As Double = SafeDbl(group.Item("longitude_min"))
                    Dim LonMax As Double = SafeDbl(group.Item("longitude_max"))
                    If LatMin > LatMax Then
                        Dim t = LatMin
                        LatMin = LatMax
                        LatMax = t
                        AppendText(Form1.TextBox1, $"Warning: LatMin > LatMax for {group.Item("name")}. Values swapped.{vbCrLf}")
                    End If
                    If LonMin > 180 Then LonMin -= 360 ' normalize to -180 to 180
                    If LonMax > 180 Then LonMax -= 360 ' normalize to -180 to 180
                    If LonMin <= LonMax Then
                        ' normal case, nothing to do
                    Else
                        Dim gap = Math.Abs(LonMin - LonMax)
                        If gap > 180 Then
                            ' AM crossing
                        Else
                            ' probably they have wrapped around the wrong way. Swap them
                            Dim t = LonMin
                            LonMin = LonMax
                            LonMax = t
                            AppendText(Form1.TextBox1, $"Warning: LonMin > LonMax for {group.Item("name")}. Values swapped.{vbCrLf}")
                        End If
                    End If
                    .AddWithValue("@refno", refno)
                    .AddWithValue("@dxcc_num", group.Item("dxcc_num").ToString)
                    .AddWithValue("@name", group.Item("name").ToString)
                    .AddWithValue("@latitude_min", LatMin)
                    .AddWithValue("@latitude_max", LatMax)
                    .AddWithValue("@longitude_min", LonMin)
                    .AddWithValue("@longitude_max", LonMax)
                    .AddWithValue("@grp_region", group.Item("grp_region").ToString)
                    .AddWithValue("@whitelist", group.Item("whitelist").ToString)
                    .AddWithValue("@comment", group.Item("comment").ToString)
                End With
                sql.ExecuteNonQuery()
            Next
            sql.CommandText = "COMMIT"
            sql.ExecuteNonQuery()
            Form1.TextBox1.AppendText($"{response.Count} groups loaded{vbCrLf}")
        End Using
    End Function
    Async Function ImportIOTAIslands() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        'Dim url = $"https://www.iota-world.org/rest/get/iota/islands?api_key={IOTA_API_KEY}"
        Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=islands.json"
        Using httpClient As New System.Net.Http.HttpClient()
            httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
            Try
                Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
                httpResult.EnsureSuccessStatusCode()
                responseString = Await httpResult.Content.ReadAsStringAsync()
            Catch ex As HttpRequestException
                MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
            End Try
        End Using

        ' Extract groups data
        Dim response = JsonNode.Parse(responseString).AsArray

        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = response.Count
        End With
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "BEGIN TRANSACTION"
            sql.ExecuteNonQuery()                             ' delete existing data
            sql.CommandText = "DELETE FROM IOTA_Islands"
            sql.ExecuteNonQuery()                             ' delete existing data
            ' use prepared statement for speed
            sql.CommandText = $"INSERT INTO IOTA_Islands (`refno`,`name`,`comment`) VALUES (@refno,@name,@comment)"
            sql.Prepare()
            For Each group In response
                ' process each group
                Form1.ProgressBar1.Value += 1
                With sql.Parameters
                    .Clear()
                    .AddWithValue("@refno", group.Item("refno").ToString)
                    .AddWithValue("@name", group.Item("name").ToString)
                    .AddWithValue("@comment", group.Item("comment").ToString)
                End With
                sql.ExecuteNonQuery()
            Next
            sql.CommandText = "COMMIT"
            sql.ExecuteNonQuery()
            Form1.TextBox1.AppendText($"{response.Count} islands loaded{vbCrLf}")
        End Using
    End Function
    Async Function ImportIOTADXCCMatchesOneIOTA() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        'Dim url = $"https://www.iota-world.org/rest/get/iota/dxccmatchesoneiota?api_key={IOTA_API_KEY}"
        Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=dxcc_matches_one_iota.json"
        Using httpClient As New System.Net.Http.HttpClient()
            httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
            Try
                Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
                httpResult.EnsureSuccessStatusCode()
                responseString = Await httpResult.Content.ReadAsStringAsync()
            Catch ex As HttpRequestException
                MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
            End Try
        End Using

        ' Extract groups data
        Dim response = JsonNode.Parse(responseString).AsArray
        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = response.Count
        End With
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "BEGIN TRANSACTION"
            sql.ExecuteNonQuery()                             ' delete existing data
            sql.CommandText = "DELETE FROM IOTA_DXCC_IOTA"
            sql.ExecuteNonQuery()                             ' delete existing data
            ' use prepared statement for speed
            sql.CommandText = $"INSERT INTO IOTA_DXCC_IOTA (`refno`,`dxcc_num`) VALUES (@refno,@dxcc_num)"
            sql.Prepare()
            For Each group In response
                ' process each group
                Form1.ProgressBar1.Value += 1
                With sql.Parameters
                    .Clear()
                    .AddWithValue("@refno", group.Item("refno").ToString)
                    .AddWithValue("@dxcc_num", group.Item("dxcc_num").ToString)
                End With
                sql.ExecuteNonQuery()
            Next
            sql.CommandText = "COMMIT"
            sql.ExecuteNonQuery()
            Form1.TextBox1.AppendText($"{response.Count} DXCC matches loaded{vbCrLf}")
        End Using
    End Function

    Sub ImportPolyFromKML()

        Dim result As String

        With Form1.OpenFileDialog1
            .Filter = "KML files (*.kml)|*.kml|All files(*.*)|*.*"
            .CheckFileExists = True
            .CheckPathExists = True
            .AddExtension = True
            .DefaultExt = "kml"
            .FileName = ""
            .Title = "Select a .kml file containing the desired polygon"

            If .ShowDialog() = DialogResult.OK Then

                Dim kmlPath = .FileName

                ' ---------------------------------------------------------
                ' 1. Load KML and extract <coordinates>
                ' ---------------------------------------------------------
                Dim doc = XDocument.Load(kmlPath)
                Dim ns = doc.Root.Name.Namespace

                Dim coordText As String =
                doc.Descendants(ns + "coordinates").First().Value

                ' Clean up whitespace and illegal characters
                coordText = Regex.Replace(coordText.Trim(), "[^0-9\-\., ]", "")

                Dim coordPairs = coordText.Split({" "}, StringSplitOptions.RemoveEmptyEntries)

                ' ---------------------------------------------------------
                ' 2. Build NTS polygon
                ' ---------------------------------------------------------
                Dim gf = NtsGeometryServices.Instance.CreateGeometryFactory(4326)
                Dim coords As New List(Of Coordinate)

                For Each pair In coordPairs
                    Dim parts = pair.Split(","c)
                    If parts.Length < 2 Then Continue For

                    Dim lon = NormalizeLongitude(parts(0))
                    Dim lat = CDbl(parts(1))

                    coords.Add(New Coordinate(lon, lat))
                Next

                ' Close ring if needed
                If Not coords.First().Equals2D(coords.Last()) Then
                    coords.Add(New Coordinate(coords.First().X, coords.First().Y))
                End If

                Dim ring As LinearRing = gf.CreateLinearRing(coords.ToArray())
                Dim poly As NetTopologySuite.Geometries.Polygon = gf.CreatePolygon(ring)

                ' ---------------------------------------------------------
                ' 3. Determine decimal precision for bbox
                ' ---------------------------------------------------------
                Dim env = poly.EnvelopeInternal
                Dim fmt = "f" & DecimalsForBbox(env.MinX, env.MinY, env.MaxX, env.MaxY)

                ' ---------------------------------------------------------
                ' 4. Build poly:"..." output
                ' ---------------------------------------------------------
                Dim outPts As New List(Of String)

                For Each c In ring.Coordinates
                    outPts.Add($"{c.X.ToString(fmt)} {c.Y.ToString(fmt)}")
                Next

                result = $"poly:""{String.Join(" ", outPts)}"""

                Clipboard.SetText(result)
                MsgBox(result, vbInformation + vbOKOnly, "Result in clipboard")

            End If
        End With

    End Sub

    Public Async Function LandSquareList() As Task
        ' Import a list of squares that are land
        Const LandDataURL = "https://osmdata.openstreetmap.de/download/land-polygons-split-4326.zip"    ' remote source of land data
        Const LandDataFile = "D:\GIS Data\Land Polygons\split\land_polygons.shp"                        ' local copy of land data
        Const LandDataZip = "D:\GIS Data\Land Polygons\split\land-polygons-split-4326.zip"              ' local copy of land data
        Const readChunkSize = 1024 * 1024            ' block size of bytes read
        Dim myQueryFilter As New QueryParameters, count As Integer = 0
        Dim timer As New Stopwatch

        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = 100
        End With
        ' Check the date on the latest download file
        Dim response = Await Http.GetAsync(LandDataURL, HttpCompletionOption.ResponseHeadersRead)      ' request for header only
        If Not response.IsSuccessStatusCode Then
            MsgBox($"Error: {response.StatusCode}", vbCritical + vbOKOnly, "Error")
            Return
        End If

        ' Check dates
        Dim LandDataURLDate = response.Content.Headers.LastModified.Value.UtcDateTime
        Dim LandDataFileDate = File.GetLastWriteTimeUtc(LandDataZip)
        Dim DateStatus As String
        If LandDataURLDate > LandDataFileDate Then DateStatus = "Land data is Out of Date" Else DateStatus = "Land data is Current"
        If MsgBox($"The OSM land data is dated {LandDataURLDate.ToUniversalTime:yyyy-MM-dd hh:mm}{vbCrLf}The local copy is dated {LandDataFileDate.ToUniversalTime:yyyy-MM-dd hh:mm}{vbCrLf}{vbCrLf}Do you wish to update the data ?", vbInformation + vbYesNo, DateStatus) = vbYes Then
            Dim totalBytes = response.Content.Headers.ContentLength       ' get count of total bytes
            timer.Start()
            AppendText(Form1.TextBox1, $"Fetching {totalBytes:n0} bytes of land data from OSM ")
            Dim totalBytesRead As Long = 0      ' total bytes read todate
            ' Download the file from OSM with progress indicator
            Using contentStream = Await response.Content.ReadAsStreamAsync,
                filestream = New FileStream(LandDataZip, FileMode.Create, FileAccess.Write, FileShare.None, readChunkSize, True)
                Dim buffer(readChunkSize) As Byte       ' byte buffer
                Dim bytesRead As Integer                ' bytes read in block
                Do
                    bytesRead = Await contentStream.ReadAsync(buffer, 0, buffer.Length)     ' read a block
                    If bytesRead > 0 Then       ' block is not empty
                        Await filestream.WriteAsync(buffer, 0, bytesRead)   ' write the block to the file
                        totalBytesRead += bytesRead         ' count total bytes read
                        Dim progressPercentage As Integer = totalBytesRead / totalBytes * 100   ' calculate percentage progress
                        Form1.ProgressBar1.Value = progressPercentage         ' display progress
                    End If
                Loop Until bytesRead = 0    ' stop when no bytes read
                filestream.Close()
            End Using
            ' Now unzip downloaded file
            Using archive = ZipFile.OpenRead(LandDataZip)
                Dim targetDirectory = Path.GetDirectoryName(LandDataZip)
                For Each entry In archive.Entries
                    entry.ExtractToFile($"{targetDirectory}\{entry.Name}", True)
                Next
            End Using
            timer.Stop()
            AppendText(Form1.TextBox1, $"[{timer.ElapsedMilliseconds / 1000:f1}s]{vbCrLf}")
        Else
            Return
        End If

        ' Convert the OSM data into a list of grid squares that contain land
        timer.Restart()
        Dim Features = Await ShapefileFeatureTable.OpenAsync(LandDataFile)
        With myQueryFilter
            .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
            .ReturnGeometry = False
        End With
        Dim land = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' return all geometry
        Dim featureCount = land.Count
        AppendText(Form1.TextBox1, $"Loading {featureCount} squares into database.")
        Using connect As New SqliteConnection(DXCC_DATA)
            Dim sql As SqliteCommand, sqlDR As SqliteDataReader
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT COUNT(*) as Count FROM LAND"
            sqlDR = sql.ExecuteReader()
            sqlDR.Read()
            Dim Before As Integer = sqlDR("Count")
            sqlDR.Close()
            sql.CommandText = "BEGIN TRANSACTION"
            sql.ExecuteNonQuery()
            sql.CommandText = "DELETE FROM LAND"
            sql.ExecuteNonQuery()
            For Each feature In land
                Dim x = CDbl(feature.Attributes("x"))
                Dim y = CDbl(feature.Attributes("y"))
                sql.CommandText = $"INSERT OR REPLACE INTO LAND (gridsquare) VALUES ('{GridSquare(x, y)}')"
                sql.ExecuteNonQuery()
                UpdateProgressBar(Form1.ProgressBar1, count / featureCount * 100)
                count += 1
            Next
            sql.CommandText = "COMMIT"
            sql.ExecuteNonQuery()
            sql.CommandText = "SELECT COUNT(*) as Count FROM LAND"
            sqlDR = sql.ExecuteReader()
            sqlDR.Read()
            Dim After As Integer = sqlDR("Count")
            sqlDR.Close()
            AppendText(Form1.TextBox1, $" Grid squares before={Before}, after={After} [{timer.ElapsedMilliseconds / 1000:f1}s]{vbCrLf}")
        End Using
    End Function
    Private Function DensifyGeometry(geom As NetTopologySuite.Geometries.Geometry, maxDegrees As Double) _
    As NetTopologySuite.Geometries.Geometry

        Dim densifier = New NetTopologySuite.Densify.Densifier(geom)
        densifier.DistanceTolerance = maxDegrees
        Return densifier.GetResultGeometry()
    End Function

    Function NormalizeLongitude(longitude As Double) As Double
        ' Normalize a longitude to between -180 and +180
        Return (longitude Mod 360 + 540) Mod 360 - 180        ' normalize longitude
    End Function

End Module
