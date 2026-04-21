Imports System.Collections.Concurrent
Imports System.Net
Imports System.Net.Http
Imports System.Text.Json
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Imports NetTopologySuite
Imports NetTopologySuite.Geometries
Imports Newtonsoft.Json.Linq

Public Module OsmLand

    Private ReadOnly _cache As New ConcurrentDictionary(Of String, NetTopologySuite.Geometries.Geometry)(StringComparer.OrdinalIgnoreCase)

    Async Function ResolveOSM(osmId As Long, osmType As String) As Task(Of NetTopologySuite.Geometries.Geometry)
        ' Await the fetch
        Dim doc As JsonDocument = Await FetchOSMObject(osmId, osmType)

        ' Root → elements array
        Dim root = doc.RootElement
        Dim elements = root.GetProperty("elements")

        ' Build node dictionary
        Dim nodes = BuildNodeDict(elements)

        ' Find the main object
        Dim obj = elements.EnumerateArray().
        First(Function(e) e.GetProperty("type").GetString() = osmType AndAlso
                         e.GetProperty("id").GetInt64() = osmId)

        If osmType = "way" Then
            Return BuildWayPolygon(obj, nodes, factory)

        ElseIf osmType = "relation" Then
            Return BuildRelationPolygon(obj, elements, nodes)

        Else
            Throw New Exception("Unsupported OSM type")
        End If
    End Function

    ' Timestamp of the last Overpass request.
    ' Used to enforce a minimum 1-second delay between requests.
    Private Async Function FetchOSMObject(osmId As Long, osmType As String) As Task(Of JsonDocument)
        ' Build Overpass query
        Dim query As String = $"[out:json][timeout:300];{osmType}({osmId});(._;>;);out;"
        Debug.WriteLine($"Sending {query} to OSM")

        ' Encode query for URL
        Dim url As String = "https://overpass.kumi.systems/api/interpreter?data=" & Uri.EscapeDataString(query)

        ' Execute request using your existing HttpClient
        Dim response As HttpResponseMessage = Await Http.GetAsync(url)
        response.EnsureSuccessStatusCode()

        ' Parse JSON
        Dim stream = Await response.Content.ReadAsStreamAsync()
        Debug.WriteLine($"FetchOSMOject: result={response.ReasonPhrase}, response of {stream.Length} bytes")
        Await Task.Delay(1000) ' Enforce 1-second delay between requests
        Return Await JsonDocument.ParseAsync(stream)
    End Function

    Function BuildRelationPolygon(rel As JsonElement,
                              elements As JsonElement,
                              nodes As Dictionary(Of Long, Coordinate)) As NetTopologySuite.Geometries.Geometry

        Dim gf As GeometryFactory =
        NtsGeometryServices.Instance.CreateGeometryFactory(4326)

        ' 1. Assemble rings (each ring = List(Of Coordinate))
        Dim rings As List(Of List(Of Coordinate)) =
        AssembleRelationRings(rel, elements, nodes)

        If rings Is Nothing OrElse rings.Count = 0 Then
            Return gf.CreateGeometryCollection(Nothing)
        End If

        ' 2. Convert rings to LinearRings and compute signed area
        Dim ringAreas As New List(Of Double)
        Dim linearRings As New List(Of LinearRing)

        For Each r In rings
            ' Ensure closed
            If Not r(0).Equals2D(r(r.Count - 1)) Then
                r.Add(New Coordinate(r(0).X, r(0).Y))
            End If

            Dim coords = r.ToArray()
            Dim ringarea = NetTopologySuite.Algorithm.Area.OfRing(coords)   ' signed area
            ringAreas.Add(ringarea)

            linearRings.Add(gf.CreateLinearRing(coords))
        Next

        ' 3. Identify outer rings (CCW) and inner rings (CW)
        Dim outers As New List(Of Integer)
        Dim inners As New List(Of Integer)

        For i = 0 To linearRings.Count - 1
            If ringAreas(i) > 0 Then
                outers.Add(i)
            Else
                inners.Add(i)
            End If
        Next

        ' 4. Assign holes to the correct outer ring
        Dim outerToHoles As New Dictionary(Of Integer, List(Of LinearRing))

        For Each oi In outers
            outerToHoles(oi) = New List(Of LinearRing)
        Next

        For Each ii In inners
            Dim hole = linearRings(ii)
            Dim holePt = hole.Coordinate

            ' Find containing outer ring
            Dim parent As Integer = -1

            For Each oi In outers
                Dim shell = linearRings(oi)
                Dim poly = gf.CreatePolygon(shell)

                If poly.Contains(gf.CreatePoint(holePt)) Then
                    parent = oi
                    Exit For
                End If
            Next

            ' If no parent found, treat as independent outer (rare OSM error)
            If parent = -1 Then
                outers.Add(ii)
                outerToHoles(ii) = New List(Of LinearRing)
            Else
                outerToHoles(parent).Add(hole)
            End If
        Next

        ' 5. Build polygons
        Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)

        For Each oi In outers
            Dim shell = linearRings(oi)
            Dim holes = outerToHoles(oi).ToArray()
            polys.Add(gf.CreatePolygon(shell, holes))
        Next

        ' 6. Return Polygon or MultiPolygon
        If polys.Count = 1 Then
            Return polys(0)
        Else
            Return gf.CreateMultiPolygon(polys.ToArray())
        End If
    End Function

    Function AssembleRelationRings(rel As JsonElement,
                                   elements As JsonElement,
                                   nodes As Dictionary(Of Long, Coordinate)) _
        As List(Of List(Of Coordinate))

        Dim ways As New List(Of List(Of Coordinate))

        ' Convert each way into a list of coordinates (not closed)
        For Each member In rel.GetProperty("members").EnumerateArray()

            If member.GetProperty("type").GetString() <> "way" Then Continue For

            Dim wid = member.GetProperty("ref").GetInt64()

            ' Find the matching way in the elements array
            Dim way = elements.EnumerateArray().
                    Where(Function(e) e.GetProperty("type").GetString() = "way").
                    Single(Function(e) e.GetProperty("id").GetInt64() = wid)

            ' Extract coordinates for this way
            Dim pts = ExtractWayPoints(way, nodes)

            If pts.Count > 1 Then
                ways.Add(pts)
            End If
        Next

        ' Stitch the ways into closed rings
        Return StitchWaysIntoRings(ways)
    End Function

    Function StitchWaysIntoRings(ways As List(Of List(Of Coordinate))) As List(Of List(Of Coordinate))

        Dim rings As New List(Of List(Of Coordinate))

        ' Work on a mutable list
        Dim remaining = ways.ToList()

        While remaining.Count > 0

            ' Start a new chain
            Dim chain = New List(Of Coordinate)(remaining(0))
            remaining.RemoveAt(0)

            Dim extended As Boolean = True

            While extended
                extended = False

                For i = remaining.Count - 1 To 0 Step -1
                    Dim w = remaining(i)

                    ' chain end → w start
                    If chain.Last().Equals2D(w.First()) Then
                        chain.AddRange(w.Skip(1))
                        remaining.RemoveAt(i)
                        extended = True
                        Continue For
                    End If

                    ' chain end → w end (reverse w)
                    If chain.Last().Equals2D(w.Last()) Then
                        w.Reverse()
                        chain.AddRange(w.Skip(1))
                        remaining.RemoveAt(i)
                        extended = True
                        Continue For
                    End If

                    ' w end → chain start
                    If w.Last().Equals2D(chain.First()) Then
                        chain.InsertRange(0, w.Take(w.Count - 1))
                        remaining.RemoveAt(i)
                        extended = True
                        Continue For
                    End If

                    ' w start → chain start (reverse w)
                    If w.First().Equals2D(chain.First()) Then
                        w.Reverse()
                        chain.InsertRange(0, w.Take(w.Count - 1))
                        remaining.RemoveAt(i)
                        extended = True
                        Continue For
                    End If
                Next
            End While

            ' If chain closes, store it as a ring
            If chain.First().Equals2D(chain.Last()) Then
                rings.Add(chain)
            End If

        End While

        Return rings
    End Function

    Function ExtractWayPoints(way As JsonElement,
                              nodes As Dictionary(Of Long, MapPoint)) As List(Of MapPoint)

        Dim pts As New List(Of MapPoint)

        For Each nd In way.GetProperty("nodes").EnumerateArray()
            Dim nid = nd.GetInt64()
            If nodes.ContainsKey(nid) Then pts.Add(nodes(nid))
        Next

        Return pts
    End Function

    Function ExtractWayPoints(way As JsonElement,
                              nodes As Dictionary(Of Long, Coordinate)) _
        As List(Of Coordinate)

        Dim pts As New List(Of Coordinate)

        For Each nd In way.GetProperty("nodes").EnumerateArray()
            Dim nid = nd.GetInt64()
            If nodes.ContainsKey(nid) Then
                pts.Add(nodes(nid))
            End If
        Next

        Return pts
    End Function

    Function BuildWayPolygon(way As JsonElement,
                             nodes As Dictionary(Of Long, Coordinate),
                             gf As GeometryFactory) As NetTopologySuite.Geometries.Polygon

        Dim coords As New List(Of Coordinate)

        ' Extract coordinates for each node in the way
        For Each nd In way.GetProperty("nodes").EnumerateArray()
            Dim nid = nd.GetInt64()
            If nodes.ContainsKey(nid) Then
                coords.Add(nodes(nid))
            End If
        Next

        ' Not enough points to form a polygon
        If coords.Count < 3 Then
            Return gf.CreatePolygon()   ' empty polygon
        End If

        ' Ensure closed ring
        If Not coords(0).Equals2D(coords(coords.Count - 1)) Then
            coords.Add(New Coordinate(coords(0).X, coords(0).Y))
        End If

        ' Build LinearRing → Polygon
        Dim shell As LinearRing = gf.CreateLinearRing(coords.ToArray())
        Return gf.CreatePolygon(shell)
    End Function

    Function BuildNodeDict(elements As JsonElement) As Dictionary(Of Long, Coordinate)

        Dim dict As New Dictionary(Of Long, Coordinate)

        For Each el In elements.EnumerateArray()
            If el.GetProperty("type").GetString() = "node" Then

                Dim id = el.GetProperty("id").GetInt64()
                Dim lat = el.GetProperty("lat").GetDouble()
                Dim lon = el.GetProperty("lon").GetDouble()

                ' NTS uses Coordinate(X=lon, Y=lat)
                dict(id) = New Coordinate(lon, lat)

            End If
        Next

        Return dict
    End Function

End Module
