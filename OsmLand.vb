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

    ' Timestamp of the last Overpass request.
    ' Used to enforce a minimum 1-second delay between requests.
    Private Async Function FetchOSMObject(osmId As Long, osmType As String) As Task(Of JsonDocument)
        ' Build Overpass query
        Dim query As String = $"[out:json][timeout:300];{osmType}({osmId});(._;>;);out geom;"
        Debug.WriteLine($"Sending {query} to OSM")

        ' Encode query for URL
        ' https://overpass.openstreetmap.fr/api/interpreter
        ' https://overpass.kumi.systems/api/interpreter
        ' https://overpass.nchc.org.tw/api/interpreter

        Dim url As String = "https://overpass-api.de/api/interpreter?data=" & Uri.EscapeDataString(query)

        ' Execute request using your existing HttpClient
        Dim response As HttpResponseMessage = Await Http.GetAsync(url)
        response.EnsureSuccessStatusCode()

        ' Parse JSON
        Dim stream = Await response.Content.ReadAsStreamAsync()
        Debug.WriteLine($"FetchOSMOject: result={response.ReasonPhrase}, response of {stream.Length} bytes")
        Await Task.Delay(1000) ' Enforce 1-second delay between requests
        Return Await JsonDocument.ParseAsync(stream)
    End Function
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
            Dim ring = gf.CreateLinearRing(coords)
            linearRings.Add(ring)
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
                              nodes As Dictionary(Of Long, Coordinate)) As List(Of Coordinate)

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

    Function StitchWaysIntoRings(ways As List(Of List(Of Coordinate))) As List(Of List(Of Coordinate))

        ' Build node → list of ways index
        Dim nodeToWays As New Dictionary(Of Coordinate, List(Of List(Of Coordinate)))(New CoordinateEqualityComparer)

        For Each w In ways
            For Each c In w
                If Not nodeToWays.ContainsKey(c) Then
                    nodeToWays(c) = New List(Of List(Of Coordinate))()
                End If
                nodeToWays(c).Add(w)
            Next
        Next

        Dim rings As New List(Of List(Of Coordinate))()
        Dim usedWays As New HashSet(Of List(Of Coordinate))()

        ' Process each way as a potential starting point
        For Each startWay In ways
            If usedWays.Contains(startWay) Then Continue For

            Dim ring As New List(Of Coordinate)(startWay)
            usedWays.Add(startWay)

            Dim extended As Boolean = True

            While extended
                extended = False

                ' Try to extend at the end
                Dim endNode = ring.Last()

                If nodeToWays.ContainsKey(endNode) Then
                    For Each w In nodeToWays(endNode)
                        If usedWays.Contains(w) Then Continue For

                        ' If the way starts at endNode
                        If w.First().Equals2D(endNode) Then
                            ring.AddRange(w.Skip(1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If

                        ' If the way ends at endNode → reverse it
                        If w.Last().Equals2D(endNode) Then
                            w.Reverse()
                            ring.AddRange(w.Skip(1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If

                        ' If the way contains endNode in the middle → split it
                        Dim idx = w.FindIndex(Function(c) c.Equals2D(endNode))
                        If idx > 0 Then
                            Dim tail = w.Skip(idx).ToList()
                            ring.AddRange(tail.Skip(1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If
                    Next
                End If

                If extended Then Continue While

                ' Try to extend at the start
                Dim startNode = ring.First()

                If nodeToWays.ContainsKey(startNode) Then
                    For Each w In nodeToWays(startNode)
                        If usedWays.Contains(w) Then Continue For

                        ' If the way ends at startNode
                        If w.Last().Equals2D(startNode) Then
                            ring.InsertRange(0, w.Take(w.Count - 1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If

                        ' If the way starts at startNode → reverse it
                        If w.First().Equals2D(startNode) Then
                            w.Reverse()
                            ring.InsertRange(0, w.Take(w.Count - 1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If

                        ' If the way contains startNode in the middle → split it
                        Dim idx = w.FindIndex(Function(c) c.Equals2D(startNode))
                        If idx > 0 Then
                            Dim head = w.Take(idx + 1).ToList()
                            ring.InsertRange(0, head.Take(head.Count - 1))
                            usedWays.Add(w)
                            extended = True
                            Exit For
                        End If
                    Next
                End If

            End While

            ' Close ring if needed
            If Not ring.First().Equals2D(ring.Last()) Then
                ring.Add(New Coordinate(ring.First().X, ring.First().Y))
            End If

            rings.Add(ring)
        Next

        Return rings
    End Function

End Module
