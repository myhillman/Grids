Imports System.Collections.Concurrent
Imports System.Net
Imports System.Net.Http
Imports System.Text.Json
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Imports Newtonsoft.Json.Linq

Public Module OsmLand

    Private ReadOnly _cache As New ConcurrentDictionary(Of String, Geometry)(StringComparer.OrdinalIgnoreCase)

    Async Function ResolveOSM(osmId As Long, osmType As String) As Task(Of Geometry)
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
            Return BuildWayPolygon(obj, nodes)

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
                                     nodes As Dictionary(Of Long, MapPoint)) As Geometry

        ' Collect all ways for this relation
        Dim wayRings = AssembleRelationRings(rel, elements, nodes)

        ' Build final multipolygon
        Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)

        For Each ring In wayRings
            pb.AddPart(ring)
        Next

        Return pb.ToGeometry()
    End Function
    Function AssembleRelationRings(rel As JsonElement,
                                   elements As JsonElement,
                                   nodes As Dictionary(Of Long, MapPoint)) As List(Of List(Of MapPoint))

        Dim ways As New List(Of List(Of MapPoint))

        ' Convert each way into a list of points (not closed)
        For Each member In rel.GetProperty("members").EnumerateArray()
            If member.GetProperty("type").GetString() <> "way" Then Continue For

            Dim wid = member.GetProperty("ref").GetInt64()

            Dim way = elements.EnumerateArray().
                    Where(Function(e) e.GetProperty("type").GetString() = "way").
                    Single(Function(e) e.GetProperty("id").GetInt64() = wid)

            Dim pts = ExtractWayPoints(way, nodes)
            If pts.Count > 1 Then ways.Add(pts)
        Next

        ' Now stitch the ways into rings
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
    Function StitchWaysIntoRings(ways As List(Of List(Of MapPoint))) As List(Of List(Of MapPoint))

    Dim rings As New List(Of List(Of MapPoint))

    ' Work on a mutable list
    Dim remaining = ways.ToList()

    While remaining.Count > 0

        ' Start a new chain
        Dim chain = New List(Of MapPoint)(remaining(0))
        remaining.RemoveAt(0)

        Dim extended As Boolean = True

        While extended
            extended = False

            For i = remaining.Count - 1 To 0 Step -1
                Dim w = remaining(i)

                ' chain end → w start
                If chain.Last().IsEqual(w.First()) Then
                    chain.AddRange(w.Skip(1))
                    remaining.RemoveAt(i)
                    extended = True
                    Continue For
                End If

                ' chain end → w end (reverse w)
                If chain.Last().IsEqual(w.Last()) Then
                    w.Reverse()
                    chain.AddRange(w.Skip(1))
                    remaining.RemoveAt(i)
                    extended = True
                    Continue For
                End If

                ' w end → chain start
                If w.Last().IsEqual(chain.First()) Then
                    chain.InsertRange(0, w.Take(w.Count - 1))
                    remaining.RemoveAt(i)
                    extended = True
                    Continue For
                End If

                ' w start → chain start (reverse w)
                If w.First().IsEqual(chain.First()) Then
                    w.Reverse()
                    chain.InsertRange(0, w.Take(w.Count - 1))
                    remaining.RemoveAt(i)
                    extended = True
                    Continue For
                End If
            Next
        End While

        ' If chain closes, store it as a ring
        If chain.First().IsEqual(chain.Last()) Then
            rings.Add(chain)
        End If

    End While

    Return rings
End Function

   Function BuildWayPolygon(way As JsonElement, nodes As Dictionary(Of Long, MapPoint)) As Polygon
        Dim pts As New List(Of MapPoint)

        For Each nd In way.GetProperty("nodes").EnumerateArray()
            Dim nid = nd.GetInt64()
            pts.Add(nodes(nid))
        Next

        ' Ensure closed ring
        If Not pts.First().Equals(pts.Last()) Then
            pts.Add(pts.First())
        End If

        Return New Polygon(pts, SpatialReferences.Wgs84)
    End Function
    
    Function BuildNodeDict(elements As JsonElement) As Dictionary(Of Long, MapPoint)
        Dim dict As New Dictionary(Of Long, MapPoint)

        For Each el In elements.EnumerateArray()
            If el.GetProperty("type").GetString() = "node" Then
                Dim id = el.GetProperty("id").GetInt64()
                Dim lat = el.GetProperty("lat").GetDouble()
                Dim lon = el.GetProperty("lon").GetDouble()
                dict(id) = New MapPoint(lon, lat, SpatialReferences.Wgs84)
            End If
        Next

        Return dict
    End Function


End Module
