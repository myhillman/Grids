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
    Function BuildRelationPolygon(rel As JsonElement, elements As JsonElement, nodes As Dictionary(Of Long, MapPoint)) As Geometry
        Dim rings As New List(Of Polygon)

        For Each member In rel.GetProperty("members").EnumerateArray()
            If member.GetProperty("type").GetString() = "way" Then
                Dim wid = member.GetProperty("ref").GetInt64()

                Dim way = elements.EnumerateArray().
                First(Function(e) e.GetProperty("type").GetString() = "way" AndAlso
                                   e.GetProperty("id").GetInt64() = wid)

                rings.Add(BuildWayPolygon(way, nodes))
            End If
        Next

        If rings.Count = 1 Then
            Return rings(0)
        Else
            Return GeometryEngine.Union(rings)
        End If
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
