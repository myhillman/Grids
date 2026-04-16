Imports System.Drawing
Imports System.Linq
Imports System.Net.Http
Imports System.Runtime.CompilerServices
Imports System.Text
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Threading
Imports SQLitePCL


' ==========================
' Core OSM model
' ==========================
Public Class OSMNode
    Public Property Id As Long
    Public Property Lat As Double
    Public Property Lon As Double
End Class

Public Class OSMMember
    Public Property Type As String   ' "node", "way", "relation"
    Public Property Ref As Long
    Public Property Role As String   ' "outer", "inner", etc.
End Class

Public Class OSMObject
    Public Property Type As String   ' "node", "way", "relation"
    Public Property ID As Long
    Public Property Nodes As List(Of Long)   ' only for ways
    Public Property Node As OSMNode          ' only for nodes
    Public Property Members As List(Of OSMMember) ' For relations
    Public Property Tags As Dictionary(Of String, String)
    Public Property Geometry As List(Of OSMNode)

End Class

' Simple polygon container
Public Class OSMPolygon
    Public Property Outer As List(Of PointD)
    Public Property Holes As New List(Of List(Of PointD))
    Public Sub New()
        Outer = New List(Of PointD)()
        Holes = New List(Of List(Of PointD))()
    End Sub
End Class

' ==========================
' Overpass query builder
' ==========================
Public Module OverpassBuilder

    ' ============================================================
    '  Overpass Query Builder (Hardened, Modern, Safe)
    ' ============================================================

    Public Function BuildSelector(filter As String) As String
        If String.IsNullOrWhiteSpace(filter) Then Return ""

        ' If it's not a simple selector, return it unchanged.
        If Not IsSimpleSelector(filter) Then
            Return filter.Trim()
        End If

        ' --- Simple selector normalization ---
        Dim f = filter.Trim()

        ' Remove trailing semicolon if present
        If f.EndsWith(";") Then f = f.Substring(0, f.Length - 1)

        ' Remove any accidental bbox/poly fragments
        ' (legacy cleanup)
        Dim idx = f.IndexOf("("c)
        If idx > -1 AndAlso f.Contains(")") Then
            ' Only strip if it's a bbox/poly, not a filter
            Dim inside = f.Substring(idx + 1, f.LastIndexOf(")"c) - idx - 1)
            If inside.Contains(",") AndAlso inside.Count(Function(c) c = ","c) >= 3 Then
                f = f.Substring(0, idx)
            End If
        End If

        Return f
    End Function

    Public Function IsSimpleSelector(filter As String) As Boolean
        If String.IsNullOrWhiteSpace(filter) Then Return False

        Dim f = filter.Trim()

        ' --- Immediately reject anything that is clearly a full program ---
        If f.Contains("->.") Then Return False          ' set assignment
        If f.Contains(";") AndAlso f.Count(Function(c) c = ";"c) > 1 Then Return False
        If f.Contains("(") AndAlso f.Contains(")") AndAlso f.Contains(";") Then Return False
        If f.Contains(".") AndAlso f.Contains(";") Then Return False

        ' --- Reject selector blocks ---
        If f.StartsWith("(") AndAlso f.EndsWith(")") Then Return False

        ' --- Reject multi-line queries ---
        If f.Contains(vbCr) OrElse f.Contains(vbLf) Then Return False

        ' --- Reject anything with recursion or set operations ---
        If f.Contains("._") OrElse f.Contains(">>") OrElse f.Contains("<<") Then Return False
        If f.Contains("- .") OrElse f.Contains("+ .") Then Return False

        ' --- If it reaches here, it's a simple selector ---
        Return True
    End Function
    Private Function QuoteValue(value As String) As String
        Return """" & value.Replace("""", "\""") & """"
    End Function



    ' ============================================================
    '  Geometry Normalization
    ' ============================================================

    Private Function NormalizeGeometry(geometry As String) As String
        If String.IsNullOrWhiteSpace(geometry) Then Return ""

        Dim g = geometry.Trim()

        ' bbox: "minLat,minLon,maxLat,maxLon"
        If g.Contains(",") AndAlso g.Count(Function(c) c = ","c) = 3 Then
            Return $"({g})"
        End If

        ' poly: handled inside InjectGeometry
        If g.StartsWith("poly:", StringComparison.OrdinalIgnoreCase) Then
            Return g
        End If

        Return ""
    End Function


    ' ============================================================
    '  Final Query Builder (Hardened)
    ' ============================================================

    Public Function BuildOverpassQuery(queryCore As String, bbox As String) As String
        Dim sb As New StringBuilder

        ' --- Standard header ---
        sb.AppendLine("[out:json][timeout:300];")

        ' --- Detect query type ---
        Dim isFullProgram As Boolean =
        queryCore.Contains("->.") OrElse
        queryCore.Contains(";") AndAlso queryCore.Split(";"c).Length > 2

        Dim isSimpleSelector As Boolean =
        Not isFullProgram AndAlso
        Not queryCore.Contains("(") AndAlso
        Not queryCore.Contains(")")

        ' --- CASE 1: Simple selector (e.g. rel["ISO3166-1"="CA"]) ---
        If isSimpleSelector Then
            If Not String.IsNullOrWhiteSpace(bbox) Then
                sb.AppendLine($"{queryCore}({bbox});")
            Else
                sb.AppendLine($"{queryCore};")
            End If



            ' --- CASE 2: Full Overpass program (China, HK/MO subtraction, etc.) ---
        ElseIf isFullProgram Then
            ' Do NOT modify the stored script
            sb.AppendLine(queryCore)

            ' Apply bbox ONLY to the final expansion
            If Not String.IsNullOrWhiteSpace(bbox) Then
                sb.AppendLine($"({bbox});")
            End If

        Else

            ' --- CASE 3: Selector block (rare, but supported) ---
            ' Example: ( rel[…]; way[…]; )->.r;
            sb.AppendLine(queryCore)

            If Not String.IsNullOrWhiteSpace(bbox) Then
                sb.AppendLine($"({bbox});")
            End If
        End If

        sb.AppendLine("(._; rel(r); way(r); node(w););")
        sb.AppendLine("out geom 0.001;")

        Return sb.ToString()
    End Function


    ' ============================================================
    '  Sanitizer (Prevents ALL Bad Request syntax errors)
    ' ============================================================

    Private Function CleanOverpassQuery(q As String) As String
        Dim s As String = q

        s = s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        While s.Contains(";;")
            s = s.Replace(";;", ";")
        End While

        s = s.Replace(";)", ")")
        s = s.Replace(";" & vbLf, vbLf)

        Dim lines = s.Split(vbLf)
        Dim cleaned As New List(Of String)

        For Each line In lines
            If line.Trim() <> "" Then cleaned.Add(line)
        Next

        Return String.Join(vbLf, cleaned).Trim()
    End Function

    ' Ensure selector ends with a semicolon
    Private Function EnsureSemicolon(s As String) As String
        Dim t = s.TrimEnd()
        If Not t.EndsWith(";"c) Then t &= ";"
        Return t
    End Function

End Module

' ==========================
' Overpass ingest (JSON → OSMObject)
' ==========================
' NOTE: This assumes you have a JSON library that gives you dynamic-like access.
' Replace `dynamic` with your actual JSON type and adjust indexing accordingly.
Public Module Overpass

    Public Function ParseOSMResponse(root As JsonElement) As List(Of OSMObject)
        Dim results As New List(Of OSMObject), elements As JsonElement

        If Not root.TryGetProperty("elements", elements) Then
            Return results
        End If

        For Each el In elements.EnumerateArray()
            Dim t As String = el.GetProperty("type").GetString()
            Dim id As Long = el.GetProperty("id").GetInt64()

            ' -------------------------
            ' Parse tags (common to all)
            ' -------------------------
            Dim tags As Dictionary(Of String, String) = Nothing
            Dim tagsProp As JsonElement

            If el.TryGetProperty("tags", tagsProp) Then
                tags = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                For Each tag In tagsProp.EnumerateObject()
                    tags(tag.Name) = tag.Value.GetString()
                Next
            End If

            Select Case t

        ' ============================================================
        ' NODE
        ' ============================================================
                Case "node"
                    Dim n As New OSMNode With {
                    .Id = id,
                    .Lat = el.GetProperty("lat").GetDouble(),
                    .Lon = el.GetProperty("lon").GetDouble()
                }

                    results.Add(New OSMObject With {
                    .Type = "node",
                    .ID = id,
                    .Node = n,
                    .Tags = tags
                })

' ============================================================
' WAY
' ============================================================
                Case "way"
                    Dim nds As New List(Of Long)
                    Dim nodesProp As JsonElement

                    If el.TryGetProperty("nodes", nodesProp) Then
                        For Each nd In nodesProp.EnumerateArray()
                            nds.Add(nd.GetInt64())
                        Next
                    End If

                    ' ---- NEW: Parse geometry if present ----
                    Dim geom As List(Of OSMNode) = Nothing
                    Dim geomProp As JsonElement

                    If el.TryGetProperty("geometry", geomProp) Then
                        geom = New List(Of OSMNode)
                        For Each g In geomProp.EnumerateArray()
                            geom.Add(New OSMNode With {
                .Lat = g.GetProperty("lat").GetDouble(),
                .Lon = g.GetProperty("lon").GetDouble()
            })
                        Next
                    End If

                    results.Add(New OSMObject With {
        .Type = "way",
        .ID = id,
        .Nodes = nds,
        .Geometry = geom,   ' <-- NEW
        .Tags = tags
    })

        ' ============================================================
        ' RELATION
        ' ============================================================
                Case "relation"
                    Dim members As New List(Of OSMMember)
                    Dim memProp As JsonElement

                    If el.TryGetProperty("members", memProp) Then
                        For Each m In memProp.EnumerateArray()
                            Dim role As String = ""
                            Dim roleProp As JsonElement

                            If m.TryGetProperty("role", roleProp) Then
                                role = roleProp.GetString()
                            End If

                            members.Add(New OSMMember With {
                            .Type = m.GetProperty("type").GetString(),
                            .Ref = m.GetProperty("ref").GetInt64(),
                            .Role = role
                        })
                        Next
                    End If

                    results.Add(New OSMObject With {
                    .Type = "relation",
                    .ID = id,
                    .Members = members,
                    .Tags = tags
                })

            End Select
        Next

        Return results
    End Function

End Module

' ==========================
' Geometry builder
' ==========================
Public Module GeometryBuilder

    ' Optional: simple 1 req/sec throttle
    Private lastRequest As DateTime = DateTime.MinValue
    Private ReadOnly throttleLock As New Object()

    Private Async Function ThrottleAsync() As Task
        Dim delay As TimeSpan = TimeSpan.Zero

        SyncLock throttleLock
            Dim nextAllowed = lastRequest.AddMilliseconds(2000)

            If nextAllowed > Now Then
                delay = nextAllowed - Now
            End If
        End SyncLock

        If delay > TimeSpan.Zero Then
            Await Task.Delay(delay)
        End If

        SyncLock throttleLock
            lastRequest = DateTime.UtcNow
        End SyncLock
    End Function

    Public Async Function RunQueryAsync(query As String) As Task(Of JsonDocument)
        Dim sw As New Stopwatch

        Await ThrottleAsync()

        'Dim url As String = "https://overpass.kumi.systems/api/interpreter"
        'Dim url As String = "https://overpass-api.de/api/interpreter"
        Dim url As String = "https://overpass.openstreetmap.fr/api/interpreter"

        Debug.WriteLine($"query={query}")
        ' Overpass requires: data=<query>
        Dim content As New StringContent("data=" & query, Encoding.UTF8, "application/x-www-form-urlencoded")
        sw.Start()
        Dim response = Await Http.PostAsync(url, content)
        ' If Overpass returns 400/429/etc, throw

        response.EnsureSuccessStatusCode()

        Dim jsonText = Await response.Content.ReadAsStringAsync()
        Dim elapsed = sw.Elapsed
        Debug.WriteLine($"{jsonText.Count} bytes returned from OSM after {elapsed.Seconds:f1}")
        If jsonText.TrimStart().StartsWith("<"c) Then
            Throw New Exception("Overpass returned HTML instead of JSON: " & jsonText)
        End If
        ' Parse JSON into JsonDocument
        Return JsonDocument.Parse(jsonText)
    End Function

End Module
