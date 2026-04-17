Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class PointD
    Public Lat As Double
    Public Lon As Double
    Public Sub New(Lat As Double, Lon As Double)
        Me.Lat = Lat
        Me.Lon = Lon
    End Sub
End Class
Public Class Ring
    Public Property Points As List(Of PointD)
    Public Property IsHole As Boolean = False

    Public Sub New()
        Points = New List(Of PointD)
    End Sub
    Public Sub Reverse()
        Points.Reverse()
    End Sub
End Class
Public Module helpers
    Public Const DXCC_DATA = "data Source=DXCC.sqlite"     ' the DXCC database
    Public Http As HttpClient
    Public HttpHandler As HttpClientHandler
    Public Sub InitHttp()
        HttpHandler = New HttpClientHandler()

        Http = New HttpClient(HttpHandler) With {
        .Timeout = TimeSpan.FromMinutes(10)}

        Http.DefaultRequestHeaders.UserAgent.Clear()
        Http.DefaultRequestHeaders.UserAgent.ParseAdd("DXCCMapper/1.0 (contact: myhillman@gmail.com)")
    End Sub

    Public Delegate Sub SetTextCallback(tb As System.Windows.Forms.TextBox, ByVal text As String)

    Public Sub SetText(tb As System.Windows.Forms.TextBox, ByVal text As String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf SetText), New Object() {tb, text})
        Else
            tb.Text = text
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Public Sub AppendText(tb As System.Windows.Forms.TextBox, ByVal text As String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf AppendText), New Object() {tb, text})
        Else
            tb.AppendText(text)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub
    Public Function IsDBNullorEmpty(ByVal value As Object) As Boolean
        ' test if DB field is null or empty
        Return IsDBNull(value) OrElse value = ""
    End Function
    Public Function SafeStr(value As Object) As String
        If value Is Nothing OrElse value Is DBNull.Value Then Return ""
        Return value.ToString()
    End Function
    Public Function SafeDbl(value As Object) As Double
        If value Is Nothing OrElse value Is DBNull.Value Then Return 0
        Return CDbl(value.ToString())
    End Function
    Public Function CoIncident(a As MapPoint, b As MapPoint) As Boolean
        ' test if points are coincident
        Debug.Assert(a.SpatialReference.Wkid = b.SpatialReference.Wkid, "Spatial references must be the same")
        If Math.Sign(a.Y) = Math.Sign(b.Y) And Math.Abs(a.Y) >= 89.5 And Math.Abs(b.Y) >= 89.5 Then Return True        ' at the poles, longitude is irrelevant
        If a.Y = b.Y And Math.Abs(a.X) = 180 And Math.Abs(b.X) = 180 Then Return True           ' +180 and -180 the same point
        Return a.IsEqual(b)
    End Function
    Public Function SQLescape(st As String) As String
        ' escape special characters for SQL
        Return st.Replace("'", "''")
    End Function
    Public Function KMLescape(st As String) As String
        ' escape special characters for KML
        Dim escapes As New Dictionary(Of String, String) From {
            {"<", "&lt;"},
            {">", "&gt;"},
            {"&", "&amp;"},
            {"""", "&quot;"},
            {"\'", "&apos;"}
            }

        Dim result As String
        result = st
        For Each s In escapes
            result = result.Replace(s.Key, s.Value)
        Next
        Return result
    End Function
    Public Function DegtoRad(deg As Double) As Double
        ' Convert degrees to radians
        Return deg * Math.PI / 180
    End Function
    Public Function RadtoDeg(rad As Double) As Double
        ' Convert radians to degrees
        Return rad * 180 / Math.PI
    End Function
    Public Function LoadPolygon(jsonText As String) As Polygon
        If String.IsNullOrWhiteSpace(jsonText) Then
            Return Nothing
        End If
        Dim geom As Geometry = GeoJsonToGeometry(jsonText)
        Dim poly As Polygon = TryCast(geom, Polygon)
        Return poly

    End Function

    Public Function Hyperlink(links As String) As String
        ' convert a list of hyperlinks, separated by semi-colon, and return html hyperlinks separated by <br>
        Dim hyperlinkList As New List(Of String)
        Dim hyperlinks = links.Split(";").ToList
        For ndx = 0 To hyperlinks.Count - 1
            hyperlinks(ndx) = Regex.Replace(hyperlinks(ndx), "[^0-9a-zA-Z\-\:\./]", "")     ' remove noise characters
        Next
        For Each link In hyperlinks
            Dim matches = Regex.Match(link, "^https.*?-(.*)$")
            hyperlinkList.Add($"<a href=""{link}"">{matches.Groups(1)}</a>")
        Next
        Return String.Join(";<br>", hyperlinkList)
    End Function
    Public Function EnsureClockwise(ring As List(Of MapPoint)) As List(Of MapPoint)
        ' Compute signed area (shoelace)
        Dim area As Double = 0
        For i = 0 To ring.Count - 2
            area += (ring(i + 1).X - ring(i).X) * (ring(i + 1).Y + ring(i).Y)
        Next

        ' ArcGIS outer rings must be clockwise (area < 0)
        If area > 0 Then
            ring.Reverse()
        End If

        Return ring
    End Function

    Public Function PolygonArea(ring As IEnumerable(Of MapPoint)) As Double
        ' Shoelace formula for signed polygon area
        ' https://en.wikipedia.org/wiki/Shoelace_formula Triangle formula
        ' Positive = CCW (outer), Negative = CW (inner)
        Dim pts = ring.ToList()
        Dim area As Double = 0

        If pts.Count > 2 Then
            For i = 0 To pts.Count - 1
                Dim j = (i + 1) Mod pts.Count
                area += pts(i).X * pts(j).Y - pts(j).X * pts(i).Y
            Next
        End If

        Return area / 2
    End Function
    Public Function PolygonArea(coords As List(Of Double())) As Double
        Dim area As Double = 0
        Dim n = coords.Count

        If n < 3 Then Return 0

        For i = 0 To n - 1
            Dim j = (i + 1) Mod n
            area += coords(i)(0) * coords(j)(1) - coords(j)(0) * coords(i)(1)
        Next

        Return area / 2.0
    End Function

    Function DensifyCleanMergeOrient(bbox As Geometry, degrees As Double) As Polygon
        Dim poly As Polygon

        '---------------------------------------------------------
        ' 1. Convert Envelope → Polygon (always CCW)
        '---------------------------------------------------------
        If TypeOf bbox Is Envelope Then
            Dim env = CType(bbox, Envelope)
            Dim pb As New PolygonBuilder(env.SpatialReference)

            pb.AddPoint(env.XMin, env.YMax) ' top-left
            pb.AddPoint(env.XMin, env.YMin) ' bottom-left
            pb.AddPoint(env.XMax, env.YMin) ' bottom-right
            pb.AddPoint(env.XMax, env.YMax) ' top-right
            pb.AddPoint(env.XMin, env.YMax) ' close

            poly = pb.ToGeometry()
        ElseIf TypeOf bbox Is Polygon Then
            poly = CType(bbox, Polygon)
        Else
            ' Unexpected type → return empty polygon
            Return New PolygonBuilder(SpatialReferences.Wgs84).ToGeometry()
        End If

        '---------------------------------------------------------
        ' 2. Densify (Runtime 200.x returns Geometry, not Polygon)
        '---------------------------------------------------------
        Dim g As Geometry = GeometryEngine.Densify(poly, degrees)

        If g.GeometryType <> GeometryType.Polygon Then
            ' Should not happen, but fallback
            Return poly
        End If

        Dim d As Polygon = CType(g, Polygon)

        '---------------------------------------------------------
        ' 3. Merge all parts into a single ring
        '    (200.x densify often splits long edges)
        '---------------------------------------------------------
        Dim merged As New List(Of MapPoint)

        For Each part In d.Parts
            Dim pts = part.Points
            For i = 0 To pts.Count - 2   ' skip closing point
                merged.Add(pts(i))
            Next
        Next

        '---------------------------------------------------------
        ' 4. Clean densify artifacts
        '---------------------------------------------------------
        merged = CleanRingPoints(merged)

        '---------------------------------------------------------
        ' 5. Normalize ring order (stabilize starting vertex)
        '---------------------------------------------------------
        merged = NormalizeRingOrder(merged)

        '---------------------------------------------------------
        ' 6. Fix orientation (KML outer rings must be CCW)
        '---------------------------------------------------------
        Dim area = PolygonArea(merged)
        If area < 0 Then merged.Reverse()

        '---------------------------------------------------------
        ' 7. Rebuild polygon
        '---------------------------------------------------------
        Dim pbFinal As New PolygonBuilder(poly.SpatialReference)
        pbFinal.AddPart(merged)
        Return pbFinal.ToGeometry()
    End Function


    Function NormalizeRingOrder(pts As List(Of MapPoint)) As List(Of MapPoint)
        ' Find lexicographically smallest point (lon, then lat)
        Dim minIndex = 0
        For i = 1 To pts.Count - 1
            If pts(i).X < pts(minIndex).X OrElse
           (pts(i).X = pts(minIndex).X AndAlso pts(i).Y < pts(minIndex).Y) Then
                minIndex = i
            End If
        Next

        ' Rotate ring so smallest point is first
        Dim rotated As New List(Of MapPoint)
        For i = 0 To pts.Count - 1
            rotated.Add(pts((minIndex + i) Mod pts.Count))
        Next

        ' Ensure closure
        If rotated(0).X <> rotated(rotated.Count - 1).X OrElse
       rotated(0).Y <> rotated(rotated.Count - 1).Y Then
            rotated.Add(rotated(0))
        End If

        Return rotated
    End Function

    Public Function ConvertOsmToEsri(polys As List(Of OSMPolygon)) As Polygon
        Dim builder As New Esri.ArcGISRuntime.Geometry.PolygonBuilder(SpatialReferences.Wgs84)

        For Each poly In polys

            ' OUTER RING
            Dim outerPc As New PointCollection(SpatialReferences.Wgs84)
            For Each pt In poly.Outer
                outerPc.Add(New MapPoint(pt.Lon, pt.Lat, SpatialReferences.Wgs84))
            Next
            builder.AddPart(outerPc)

            ' HOLES
            For Each hole In poly.Holes
                Dim holePc As New PointCollection(SpatialReferences.Wgs84)
                For Each pt In hole
                    holePc.Add(New MapPoint(pt.Lon, pt.Lat, SpatialReferences.Wgs84))
                Next
                builder.AddPart(holePc)
            Next

        Next

        Return builder.ToGeometry()
    End Function
    Public Function NormalizeRingTo180(pts As List(Of MapPoint)) As List(Of MapPoint)
        Dim out As New List(Of MapPoint)
        If pts.Count = 0 Then Return out

        ' Start with first point
        out.Add(pts(0))

        For i = 1 To pts.Count - 1
            Dim prev = out(out.Count - 1)
            Dim curr = pts(i)
            Dim x = curr.X

            ' ---------------------------------------------------------
            ' 1. Maintain continuity (no jumps > 180°)
            ' ---------------------------------------------------------
            While x - prev.X > 180
                x -= 360
            End While
            While x - prev.X < -180
                x += 360
            End While

            ' ---------------------------------------------------------
            ' 2. Clamp into [-180, 180] ONLY if outside range
            '    (do NOT unwrap or shift valid longitudes)
            ' ---------------------------------------------------------
            If x > 180 Then x -= 360
            If x < -180 Then x += 360

            out.Add(New MapPoint(x, curr.Y, curr.SpatialReference))
        Next

        ' ---------------------------------------------------------
        ' 3. Enforce CCW orientation (outer ring)
        ' ---------------------------------------------------------
        If PolygonArea(out) < 0 Then
            out.Reverse()
        End If

        Return out
    End Function
    ''' <summary>
    ''' Normalizes ring orientation for a polygon: outers CCW, inners CW.
    ''' Uses point-in-polygon classification and basic antimeridian handling.
    ''' </summary>
    ' Fix ring orientation and classify rings into outers/inners,
    ' then return a single Polygon (multipolygon if needed).
    Public Function FixRingOrientation(poly As Polygon) As Polygon
        If poly Is Nothing OrElse poly.Parts.Count = 0 Then Return poly
        Dim sr = poly.SpatialReference
        ' 1. Extract rings
        Dim rings As New List(Of List(Of MapPoint))
        For Each part In poly.Parts
            rings.Add(part.Points.ToList())
        Next

        ' 2. Build metadata
        Dim n = rings.Count
        Dim areas(n - 1) As Double
        Dim polys(n - 1) As Polygon
        Dim interiorPts(n - 1) As MapPoint

        For i = 0 To n - 1
            areas(i) = PolygonArea(rings(i))
            polys(i) = BuildPolygonFromRing(rings(i), sr)
            interiorPts(i) = GeometryEngine.LabelPoint(polys(i))   ' ⭐ guaranteed interior point
        Next

        ' 3. Sort indices by absolute area (largest first)
        Dim order = Enumerable.Range(0, n).OrderByDescending(Function(i) Math.Abs(areas(i))).ToList()

        ' parent(i) = index of containing ring, or -1 if none
        Dim parent(n - 1) As Integer
        For i = 0 To n - 1
            parent(i) = -1
        Next

        ' 4. Containment-based classification
        For Each i In order
            Dim childPt = interiorPts(i)

            For Each j In order
                If j = i Then Continue For

                ' ⭐ Reliable containment test
                If GeometryEngine.Contains(polys(j), childPt) Then
                    ' If j contains i, j is the parent
                    parent(i) = j
                End If
            Next
        Next

        ' 5. Group into outers with their inners
        Dim outerToInners As New Dictionary(Of Integer, List(Of Integer))
        For i = 0 To n - 1
            If parent(i) = -1 Then
                If Not outerToInners.ContainsKey(i) Then
                    outerToInners(i) = New List(Of Integer)()
                End If
            Else
                If Not outerToInners.ContainsKey(parent(i)) Then
                    outerToInners(parent(i)) = New List(Of Integer)()
                End If
                outerToInners(parent(i)).Add(i)
            End If
        Next

        ' 6. Build result polygon (may be multipolygon)
        Dim resultPolys As New List(Of Polygon)

        For Each kvp In outerToInners
            Dim outerIdx = kvp.Key
            Dim innerIdxs = kvp.Value

            Dim pb As New PolygonBuilder(sr)

            ' Outer → enforce CCW
            Dim outerRing = New List(Of MapPoint)(rings(outerIdx))
            If PolygonArea(outerRing) < 0 Then outerRing.Reverse()
            pb.AddPart(outerRing)

            ' Inners → enforce CW
            For Each ii In innerIdxs
                Dim innerRing = New List(Of MapPoint)(rings(ii))
                If PolygonArea(innerRing) > 0 Then innerRing.Reverse()
                pb.AddPart(innerRing)
            Next

            resultPolys.Add(pb.ToGeometry())
        Next

        ' 7. If multiple outers, union into one multipolygon
        Dim result As Geometry = resultPolys(0)
        For i = 1 To resultPolys.Count - 1
            result = GeometryEngine.Union(result, resultPolys(i))
        Next
        Return CType(result, Polygon)
    End Function

    Public Function FixRingOrientationRings(poly As Polygon) As List(Of List(Of Double()))
        Dim rings As New List(Of List(Of Double()))()

        ' Extract coordinates NOW, before ArcGIS can rewrite anything
        For Each part In poly.Parts
            Dim ring As New List(Of Double())()
            For Each pt In part.Points
                ring.Add(New Double() {pt.X, pt.Y})
            Next
            rings.Add(ring)
        Next

        ' Orientation
        For Each ring In rings
            If PolygonArea(ring) < 0 Then
                ring.Reverse()
            End If
        Next

        Return rings
    End Function


    Private Class OuterRing
        Public Property Outer As List(Of MapPoint)
        Public Property OuterPolygon As Polygon
        Public Property OuterEnv As Envelope
        Public Property Inners As List(Of List(Of MapPoint))
    End Class

    Public Function CleanPoly(poly As Polygon) As Polygon
        ' Used to clean up the mess made by Densify
        Dim pb As New PolygonBuilder(poly.SpatialReference)

        For Each part In poly.Parts
            Dim pts = part.Points.ToList()
            Dim cleaned = CleanRingPoints(pts)
            pb.AddPart(cleaned)
        Next

        Return pb.ToGeometry()
    End Function
    Function CleanRingPoints(pts As IList(Of MapPoint)) As List(Of MapPoint)
        Dim cleaned As New List(Of MapPoint)

        ' Remove duplicate consecutive points
        For i = 0 To pts.Count - 1
            Dim p = pts(i)
            Dim prev = If(i = 0, Nothing, pts(i - 1))

            If prev IsNot Nothing AndAlso p.X = prev.X AndAlso p.Y = prev.Y Then
                Continue For
            End If

            cleaned.Add(p)
        Next

        ' Remove duplicate last point if same as first
        If cleaned.Count > 1 Then
            Dim first = cleaned(0)
            Dim last = cleaned(cleaned.Count - 1)
            If first.X = last.X AndAlso first.Y = last.Y Then
                cleaned.RemoveAt(cleaned.Count - 1)
            End If
        End If

        ' Ensure ring is closed
        cleaned.Add(cleaned(0))

        Return cleaned
    End Function

    Function MergeParts(poly As Polygon) As Polygon
        Dim allPts As New List(Of MapPoint)

        For Each part In poly.Parts
            ' Skip the closing point of each part
            For i = 0 To part.Points.Count - 2
                allPts.Add(part.Points(i))
            Next
        Next

        ' Close the ring
        allPts.Add(allPts(0))

        Dim pb As New PolygonBuilder(poly.SpatialReference)
        pb.AddPart(allPts)
        Return pb.ToGeometry()
    End Function

    ' Helper: signed area (WGS84 ok for orientation)
    Private Function PolygonArea(ring As List(Of MapPoint)) As Double
        Dim sum As Double = 0
        For i = 0 To ring.Count - 2
            sum += (ring(i).X * ring(i + 1).Y) - (ring(i + 1).X * ring(i).Y)
        Next
        Return sum / 2.0
    End Function

    ' Helper: build a Polygon from a single ring
    Private Function BuildPolygonFromRing(ring As List(Of MapPoint), sr As SpatialReference) As Polygon
        Dim pb As New PolygonBuilder(sr)
        pb.AddPart(ring)
        Return pb.ToGeometry()
    End Function




    ''' <summary>
    ''' Normalizes ring longitudes to reduce antimeridian issues by shifting coordinates
    ''' into a consistent range when a ring clearly crosses the 180° meridian.
    ''' </summary>
    Private Sub NormalizeRingsAcrossAntimeridian(rings As List(Of List(Of MapPoint)))
        For Each ring In rings
            If ring.Count < 2 Then Continue For

            Dim minLon = ring.Min(Function(p) p.X)
            Dim maxLon = ring.Max(Function(p) p.X)

            ' Heuristic: if span > 180°, assume antimeridian crossing
            If (maxLon - minLon) > 180 Then

                ' Count points on each side of the prime meridian
                Dim positives = ring.Where(Function(p) p.X > 0).Count()
                Dim negatives = ring.Count - positives

                Dim shift As Double = If(positives >= negatives, 360.0, -360.0)

                For i = 0 To ring.Count - 1
                    Dim p = ring(i)
                    If shift > 0 AndAlso p.X < 0 Then
                        ring(i) = New MapPoint(p.X + 360.0, p.Y, p.SpatialReference)
                    ElseIf shift < 0 AndAlso p.X > 0 Then
                        ring(i) = New MapPoint(p.X - 360.0, p.Y, p.SpatialReference)
                    End If
                Next
            End If
        Next
    End Sub

    Function Avoid180(geom As Polygon) As Polygon
        ' Nudge geometry away from 180 if it is close to avoid ArcGIS processing problems. We can get away with a small nudge because the data is very coarse and we will be generalizing it anyway
        Dim builder As New PolygonBuilder(SpatialReferences.Wgs84)
        Const EPS = 0.0001
        For Each part In geom.Parts
            Dim newPart As New List(Of MapPoint)
            For Each pt In part.Points
                Dim lon = pt.X
                Dim lat = pt.Y
                If lon > 180 - EPS Then lon = 180 - EPS
                If lon < -180 + EPS Then lon = -180 + EPS
                newPart.Add(New MapPoint(lon, lat, SpatialReferences.Wgs84))
            Next
            builder.AddPart(newPart)
        Next
        Return builder.ToGeometry()
    End Function
    ' ====================================================================
    ' FUNCTION: HashText
    '
    ' PURPOSE:
    '   Produces a SHA-256 hash of a string.
    '
    ' WHY:
    '   • Allows change detection
    '   • Ensures stable identifiers even if expression order changes
    ' ====================================================================
    Public Function HashText(input As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes = Encoding.UTF8.GetBytes(input)
            Dim hash = sha.ComputeHash(bytes)
            Return Convert.ToHexString(hash)
        End Using
    End Function

    ' ====================================================================
    ' GeoJson support
    ' ====================================================================

    ' Provides:
    '   GeometryToGeoJson - convert ArcGis geometry to GeoJson string
    '   GeoJsonToGeometry - convert GeoJson string to ArcGis geometry

    Public Function GeometryToGeoJson(geom As Geometry) As String

        Select Case geom.GeometryType

            Case GeometryType.Point
                Dim p = CType(geom, MapPoint)
                Return JsonConvert.SerializeObject(New With {
                .type = "Point",
                .coordinates = New Double() {p.X, p.Y}
            })

            Case GeometryType.Multipoint
                Dim mp = CType(geom, Multipoint)
                Dim coords = mp.Points.Select(Function(pt) New Double() {pt.X, pt.Y}).ToList()
                Return JsonConvert.SerializeObject(New With {
                .type = "MultiPoint",
                .coordinates = coords
            })

            Case GeometryType.Polyline
                Dim pl = CType(geom, Polyline)
                Dim lines = New List(Of List(Of Double()))()

                For Each part In pl.Parts
                    lines.Add(part.Points.Select(Function(pt) New Double() {pt.X, pt.Y}).ToList())
                Next

                If lines.Count = 1 Then
                    Return JsonConvert.SerializeObject(New With {
                    .type = "LineString",
                    .coordinates = lines(0)
                })
                Else
                    Return JsonConvert.SerializeObject(New With {
                    .type = "MultiLineString",
                    .coordinates = lines
                })
                End If

            Case GeometryType.Polygon
                Dim poly = CType(geom, Polygon)

                ' orientedRings is now List(Of List(Of Double()))
                Dim orientedRings = FixRingOrientationRings(poly)

                Dim rings As New List(Of List(Of Double()))()

                For Each ring In orientedRings
                    Dim coords As New List(Of Double())()

                    For Each pt In ring
                        ' pt is Double() now
                        coords.Add(New Double() {pt(0), pt(1)})
                    Next

                    ' Ensure closure
                    If coords.Count > 0 Then
                        Dim first = coords(0)
                        Dim last = coords(coords.Count - 1)
                        If first(0) <> last(0) OrElse first(1) <> last(1) Then
                            coords.Add(first)
                        End If
                    End If

                    rings.Add(coords)
                Next

                Return JsonConvert.SerializeObject(New With {
                    .type = "Polygon",
                    .coordinates = rings
                })


            Case Else
                Throw New Exception("Unsupported geometry type: " & geom.GeometryType.ToString())
        End Select
    End Function

    Public Function GeoJsonToGeometry(json As String) As Geometry
        Dim jo = JObject.Parse(json)
        Dim t = jo("type").ToString()
        Dim sr = SpatialReferences.Wgs84

        Select Case t

            Case "Point"
                Dim c = jo("coordinates")
                Return New MapPoint(c(0).ToObject(Of Double)(), c(1).ToObject(Of Double)(), sr)

            Case "MultiPoint"
                Dim pts = jo("coordinates").Select(
                Function(c) New MapPoint(c(0).ToObject(Of Double)(), c(1).ToObject(Of Double)(), sr)
            )
                Return New Multipoint(pts, sr)

            Case "LineString"
                Return ImportLineString(jo("coordinates"), sr)

            Case "MultiLineString"
                Return ImportMultiLineString(jo("coordinates"), sr)

            Case "Polygon"
                Return ImportPolygon(jo("coordinates"), sr)

            Case "MultiPolygon"
                Return ImportMultiPolygon(jo("coordinates"), sr)

            Case Else
                Throw New Exception("Unsupported GeoJSON type: " & t)
        End Select
    End Function
    Private Function ImportLineString(coords As JToken, sr As SpatialReference) As Polyline
        Dim pb As New PolylineBuilder(sr)
        Dim pts = coords.Select(Function(c) New MapPoint(c(0), c(1), sr)).ToList()
        pb.AddPart(pts)
        Return pb.ToGeometry()
    End Function
    Private Function ImportMultiLineString(coords As JToken, sr As SpatialReference) As Polyline
        Dim pb As New PolylineBuilder(sr)

        For Each line In coords
            Dim pts = line.Select(Function(c) New MapPoint(c(0), c(1), sr)).ToList()
            pb.AddPart(pts)
        Next

        Return pb.ToGeometry()
    End Function
    Private Function ImportPolygon(coords As JToken, sr As SpatialReference) As Polygon
        Dim pb As New PolygonBuilder(sr)

        For Each ring In coords
            Dim pts = ring.Select(Function(c) New MapPoint(c(0), c(1), sr)).ToList()
            pb.AddPart(pts)
        Next

        Return pb.ToGeometry()
    End Function
    Private Function ImportMultiPolygon(coords As JToken, sr As SpatialReference) As Polygon
        Dim pb As New PolygonBuilder(sr)

        For Each poly In coords
            For Each ring In poly
                Dim pts = ring.Select(Function(c) New MapPoint(c(0), c(1), sr)).ToList()
                pb.AddPart(pts)
            Next
        Next

        Return pb.ToGeometry()
    End Function

End Module
