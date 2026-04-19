Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Module GeomLibrary
    public DebugOutput As Boolean=false
    Public Function CoIncident(a As MapPoint, b As MapPoint) As Boolean
        ' test if points are coincident
        Debug.Assert(a.SpatialReference.Wkid = b.SpatialReference.Wkid, "Spatial references must be the same")
        If Math.Sign(a.Y) = Math.Sign(b.Y) And Math.Abs(a.Y) >= 89.5 And Math.Abs(b.Y) >= 89.5 Then Return True        ' at the poles, longitude is irrelevant
        If a.Y = b.Y And Math.Abs(a.X) = 180 And Math.Abs(b.X) = 180 Then Return True           ' +180 and -180 the same point
        Return a.IsEqual(b)
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
    
    Function DensifyCleanOrient(bbox As Geometry, degrees As Double) As Polygon
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
        Return New PolygonBuilder(SpatialReferences.Wgs84).ToGeometry()
    End If

    '---------------------------------------------------------
    ' 2. Densify (Runtime 200.x returns Geometry, not Polygon)
    '---------------------------------------------------------
    Dim g As Geometry = GeometryEngine.Densify(poly, degrees)

    If g.GeometryType <> GeometryType.Polygon Then
        Return poly
    End If

    Dim d As Polygon = CType(g, Polygon)

    '---------------------------------------------------------
    ' 3. CLEAN + ORIENT EACH PART INDIVIDUALLY
    '    (NO MERGE — preserves antimeridian structure)
    '---------------------------------------------------------
    Dim pbFinal As New PolygonBuilder(poly.SpatialReference)

    For Each part In d.Parts
        Dim pts = part.Points.ToList()

        ' Clean densify artifacts
        pts = CleanRingPoints(pts)

        ' Normalize ring order
        pts = NormalizeRingOrder(pts)

        ' Fix orientation (outer rings CCW)
        Dim area = PolygonArea(pts)
        If area < 0 Then pts.Reverse()

        ' Add cleaned part
        pbFinal.AddPart(pts)
    Next

    Return pbFinal.ToGeometry()
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
            areas(i) = PolygonAreaAMsafe(rings(i))
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
            If PolygonAreaAMsafe(outerRing) < 0 Then outerRing.Reverse()
            pb.AddPart(outerRing)

            ' Inners → enforce CW
            For Each ii In innerIdxs
                Dim innerRing = New List(Of MapPoint)(rings(ii))
                If PolygonAreaAMsafe(innerRing) > 0 Then innerRing.Reverse()
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

    ' Helper: signed area (WGS84 ok for orientation)
    ' +ve area is CCW which means outer
    Private Function PolygonArea(ring As List(Of MapPoint)) As Double
        Dim sum As Double = 0
        For i = 0 To ring.Count - 2
            sum += (ring(i).X * ring(i + 1).Y) - (ring(i + 1).X * ring(i).Y)
        Next
        Return sum / 2.0
    End Function
    public Function PolygonAreaAMSafe(ring As List(Of MapPoint)) As Double
        ' 1. Make a working copy
        Dim pts As New List(Of MapPoint)
        For Each p In ring
            pts.Add(New MapPoint(p.X, p.Y, p.SpatialReference))
        Next

        ' 2. Ensure closed
        If pts.First().X <> pts.Last().X OrElse pts.First().Y <> pts.Last().Y Then
            pts.Add(pts.First())
        End If

        ' 3. Unwrap longitudes (make continuous)
        For i = 1 To pts.Count - 1
            Dim prev = pts(i - 1)
            Dim cur = pts(i)
            Dim dx = cur.X - prev.X

            If dx > 180 Then
                pts(i) = New MapPoint(cur.X - 360, cur.Y, cur.SpatialReference)
            ElseIf dx < -180 Then
                pts(i) = New MapPoint(cur.X + 360, cur.Y, cur.SpatialReference)
            End If
        Next

        ' 4. Shoelace formula (now safe)
        Dim sum As Double = 0
        For i = 0 To pts.Count - 2
            sum += (pts(i).X * pts(i + 1).Y) -
                   (pts(i + 1).X * pts(i).Y)
        Next

        Return sum / 2.0
    End Function

    ' Helper: build a Polygon from a single ring
    Private Function BuildPolygonFromRing(ring As List(Of MapPoint), sr As SpatialReference) As Polygon
        Dim pb As New PolygonBuilder(sr)
        pb.AddPart(ring)
        Return pb.ToGeometry()
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

    ' ---------------------------------------------------------
    '  BBOX/POLY PARSER (lon lat order)
    ' ---------------------------------------------------------
Public Function ParseBBoxOrPoly(value As String) As Geometry
    Dim st = value.Trim()

    ' ---------------------------------------------------------
    ' 1. BBOX CASE (4 numbers)
    ' ---------------------------------------------------------
    Dim parts = st.Split(","c).Select(Function(x) x.Trim()).ToList()
    If parts.Count = 4 Then

        Dim minLon = Double.Parse(parts(0))
        Dim minLat = Double.Parse(parts(1))
        Dim maxLon = Double.Parse(parts(2))
        Dim maxLat = Double.Parse(parts(3))

        Dim rings As New List(Of List(Of MapPoint))

        ' Normal bbox (no AM crossing)
        If maxLon >= minLon AndAlso (maxLon - minLon) <= 180 Then
            rings.Add(New List(Of MapPoint) From {
                New MapPoint(minLon, minLat),
                New MapPoint(maxLon, minLat),
                New MapPoint(maxLon, maxLat),
                New MapPoint(minLon, maxLat),
                New MapPoint(minLon, minLat)
            })

        Else
            ' AM-crossing bbox → split into two rectangles

            ' WEST rectangle: minLon → -180
            rings.Add(New List(Of MapPoint) From {
                New MapPoint(minLon, minLat),
                New MapPoint(-180, minLat),
                New MapPoint(-180, maxLat),
                New MapPoint(minLon, maxLat),
                New MapPoint(minLon, minLat)
            })

            ' EAST rectangle: 180 → maxLon
            rings.Add(New List(Of MapPoint) From {
                New MapPoint(180, minLat),
                New MapPoint(maxLon, minLat),
                New MapPoint(maxLon, maxLat),
                New MapPoint(180, maxLat),
                New MapPoint(180, minLat)
            })
        End If

        ' Build ONE polygon with multiple parts
        Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
        For Each r In rings
            pb.AddPart(NormalizeRingCCW(r))
        Next

        Return pb.ToGeometry()
    End If

    ' ---------------------------------------------------------
    ' 2. POLY CASE (poly:"...")
    ' ---------------------------------------------------------
    Dim rPoly As New Regex("^poly:""(.+)""$", RegexOptions.IgnoreCase)
    Dim matches = rPoly.Match(st)
    If matches.Success Then

        Dim coords = matches.Groups(1).Value.Split(" "c).
            Select(Function(x) x.Trim()).
            Where(Function(x) x.Length > 0).
            ToList()

        If coords.Count Mod 2 <> 0 Then
            Throw New Exception($"Unrecognised bbox pattern {st}")
        End If

        ' Build raw ring
        Dim pts As New List(Of MapPoint)
        For i = 0 To coords.Count - 1 Step 2
            Dim lon = Double.Parse(coords(i))
            Dim lat = Double.Parse(coords(i + 1))
            pts.Add(New MapPoint(lon, lat, SpatialReferences.Wgs84))
        Next

        ' Close ring
        If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))

        ' Normalize continuity + CCW
        pts = NormalizeRingCCW(pts)

        ' ---- NEW: RANGE-BASED AM DETECTION ----
        Dim minLon = pts.Min(Function(p) p.X)
        Dim maxLon = pts.Max(Function(p) p.X)

        Dim rings As New List(Of List(Of MapPoint))

        If minLon < 180 AndAlso maxLon > 180 Then
            ' ---- CLIP INTO TWO PARTS ----
            Dim west = ClipRingToXMax(pts, 180)
            Dim east = ClipRingToXMin(pts, 180)

            If west.Count >= 4 Then rings.Add(west)
            If east.Count >= 4 Then rings.Add(east)

        Else
            ' No AM crossing → keep as single ring
            rings.Add(pts)
        End If

        ' Build ONE polygon with multiple parts
        Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
        For Each r In rings
            pb.AddPart(NormalizeRingCCW(r))
        Next

        Return pb.ToGeometry()
    End If

    Throw New Exception($"Unrecognised bbox pattern {st}")
End Function
    Private Function ClipRingToXMax(ring As List(Of MapPoint), xMax As Double) As List(Of MapPoint)
        Dim output As New List(Of MapPoint)
        If ring.Count = 0 Then Return output

        Dim s = ring(ring.Count - 1) ' start with last point

        For Each e In ring
            Dim sIn = (s.X <= xMax)
            Dim eIn = (e.X <= xMax)

            If sIn AndAlso eIn Then
                ' both inside → keep E
                output.Add(e)

            ElseIf sIn AndAlso Not eIn Then
                ' leaving → add intersection only
                Dim t = (xMax - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New MapPoint(xMax, y, s.SpatialReference))

            ElseIf Not sIn AndAlso eIn Then
                ' entering → add intersection + E
                Dim t = (xMax - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New MapPoint(xMax, y, s.SpatialReference))
                output.Add(e)
            End If

            s = e
        Next

        ' close ring
        If output.Count > 0 Then
            Dim f = output(0)
            Dim l = output(output.Count - 1)
            If Not f.IsEqual(l) Then output.Add(f)
        End If

        Return output
    End Function
    Private Function ClipRingToXMin(ring As List(Of MapPoint), xMin As Double) As List(Of MapPoint)
        Dim output As New List(Of MapPoint)
        If ring.Count = 0 Then Return output

        Dim s = ring(ring.Count - 1)

        For Each e In ring
            Dim sIn = (s.X >= xMin)
            Dim eIn = (e.X >= xMin)

            If sIn AndAlso eIn Then
                ' both inside → keep E
                output.Add(e)

            ElseIf sIn AndAlso Not eIn Then
                ' leaving → add intersection only
                Dim t = (xMin - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New MapPoint(xMin, y, s.SpatialReference))

            ElseIf Not sIn AndAlso eIn Then
                ' entering → add intersection + E
                Dim t = (xMin - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New MapPoint(xMin, y, s.SpatialReference))
                output.Add(e)
            End If

            s = e
        Next

        ' close ring
        If output.Count > 0 Then
            Dim f = output(0)
            Dim l = output(output.Count - 1)
            If Not f.IsEqual(l) Then output.Add(f)
        End If

        Return output
    End Function


    Private Function NormalizeRingCCW(pts As List(Of MapPoint)) As List(Of MapPoint)
        ' 1. Copy
        Dim fixed As New List(Of MapPoint)
        For Each p In pts
            fixed.Add(New MapPoint(p.X, p.Y, p.SpatialReference))
        Next

        ' 2. Ensure closed
        If fixed.First().X <> fixed.Last().X OrElse fixed.First().Y <> fixed.Last().Y Then
            fixed.Add(fixed.First())
        End If

        ' 3. Unwrap longitudes (remove >180° jumps)
        For i = 1 To fixed.Count - 1
            Dim prev = fixed(i - 1)
            Dim cur = fixed(i)
            Dim dx = cur.X - prev.X

            If dx > 180 Then
                fixed(i) = New MapPoint(cur.X - 360, cur.Y, cur.SpatialReference)
            ElseIf dx < -180 Then
                fixed(i) = New MapPoint(cur.X + 360, cur.Y, cur.SpatialReference)
            End If
        Next

        ' 4. Signed area (shoelace)
        Dim area As Double = 0
        Dim n = fixed.Count
        For i = 0 To n - 1
            Dim j = (i + 1) Mod n
            area += (fixed(i).X * fixed(j).Y) -
                    (fixed(j).X * fixed(i).Y)
        Next
        area /= 2.0

        ' 5. Enforce CCW
        If area < 0 Then
            fixed.Reverse()
        End If

        ' 6. DO NOT re-wrap here
        Return fixed
    End Function


    Private Function NudgeLon(lon As Double) As Double
        If lon < -179.999 Then Return -179.999
        If lon > 179.999 Then Return 179.999
        If lon = -180 Then Return -179.999
        If lon = 180 Then Return 179.999
        Return lon
    End Function

Public Function SplitRingAtAntimeridian(ring As List(Of MapPoint)) As List(Of List(Of MapPoint))
    Const EPS As Double = 1e-9

    Dim parts As New List(Of List(Of MapPoint))
    Dim current As New List(Of MapPoint)

    If ring Is Nothing OrElse ring.Count = 0 Then Return parts

    ' Start with first vertex
    current.Add(ring(0))

    ' ---------------------------------------------------------
    ' 1. RAW-SPACE crossing detection + splitting
    ' ---------------------------------------------------------
    For i = 1 To ring.Count - 1
        Dim p1 = current.Last()
        Dim p2 = ring(i)

        Dim lon1 = p1.X
        Dim lon2 = p2.X
        Dim lat1 = p1.Y
        Dim lat2 = p2.Y

        Dim delta_raw = lon2 - lon1

        If Math.Abs(delta_raw) > 180 + EPS Then
            ' crossing at +180 or -180
            Dim target As Double = If(delta_raw > 0, 180.0, -180.0)

            ' intersection in RAW space
            Dim t = (target - lon1) / (lon2 - lon1)
            Dim y = lat1 + t * (lat2 - lat1)

            ' end current part at target
            current.Add(New MapPoint(target, y))
            parts.Add(New List(Of MapPoint)(current))

            ' start new part at opposite meridian
            current = New List(Of MapPoint)
            current.Add(New MapPoint(-target, y))
        End If

        current.Add(p2)
    Next

    parts.Add(current)

    ' ---------------------------------------------------------
    ' 2. UNWRAP each part independently (continuous longitude)
    ' ---------------------------------------------------------
    Dim unwrappedParts As New List(Of List(Of MapPoint))

    For Each prt In parts
        Dim uw As New List(Of MapPoint)
        Dim prev = prt(0).X
        Dim offset As Double = 0

        uw.Add(New MapPoint(prev, prt(0).Y))

        For i = 1 To prt.Count - 1
            Dim lon = prt(i).X
            Dim lat = prt(i).Y

            Dim delta = lon - prev

            If delta > 180 Then offset -= 360
            If delta < -180 Then offset += 360

            uw.Add(New MapPoint(lon + offset, lat))

            prev = lon
        Next

        unwrappedParts.Add(uw)
    Next

    ' ---------------------------------------------------------
    ' 3. CLOSE (but DO NOT renormalize)
    ' ---------------------------------------------------------
    Dim finalParts As New List(Of List(Of MapPoint))

    For Each prt In unwrappedParts
        If prt.Count < 3 Then Continue For

        Dim f = prt(0)
        Dim l = prt(prt.Count - 1)

        If Not (f.X = l.X AndAlso f.Y = l.Y) Then
            prt.Add(New MapPoint(f.X, f.Y))
        End If

        finalParts.Add(prt)
    Next

    Return finalParts
End Function


    public Function WrapBackTo180(ring As List(Of MapPoint)) As List(Of MapPoint)
        Dim out As New List(Of MapPoint)

        For Each p In ring
            Dim x = p.X
            While x > 180 : x -= 360 : End While
            While x < -180 : x += 360 : End While
            out.Add(New MapPoint(x, p.Y, p.SpatialReference))
        Next

        Return out
    End Function


   ' ring winding is not guaranteed in source data, so we need to classify rings as outer vs inner, and then orient them correctly (outer CCW, inner CW) to ensure consistent geometry across all sources. This is done by building a containment hierarchy of rings, and classifying by depth parity (outer = even depth, inner = odd depth).
    Public Class ClassifiedRing
        Public Property Part As ReadOnlyPart
        Public Property Depth As Integer
        Public Property IsOuter As Boolean
        Public Property Parent As ClassifiedRing
        Public Property Children As New List(Of ClassifiedRing)
    End Class


    Private Sub AssignDepth(r As ClassifiedRing, depth As Integer)
        r.Depth = depth
        For Each c In r.Children
            AssignDepth(c, depth + 1)
        Next
    End Sub
    Private Function NormalizeLon(p As MapPoint) As MapPoint
        Dim x = p.X
        If x < 0 Then x += 360
        Return New MapPoint(x, p.Y, SpatialReferences.Wgs84)
    End Function

End Module
