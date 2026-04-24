Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports NetTopologySuite.Algorithm
Imports NetTopologySuite.Geometries

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

    ' Helper: signed area (WGS84 ok for orientation)
    ' +ve area is CCW which means outer
    Private Function PolygonArea(ring As List(Of MapPoint)) As Double
        Dim sum As Double = 0
        For i = 0 To ring.Count - 2
            sum += (ring(i).X * ring(i + 1).Y) - (ring(i + 1).X * ring(i).Y)
        Next
        Return sum / 2.0
    End Function

    ' ---------------------------------------------------------
    '  BBOX/POLY PARSER (lon lat order)
    ' ---------------------------------------------------------

    Public Function ParseBBoxOrPoly(value As String) As NetTopologySuite.Geometries.Geometry
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

            ' FIX: ensure latitudes are ordered correctly
            If minLat > maxLat Then
                Dim t = minLat
                minLat = maxLat
                maxLat = t
            End If
            Dim rings As New List(Of List(Of Coordinate))

            ' Normal bbox (no AM crossing)
            If maxLon >= minLon AndAlso (maxLon - minLon) <= 180 Then
                rings.Add(New List(Of Coordinate) From {
                New Coordinate(minLon, minLat),
                New Coordinate(maxLon, minLat),
                New Coordinate(maxLon, maxLat),
                New Coordinate(minLon, maxLat),
                New Coordinate(minLon, minLat)
            })
            Else
                ' AM-crossing bbox → split into two rectangles

                ' WEST rectangle: minLon → -180
                rings.Add(New List(Of Coordinate) From {
                New Coordinate(minLon, minLat),
                New Coordinate(-180, minLat),
                New Coordinate(-180, maxLat),
                New Coordinate(minLon, maxLat),
                New Coordinate(minLon, minLat)
            })

                ' EAST rectangle: 180 → maxLon
                rings.Add(New List(Of Coordinate) From {
                New Coordinate(180, minLat),
                New Coordinate(maxLon, minLat),
                New Coordinate(maxLon, maxLat),
                New Coordinate(180, maxLat),
                New Coordinate(180, minLat)
            })
            End If

            Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)
            For Each r In rings
                Dim norm = NormalizeRingCCW(r)
                Dim shell = factory.CreateLinearRing(norm.ToArray())
                polys.Add(factory.CreatePolygon(shell))
            Next

            If polys.Count = 1 Then
                Return polys(0)
            Else
                Return factory.CreateMultiPolygon(polys.ToArray())
            End If
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
            Dim pts As New List(Of Coordinate)
            For i = 0 To coords.Count - 1 Step 2
                Dim lon = Double.Parse(coords(i))
                Dim lat = Double.Parse(coords(i + 1))
                pts.Add(New Coordinate(lon, lat))
            Next

            ' Close ring
            If Not pts(0).Equals2D(pts.Last()) Then pts.Add(New Coordinate(pts(0).X, pts(0).Y))

            ' Normalize continuity + CCW
            pts = NormalizeRingCCW(pts)

            ' ---- RANGE-BASED AM DETECTION ----
            Dim minLon = pts.Min(Function(p) p.X)
            Dim maxLon = pts.Max(Function(p) p.X)

            Dim rings As New List(Of List(Of Coordinate))

            If minLon > maxLon Then
                ' ---- CLIP INTO TWO PARTS ----
                Dim west = ClipRingToXMax(pts, 180)
                Dim east = ClipRingToXMin(pts, 180)

                If west.Count >= 4 Then rings.Add(west)
                If east.Count >= 4 Then rings.Add(east)
            Else
                ' No AM crossing → keep as single ring
                rings.Add(pts)
            End If

            Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)
            For Each r In rings
                Dim norm = NormalizeRingCCW(r)
                Dim shell = factory.CreateLinearRing(norm.ToArray())
                polys.Add(factory.CreatePolygon(shell))
            Next

            If polys.Count = 1 Then
                Return polys(0)
            Else
                Return factory.CreateMultiPolygon(polys.ToArray())
            End If
        End If

        Throw New Exception($"Unrecognised bbox pattern {st}")
    End Function

    Private Function ClipRingToXMax(ring As List(Of Coordinate), xMax As Double) As List(Of Coordinate)
        Dim output As New List(Of Coordinate)
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
                output.Add(New Coordinate(xMax, y))

            ElseIf Not sIn AndAlso eIn Then
                ' entering → add intersection + E
                Dim t = (xMax - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New Coordinate(xMax, y))
                output.Add(e)
            End If

            s = e
        Next

        ' close ring
        If output.Count > 0 Then
            Dim f = output(0)
            Dim l = output(output.Count - 1)
            If Not f.Equals2D(l) Then output.Add(New Coordinate(f.X, f.Y))
        End If

        Return output
    End Function

    Private Function ClipRingToXMin(ring As List(Of Coordinate), xMin As Double) As List(Of Coordinate)
        Dim output As New List(Of Coordinate)
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
                output.Add(New Coordinate(xMin, y))

            ElseIf Not sIn AndAlso eIn Then
                ' entering → add intersection + E
                Dim t = (xMin - s.X) / (e.X - s.X)
                Dim y = s.Y + t * (e.Y - s.Y)
                output.Add(New Coordinate(xMin, y))
                output.Add(e)
            End If

            s = e
        Next

        ' close ring
        If output.Count > 0 Then
            Dim f = output(0)
            Dim l = output(output.Count - 1)
            If Not f.Equals2D(l) Then output.Add(New Coordinate(f.X, f.Y))
        End If

        Return output
    End Function

    Private Function NormalizeRingCCW(ring As List(Of Coordinate)) As List(Of Coordinate)
        ' Ensure closed
        If Not ring(0).Equals2D(ring(ring.Count - 1)) Then
            ring.Add(New Coordinate(ring(0).X, ring(0).Y))
        End If

        Dim arr = ring.ToArray()
        ' NTS Orientation.IsCCW returns True for CCW
        If Not Orientation.IsCCW(arr) Then
            Array.Reverse(arr)
        End If

        Return arr.ToList()
    End Function

    Public Function GridSquare(lon As Double, lat As Double) As String
        ' -------------------------------
        ' 1. Validate input ranges
        ' -------------------------------
        If Double.IsNaN(lon) OrElse Double.IsNaN(lat) Then
            Throw New ArgumentException("Latitude/longitude cannot be NaN.")
        End If

        If lon < -180 OrElse lon > 180 Then
            Throw New ArgumentOutOfRangeException(NameOf(lon),
            $"Longitude {lon} is out of range (-180 to 180).")
        End If

        If lat < -90 OrElse lat > 90 Then
            Throw New ArgumentOutOfRangeException(NameOf(lat),
            $"Latitude {lat} is out of range (-90 to 90).")
        End If

        ' -------------------------------
        ' 2. Constants
        ' -------------------------------
        Const upper = "ABCDEFGHIJKLMNOPQR"   ' 18 letters
        Const numbers = "0123456789"

        ' -------------------------------
        ' 3. Convert to Maidenhead space
        ' -------------------------------
        Dim adjLon = lon + 180.0
        Dim adjLat = lat + 90.0

        ' -------------------------------
        ' 4. Compute field (letters)
        ' -------------------------------
        Dim lonFieldIndex = CInt(Math.Floor(adjLon / 20.0))   ' 0–17
        Dim latFieldIndex = CInt(Math.Floor(adjLat / 10.0))   ' 0–17

        ' Safety clamp (should never trigger, but protects against rounding edge cases)
        lonFieldIndex = Math.Max(0, Math.Min(17, lonFieldIndex))
        latFieldIndex = Math.Max(0, Math.Min(17, latFieldIndex))

        Dim lonField = upper(lonFieldIndex)
        Dim latField = upper(latFieldIndex)

        ' -------------------------------
        ' 5. Compute square (digits)
        ' -------------------------------
        Dim lonSquareIndex = CInt(Math.Floor((adjLon Mod 20.0) / 2.0))   ' 0–9
        Dim latSquareIndex = CInt(Math.Floor(adjLat Mod 10.0))           ' 0–9

        lonSquareIndex = Math.Max(0, Math.Min(9, lonSquareIndex))
        latSquareIndex = Math.Max(0, Math.Min(9, latSquareIndex))

        Dim lonSquare = numbers(lonSquareIndex)
        Dim latSquare = numbers(latSquareIndex)

        ' -------------------------------
        ' 6. Return 4‑character Maidenhead
        ' -------------------------------
        Return $"{lonField}{latField}{lonSquare}{latSquare}"
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

End Module
