Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite

Module dxccSources
	Public Class DxccSource
		Public Property DXCCnum As Integer
		Public Property name As String
		Public Property source As String
		Public Property rule As String
		Public Property tolerance_m As Integer
		Public Property bbox As String
		Public Property notes As String
	End Class
	Public Async Sub BuildAndSaveAllDxccGeometries()
		Dim count As Integer = 0, success As Integer = 0
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			connect.CreateFunction(name:="hashV", function:=Function(values As Object()) HashVariadic(values))

			Dim sql = "SELECT DXCCnum, Entity FROM DXCC WHERE (`hash` <> hashV(`rule`,`bbox`,`tolerance_m`,`geometry`) OR geometry is NULL) AND `Deleted`=0 AND `DXCCnum`<>999 ORDER BY `Entity`"

			Using cmd As New SqliteCommand(sql, connect)
				Using r = cmd.ExecuteReader()
					While r.Read()
						count += 1
						Debug.WriteLine($"Getting geometry for {r("Entity")}")
						Dim id = r("DXCCnum")
						Try
							Dim result = Await BuildAndSaveDxccGeometry(id, connect)
							If result Then
								Debug.WriteLine($"Built geometry for DXCC ID: {id}")
								success += 1
							End If
						Catch ex As Exception
							Debug.WriteLine($"Error building geometry for DXCC ID {id}: {ex.Message}")
						End Try
					End While
				End Using
			End Using
		End Using
		Debug.WriteLine($"Done. Processed {count} DXCC entries. Successfully built {success} geometries.")
	End Sub
	Private Function HashVariadic(values As Object()) As String
		If values Is Nothing OrElse values.Length = 0 Then
			Return Nothing
		End If

		Dim parts As New List(Of String)

		For Each v In values
			If v IsNot Nothing AndAlso v IsNot DBNull.Value Then
				parts.Add(v.ToString())
			Else
				parts.Add("")   ' or skip, depending on your hashing rules
			End If
		Next

		Dim combined = String.Join("|", parts)
		Return HashText(combined)
	End Function


	Private Async Function BuildAndSaveDxccGeometry(dxccId As Integer, connect As SqliteConnection) As Task(Of Boolean)
		Try
			Dim src = LoadDxccSource(dxccId, connect)
			Dim geom = Await BuildDXCCGeometryAsync(src)
			If geom Is Nothing OrElse geom.IsEmpty Then
				Debug.WriteLine($"No geometry built for DXCC ID {dxccId}. Skipping save.")
				Return False
			End If
			SaveDxccGeometry(dxccId, geom, connect)
		Catch ex As Exception
			Debug.WriteLine($"Error building geometry for DXCC ID {dxccId}: {ex.Message}")
			Return False
		End Try
		Return True
	End Function

	Private Function LoadDxccSource(dxccId As Integer, conn As SqliteConnection) As DxccSource

		Dim sql = "SELECT DXCCnum, Entity, source, rule, bbox, tolerance_m, notes FROM DXCC WHERE DXCCnum=@id"

		Using cmd As New SqliteCommand(sql, conn)
			cmd.Parameters.AddWithValue("@id", dxccId)

			Using r = cmd.ExecuteReader()
				If Not r.Read() Then Throw New Exception("DXCC not found: " & dxccId)

				Return New DxccSource With {
					.DXCCnum = r("DXCCnum"),
					.name = SafeStr(r("Entity")),
					.source = SafeStr(r("source")),
					.rule = SafeStr(r("rule")),
					.tolerance_m = If(r.IsDBNull(r.GetOrdinal("tolerance_m")), 0, r.GetInt32(r.GetOrdinal("tolerance_m"))),
					.bbox = SafeStr(r("bbox")),
					.notes = SafeStr(r("notes"))
				}
			End Using
		End Using
	End Function

	Private Async Function BuildFromOSM(osmName As String) As Task(Of Geometry)
		' simple parser to start. ways start with W.
		Dim osmId = osmName
		Dim osmType As String = "relation"  ' default to relation, which is the most common type for countries. This means that if a user just enters "France" it will resolve to the relation, which is what we want in 99% of cases. If they want a specific way, they can enter "W.waysid"
		If osmName.StartsWith("W", StringComparison.OrdinalIgnoreCase) Then
			osmType = "way"
			osmId = osmName.Substring(1)
		End If

		Dim g = Await OsmLand.ResolveOSM(osmId, osmType)

		If g Is Nothing Then
			Throw New Exception("OSM geometry not found: " & osmName)
		End If

		'Return SimplifyGeometry(g, tol)
		Return g
	End Function
	Private Function BuildFromFake(fakeName As String) As Geometry
		' Some geometries just dont exist anywhere, especially small rocks, reefs and atolls
		' This makes fake geometry for these
		Dim p As MapPoint, g As Geometry

		Select Case fakeName
			Case "Willis Is"
				p = New MapPoint(150.2, -16.29, SpatialReferences.Wgs84)    ' center of island
				g = GeometryEngine.Buffer(p, MetersToDegrees(p, 200))       ' its a small island

			Case "Mellish Reef"
				p = New MapPoint(155.8644, -17.4126, SpatialReferences.Wgs84)    ' center of island
				g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

			Case "St Peter & St Paul Rocks"
				p = New MapPoint(-29.335, 0.9158, SpatialReferences.Wgs84)    ' center of island
				g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

			Case "Malpelo Is"
				p = New MapPoint(-81.5833, 3.9833, SpatialReferences.Wgs84)    ' center of island
				g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

			Case Else
				Throw New Exception($"FAKE parameter {fakeName} not recognised")
		End Select

		Return g
	End Function

	Public Function MetersToDegrees(g As Geometry, meters As Double) As Double
		' Use the latitude of the geometry's center
		Dim env = g.Extent
		Dim lat = (env.YMin + env.YMax) / 2.0
		Dim metersPerDegree = 111320 * Math.Cos(lat * Math.PI / 180)
		Return meters / metersPerDegree
	End Function
	Private Sub SaveDxccGeometry(dxccId As Integer, geom As Geometry, conn As SqliteConnection)
		If geom Is Nothing OrElse geom.IsEmpty Then
			Debug.WriteLine($"Geometry for DXCC {dxccId} is empty. Not saved")
			Return      ' don't save empty geometry
		End If
		If dxccId = 171 Then
			Dim poly As Polygon = geom
			Debug.WriteLine($"DXCC {171} has area {PolygonArea(poly.Parts(0).Points)}")
		End If
		Dim GeoJson As String = GeometryToGeoJson(geom)   ' GeoJSON
		If dxccId = 171 Then
			Debug.WriteLine($"DXCC 171: GeoJson being saved = {GeoJson}")
		End If

		' Update geometry and hash in one statement
		Dim sql = "UPDATE DXCC SET geometry = $g, hash = hashV(rule, bbox, tolerance_m, $g) WHERE DXCCnum = $id"

		Using cmd As New SqliteCommand(sql, conn)
			cmd.Parameters.AddWithValue("$g", GeoJson)
			cmd.Parameters.AddWithValue("$id", dxccId)
			cmd.ExecuteNonQuery()
		End Using
	End Sub
    Public Async Function BuildDXCCGeometryAsync(src As DxccSource) As Task(Of Geometry)

		' STEP 1 — Resolve base geometry
		Dim baseGeom As Geometry = Await ResolveRuleExpressionAsync(src)
		Dim degTol = src.tolerance_m / 110540       ' convert meter tolerance to degrees
		baseGeom = GeometryEngine.Generalize(baseGeom, degTol, True)
		baseGeom = RoundGeometry(baseGeom, 6)
		If baseGeom Is Nothing Then
            Debug.WriteLine($"Failed to resolve base geometry for DXCC '{src.name}' using rule: {src.rule}")
            Return Nothing
        End If

        ' STEP 2 — Apply bbox/poly if present
        If Not String.IsNullOrWhiteSpace(src.bbox) Then

            Dim clipGeom As Geometry = ParseBBoxOrPoly(src.bbox)
            If clipGeom Is Nothing Then
                Debug.WriteLine($"Failed to resolve bbox for DXCC '{src.name}' using bbox: {src.bbox}")
                Return Nothing
            End If

            ' Diagnostics
            Dim cext = clipGeom.Extent
            Debug.WriteLine($"[CLIP] {src.name}: XMin={cext.XMin}, XMax={cext.XMax}, YMin={cext.YMin}, YMax={cext.YMax}")

            Dim bext = baseGeom.Extent
            Debug.WriteLine($"[BASE] {src.name}: XMin={bext.XMin}, XMax={bext.XMax}, YMin={bext.YMin}, YMax={bext.YMax}")

            ' Clip (no splitting here)
            baseGeom = GeometryEngine.Intersection(baseGeom, clipGeom)
        End If

        If baseGeom Is Nothing OrElse baseGeom.IsEmpty Then
            Debug.WriteLine($"Resulting geometry is empty after applying bbox {src.bbox} for DXCC '{src.name}'")
            Return Nothing
        End If

		Dim ex = baseGeom.Extent
        Debug.WriteLine($"[BASE -result] {src.name}: XMin={ex.XMin}, XMax={ex.XMax}, YMin={ex.YMin}, YMax={ex.YMax}")

        Return baseGeom
    End Function



	Public Function SimplifyGeometry(geom As Geometry, tolerance As Double) As Geometry
		' Null or trivial geometry → return as-is
		If geom Is Nothing Then Return Nothing
		If TypeOf geom Is MapPoint Then Return geom
		If TypeOf geom Is Polyline AndAlso DirectCast(geom, Polyline).Parts.Count = 0 Then Return geom
		If TypeOf geom Is Polygon AndAlso DirectCast(geom, Polygon).Parts.Count = 0 Then Return geom

		' ArcGIS Runtime 200.x Simplify() is too aggressive — do NOT use it
		' Instead use Generalize() with preserveTopology:=True
		Dim simplified As Geometry = GeometryEngine.Generalize(geom, tolerance, True)

		' Generalize can introduce floating-point drift → round coordinates
		simplified = RoundGeometry(simplified, 6)

		Return simplified
	End Function
    Public Function RoundGeometry(geom As Geometry, decimals As Integer) As Geometry
        If geom Is Nothing Then Return Nothing

        If TypeOf geom Is MapPoint Then
            Dim p = DirectCast(geom, MapPoint)
            Return New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals))
        End If

        If TypeOf geom Is Polygon Then
            Dim poly = DirectCast(geom, Polygon)
            Dim builder As New PolygonBuilder(poly.SpatialReference)

            For Each part In poly.Parts
                Dim newPart As New List(Of MapPoint)
                For Each p In part.Points
                    newPart.Add(New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals)))
                Next
                builder.AddPart(newPart)
            Next

            Return builder.ToGeometry()
        End If

        If TypeOf geom Is Polyline Then
            Dim line = DirectCast(geom, Polyline)
            Dim builder As New PolylineBuilder(line.SpatialReference)

            For Each part In line.Parts
                Dim newPart As New List(Of MapPoint)
                For Each p In part.Points
                    newPart.Add(New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals)))
                Next
                builder.AddPart(newPart)
            Next

            Return builder.ToGeometry()
        End If

        Return geom
    End Function


	' ---------------------------------------------------------
	'  RULE PARSER (UNION SUPPORT) — NOW TAKES DxccSource
	' ---------------------------------------------------------
	Private Async Function ResolveRuleExpressionAsync(src As DxccSource) As Task(Of Geometry)

		Dim parts = Tokenize(src.rule)

		If parts.Count = 0 Then
			Return Nothing
		End If

		Dim current As Geometry = Nothing

		For Each part In parts
			Dim geom = Await ResolveSingleSourceAsync(src.source, part.Token)
			If geom Is Nothing Then Continue For

			If current Is Nothing Then
				' First geometry always initializes the accumulator
				current = geom
			Else
				If part.Op = "+"c Then
					current = GeometryEngine.Union(current, geom)
				ElseIf part.Op = "-"c Then
					current = GeometryEngine.Difference(current, geom)
				End If
			End If
		Next

		Return current
	End Function

	Private Function Tokenize(expr As String) As List(Of (Op As Char, Token As String))
		Dim result As New List(Of (Char, String))
		Dim current As String = ""
		Dim currentOp As Char = "+"c   ' default for first token

		For Each ch In expr
			If ch = "+"c OrElse ch = "-"c Then
				If current.Trim().Length > 0 Then
					result.Add((currentOp, current.Trim()))
				End If
				current = ""
				currentOp = ch
			Else
				current &= ch
			End If
		Next

		If current.Trim().Length > 0 Then
			result.Add((currentOp, current.Trim()))
		End If

		Return result
	End Function

	' ---------------------------------------------------------
	'  SINGLE SOURCE RESOLVER (ASYNC)
	' ---------------------------------------------------------
	Private Async Function ResolveSingleSourceAsync(source As String,
													token As String) As Task(Of Geometry)

		Select Case source.ToUpperInvariant()

			Case "NE"
				Return Await NaturalEarth.Lookup(token)

			Case "OSM"
				Return Await BuildFromOSM(token)

			Case "FAKE"
				Return BuildFromFake(token)

			Case Else
				Throw New Exception($"Unsupported resolver type '{source}' for token '{token}'.")
		End Select
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

			' ---------------------------------------------------------
			' CASE A: Normal envelope (no AM crossing)
			' ---------------------------------------------------------
			' Normal bbox ONLY if it does NOT cross the AM
			If maxLon >= minLon AndAlso (maxLon - minLon) <= 180 Then
				Return New Envelope(minLon, minLat, maxLon, maxLat, SpatialReferences.Wgs84)
			End If

			' ---------------------------------------------------------
			' CASE B: AM-crossing bbox → build TWO rectangles
			' ---------------------------------------------------------
			Dim rawRings As New List(Of List(Of MapPoint))

			' WEST RECTANGLE: minLon → -180. make it not touch the east rectangle
			rawRings.Add(New List(Of MapPoint) From {
				New MapPoint(minLon, minLat),
				New MapPoint(NudgeLon(-180), minLat),
				New MapPoint(NudgeLon(-180), maxLat),
				New MapPoint(NudgeLon(minLon), maxLat),
				New MapPoint(NudgeLon(minLon), minLat)
		})

			' EAST RECTANGLE: 180 → maxLon (clockwise)
			rawRings.Add(New List(Of MapPoint) From {
				New MapPoint(nudgelon(180), minLat),
				New MapPoint(nudgelon(maxLon), minLat),
				New MapPoint(nudgelon(maxLon), maxLat),
				New MapPoint(nudgelon(180), maxLat),
				New MapPoint(nudgelon(180), minLat)
				})

			' ---------------------------------------------------------
			' Split each ring at AM
			' ---------------------------------------------------------
			Dim splitRings As New List(Of List(Of MapPoint))
			For Each rr In rawRings
				For Each sr In SplitRingAtAntimeridian(rr)
					splitRings.Add(sr)
				Next
			Next

			' ---------------------------------------------------------
			' Remove degenerate rings
			' ---------------------------------------------------------
			Dim cleaned As New List(Of List(Of MapPoint))
			For Each ring In splitRings

				If ring.Count < 4 Then Continue For

				Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
				For Each p In ring
					Dim key = $"{p.X:0.000000},{p.Y:0.000000}"
					seen.Add(key)
				Next
				If seen.Count < 3 Then Continue For

				cleaned.Add(ring)
			Next

			' ---------------------------------------------------------
			' Enforce clockwise orientation (ArcGIS Runtime requirement)
			' ---------------------------------------------------------
			Dim finalRings As New List(Of List(Of MapPoint))
			For Each r In cleaned
				finalRings.Add(EnsureClockwise(r))
			Next

			' ---------------------------------------------------------
			' Build polygon
			' ---------------------------------------------------------
			Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
			For Each r In finalRings
				pb.AddPart(r)
			Next

			Return pb.ToGeometry()
		End If

		' ---------------------------------------------------------
		' 2. POLY CASE (poly:"...")
		' ---------------------------------------------------------
		Dim rPoly As New Regex("^poly:""(.+)""$", RegexOptions.IgnoreCase)
		Dim matches = rPoly.Match(st)
		If matches.Success Then

			' Parse coords
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
			    Dim lon = NudgeLon(Double.Parse(coords(i)))
			    Dim lat = Double.Parse(coords(i + 1))
			    pts.Add(New MapPoint(lon, lat, SpatialReferences.Wgs84))
			Next

			' Close ring
			If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))

			' Split at AM
			Dim rings = SplitRingAtAntimeridian(pts)

			' Remove degenerate rings
			Dim cleaned As New List(Of List(Of MapPoint))
			For Each ring In rings

				If ring.Count < 4 Then Continue For

				Dim seen As New HashSet(Of String)(StringComparer.Ordinal)
				For Each p In ring
					Dim key = $"{p.X:0.000000},{p.Y:0.000000}"
					seen.Add(key)
				Next
				If seen.Count < 3 Then Continue For

				cleaned.Add(ring)
			Next

			' Build polygon
			Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
			For Each r In cleaned
				pb.AddPart(r)
			Next

			Return pb.ToGeometry()
		End If

		Throw New Exception($"Unrecognised bbox pattern {st}")
	End Function
    Private Function NudgeLon(lon As Double) As Double
        Const EPS As Double = 1E-5
        If lon <= -180 Then Return -180 + EPS
        If lon >= 180 Then Return 180 - EPS
        Return lon
    End Function

	Public Function SplitRingAtAntimeridian(ring As List(Of MapPoint)) As List(Of List(Of MapPoint))
		Dim result As New List(Of List(Of MapPoint))

		' ---------------------------------------------------------
		' 1. UNWRAP longitudes into continuous space
		'    e.g. 170, 175, 180, 185, 190, 200
		' ---------------------------------------------------------
		Dim unwrapped As New List(Of MapPoint)
		Dim prevLon As Double = ring(0).X
		Dim offset As Double = 0

		unwrapped.Add(New MapPoint(prevLon, ring(0).Y))

		For i = 1 To ring.Count - 1
			Dim lon = ring(i).X
			Dim lat = ring(i).Y

			' compute raw delta
			Dim delta = lon - prevLon

			' unwrap if jump > 180
			If delta > 180 Then
				offset -= 360
			ElseIf delta < -180 Then
				offset += 360
			End If

			Dim uwLon = lon + offset
			unwrapped.Add(New MapPoint(uwLon, lat))

			prevLon = lon
		Next

		' ---------------------------------------------------------
		' 2. Split at every crossing of ±180 in unwrapped space
		' ---------------------------------------------------------
		Dim part As New List(Of MapPoint)
		part.Add(unwrapped(0))

		For i = 1 To unwrapped.Count - 1
			Dim p1 = unwrapped(i - 1)
			Dim p2 = unwrapped(i)

			Dim lon1 = p1.X
			Dim lon2 = p2.X
			Dim lat1 = p1.Y
			Dim lat2 = p2.Y

			' Check if segment crosses ±180
			If (lon1 < 180 AndAlso lon2 > 180) OrElse (lon1 > 180 AndAlso lon2 < 180) Then
				' crossing at +180
				Dim t = (180 - lon1) / (lon2 - lon1)
				Dim y = lat1 + t * (lat2 - lat1)

				Dim c1 As New MapPoint(180, y)
				part.Add(c1)
				result.Add(New List(Of MapPoint)(part))

				part = New List(Of MapPoint)
				part.Add(New MapPoint(-180, y))
			ElseIf (lon1 < -180 AndAlso lon2 > -180) OrElse (lon1 > -180 AndAlso lon2 < -180) Then
				' crossing at -180
				Dim t = (-180 - lon1) / (lon2 - lon1)
				Dim y = lat1 + t * (lat2 - lat1)

				Dim c1 As New MapPoint(-180, y)
				part.Add(c1)
				result.Add(New List(Of MapPoint)(part))

				part = New List(Of MapPoint)
				part.Add(New MapPoint(180, y))
			End If

			part.Add(p2)
		Next

		result.Add(part)

		' ---------------------------------------------------------
		' 3. RENORMALIZE all parts back to [-180,180]
		' ---------------------------------------------------------
		Dim finalParts As New List(Of List(Of MapPoint))

		For Each r In result
			Dim nr As New List(Of MapPoint)
			For Each p In r
				nr.Add(NormalizeTo180(p))
			Next

			' close ring
			If Not (nr(0).X = nr.Last().X AndAlso nr(0).Y = nr.Last().Y) Then
				nr.Add(nr(0))
			End If

			' remove degenerate
			If nr.Count >= 4 Then finalParts.Add(nr)
		Next

		Return finalParts
	End Function


	Public Function NormalizeTo180(p As MapPoint) As MapPoint
        Dim lon = p.X
        Dim lat = p.Y

        ' Normalize into [-180, 180)
        lon = ((lon + 180) Mod 360 + 360) Mod 360 - 180

        ' SPECIAL FIX: preserve +180 exactly
        If Math.Abs(p.X - 180) < 0.0000001 Then lon = 180

        Return New MapPoint(lon, lat, p.SpatialReference)
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
