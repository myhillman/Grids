Imports System.Globalization
Imports System.IO
Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports Esri.ArcGISRuntime.Portal
Imports Esri.ArcGISRuntime.Tasks.NetworkAnalysis
Imports Ionic
Imports Microsoft.Data.Sqlite
Imports RTools_NTS.Util

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

	Private Function BuildDerived(expr As String, tol As Integer) As Geometry
		Throw New NotImplementedException("Derived rule not implemented: " & expr)
	End Function

	Private Function SimplifyGeometry(g As Geometry, tolMeters As Integer) As Geometry
		If g Is Nothing Then Return g
		If tolMeters <= 0 Then Return g

		' ArcGIS Runtime simplification
		Dim tolerance = MetersToDegrees(g, tolMeters)       ' tolerance must be degrees
		Return GeometryEngine.Generalize(g, tolerance, True)
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

		' STEP 1 — Resolve base geometry from rule (NE/OSM only)
		Dim baseGeom As Geometry = Await ResolveRuleExpressionAsync(src)
		If baseGeom Is Nothing Then
			Debug.WriteLine($"Failed to resolve base geometry for DXCC '{src.name}' using rule: {src.rule} with bbox: {src.bbox}")
			Return Nothing
		End If

		' STEP 2 — Apply bbox/poly if present
		If Not String.IsNullOrWhiteSpace(src.bbox) Then
			Dim clipGeom As Geometry = ParseBBoxOrPoly(src.bbox)
			If clipGeom Is Nothing Then
				Debug.WriteLine($"Failed to resolve bbox for DXCC '{src.name}' using bbox: {src.bbox}")
			End If
			If TypeOf clipGeom Is Envelope Then
				' FAST PATH for rectangular clipping
				baseGeom = GeometryEngine.Clip(baseGeom, CType(clipGeom, Envelope))
			Else
				' POLYGON PATH
				Dim result As Geometry = Nothing
				If src.name = "Fiji" Then
					Dim pg = TryCast(clipGeom, Polygon)
					If pg IsNot Nothing Then
						Dim partIndex As Integer = 0
						For Each part In pg.Parts
							partIndex += 1
							Dim pts = part.Points.ToList()
						Next
					End If
					For Each part In pg.Parts
						Dim pts = part.Points.ToList()
					Next
				End If
				Dim clipPoly = CType(clipGeom, Polygon)

				For Each part In clipPoly.Parts
					Dim pts = part.Points.ToList()
					Dim partPoly = New PolygonBuilder(pts, SpatialReferences.Wgs84).ToGeometry()

					Dim piece = GeometryEngine.Intersection(baseGeom, partPoly)

					If piece IsNot Nothing AndAlso Not piece.IsEmpty Then
						If result Is Nothing Then
							result = piece
						Else
							result = GeometryEngine.Union(result, piece)
						End If
					End If
				Next

				baseGeom = result
			End If
		End If
		If baseGeom Is Nothing Then
			Debug.WriteLine($"Resulting geometry is empty after applying bbox {src.bbox} for DXCC '{src.name}'")
		End If
		' STEP 3 — simplification using tolerance_m
		baseGeom = SimplifyGeometry(baseGeom, src.tolerance_m)

		Return baseGeom
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
		' Build antimeridian cut line once
		' ---------------------------------------------------------
		Dim antiBuilder As New PolylineBuilder(SpatialReferences.Wgs84)
		antiBuilder.AddPoint(180, -90)
		antiBuilder.AddPoint(180, 90)
		Dim anti As Polyline = antiBuilder.ToGeometry()

		' ---------------------------------------------------------
		' 1. BBOX CASE (4 numbers)
		' ---------------------------------------------------------
		Dim parts = st.Split(","c).Select(Function(x) x.Trim()).ToList()
		If parts.Count = 4 Then

			Dim minLon = Double.Parse(parts(0))
			Dim minLat = Double.Parse(parts(1))
			Dim maxLon = Double.Parse(parts(2))
			Dim maxLat = Double.Parse(parts(3))

			' CASE A: Normal envelope (fast path)
			If minLon <= maxLon Then
				Return New Envelope(minLon, minLat, maxLon, maxLat, SpatialReferences.Wgs84)
			End If

			' CASE B: Crosses antimeridian → build TWO rectangles explicitly
			Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)

			' West part: minLon → 180
			pb.AddPart(New List(Of MapPoint) From {
			New MapPoint(minLon, minLat),
			New MapPoint(180, minLat),
			New MapPoint(180, maxLat),
			New MapPoint(minLon, maxLat),
			New MapPoint(minLon, minLat)
		})

			' East part: -180 → maxLon
			pb.AddPart(New List(Of MapPoint) From {
			New MapPoint(-180, minLat),
			New MapPoint(maxLon, minLat),
			New MapPoint(maxLon, maxLat),
			New MapPoint(-180, maxLat),
			New MapPoint(-180, minLat)
		})

			Dim bboxPoly = pb.ToGeometry()

			' Ensure CCW orientation
			Dim list As New List(Of Geometry) From {bboxPoly}
			Return NormalizeCut(list)
		End If

		' ---------------------------------------------------------
		' 2. POLY CASE (poly:"...")
		' ---------------------------------------------------------
		Dim r As New Regex("^poly:""(.+)""$", RegexOptions.IgnoreCase)
		Dim matches = r.Match(st)
		If matches.Success Then

			' Parse coords
			Dim coords = matches.Groups(1).Value.Split(" "c).
					Select(Function(x) x.Trim()).
					Where(Function(x) x.Length > 0).
					ToList()

			If coords.Count Mod 2 <> 0 Then
				Throw New Exception($"Unrecognised bbox pattern {st}")
			End If

			' Build raw ring (NO normalization yet)
			Dim pts As New List(Of MapPoint)
			For i = 0 To coords.Count - 1 Step 2
				pts.Add(New MapPoint(Double.Parse(coords(i)),
								 Double.Parse(coords(i + 1)),
								 SpatialReferences.Wgs84))
			Next

			If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))

			Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
			pb.AddPart(pts)
			Dim rawPoly = pb.ToGeometry()

			' CUT FIRST (before normalization)
			Dim cut = GeometryEngine.Cut(rawPoly, anti)

			Dim out As New PolygonBuilder(SpatialReferences.Wgs84)

			If cut Is Nothing OrElse cut.Count = 0 Then
				' No cut → normalize whole ring
				Dim norm = NormalizeRingTo180(pts)
				out.AddPart(norm)
			Else
				' Normalize each cut part separately
				For Each g In cut
					Dim pg = TryCast(g, Polygon)
					If pg IsNot Nothing Then
						For Each part In pg.Parts
							Dim ring = part.Points.ToList()
							Dim norm = NormalizeRingTo180(ring)
							out.AddPart(norm)
						Next
					End If
				Next
			End If

			Return out.ToGeometry()
		End If

		Throw New Exception($"Unrecognised bbox pattern {st}")
	End Function


	Function NormalizeCut(cut As IReadOnlyList(Of Geometry)) As Geometry
		Dim out As New PolygonBuilder(SpatialReferences.Wgs84)
		For Each g In cut
			Dim pg = TryCast(g, Polygon)
			If pg IsNot Nothing Then
				For Each part In pg.Parts

					' Convert to mutable list
					Dim pts = part.Points.ToList

					' Ensure closed
					If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))

					' Compute signed area
					Dim area = PolygonArea(pts)

					' Reverse if clockwise (negative area)
					If area < 0 Then
						pts.Reverse()
					End If

					out.AddPart(pts)
				Next
			End If
		Next

		Return out.ToGeometry()
	End Function

	' ring winding is not guaranteed in source data, so we need to classify rings as outer vs inner, and then orient them correctly (outer CCW, inner CW) to ensure consistent geometry across all sources. This is done by building a containment hierarchy of rings, and classifying by depth parity (outer = even depth, inner = odd depth).
	Public Class ClassifiedRing
		Public Property Part As ReadOnlyPart
		Public Property Depth As Integer
		Public Property IsOuter As Boolean
		Public Property Parent As ClassifiedRing
		Public Property Children As New List(Of ClassifiedRing)
	End Class
	Public Function NormalizePolygon(poly As Polygon) As Polygon
		If poly Is Nothing OrElse poly.IsEmpty Then Return Nothing
		' 1. Classify rings (outer vs inner)
		Dim rings = ClassifyPolygonRings(poly)   ' from earlier code

		' 2. Build a new polygon with corrected orientation
		Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)

		For Each r In rings

			Dim pts = r.Part.Points.ToList()

			Dim area = PolygonArea(pts)

			If r.IsOuter Then
				' outer rings must be CCW → positive area
				If area < 0 Then pts.Reverse()
			Else
				' inner rings must be CW → negative area
				If area > 0 Then pts.Reverse()
			End If

			pb.AddPart(pts)
		Next

		Return pb.ToGeometry()
	End Function

	Public Function ClassifyPolygonRings(poly As Polygon) As List(Of ClassifiedRing)

		Dim rings As New List(Of ClassifiedRing)

		' 1. Collect rings
		For Each part In poly.Parts
			rings.Add(New ClassifiedRing With {.Part = part})
		Next

		' 2. Determine containment relationships
		For i = 0 To rings.Count - 1
			Dim child = rings(i)
			Dim bestParent As ClassifiedRing = Nothing

			For j = 0 To rings.Count - 1
				If i = j Then Continue For

				Dim candidate = rings(j)

				If RingContainsRing(candidate.Part, child.Part) Then
					If bestParent Is Nothing Then
						bestParent = candidate
					Else
						' choose the smallest containing ring
						If Math.Abs(PolygonArea(candidate.Part.Points)) < Math.Abs(PolygonArea(bestParent.Part.Points)) Then
							bestParent = candidate
						End If
					End If
				End If
			Next

			child.Parent = bestParent
			If bestParent IsNot Nothing Then
				bestParent.Children.Add(child)
			End If
		Next

		' 3. Assign depths
		For Each r In rings
			If r.Parent Is Nothing Then
				AssignDepth(r, 0)
			End If
		Next

		' 4. Classify outer/inner by depth parity
		For Each r In rings
			r.IsOuter = (r.Depth Mod 2 = 0)
		Next

		Return rings
	End Function
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

	Private Function RingContainsRing(outerPart As ReadOnlyPart, innerPart As ReadOnlyPart) As Boolean

		' Normalize both rings into 0..360 domain
		Dim outerPts = outerPart.Points.Select(Function(p) NormalizeLon(p)).ToList()
		Dim innerPts = innerPart.Points.Select(Function(p) NormalizeLon(p)).ToList()

		Dim testPoint = innerPts(0)

		Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
		For Each p In outerPts
			pb.AddPoint(p)
		Next

		' Close ring if needed
		If Not outerPts(0).IsEqual(outerPts(outerPts.Count - 1)) Then
			pb.AddPoint(outerPts(0))
		End If

		Dim outerPoly = pb.ToGeometry()

		Return GeometryEngine.Contains(outerPoly, testPoint)
	End Function

End Module
