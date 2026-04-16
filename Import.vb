Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text.Json
Imports System.Text.Json.Nodes
Imports System.Text.RegularExpressions
Imports System.Web
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry

Imports HtmlAgilityPack
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualBasic.FileIO
Imports PolylinerNet

Module Import
	' Module contains all import functions
	Sub ImportISO3166()
		' Import ISO3166-1 country codes
		Dim sql As SqliteCommand, updated As Integer = 0
		Using connect As New SqliteConnection(DXCC_DATA),
			  ISO3166 As New TextFieldParser($"{Application.StartupPath}\ISO3166.csv")
			connect.Open()
			sql = connect.CreateCommand
			With ISO3166
				.TextFieldType = FileIO.FieldType.Delimited
				.SetDelimiters(",")
			End With
			Dim lines As Integer = 1
			Try
				sql.CommandText = $"DELETE FROM `ISO31661`"    ' remove existing data
				sql.ExecuteNonQuery()

				While Not ISO3166.EndOfData
					If lines > 1 Then
						Dim currentRow = ISO3166.ReadFields
						sql.CommandText = $"INSERT INTO `ISO31661`(`Entity`,`Code`) VALUES('{SQLescape(currentRow(0))}','{currentRow(1)}')"
						updated += sql.ExecuteNonQuery
					End If
					lines += 1
				End While
			Catch ex As MalformedLineException
				MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
			End Try
			AppendText(Form1.TextBox1, $"{updated} ISO6133 codes imported{vbCrLf}")
		End Using
	End Sub
	Sub ImportEUASBorder()
		' Import the border kml file for EU/AS. It has 464 linestrings and 13 digit precision
		' We can compress it to one linestring and 4 dec places. Reduces size to 10%
		Dim plb As New PolylineBuilder(SpatialReferences.Wgs84)

		Dim doc = XDocument.Load($"{Application.StartupPath}\border.kml")    ' read the XML
		Dim ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
		Dim linestrings = doc.Descendants(ns + "LineString")       ' find all the linestrings
		AppendText(Form1.TextBox1, $"{linestrings.Count} linestrings loaded{vbCrLf}")
		For Each coordinates In linestrings.Descendants(ns + "coordinates")
			Dim coords As List(Of String), linestr As New Part(SpatialReferences.Wgs84)
			Dim value = coordinates.Value
			value = Regex.Replace(value, "[^0-9\-\., ]", "")     ' remove noise characters
			value = Trim(value)
			coords = Split(value, " ").ToList   ' bust into coordinate pairs
			linestr.Clear()
			For Each coord In coords
				Dim points = Split(coord, ",")
				linestr.AddPoint(points(0), points(1))
			Next
			plb.AddPart(linestr)
		Next
		' make polygon out of parts
		Dim poly As New Polygon(plb.Parts)
		Dim polyGeneralized = Generalize(poly)   ' reduce it in size
		' Now produce KML for border
		Using kml As New StreamWriter($"{Application.StartupPath}\KML\BorderEUAS.kml", False)
			kml.WriteLine(KMLheader)
			kml.WriteLine("<Placemark><styleUrl>#red</styleUrl>")
			kml.WriteLine("<name>Boundary between European and Asiatic Russia</name>")
			kml.WriteLine("<description>Generally accepted border between Europe and Asia.</description>")
			kml.WriteLine("<MultiGeometry>")
			For Each p In polyGeneralized.Parts
				kml.WriteLine("<LineString>")
				KMLcoordinates(kml, p.Points.ToList, 4)
				kml.WriteLine("</LineString>")
			Next
			kml.WriteLine("</MultiGeometry>")
			kml.WriteLine("</Placemark>")
			kml.WriteLine(KMLfooter)
		End Using
		AppendText(Form1.TextBox1, $"Done{vbCrLf}")
	End Sub
	Public Class zonedata
		Public Property Id As String
		Public Property Polyline As String
		Public Property Mask As String
		Public Property Color As String
		Public Property Lat As String
		Public Property Lon As String
		Public Property Description As String

		Public Sub New(data As List(Of String))
			Id = data(0)
			Polyline = data(1)
			Mask = data(2)
			Color = data(3)
			Lat = data(4)
			Lon = data(5)
			Description = data(6)
		End Sub
	End Class
	Async Sub ImportCQITUZones()

		Dim ZoneCheck As New Dictionary(Of String, String) From {
		{"CQ", "https://zone-check.eu/includes/zones.cq.js"},
		{"ITU", "https://zone-check.eu/includes/zones.itu.js"}
	}
		' See https://developers.google.com/maps/documentation/utilities/polylinealgorithm

		Using connect As New SqliteConnection(DXCC_DATA)
			Await connect.OpenAsync()

			Using httpClient As New HttpClient(),
			  zonedataWriter As New StreamWriter(Path.Combine(Application.StartupPath, "Zonedata.vb"), False)

				httpClient.Timeout = TimeSpan.FromMinutes(10)
				httpClient.DefaultRequestHeaders.UserAgent.ParseAdd("Mozilla/5.0")

				For Each zc In ZoneCheck

					AppendText(Form1.TextBox1, $"Getting {zc.Key} data from {zc.Value}{vbCrLf}")

					' ---------------------------------------------------------
					' 1. Download JS file
					' ---------------------------------------------------------
					Dim responseString As String = ""
					Try
						Using httpResult = Await httpClient.GetAsync(zc.Value, HttpCompletionOption.ResponseHeadersRead)
							httpResult.EnsureSuccessStatusCode()
							Using stream = Await httpResult.Content.ReadAsStreamAsync()
								Using reader As New StreamReader(stream)
									responseString = Await reader.ReadToEndAsync()
								End Using
							End Using
						End Using
					Catch ex As Exception
						MsgBox(ex.Message, vbCritical, "zone-check request error")
						Continue For
					End Try

					AppendText(Form1.TextBox1, $"{responseString.Length} bytes retrieved{vbCrLf}")

					' ---------------------------------------------------------
					' 2. Parse JS array 
					' ---------------------------------------------------------
					Dim zones As List(Of zonedata) = ParseZoneArray(zc.Key, responseString)
					AppendText(Form1.TextBox1, $"{zones.Count} {zc.Key} zones retrieved{vbCrLf}")

					' ---------------------------------------------------------
					' 3. Prepare SQL
					' ---------------------------------------------------------
					Using tx = connect.BeginTransaction(),
					  cmd As SqliteCommand = connect.CreateCommand()

						cmd.Transaction = tx

						' Delete existing rows
						cmd.CommandText = "DELETE FROM ZoneLines WHERE Type = @type"
						cmd.Parameters.AddWithValue("@type", zc.Key)
						Await cmd.ExecuteNonQueryAsync()
						cmd.Parameters.Clear()

						' Insert command
						cmd.CommandText = "INSERT INTO ZoneLines (Type, zone, geometry) VALUES (@type, @zone, @geom)"

						Dim polyliner As New Polyliner()

						' ---------------------------------------------------------
						' 4. Decode polylines and insert polygons
						' ---------------------------------------------------------
						For Each z In zones

							Dim coords = polyliner.Decode(z.Polyline)
							Dim linestr As New List(Of MapPoint)
							' nudge coordinates away from anti-meridian if needed to avoid problems with ArcGIS geometry processing. We can get away with a small nudge because the data is very coarse and we will be generalizing it anyway
							For Each c In coords
								Dim lon = c.Longitude
								Dim lat = c.Latitude
								linestr.Add(New MapPoint(lon, lat, SpatialReferences.Wgs84))
							Next

							' Close polygon if needed
							If Not CoIncident(linestr.First, linestr.Last) Then
								linestr.Add(linestr.First)
							End If

							Dim poly As New Polygon(linestr)
							poly = poly.Generalize(0.2, False)
							poly = poly.Simplify()

							If poly Is Nothing OrElse poly.IsEmpty Then Continue For

							' round the polygon to avoid 180
							Dim builder As New PolygonBuilder(SpatialReferences.Wgs84)
							Const EPS = 0.0001
							For Each part In poly.Parts
								Dim newPart As New List(Of MapPoint)

								For Each pt In part.Points
									Dim lon = Math.Round(pt.X, 4)
									Dim lat = Math.Round(pt.Y, 4)

									' Clamp to safe domain
									If lon > 180 - EPS Then lon = 180 - EPS
									If lon < -180 + EPS Then lon = -180 + EPS

									newPart.Add(New MapPoint(lon, lat, SpatialReferences.Wgs84))
								Next

								builder.AddPart(newPart)
							Next
							poly = builder.ToGeometry()

							cmd.Parameters.Clear()
							cmd.Parameters.AddWithValue("@type", zc.Key)
							cmd.Parameters.AddWithValue("@zone", CInt(z.Id))
							cmd.Parameters.AddWithValue("@geom", GeometryToGeoJson(poly))

							Await cmd.ExecuteNonQueryAsync()
						Next

						Await tx.CommitAsync()
					End Using
				Next
			End Using
		End Using

		AppendText(Form1.TextBox1, $"Done{vbCrLf}")
	End Sub

	' ParseZoneArray
	' ---------------
	' Extracts and parses the CQ or ITU zone array from the downloaded JavaScript file.
	' The JS files contain arrays like:
	'   var cqzones = [
	'       ['01','polyline','mask','#006600','65','-140','desc'],
	'       ...
	'   ];
	'
	' This parser:
	'   • Locates the correct JS variable (cqzones or ituzones)
	'   • Extracts ONLY the array text (from first "[" to matching "]]")
	'   • Walks the array with a small state machine
	'   • Handles single‑quoted strings safely (no JSON, no JS eval)
	'   • Builds zonedata objects from each inner array
	Private Function ParseZoneArray(target As String, js As String) As List(Of zonedata)
		Dim varName = $"{target.ToLower()}zones"
		Dim varIdx = js.IndexOf($"var {varName}", StringComparison.OrdinalIgnoreCase)
		If varIdx = -1 Then Throw New Exception($"Variable {varName} not found")

		Dim openIdx = js.IndexOf("["c, varIdx)
		If openIdx = -1 Then Throw New Exception($"Opening [ for {varName} not found")

		Dim depth As Integer = 0
		Dim inString As Boolean = False
		Dim escapeNext As Boolean = False
		Dim closeIdx As Integer = -1

		For i = openIdx To js.Length - 1
			Dim ch = js(i)

			If escapeNext Then
				escapeNext = False
				Continue For
			End If

			If ch = "\"c Then
				escapeNext = True
				Continue For
			End If

			If ch = "'"c Then
				inString = Not inString
				Continue For
			End If

			If Not inString Then
				If ch = "["c Then depth += 1
				If ch = "]"c Then
					depth -= 1
					If depth = 0 Then
						closeIdx = i
						Exit For
					End If
				End If
			End If
		Next

		If closeIdx = -1 Then Throw New Exception($"End of {varName} array not found")

		Dim jsarray = js.Substring(openIdx, closeIdx - openIdx + 1)

		' strip leading [[ and trailing ]];
		Dim inner = jsarray.Substring(2, jsarray.Length - 2 - 3)  ' remove "[[" and "]];"

		Dim entries = inner.Split(New String() {"],["}, StringSplitOptions.None)

		Dim zones As New List(Of zonedata)

		For Each entry In entries
			' entry looks like: '01','poly','mask','color',lat,lon,'desc'
			Dim fields = entry.Split(","c)

			If fields.Length < 7 Then Continue For

			Dim id = fields(0).Trim().Trim("'"c)
			Dim poly = fields(1).Trim().Trim("'"c)
			Dim mask = fields(2).Trim().Trim("'"c)
			Dim color = fields(3).Trim().Trim("'"c)
			Dim lat = fields(4).Trim()
			Dim lon = fields(5).Trim()
			Dim desc = String.Join(",", fields.Skip(6)).Trim().Trim("'"c)

			Dim data As New List(Of String) From {
			id, poly, mask, color, lat, lon, desc
		}

			zones.Add(New zonedata(data))
		Next

		Return zones
	End Function

	Async Function ImportTimeZones() As Task
		' Import a timezone database
		Dim myQueryFilter As New QueryParameters

		Dim Features = Await ShapefileFeatureTable.OpenAsync(
		"D:\GIS Data\Natural Earth\Timezones\ne_10m_time_zones.shp")

		With myQueryFilter
			.OutSpatialReference = SpatialReferences.Wgs84
			.ReturnGeometry = True
		End With

		Dim Timezones = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)

		With Form1.ProgressBar1
			.Minimum = 0
			.Value = 0
			.Maximum = Timezones.Count
		End With

		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()

			' Clear table
			Using cmd = connect.CreateCommand()
				cmd.CommandText = "DELETE FROM Timezones"
				cmd.ExecuteNonQuery()
			End Using

			' Prepare INSERT command once
			Using cmd = connect.CreateCommand()
				cmd.CommandText =
				"INSERT INTO Timezones (name, places, color, geometry) " &
				"VALUES (@name, @places, @color, @geometry)"

				Dim pName = cmd.CreateParameter()
				pName.ParameterName = "@name"
				cmd.Parameters.Add(pName)

				Dim pPlaces = cmd.CreateParameter()
				pPlaces.ParameterName = "@places"
				cmd.Parameters.Add(pPlaces)

				Dim pColor = cmd.CreateParameter()
				pColor.ParameterName = "@color"
				cmd.Parameters.Add(pColor)

				Dim pGeom = cmd.CreateParameter()
				pGeom.ParameterName = "@geometry"
				cmd.Parameters.Add(pGeom)

				' Sort by name
				Dim tzList = Timezones.Select(Function(f) f).OrderBy(Function(f) CSng(f.Attributes("name"))).ToList()

				For Each timezone In tzList
					Dim name = timezone.Attributes("name")
					Dim places = timezone.Attributes("places")
					Dim color = timezone.Attributes("map_color6")

					' --- GEOMETRY PIPELINE ---
					Dim poly As Polygon = timezone.Geometry
					'Debug.WriteLine($"""{name}"",""{places}"",{poly.Parts.Count}")
					poly = poly.Simplify()

					Dim polyGeneralized As Polygon = poly.Generalize(0.1, True)
					If polyGeneralized.IsEmpty Then polyGeneralized = poly
					polyGeneralized = FixRingOrientation(polyGeneralized)

					' --- PARAMETER ASSIGNMENT ---
					pName.Value = name
					pPlaces.Value = places
					pColor.Value = color
					pGeom.Value = GeometryToGeoJson(polyGeneralized)

					cmd.ExecuteNonQuery()

					Form1.ProgressBar1.Value += 1
				Next
			End Using
		End Using

		AppendText(Form1.TextBox1, $"{Timezones.Count} timezones imported{vbCrLf}")
	End Function


	Async Function ImportIARURegions() As Task
		' Import IARU regions. Data from Tim Makins (EI8IC)
		Dim myQueryFilter As New QueryParameters, sql As SqliteCommand, count As Integer = 0, GeneralizeDistance As Integer = 1000
		Dim lines As New PolylineBuilder(SpatialReferences.Wgs84), OriginalCount As Integer = 0, DensifyCount As Integer = 0, GeneralizeCount As Integer = 0

		Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\IARU\IARU Region Lines_50m_No_Edges_Simplified.shp")
		With myQueryFilter
			.OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
			.ReturnGeometry = True
		End With
		Dim Regions = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query
		AppendText(Form1.TextBox1, $"{Regions.Count} IARU boundary lines loaded{vbCrLf}")
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			sql = connect.CreateCommand
			sql.CommandText = "DELETE FROM IARU"       ' remove existing data
			sql.ExecuteNonQuery()

			For Each Reg In Regions
				Dim geom As Polyline = Reg.Geometry
				For Each line In geom.Parts
					geom = NormalizeCentralMeridian(geom)         ' normalize in case line crosses anti-meridian (one does)
					lines.AddPart(line)
				Next
			Next
			' Now add all lines to database
			Dim geometry = lines.ToGeometry
			For Each prt In geometry.Parts
				OriginalCount += prt.Points.Count
			Next
			geometry = geometry.Densify(5).Simplify
			For Each prt In geometry.Parts
				DensifyCount += prt.Points.Count
			Next
			'Dim GeneralizeAngle As Double = Form1.RadtoDeg(Math.Asin(GeneralizeDistance / Form1.EARTH_RADIUS))   ' angle for Generalize
			'geometry = geometry.Generalize(GeneralizeAngle, True)
			'For Each prt In geometry.Parts
			'    GeneralizeCount += prt.Points.Count
			'Next
			For Each line In geometry.Parts
				count += 1
				Dim geo = New Polyline(line.Points).ToJson
				sql.CommandText = $"INSERT INTO IARU (line,geometry) VALUES ({count},'{geo}')"
				sql.ExecuteNonQuery()
			Next
			AppendText(Form1.TextBox1, $"Linestrings {geometry.Parts.Count}, Original {OriginalCount}, Densify {DensifyCount}, Generalize {GeneralizeCount}{vbCrLf}")
		End Using
	End Function

	Async Function ImportAntarctica() As Task
		' Get the country boundary for Antarctica. No useful one in OSM
		Dim qp As New QueryParameters
		Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\World countries generalized\World_Countries_Generalized.shp")
		Using connect As New SqliteConnection(DXCC_DATA)
			Dim sql As SqliteCommand, sqlDR As SqliteDataReader
			connect.Open()
			sql = connect.CreateCommand
			' get the bounding polygon
			sql.CommandText = "SELECT * FROM DXCC WHERE Entity='Antarctica'"
			sqlDR = sql.ExecuteReader
			sqlDR.Read()
			Dim Antarctica As Polygon = ParseBox(sqlDR("bbox"))
			sqlDR.Close()
			With qp
				.WhereClause = "COUNTRY='Antarctica'"            ' get all features
				.Geometry = Antarctica
				.ReturnGeometry = True
				.OutSpatialReference = SpatialReferences.Wgs84
				.SpatialRelationship = SpatialRelationship.Intersects
			End With
			Dim fqr = Await Features.QueryFeaturesAsync(qp)
			Dim plb As New Esri.ArcGISRuntime.Geometry.PolygonBuilder(SpatialReferences.Wgs84)
			For Each f In fqr
				Dim geom As Polygon = f.Geometry
				For Each p In geom.Parts
					plb.AddPart(p)
				Next
			Next
			Dim poly As New Polygon(plb.Parts)
			'poly = poly.Intersection(Antarctica)        ' remove unwanted bits
			Dim polyGeneralized = GeneralizeByPart(poly)   ' reduce it in size
			Dim geometry = GeometryToGeoJson(polyGeneralized)           ' convert to geojson
			' Save in database
			sql.CommandText = "UPDATE DXCC SET geometry=@geom WHERE Entity=@entity"
			sql.Parameters.Clear()
			sql.Parameters.AddWithValue("@geom", geometry)
			sql.Parameters.AddWithValue("@entity", "Antarctica")
			sql.ExecuteNonQuery()

		End Using
		AppendText(Form1.TextBox1, $"Done{vbCrLf}")
	End Function

	Async Function ImportAntarcticBases() As Task
		Dim responseString As String = ""
		Using httpClient As New System.Net.Http.HttpClient()
			httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
			Dim url = "https://www.coolantarctica.com/Community/antarctic_bases.php"
			Try
				Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
				httpResult.EnsureSuccessStatusCode()
				responseString = Await httpResult.Content.ReadAsStringAsync()
			Catch ex As HttpRequestException
				MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
			End Try
		End Using

		Dim sql As SqliteCommand
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			sql = connect.CreateCommand
			' Delete existing data
			sql.CommandText = "DELETE FROM Antarctic"
			sql.ExecuteNonQuery()
			' Dim allowedChars As String = "[^0-9a-zA-Z\-\:\./°\'\""]"
			Dim replacements As New Dictionary(Of String, String) From {
			{"&deg;", "°"},
			{"&#39;", "'"},
			{"&quot;", """"}
			}
			Dim allowedChars As String = "(\t|\n|&nbsp;)"
			Dim htmldoc As New HtmlDocument()
			htmldoc.LoadHtml(responseString)
			Dim table = htmldoc.DocumentNode.SelectSingleNode("//table")
			Dim rows = table.SelectNodes("tr")
			With Form1.ProgressBar1
				.Minimum = 0
				.Value = 0
				.Maximum = rows.Count - 1
			End With
			For row = 0 To rows.Count - 2
				Form1.ProgressBar1.Value += 1
				Dim bgcolor = rows(row).SelectNodes("td").Item(0).Attributes("bgcolor")
				Dim open As String
				If bgcolor Is Nothing Then open = "Summer only" Else open = "Year round"
				Dim Name = HttpUtility.HtmlDecode(rows(row).SelectNodes("td").Item(0).InnerText)
				Name = Regex.Replace(Name, allowedChars, "")
				Dim nation = HttpUtility.HtmlDecode(rows(row).SelectNodes("td").Item(1).InnerText)
				nation = Regex.Replace(nation, allowedChars, "")
				Dim coordinates = HttpUtility.HtmlDecode(rows(row).SelectNodes("td").Item(2).InnerText)
				coordinates = Regex.Replace(coordinates, allowedChars, "")
				For Each rep In replacements
					coordinates = coordinates.Replace(rep.Key, rep.Value)
				Next
				Dim coord As MapPoint = CoordinateFormatter.FromLatitudeLongitude(coordinates, SpatialReferences.Wgs84)  ' DD MM SS to decimal
				Dim situation = HttpUtility.HtmlDecode(rows(row).SelectNodes("td").Item(3).InnerText)
				situation = Regex.Replace(situation, allowedChars, "")
				Dim altitude = HttpUtility.HtmlDecode(rows(row).SelectNodes("td").Item(4).InnerText)
				altitude = Regex.Replace(altitude, allowedChars, "")
				sql.CommandText = "
                            INSERT INTO Antarctic (open, name, nation, coordinates, situation, altitude)
                            VALUES (@open, @name, @nation, @coords, @situation, @altitude)
                        "

				sql.Parameters.Clear()
				sql.Parameters.AddWithValue("@open", open)
				sql.Parameters.AddWithValue("@name", Name)
				sql.Parameters.AddWithValue("@nation", nation)
				sql.Parameters.AddWithValue("@coords", coord.ToJson)
				sql.Parameters.AddWithValue("@situation", situation)
				sql.Parameters.AddWithValue("@altitude", altitude)

				sql.ExecuteNonQuery()

			Next
		End Using
	End Function

	Async Function ImportIOTAGroups() As Task
		Dim responseString As String = "", sql As SqliteCommand
		' Get JSON data for all IOTA groups
		'Dim url = $"https://www.iota-world.org/rest/get/iota/groups?api_key={IOTA_API_KEY}"
		Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=groups.json"
		Using httpClient As New System.Net.Http.HttpClient()
			httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
			Try
				Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
				httpResult.EnsureSuccessStatusCode()
				responseString = Await httpResult.Content.ReadAsStringAsync()
			Catch ex As HttpRequestException
				MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
			End Try
		End Using

		' Extract groups data
		Dim response = JsonNode.Parse(responseString).AsArray
		With Form1.ProgressBar1
			.Minimum = 0
			.Value = 0
			.Maximum = response.Count
		End With
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			sql = connect.CreateCommand
			sql.CommandText = "BEGIN TRANSACTION"
			sql.ExecuteNonQuery()                             ' delete existing data
			sql.CommandText = "DELETE FROM IOTA_Groups"
			sql.ExecuteNonQuery()                             ' delete existing data
			' use prepared statement for speed
			sql.CommandText = $"INSERT INTO IOTA_Groups 
        (`refno`,`name`,`dxcc_num`,latitude_max,latitude_min,longitude_max,longitude_min,grp_region,whitelist,comment) VALUES (@refno,@name,@dxcc_num,@latitude_max,@latitude_min,@longitude_max,@longitude_min,@grp_region,@whitelist,@comment)
"
			sql.Prepare()
			For Each group In response
				' process each group
				Dim refno = SafeStr(group.Item("refno"))
				Form1.ProgressBar1.Value += 1
				With sql.Parameters
					.Clear()
					' Sometimes they have things backwards so fix it if necessary
					Dim LatMin As Double = SafeDbl(group.Item("latitude_min"))
					Dim LatMax As Double = SafeDbl(group.Item("latitude_max"))
					Dim LonMin As Double = SafeDbl(group.Item("longitude_min"))
					Dim LonMax As Double = SafeDbl(group.Item("longitude_max"))
					If LatMin > LatMax Then
						Dim t = LatMin
						LatMin = LatMax
						LatMax = t
						AppendText(Form1.TextBox1, $"Warning: LatMin > LatMax for {group.Item("name")}. Values swapped.{vbCrLf}")
					End If
					If LonMin > 180 Then LonMin -= 360 ' normalize to -180 to 180
					If LonMax > 180 Then LonMax -= 360 ' normalize to -180 to 180
					If LonMin <= LonMax Then
						' normal case, nothing to do
					Else
						Dim gap = Math.Abs(LonMin - LonMax)
						If gap > 180 Then
							' AM crossing
						Else
							' probably they have wrapped around the wrong way. Swap them
							Dim t = LonMin
							LonMin = LonMax
							LonMax = t
							AppendText(Form1.TextBox1, $"Warning: LonMin > LonMax for {group.Item("name")}. Values swapped.{vbCrLf}")
						End If
					End If
					.AddWithValue("@refno", refno)
					.AddWithValue("@dxcc_num", group.Item("dxcc_num").ToString)
					.AddWithValue("@name", group.Item("name").ToString)
					.AddWithValue("@latitude_min", LatMin)
					.AddWithValue("@latitude_max", LatMax)
					.AddWithValue("@longitude_min", LonMin)
					.AddWithValue("@longitude_max", LonMax)
					.AddWithValue("@grp_region", group.Item("grp_region").ToString)
					.AddWithValue("@whitelist", group.Item("whitelist").ToString)
					.AddWithValue("@comment", group.Item("comment").ToString)
				End With
				sql.ExecuteNonQuery()
			Next
			sql.CommandText = "COMMIT"
			sql.ExecuteNonQuery()
			Form1.TextBox1.AppendText($"{response.Count} groups loaded{vbCrLf}")
		End Using
	End Function
	Async Function ImportIOTAIslands() As Task
		Dim responseString As String = "", sql As SqliteCommand
		' Get JSON data for all IOTA groups
		'Dim url = $"https://www.iota-world.org/rest/get/iota/islands?api_key={IOTA_API_KEY}"
		Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=islands.json"
		Using httpClient As New System.Net.Http.HttpClient()
			httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
			Try
				Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
				httpResult.EnsureSuccessStatusCode()
				responseString = Await httpResult.Content.ReadAsStringAsync()
			Catch ex As HttpRequestException
				MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
			End Try
		End Using

		' Extract groups data
		Dim response = JsonNode.Parse(responseString).AsArray

		With Form1.ProgressBar1
			.Minimum = 0
			.Value = 0
			.Maximum = response.Count
		End With
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			sql = connect.CreateCommand
			sql.CommandText = "BEGIN TRANSACTION"
			sql.ExecuteNonQuery()                             ' delete existing data
			sql.CommandText = "DELETE FROM IOTA_Islands"
			sql.ExecuteNonQuery()                             ' delete existing data
			' use prepared statement for speed
			sql.CommandText = $"INSERT INTO IOTA_Islands (`refno`,`name`,`comment`) VALUES (@refno,@name,@comment)"
			sql.Prepare()
			For Each group In response
				' process each group
				Form1.ProgressBar1.Value += 1
				With sql.Parameters
					.Clear()
					.AddWithValue("@refno", group.Item("refno").ToString)
					.AddWithValue("@name", group.Item("name").ToString)
					.AddWithValue("@comment", group.Item("comment").ToString)
				End With
				sql.ExecuteNonQuery()
			Next
			sql.CommandText = "COMMIT"
			sql.ExecuteNonQuery()
			Form1.TextBox1.AppendText($"{response.Count} islands loaded{vbCrLf}")
		End Using
	End Function
	Async Function ImportIOTADXCCMatchesOneIOTA() As Task
		Dim responseString As String = "", sql As SqliteCommand
		' Get JSON data for all IOTA groups
		'Dim url = $"https://www.iota-world.org/rest/get/iota/dxccmatchesoneiota?api_key={IOTA_API_KEY}"
		Dim url = $"https://www.iota-world.org/islands-on-the-air/downloads/download-file.html?path=dxcc_matches_one_iota.json"
		Using httpClient As New System.Net.Http.HttpClient()
			httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
			Try
				Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
				httpResult.EnsureSuccessStatusCode()
				responseString = Await httpResult.Content.ReadAsStringAsync()
			Catch ex As HttpRequestException
				MsgBox($"{ex.Message}{vbCrLf}url={url}", vbCritical + vbOKOnly, "Retrieve error")
			End Try
		End Using

		' Extract groups data
		Dim response = JsonNode.Parse(responseString).AsArray
		With Form1.ProgressBar1
			.Minimum = 0
			.Value = 0
			.Maximum = response.Count
		End With
		Using connect As New SqliteConnection(DXCC_DATA)
			connect.Open()
			sql = connect.CreateCommand
			sql.CommandText = "BEGIN TRANSACTION"
			sql.ExecuteNonQuery()                             ' delete existing data
			sql.CommandText = "DELETE FROM IOTA_DXCC_IOTA"
			sql.ExecuteNonQuery()                             ' delete existing data
			' use prepared statement for speed
			sql.CommandText = $"INSERT INTO IOTA_DXCC_IOTA (`refno`,`dxcc_num`) VALUES (@refno,@dxcc_num)"
			sql.Prepare()
			For Each group In response
				' process each group
				Form1.ProgressBar1.Value += 1
				With sql.Parameters
					.Clear()
					.AddWithValue("@refno", group.Item("refno").ToString)
					.AddWithValue("@dxcc_num", group.Item("dxcc_num").ToString)
				End With
				sql.ExecuteNonQuery()
			Next
			sql.CommandText = "COMMIT"
			sql.ExecuteNonQuery()
			Form1.TextBox1.AppendText($"{response.Count} DXCC matches loaded{vbCrLf}")
		End Using
	End Function

	Sub ImportPolyFromKML()
		' Import an OSM poly specification from a polygon in a KML file
		Dim result As String
		Dim polyPoints As New List(Of String)
		With Form1.OpenFileDialog1
			.Filter = "KML files (*.kml)|*.kml|All files(*.*)|*.*"
			.CheckFileExists = True
			.CheckPathExists = True
			.OkRequiresInteraction = True
			.AddExtension = True
			.DefaultExt = "kml"
			.AddToRecent = True
			.FileName = ""
			.Title = "Select a .kml file containing the desired polygon"
			If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
				Dim sr = .FileName
				' Open KML file as XML
				Dim doc = XDocument.Load(sr)    ' read the XML
				Dim ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
				Dim polygon = doc.Descendants(ns + "coordinates").Value
				polygon = Regex.Replace(polygon.Trim, "[^0-9\-\., ]", "")        ' find polygon list of coordinates
				Dim s = Split(polygon, " ")
				For Each coord In s
					Dim c = Split(coord, ",")
					Dim X = CDbl(c(1)) Mod 90
					Dim Y = NormalizeLongitude(c(0))       ' normalize longitude
					polyPoints.Add($"{X:f1} {Y:f1}")
				Next
				result = $"poly:""{Strings.Join(polyPoints.ToArray, " ")}"""
				Clipboard.SetText(result)           ' copy to clipboard
				MsgBox(result, vbInformation + vbOK, "Result in clipboard")        ' add closing double quote
			End If
		End With
	End Sub
	Public Async Function LandSquareList() As Task
		' Import a list of squares that are land
		Const LandDataURL = "https://osmdata.openstreetmap.de/download/land-polygons-split-4326.zip"    ' remote source of land data
		Const LandDataFile = "D:\GIS Data\Land Polygons\split\land_polygons.shp"                        ' local copy of land data
		Const LandDataZip = "D:\GIS Data\Land Polygons\split\land-polygons-split-4326.zip"              ' local copy of land data
		Const readChunkSize = 1024 * 1024            ' block size of bytes read
		Dim myQueryFilter As New QueryParameters, count As Integer = 0
		Dim timer As New Stopwatch

		With Form1.ProgressBar1
			.Minimum = 0
			.Value = 0
			.Maximum = 100
		End With
		' Check the date on the latest download file
		Dim response = Await Http.GetAsync(LandDataURL, HttpCompletionOption.ResponseHeadersRead)      ' request for header only
		If Not response.IsSuccessStatusCode Then
			MsgBox($"Error: {response.StatusCode}", vbCritical + vbOKOnly, "Error")
			Return
		End If

		' Check dates
		Dim LandDataURLDate = response.Content.Headers.LastModified.Value.UtcDateTime
		Dim LandDataFileDate = File.GetLastWriteTimeUtc(LandDataZip)
		Dim DateStatus As String
		If LandDataURLDate > LandDataFileDate Then DateStatus = "Land data is Out of Date" Else DateStatus = "Land data is Current"
		If MsgBox($"The OSM land data is dated {LandDataURLDate.ToUniversalTime:yyyy-MM-dd hh:mm}{vbCrLf}The local copy is dated {LandDataFileDate.ToUniversalTime:yyyy-MM-dd hh:mm}{vbCrLf}{vbCrLf}Do you wish to update the data ?", vbInformation + vbYesNo, DateStatus) = vbYes Then
			Dim totalBytes = response.Content.Headers.ContentLength       ' get count of total bytes
			timer.Start()
			AppendText(Form1.TextBox1, $"Fetching {totalBytes:n0} bytes of land data from OSM ")
			Dim totalBytesRead As Long = 0      ' total bytes read todate
			' Download the file from OSM with progress indicator
			Using contentStream = Await response.Content.ReadAsStreamAsync,
				filestream = New FileStream(LandDataZip, FileMode.Create, FileAccess.Write, FileShare.None, readChunkSize, True)
				Dim buffer(readChunkSize) As Byte       ' byte buffer
				Dim bytesRead As Integer                ' bytes read in block
				Do
					bytesRead = Await contentStream.ReadAsync(buffer, 0, buffer.Length)     ' read a block
					If bytesRead > 0 Then       ' block is not empty
						Await filestream.WriteAsync(buffer, 0, bytesRead)   ' write the block to the file
						totalBytesRead += bytesRead         ' count total bytes read
						Dim progressPercentage As Integer = totalBytesRead / totalBytes * 100   ' calculate percentage progress
						Form1.ProgressBar1.Value = progressPercentage         ' display progress
					End If
				Loop Until bytesRead = 0    ' stop when no bytes read
				filestream.Close()
			End Using
			' Now unzip downloaded file
			Using archive = ZipFile.OpenRead(LandDataZip)
				Dim targetDirectory = Path.GetDirectoryName(LandDataZip)
				For Each entry In archive.Entries
					entry.ExtractToFile($"{targetDirectory}\{entry.Name}", True)
				Next
			End Using
			timer.Stop()
			AppendText(Form1.TextBox1, $"[{timer.ElapsedMilliseconds / 1000:f1}s]{vbCrLf}")
		Else
			Return
		End If

		' Convert the OSM data into a list of grid squares that contain land
		timer.Restart()
		Dim Features = Await ShapefileFeatureTable.OpenAsync(LandDataFile)
		With myQueryFilter
			.OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
			.ReturnGeometry = False
		End With
		Dim land = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' return all geometry
		Dim featureCount = land.Count
		AppendText(Form1.TextBox1, $"Loading {featureCount} squares into database.")
		Using connect As New SqliteConnection(DXCC_DATA)
			Dim sql As SqliteCommand, sqlDR As SqliteDataReader
			connect.Open()
			sql = connect.CreateCommand
			sql.CommandText = "SELECT COUNT(*) as Count FROM LAND"
			sqlDR = sql.ExecuteReader()
			sqlDR.Read()
			Dim Before As Integer = sqlDR("Count")
			sqlDR.Close()
			sql.CommandText = "BEGIN TRANSACTION"
			sql.ExecuteNonQuery()
			sql.CommandText = "DELETE FROM LAND"
			sql.ExecuteNonQuery()
			For Each feature In land
				Dim pnt = New MapPoint(CDbl(feature.Attributes("x")), CDbl(feature.Attributes("y")), SpatialReferences.Wgs84)
				sql.CommandText = $"INSERT OR REPLACE INTO LAND (gridsquare) VALUES ('{GridSquare(pnt)}')"
				sql.ExecuteNonQuery()
				UpdateProgressBar(Form1.ProgressBar1, count / featureCount * 100)
				count += 1
			Next
			sql.CommandText = "COMMIT"
			sql.ExecuteNonQuery()
			sql.CommandText = "SELECT COUNT(*) as Count FROM LAND"
			sqlDR = sql.ExecuteReader()
			sqlDR.Read()
			Dim After As Integer = sqlDR("Count")
			sqlDR.Close()
			AppendText(Form1.TextBox1, $" Grid squares before={Before}, after={After} [{timer.ElapsedMilliseconds / 1000:f1}s]{vbCrLf}")
		End Using
	End Function
	Function NormalizeLongitude(longitude As Double) As Double
		' Normalize a longitude to between -180 and +180
		Return (longitude Mod 360 + 540) Mod 360 - 180        ' normalize longitude
	End Function

End Module
