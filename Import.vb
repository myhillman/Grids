Imports System.Collections.Immutable
Imports System.IO
Imports System.Net.Http
Imports System.Text.Json.Nodes
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Xml
Imports System.Xml.XPath
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Grids.Form1
Imports HtmlAgilityPack
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualBasic.FileIO
Imports PolylinerNet
Module Import
    ' Module contains all import functions
    Sub ImportISO3166()
        ' Import ISO3166-1 country codes
        Dim sql As SqliteCommand, updated As Integer = 0
        Using connect As New SqliteConnection(Form1.DXCC_DATA),
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
                        sql.CommandText = $"INSERT INTO `ISO31661`(`Entity`,`Code`) VALUES('{Form1.SQLescape(currentRow(0))}','{currentRow(1)}')"
                        updated += sql.ExecuteNonQuery
                    End If
                    lines += 1
                End While
            Catch ex As MalformedLineException
                MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
            End Try
            Form1.AppendText(Form1.TextBox1, $"{updated} ISO6133 codes imported{vbCrLf}")
        End Using
    End Sub
    Sub ImportEUASBorder()
        ' Import the border kml file for EU/AS. It has 464 linestrings and 13 digit precision
        ' We can compress it to one linestring and 4 dec places. Reduces size to 10%
        Dim plb As New PolylineBuilder(SpatialReferences.Wgs84)

        Dim doc = XDocument.Load($"{Application.StartupPath}\border.kml")    ' read the XML
        Dim ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
        Dim linestrings = doc.Descendants(ns + "LineString")       ' find all the linestrings
        Form1.AppendText(Form1.TextBox1, $"{linestrings.Count} linestrings loaded{vbCrLf}")
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
        Dim polyGeneralized = Form1.Generalize(poly)   ' reduce it in size
        ' Now produce KML for border
        Using kml As New StreamWriter($"{Application.StartupPath}\KML\BorderEUAS.kml", False)
            kml.WriteLine(Form1.KMLheader)
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
            kml.WriteLine(Form1.KMLfooter)
        End Using
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
    Class zonedata
        Sub New(data As List(Of String))
            fields = data
        End Sub

        ' fields of zone data
        ' 0 = zone number
        ' 1 = coordinates in Google Map format
        ' 2 = levels ?
        ' 3 = fill color
        ' 4 = label lat
        ' 5 = label lng
        ' 6 = text ?
        Property fields As List(Of String)
    End Class
    Enum state_enum
        idle
        arrayStart      ' start of array ([)
        recordStart     ' start of record ([)
        inField          ' start of field (')
        escape        ' in field character escape (\)
        fieldEnd        ' end of field (')
        recordEnd       ' end of record (])
        fieldseparator  ' field separator (,)
        recordseparator ' record separator (,)
        arrayEnd        ' end of array (])
    End Enum
    Async Sub ImportCQITUZones()
        ' Import CQ & ITU zone data from https://zone-check.eu/
        Dim responseString As String, zones As New List(Of zonedata), state As state_enum, field As String
        Dim ZoneCheck As New Dictionary(Of String, String) From {{"CQ", "https://zone-check.eu/includes/zones.cq.js"}, {"ITU", "https://zone-check.eu/includes/zones.itu.js"}}
        Dim sql As SqliteCommand

        ' Scan a javascript file and extract zone definitions therein. Data is stored as a javascript array
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand

            Using httpClient As New System.Net.Http.HttpClient(),
                 Zonedata As New StreamWriter($"{Application.StartupPath}\Zonedata.vb", False)
                httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
                For Each zc In ZoneCheck
                    Form1.AppendText(Form1.TextBox1, $"Getting {zc.Key} data from {zc.Value}{vbCrLf}")
                    Dim url As New Uri(zc.Value)
                    Try
                        Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.GetAsync(url)
                        httpResult.EnsureSuccessStatusCode()
                        responseString = Await httpResult.Content.ReadAsStringAsync()
                        Form1.AppendText(Form1.TextBox1, $"{responseString.Length} bytes retrieved{vbCrLf}")
                    Catch ex As HttpRequestException
                        MsgBox($"{ex.Message}", vbCritical + vbOKOnly, "zone-check request error")
                    End Try
                    state = state_enum.idle
                    zones.Clear()
                    field = ""
                    Dim index As Integer = 0
                    Dim fields = New List(Of String)
                    ' parse array into records using FSM
                    ' var zones = [['<zone>','<coords>','<field>'...],[ ... ]]
                    For Each ch In responseString
                        Select Case state
                            Case state_enum.idle
                                If ch = "[" Then state = state_enum.arrayStart
                            Case state_enum.arrayStart
                                If ch = "[" Then state = state_enum.recordStart Else state = state_enum.idle
                            Case state_enum.recordStart
                                If ch = "'" Then state = state_enum.inField Else Throw New Exception($"State {state} ""'"" expected at char {index}")
                            Case state_enum.inField
                                If ch = "\" Then
                                    state = state_enum.escape
                                ElseIf ch = "'" Then
                                    state = state_enum.fieldEnd
                                    fields.Add(field)
                                    field = ""
                                Else
                                    field &= ch
                                End If
                            Case state_enum.escape
                                field &= ch
                                state = state_enum.inField
                            Case state_enum.fieldEnd
                                If ch = "," Then
                                    state = state_enum.fieldseparator
                                ElseIf ch = "]" Then
                                    state = state_enum.recordEnd
                                    zones.Add(New zonedata(fields))
                                    fields = New List(Of String)
                                Else
                                    Throw New Exception($"State {state} "", or ]"" expected at char {index}")
                                End If
                            Case state_enum.fieldseparator
                                If ch = "'" Then state = state_enum.inField Else Throw New Exception($"State {state} ""'"" expected at char {index}")
                            Case state_enum.recordseparator
                                If ch = "[" Then state = state_enum.recordStart Else Throw New Exception($"State {state} ""["" expected at char {index}")
                            Case state_enum.recordEnd
                                If ch = "]" Then
                                    Exit For ' Done
                                ElseIf ch = "," Then
                                    state = state_enum.recordseparator
                                End If
                            Case state_enum.arrayEnd
                                Exit For ' Done
                        End Select
                        index += 1
                    Next
                    Form1.AppendText(Form1.TextBox1, $"{zones.Count} {zc.Key} zones retrieved{vbCrLf}")
                    ' decode Google coordinates
                    ' data format https://developers.google.com/maps/documentation/utilities/polylinealgorithm
                    Dim polyliner = New Polyliner()
                    sql.CommandText = $"BEGIN TRANSACTION"      ' for speed
                    sql.ExecuteNonQuery()
                    sql.CommandText = $"DELETE FROM ZoneLines WHERE type='{zc.Key}'" ' Delete existing data
                    sql.ExecuteNonQuery()
                    For Each z In zones
                        Dim coords = polyliner.Decode(z.fields(1))
                        Dim linestr As New List(Of MapPoint)
                        For Each c In coords
                            linestr.Add(New MapPoint(c.Longitude, c.Latitude, SpatialReferences.Wgs84))
                        Next
                        If Not CoIncident(linestr.First, linestr.Last) Then linestr.Add(linestr.First)    ' close polygon
                        Dim poly As New Polygon(linestr)  ' create polygon
                        poly.Generalize(0.2, False)                  ' reduce number of points
                        poly = poly.Simplify                        ' ensure polygons have correct winding
                        ' Insert into database
                        sql.CommandText = $"INSERT INTO ZoneLines (Type,zone,geometry) VALUES ('{zc.Key}',{CInt(z.fields(0))},'{poly.ToJson}')"
                        sql.ExecuteNonQuery()
                    Next
                    sql.CommandText = $"COMMIT"      ' for speed
                    sql.ExecuteNonQuery()
                Next
            End Using
        End Using
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
    Async Function ImportTimeZones() As Task
        ' Import a timezone database
        Dim myQueryFilter As New QueryParameters, sql As SqliteCommand

        Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\Timezones\ne_10m_time_zones.shp")
        With myQueryFilter
            .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
            .ReturnGeometry = True
        End With
        Dim Timezones = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "DELETE FROM Timezones"       ' remove existing data
            sql.ExecuteNonQuery()
            For Each timezone In Timezones
                Dim poly = timezone.Geometry        ' get the geometry
                poly = poly.Simplify
                Dim polyGeneralized = poly.Generalize(0.1, True)   ' make it smaller
                If polyGeneralized.IsEmpty Then polyGeneralized = poly    ' if if disappeared when Generalized, restore it
                sql.CommandText = $"INSERT INTO Timezones (name,places,color,geometry) VALUES ('{Form1.SQLescape(timezone.Attributes("name"))}','{Form1.SQLescape(timezone.Attributes("places"))}',{timezone.Attributes("map_color6")},'{polyGeneralized.ToJson}')"
                sql.ExecuteNonQuery()
            Next
        End Using
        Form1.AppendText(Form1.TextBox1, $"{Timezones.Count} timezones imported{vbCrLf}")
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
        Form1.AppendText(Form1.TextBox1, $"{Regions.Count} IARU boundary lines loaded{vbCrLf}")
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
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
            geometry = geometry.Densify(5)
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
            Form1.AppendText(Form1.TextBox1, $"Linestrings {geometry.Parts.Count}, Original {OriginalCount}, Densify {DensifyCount}, Generalize {GeneralizeCount}{vbCrLf}")
        End Using
    End Function

    Async Function ImportAntarctica() As Task
        ' Get the country boundary for Antarctica. No useful one in OSM
        Dim qp As New QueryParameters
        Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\World countries generalized\World_Countries_Generalized.shp")
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            Dim sql As SqliteCommand, sqlDR As SqliteDataReader
            connect.Open()
            sql = connect.CreateCommand
            ' get the bounding polygon
            sql.CommandText = "SELECT * FROM DXCC WHERE Entity='Antarctica'"
            sqlDR = sql.ExecuteReader
            sqlDR.Read()
            Dim Antarctica As Polygon = Form1.ParseBox(sqlDR("bbox"))
            sqlDR.Close()
            With qp
                .WhereClause = "COUNTRY='Antarctica'"            ' get all features
                .Geometry = Antarctica
                .ReturnGeometry = True
                .OutSpatialReference = SpatialReferences.Wgs84
                .SpatialRelationship = SpatialRelationship.Intersects
            End With
            Dim fqr = Await Features.QueryFeaturesAsync(qp)
            Dim plb As New PolygonBuilder(SpatialReferences.Wgs84)
            For Each f In fqr
                Dim geom As Polygon = f.Geometry
                For Each p In geom.Parts
                    plb.AddPart(p)
                Next
            Next
            Dim poly As New Polygon(plb.Parts)
            'poly = poly.Intersection(Antarctica)        ' remove unwanted bits
            Dim polyGeneralized = Form1.GeneralizeByPart(poly)   ' reduce it in size
            Dim geometry = polyGeneralized.ToJson           ' convert to json
            ' Save in database

            sql.CommandText = $"UPDATE DXCC SET geometry='{geometry}' WHERE Entity='Antarctica'"
            sql.ExecuteNonQuery()
        End Using
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
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
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
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
                sql.CommandText = $"INSERT INTO Antarctic (open,name,nation,coordinates,situation,altitude) VALUES ('{open}','{Form1.SQLescape(Name)}','{Form1.SQLescape(nation)}','{coord.ToJson}','{situation}','{altitude}')"
                sql.ExecuteNonQuery()
            Next
        End Using
    End Function

    Async Function ImportIOTAGroups() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        Dim url = $"https://www.iota-world.org/rest/get/iota/groups?api_key={IOTA_API_KEY}"
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
        Dim response = JsonNode.Parse(responseString)
        If response!status <> "ok" Then
            MsgBox($"Error retrieving IOTA data ({response!status})", vbCritical + vbOKOnly, "No data")
        Else
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = response!content.AsArray.Count
            End With
            Using connect As New SqliteConnection(Form1.DXCC_DATA)
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
                For Each group In response!content.AsArray
                    ' process each group
                    Form1.ProgressBar1.Value += 1
                    With sql.Parameters
                        .Clear()
                        .AddWithValue("@refno", group.Item("refno").ToString)
                        .AddWithValue("@dxcc_num", group.Item("dxcc_num").ToString)
                        .AddWithValue("@name", group.Item("name").ToString)
                        .AddWithValue("@latitude_min", group.Item("latitude_min").ToString)
                        .AddWithValue("@latitude_max", group.Item("latitude_max").ToString)
                        .AddWithValue("@longitude_min", group.Item("longitude_min").ToString)
                        .AddWithValue("@longitude_max", group.Item("longitude_max").ToString)
                        .AddWithValue("@grp_region", group.Item("grp_region").ToString)
                        .AddWithValue("@whitelist", group.Item("whitelist").ToString)
                        .AddWithValue("@comment", group.Item("comment").ToString)
                    End With
                    sql.ExecuteNonQuery()
                Next
                sql.CommandText = "COMMIT"
                sql.ExecuteNonQuery()
            End Using
        End If
    End Function
    Async Function ImportIOTAIslands() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        Dim url = $"https://www.iota-world.org/rest/get/iota/islands?api_key={IOTA_API_KEY}"
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
        Dim response = JsonNode.Parse(responseString)
        If response!status <> "ok" Then
            MsgBox($"Error retrieving IOTA data ({response!status})", vbCritical + vbOKOnly, "No data")
        Else
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = response!content.AsArray.Count
            End With
            Using connect As New SqliteConnection(Form1.DXCC_DATA)
                connect.Open()
                sql = connect.CreateCommand
                sql.CommandText = "BEGIN TRANSACTION"
                sql.ExecuteNonQuery()                             ' delete existing data
                sql.CommandText = "DELETE FROM IOTA_Islands"
                sql.ExecuteNonQuery()                             ' delete existing data
                ' use prepared statement for speed
                sql.CommandText = $"INSERT INTO IOTA_Islands (`refno`,`name`,`comment`) VALUES (@refno,@name,@comment)"
                sql.Prepare()
                For Each group In response!content.AsArray
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
            End Using
        End If
    End Function
    Async Function ImportIOTADXCCMatchesOneIOTA() As Task
        Dim responseString As String = "", sql As SqliteCommand
        ' Get JSON data for all IOTA groups
        Dim url = $"https://www.iota-world.org/rest/get/iota/dxccmatchesoneiota?api_key={IOTA_API_KEY}"
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
        Dim response = JsonNode.Parse(responseString)
        If response!status <> "ok" Then
            MsgBox($"Error retrieving IOTA data ({response!status})", vbCritical + vbOKOnly, "No data")
        Else
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = response!content.AsArray.Count
            End With
            Using connect As New SqliteConnection(Form1.DXCC_DATA)
                connect.Open()
                sql = connect.CreateCommand
                sql.CommandText = "BEGIN TRANSACTION"
                sql.ExecuteNonQuery()                             ' delete existing data
                sql.CommandText = "DELETE FROM IOTA_DXCC_IOTA"
                sql.ExecuteNonQuery()                             ' delete existing data
                ' use prepared statement for speed
                sql.CommandText = $"INSERT INTO IOTA_DXCC_IOTA (`refno`,`dxcc_num`) VALUES (@refno,@dxcc_num)"
                sql.Prepare()
                For Each group In response!content.AsArray
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
            End Using
        End If
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
    Function NormalizeLongitude(longitude As Double) As Double
        ' Normalize a longitude to between -180 and +180
        Return (longitude Mod 360 + 540) Mod 360 - 180        ' normalize longitude
    End Function

End Module
