Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Imports Microsoft.VisualBasic.FileIO
Imports System.IO
Imports System.Net.Http
Imports System.Text.Json.Nodes
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Xml
Imports System.Xml.XPath
Imports HtmlAgilityPack

Module Import
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
        Dim nsmgr As New XmlNamespaceManager(New NameTable)
        nsmgr.AddNamespace("x", ns.NamespaceName)
        Dim linestrings = doc.XPathSelectElements("//x:LineString/x:coordinates", nsmgr)       ' find all the linestrings
        Form1.AppendText(Form1.TextBox1, $"{linestrings.Count} linestrings loaded{vbCrLf}")
        For Each coordinates In linestrings
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
    Sub ImportCQITUZones()
        ' Import CQ & ITU zone data
        Dim sql As SqliteCommand

        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            ' Delete existing data
            sql.CommandText = $"DELETE FROM ZoneLines"
            sql.ExecuteNonQuery()
            sql.CommandText = $"DELETE FROM ZoneLabels"
            sql.ExecuteNonQuery()
            ' Load CQ zone data
            Dim doc = XDocument.Load($"{Application.StartupPath}\CQ.kml")    ' read the XML
            Dim ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
            Dim nsmgr As New XmlNamespaceManager(New NameTable)
            nsmgr.AddNamespace("x", ns.NamespaceName)
            ' First extract 39 linestrings
            Dim folders = doc.XPathSelectElement("//x:Folder", nsmgr)   ' the top level folder
            Dim folder = folders.XPathSelectElement("x:Folder[1]", nsmgr)       ' find all the linestrings
            'Dim CQboundaries = boundaries.XPathSelectElements("x:Placemark", nsmgr)
            Dim CQboundaries = From p In folder.Descendants Where p.Name.LocalName = "Placemark" Select p
            Form1.AppendText(Form1.TextBox1, $"Retrieved {CQboundaries.Count} CQ zone boundaries{vbCrLf}")
            For Each boundary In CQboundaries
                Dim name = boundary.XPathSelectElement("x:name", nsmgr).Value
                Dim NameSplit = name.Split("-")
                Dim line = CInt(NameSplit(1))
                Dim coordinates = boundary.XPathSelectElement("x:LineString/x:coordinates", nsmgr).Value
                coordinates = Regex.Replace(coordinates, "[^0-9\-\., ]", "")     ' remove noise characters
                coordinates = Trim(coordinates)
                Dim coords = Split(coordinates, " ").ToList   ' bust into coordinate pairs
                Dim linestr As New List(Of MapPoint)
                For Each coord In coords
                    Dim points = Split(coord, ",")
                    linestr.Add(New MapPoint(points(0), points(1), SpatialReferences.Wgs84))
                Next
                Dim linestring As New Multipoint(linestr)   ' create multipoint linestring
                ' Insert into database
                sql.CommandText = $"INSERT INTO ZoneLines (Type,line,geometry) VALUES ('CQ',{line},'{linestring.ToJson}')"
                sql.ExecuteNonQuery()
            Next

            ' Now extract 40 zone labels
            Dim labels = folders.XPathSelectElement("x:Folder[2]", nsmgr)       ' find all the labels
            Dim CQlabels = From p In labels.Descendants Where p.Name.LocalName = "Placemark" Select p
            Form1.AppendText(Form1.TextBox1, $"Retrieved {CQlabels.Count} CQ zone labels{vbCrLf}")
            For Each label In CQlabels
                Dim name = label.XPathSelectElement("x:name", nsmgr).Value
                Dim NameSplit = name.Split(" ")
                Dim Zone = CInt(NameSplit(1))
                Dim coordinates = label.XPathSelectElement("x:Point/x:coordinates", nsmgr).Value
                coordinates = Regex.Replace(coordinates, "[^0-9\-\., ]", "")     ' remove noise characters
                coordinates = Trim(coordinates)
                Dim coords = Split(coordinates, ",").ToList   ' bust into coordinates
                Dim pnt As New MapPoint(coords(0), coords(1), SpatialReferences.Wgs84)
                ' Insert into database
                sql.CommandText = $"INSERT INTO ZoneLabels (Type,zone,geometry) VALUES ('CQ',{Zone},'{pnt.ToJson}')"
                sql.ExecuteNonQuery()
            Next

            ' Load ITU zone data
            doc = XDocument.Load($"{Application.StartupPath}\ITU.kml")    ' read the XML
            ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
            nsmgr = New XmlNamespaceManager(New NameTable)
            nsmgr.AddNamespace("x", ns.NamespaceName)
            ' First extract linestrings
            folders = doc.XPathSelectElement("//x:Folder", nsmgr)   ' the top level folder
            folder = folders.XPathSelectElement("x:Folder[1]", nsmgr)       ' find all the linestrings

            Dim ITUboundaries = From p In folder.Descendants Where p.Name.LocalName = "Placemark" Select p
            Form1.AppendText(Form1.TextBox1, $"Retrieved {ITUboundaries.Count} ITU zone boundaries{vbCrLf}")
            For Each boundary In ITUboundaries
                Dim name = boundary.XPathSelectElement("x:name", nsmgr).Value
                Dim NameSplit = name.Split("-")
                Dim line = CInt(NameSplit(1))
                Dim coordinates = boundary.XPathSelectElement("x:LineString/x:coordinates", nsmgr).Value
                coordinates = Regex.Replace(coordinates, "[^0-9\-\., ]", "")     ' remove noise characters
                coordinates = Trim(coordinates)
                Dim coords = Split(coordinates, " ").ToList   ' bust into coordinate pairs
                Dim linestr As New List(Of MapPoint)
                For Each coord In coords
                    Dim points = Split(coord, ",")
                    linestr.Add(New MapPoint(points(0), points(1), SpatialReferences.Wgs84))
                Next
                Dim linestring As New Multipoint(linestr)   ' create multipoint linestring
                ' Insert into database
                sql.CommandText = $"INSERT INTO ZoneLines (Type,line,geometry) VALUES ('ITU',{line},'{linestring.ToJson}')"
                sql.ExecuteNonQuery()
            Next

            ' Now extract 40 zone labels
            labels = folders.XPathSelectElement("x:Folder[2]", nsmgr)       ' find all the labels
            Dim ITUlabels = From p In labels.Descendants Where p.Name.LocalName = "Placemark" Select p
            Form1.AppendText(Form1.TextBox1, $"Retrieved {ITUlabels.Count} ITU zone labels{vbCrLf}")
            For Each label In ITUlabels
                Dim name = label.XPathSelectElement("x:name", nsmgr).Value
                Dim NameSplit = name.Split(" ")
                Dim Zone = CInt(NameSplit(1))
                Dim coordinates = label.XPathSelectElement("x:Point/x:coordinates", nsmgr).Value
                coordinates = Regex.Replace(coordinates, "[^0-9\-\., ]", "")     ' remove noise characters
                coordinates = Trim(coordinates)
                Dim coords = Split(coordinates, ",").ToList   ' bust into coordinates
                Dim pnt As New MapPoint(coords(0), coords(1), SpatialReferences.Wgs84)
                ' Insert into database
                sql.CommandText = $"INSERT INTO ZoneLabels (Type,zone,geometry) VALUES ('ITU',{Zone},'{pnt.ToJson}')"
                sql.ExecuteNonQuery()
            Next
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
        ' Import a timezone database
        Dim myQueryFilter As New QueryParameters, sql As SqliteCommand, count As Integer = 0

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
            Dim GeneralizeAngle As Double = Form1.RadtoDeg(Math.Asin(20 * 1000 / Form1.EARTH_RADIUS))   ' angle for Generalize
            For Each Reg In Regions
                count += 1
                Dim geom As Polyline = Reg.Geometry
                Dim OriginalCount As Integer = 0
                Dim DensifyCount As Integer = 0
                Dim GeneralizeCount As Integer = 0
                For Each line In geom.Parts
                    OriginalCount += line.PointCount
                Next
                geom = geom.Densify(2)    ' make sure there is a point every 2 degrees
                For Each line In geom.Parts
                    DensifyCount += line.PointCount
                Next
                geom = GeometryEngine.Generalize(geom, GeneralizeAngle, True)   ' Generalize a polygon to reduce it in size
                For Each line In geom.Parts
                    GeneralizeCount += line.PointCount
                Next
                If geom.IsEmpty Then geom = Reg.Geometry
                If geom.LengthGeodetic > 0 Then sql.CommandText = $"INSERT INTO IARU (line,geometry) VALUES ({count},'{geom.ToJson}')"
                sql.ExecuteNonQuery()
                Form1.AppendText(Form1.TextBox1, $"{count} Parts {geom.Parts.Count}, Original {OriginalCount}, Densify {DensifyCount}, Generalize {GeneralizeCount}{vbCrLf}")
            Next
        End Using
    End Function

    Async Function ImportAntarctica() As Task
        ' Get the country boundary for Antarctica. No useful one in OSM
        Dim qp As New QueryParameters
        Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\World countries generalized\World_Countries_Generalized.shp")
        With qp
            .WhereClause = "COUNTRY='Antarctica'"            ' get all features
            .ReturnGeometry = True
        End With
        Dim fqr = Await Features.QueryFeaturesAsync(qp)
    End Function

    Async Function ImportAntarcticBases() As Task
        Dim responseString As String
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

    Const IOTA_API_KEY = "Y3935PWYAQZYP3HVZ7QA"
    Async Function ImportIOTAGroups() As Task
        Dim responseString As String, sql As SqliteCommand
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
                sql.ExecuteNonQuery()                             ' delete existing data
            End Using
        End If
    End Function
    Async Function ImportIOTAIslands() As Task
        Dim responseString As String, sql As SqliteCommand
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
                sql.ExecuteNonQuery()                             ' delete existing data
            End Using
        End If
    End Function
    Async Function ImportIOTADXCCMatchesOneIOTA() As Task
        Dim responseString As String, sql As SqliteCommand
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
                sql.ExecuteNonQuery()                             ' delete existing data
            End Using
        End If
    End Function
End Module
