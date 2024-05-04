Imports System.Collections.Immutable
Imports System.Diagnostics.Contracts
Imports System.IO
Imports System.IO.Compression
Imports System.Net.Http
Imports System.Text.Json.Nodes
Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite

Enum Winding
    Outer = 1
    Inner = 2
End Enum
Public Class Form1
    Public Const GridFieldX = 20, GridFieldY = 10    ' size of gridfield in degrees in X,Y
    Public Const GridSquareX = 2, GridSquareY = 1    ' size of gridsquare in degrees in X,Y
    Public Const DXCC_DATA = "data Source=DXCC.sqlite"     ' the DXCC database
    Public Const EARTH_RADIUS = 6371000    ' radius of earth in meters
    Const CLOSENESS = 1000     ' distance for Generalize in meters
    Public KMLheader As String
    Public Const KMLfooter = "</Document></kml>"       ' standard footer for kml file
    Public ColourMapping = {"", "red", "green", "blue", "yellow", "cyan", "magenta", "white"}   ' colours for polygons
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Make standard KML header for all files
        KMLheader = $"<?xml version='1.0' encoding='UTF-8'?>
<kml xmlns = ""http://www.opengis.net/kml/2.2"" xmlns:gx=""http://www.google.com/kml/ext/2.2"">
<Document><open>1</open>
<Style id = ""boundary""><LineStyle><color>ffff0000</color><width>3</width></LineStyle><PolyStyle><color>7Fff0000</color></PolyStyle></Style>
<Placemark><name>DXCC Map Of the World by VK3OHM</name>
<description><![CDATA[Copyright: The data included in this document is from www.openstreetmap.org. The data is made available under ODbL.
    Data extraction by Marc Hillman (VK3OHM).<br><br>
The main purpose of this data is to display the boundaries of every DXCC entity. Adjacent entities have different colours.<br><br>
There are some additional folders which are closed by default. You must open them to see the contents. They are:<br><br>
<table border=1>
<tr><th>Folder</th><th>Description</th></tr>
<tr><td>DXCC Entities</td><td>Polygons displaying the boundaries of all DXCC entities.</td></tr>
<tr><td>Prefixes</td><td>The ARRL prefix for each entity is displayed in the center of the entity.</td></tr>
<tr><td>Grid Squares</td><td>The boundary of every grid square that intersects with an entity is displayed. The 4-character grid square code is displayed in the center of the grid square. This folder is searchable, so you can use it to locate any gridsquare.</td></tr>
<tr><td>CQ Zones</td><td>CQ magazine (now defunct) zones used for Worked All Zones (WAZ) award now administered by ARRL.(Using data created by Francesco Crosilla IV3TMM (SK))</td></tr>
<tr><td>ITU Zones</td><td>International Telegraphic Union (ITU) zones. (Using data created by Francesco Crosilla IV3TMM (SK))</td></tr>
<tr><td>IARU regions</td><td>International Amateur Radio Union (IARU) regions. (Using data created by Tim Makins (EI8IC))</td></tr>
<tr><td>Timezones</td><td>World time zones</td></tr>
<tr><td>Antarctic bases</td><td>Location and basic details of Antarctic bases.</td></tr>
<tr><td>Bounding boxes</td><td>To extract the entity data it was sometimes necessary to use a bounding box to filter the returned geometry. Open this folder will display those boxes. They are not much use in normal operation, but are used for debugging.</td></tr>
</table>
<br>Data created {Now:R}
]]>
</description><gx:balloonVisibility>1</gx:balloonVisibility></Placemark>
<Style id=""red""><PolyStyle><color>9F0000Ff</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ff0000ff</color><width>2</width></LineStyle></Style>
<Style id=""green""><PolyStyle><color>9F00Ff00</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ff00ff00</color><width>2</width></LineStyle></Style>
<Style id=""blue""><PolyStyle><color>9Fff0000</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ffff0000</color><width>2</width></LineStyle></Style>
<Style id=""yellow""><PolyStyle><color>9F00Ffff</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ff00ffff</color><width>2</width></LineStyle></Style>
<Style id =""cyan""><PolyStyle><color>9Fff00ff</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ffff00ff</color><width>2</width></LineStyle></Style>
<Style id=""magenta""><PolyStyle><color>9Fffff00</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ffffff00</color><width>2</width></LineStyle></Style>
<Style id=""white""><PolyStyle><color>9Fffffff</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ffffffff</color><width>2</width></LineStyle></Style>
"
        ' highlight polygons when hovering
        For i = 1 To ColourMapping.Length - 1
            KMLheader &= $"<StyleMap id='boundary_{i}'>{vbCrLf}"
            KMLheader &= $"<Pair><key>normal</key><styleUrl>#{ColourMapping(i)}</styleUrl></Pair>{vbCrLf}"
            KMLheader &= $"<Pair><key>highlight</key><styleUrl>#white</styleUrl></Pair>{vbCrLf}"
            KMLheader &= $"</StyleMap>{vbCrLf}"
        Next
    End Sub
    Private Function RangeString(number_array As List(Of Integer)) As String
        ' convert a list of numbers e.g. 11,16,12,17,18,15,22,23,24    to     11-12,15-18,22-24
        Dim number As Integer, previous_number As Integer, range As Boolean, result As String
        Debug.Assert(number_array.Count > 0, "0 length list")
        number_array.Sort()
        ' Loop through array And build range string
        previous_number = number_array(0)       ' start with first number
        range = False                           ' not in a contiguous range of numbers
        result = $"{previous_number:00}"         ' start with first number in list
        For i = 1 To number_array.Count - 1
            number = number_array(i)
            If number = previous_number + 1 Then
                range = True
            Else
                If range Then
                    result &= $"-{previous_number:00}"
                    range = False
                End If
                result &= $", {number:00}"
            End If
            previous_number = number
        Next
        If range Then result &= $"-{previous_number:00}"
        Return result
    End Function
    Shared Function GridSquare(p As MapPoint) As String
        ' Convert a mappoint to a gridsquare
        ' https://ham.stackexchange.com/questions/221/how-can-one-convert-from-lat-long-to-grid-square
        Const upper = "ABCDEFGHIJKLMNOPQR"
        Const numbers = "0123456789"
        Dim lon As Double = p.X + 180, lat As Double = p.Y + 90
        Dim lat_sq = upper(Math.Floor(lat / GridFieldY))
        Dim lon_sq = upper(Math.Floor(lon / GridFieldX))
        Dim lat_field = numbers(Math.Floor(lat Mod 10))
        Dim lon_field = numbers(Math.Floor((lon / 2) Mod 10))
        Return $"{lon_sq}{lat_sq}{lon_field}{lat_field}"
    End Function

    Private Async Sub UseShapefileToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles UseShapefileToolStripMenuItem1.Click
        'Dim json As String = File.ReadAllText("world-administrative-boundaries.geojson")
        'Boundaries = FeatureCollection.FromJson(json)
        Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\World countries generalized\World_Countries_Generalized.shp")
        Dim Range As String, prefix As String
        Dim csvfields As New List(Of String), square_list As New List(Of Integer), field As String, name As String
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, locators As New List(Of String)
        Dim sql1 As SqliteCommand, SQLdr1 As SqliteDataReader
        Dim PrefixesFound As New List(Of String)
        Dim DXCC As Integer, ProductionGrids As New List(Of String), squares As New List(Of String)
        ' List mapping world administrative area names to DXCC names
        Dim NameChanges As New Dictionary(Of String, String) From {
            {"Palestinian Territory", "Palestine"},
            {"Vatican City", "Vatican"},
            {"Antigua and Barbuda", "Antigua & Barbuda"},
            {"Azores Islands", "Azores"},
            {"Bosnia and Herzegovina", "Bosnia-Herzegovina"},
            {"Bouvet Island", "Bouvet"},
            {"British Virgin Islands", "British Virgin Is"},
            {"Brunei Darussalam", "Brunei"},
            {"Burkina Faso", "Burkina-Faso"},
            {"Cabo Verde", "Cape Verde"},
            {"Canarias", "Canary Is"},
            {"Cayman Islands", "Cayman Is"},
            {"Christmas Island", "Christmas Is"},
            {"Cocos Islands", "Cocos Is"},
            {"Cocos (Keeling) Islands", "Cocos-Keeling Is"},
            {"Congo", "Republic of the Congo"},
            {"Côte d'Ivoire", "Cote d'Ivoire"},
            {"Congo DRC", "Dem. Republic of the Congo"},
            {"Falkland Islands", "Falkland Is"},
            {"Faroe Islands", "Faroe Is"},
            {"Gambia", "The Gambia"},
            {"Glorioso Islands", "Glorioso Is"},
            {"Guantanamo", "Guantanamo Bay"},
            {"Heard Island and McDonald Islands", "Heard Is"},
            {"Iran (Islamic Republic of)", "Iran"},
            {"Juan De Nova Island", "Juan de Nova"},
            {"Madeira", "Madeira Is"},
            {"Marshall Islands", "Marshall Is"},
            {"Mauritius", "Mauritius Is"},
            {"Micronesia (Federated States of)", "Micronesia"},
            {"Midway Is.", "Midway Is"},
            {"Moldova, Republic of", "Moldova"},
            {"Netherlands Antilles", "Bonaire, Curacao"},
            {"North Korea", "Democratic People's Republic of Korea"},
            {"Northern Mariana Islands", "Mariana Is"},
            {"Norfolk Island", "Norfolk Is"},
            {"Pitcairn", "Pitcairn Is"},
            {"Réunion", "Reunion"},
            {"Saba", "Saba & St. Eustatius"},
            {"Saint Barthelemy", "St. Barthelemy"},
            {"Saint Eustatius", "St. Eustatius"},
            {"Saint Helena", "St Helena Is"},
            {"Saint Kitts and Nevis", "St Kitts & Nevis"},
            {"Saint Lucia", "St Lucia"},
            {"Saint Pierre and Miquelon", "St. Pierre & Miquelon"},
            {"Saint Vincent and the Grenadines", "St Vincent"},
            {"Sao Tome and Principe", "Sao Tome & Principe"},
            {"Slovakia", "Slovak Republic"},
            {"Solomon Islands", "Solomon Is"},
            {"South Africa", "Republic of South Africa"},
            {"South Korea", "Republic of Korea"},
            {"Svalbard", "Svalbard Is"},
            {"Turkiye", "Turkey"},
            {"Eswatini", "Kingdom of eSwatini"},
            {"Syrian Arab Republic", "Syria"},
            {"Tokelau", "Tokelau Is"},
            {"Trinidad and Tobago", "Trinidad & Tobago"},
            {"Turks and Caicos Islands", "Turks & Caicos Is"},
            {"United Republic of Tanzania", "Tanzania"},
            {"US Virgin Islands", "US Virgin Is"},
            {"United States of America", "United States"},
            {"Wallis and Futuna", "Wallis & Futuna"}
            }

        PrefixesFound.Clear()
        Dim myQueryFilter As New QueryParameters
        With myQueryFilter
            .OutSpatialReference = SpatialReferences.Wgs84     ' results in WGS84
            .ReturnGeometry = True
        End With
        myQueryFilter.OrderByFields.Add(New Data.OrderBy("country", Data.SortOrder.Ascending))         ' sort the results by name
        Dim Countries = Await Features.QueryFeaturesAsync(myQueryFilter).ConfigureAwait(False)           ' run query

        Using connect As New SqliteConnection(DXCC_DATA),
            gridlist As New StreamWriter($"{Application.StartupPath}\GridList.txt", False),
             missing As New StreamWriter($"{Application.StartupPath}\missing.txt", False)
            connect.Open()  ' open database
            sql = connect.CreateCommand
            sql1 = connect.CreateCommand
            For Each country In Countries
                name = country.Attributes("country")
                Dim value As String = Nothing
                If NameChanges.TryGetValue(name, value) Then name = value
                ' Find the prefix
                sql.CommandText = $"SELECT * FROM DXCC WHERE Entity='{name.Replace("'", "''")}' AND Deleted=0"
                SQLdr = sql.ExecuteReader()
                If SQLdr.HasRows Then
                    SQLdr.Read()
                    prefix = SQLdr("prefix")
                    PrefixesFound.Add($"'{prefix}'")      ' remember prefixes found
                    ProductionGrids.Clear()
                    ' Make a list of grids in production system
                    DXCC = SQLdr("DXCCnum")
                    sql1.CommandText = $"SELECT * FROM DXCCtoGRID WHERE DXCC={DXCC} ORDER BY GRID"
                    SQLdr1 = sql1.ExecuteReader()
                    While SQLdr1.Read
                        ProductionGrids.Add(SQLdr1("GRID"))
                    End While
                    SQLdr1.Close()
                    ' produce gridsquare list
                    Dim env = country.Geometry.Extent         '   find extents of country
                    ' Find top left and bottom right corner of square
                    Dim TopLeft = New MapPoint(Int(env.XMin / GridSquareX) * GridSquareX, Int(env.YMin / GridSquareY) * GridSquareY, SpatialReferences.Wgs84)
                    Dim BottomRight = New MapPoint(Int((env.XMax + GridSquareX) / GridSquareX) * GridSquareX, Int((env.YMax + GridSquareY) / GridSquareY) * GridSquareY, SpatialReferences.Wgs84)
                    ' test each gridsquare in the extent to see if it intersects with country
                    Dim x As Integer, y As Integer
                    squares.Clear()
                    x = TopLeft.X
                    While x < BottomRight.X
                        y = TopLeft.Y
                        While y < BottomRight.Y
                            Dim square = New Envelope(x, y, x + GridSquareX, y + GridSquareY, SpatialReferences.Wgs84)   ' the square we are testing
                            Dim SquareBuffered = GeometryEngine.BufferGeodetic(square, -1000, LinearUnits.Meters)        ' reduce the square by 100 on all sides
                            If GeometryEngine.Intersects(SquareBuffered, country.Geometry) Then
                                squares.Add(GridSquare(New MapPoint(x, y, SpatialReferences.Wgs84)))    ' add to list of gridsquares
                            End If
                            y += GridSquareY
                        End While
                        x += GridSquareX
                    End While
                    squares.Sort()

                    If squares.Count = 0 Then
                        gridlist.WriteLine($"// Found no grids for {name}")
                    Else
                        ' export grids
                        csvfields.Clear()
                        If name.Contains(" "c) Then name = $"""{name}"""     ' escape spaces
                        csvfields.Add(name)  ' name
                        csvfields.Add(prefix)     ' prefix
                        square_list.Clear()
                        Dim previous_field As String = squares(0).Substring(0, 2)      ' start with first field
                        Dim sqr = Int(squares(0).Substring(2, 2))   ' the first square
                        square_list.Add(sqr)
                        locators.Clear()
                        For i = 1 To squares.Count - 1
                            field = squares(i).Substring(0, 2)      ' the field
                            sqr = Int(squares(i).Substring(2, 2))   ' the square
                            If field <> previous_field Then
                                ' beginning of New field
                                ' Change of field - output current one
                                Range = RangeString(square_list) ' convert To shorter range format
                                locators.Add($"{previous_field} {Range}")
                                square_list.Clear()
                            End If
                            previous_field = field
                            square_list.Add(sqr)       ' add square to list
                        Next
                        ' Last entry
                        Range = RangeString(square_list) ' convert to shorter range format
                        locators.Add($"{previous_field} {Range}")
                        csvfields.Add($"""{String.Join(", ", locators)}""")
                        Dim txt = String.Join(", ", csvfields)
                        gridlist.WriteLine(txt)
                    End If
                    ' display differences between two grid lists
                    If squares.Count > 0 Then
                        Dim Production = ProductionGrids.Except(squares)    ' squares that are in production, but not world
                        Dim World = squares.Except(ProductionGrids)         ' squares that are in world, but not production
                        missing.WriteLine(name)
                        If Not Production.Except(World).Any And Not World.Except(Production).Any Then
                            missing.WriteLine("Lists match")
                        Else
                            missing.WriteLine($" Missing from production {String.Join(",", World)}")
                            missing.WriteLine($" Missing from world {String.Join(",", Production)}")
                        End If
                    End If
                Else
                    prefix = "???"      ' failed to lookup prefix
                End If
                SQLdr.Close()
                AppendText(TextBox1, $"{name}{vbCrLf}")
            Next
            ' Display prefixes not found
            Dim count As Integer = 0
            sql.CommandText = $"SELECT * FROM DXCC WHERE prefix NOT IN ({String.Join(",", PrefixesFound)}) AND Deleted=0 ORDER By Entity"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read()
                gridlist.WriteLine($"// Not found {SQLdr("Entity")} {SQLdr("prefix")}")
                count += 1
            End While
            gridlist.WriteLine($"// Total of {count} DXCC entities not found, {PrefixesFound.Count} found")
        End Using
    End Sub

    Private Async Sub UseOSMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UseOSMToolStripMenuItem.Click
        ' Retrieve country boundaries from OpenStreetMap (OSM)
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim DXCC As Integer, Entity As String, geometry As String
        Dim ProcessList As New List(Of String)        ' list of countries to process
        Dim Specials As New List(Of Integer) From {45, 247}   ' islands with no geometry

        Using connect As New SqliteConnection(DXCC_DATA)
            ' get all countries where the OSM parameters are known
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM `DXCC` WHERE `geometry` IS NULL AND `query` IS NOT NULL AND `Deleted`=0 ORDER BY `Entity`"
            SQLdr = sql.ExecuteReader()
            While SQLdr.Read
                DXCC = SQLdr("DXCCnum")
                Entity = SQLdr("Entity")
                ProcessList.Add(Entity)
            End While
            SQLdr.Close()     ' close the reader
            For Each item In ProcessList
                Await CreateGrids(connect, item)
            Next
            ' Do specials
            For Each DXCC In Specials
                sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={DXCC}"
                SQLdr = sql.ExecuteReader()
                SQLdr.Read()
                Dim pnt = New MapPoint(SQLdr("lon"), SQLdr("lat"), SpatialReferences.Wgs84)     ' create a point at the DXCC
                SQLdr.Close()
                Dim boundary = GeometryEngine.BufferGeodetic(pnt, 50000, LinearUnits.Meters)     ' create a cicular island
                geometry = boundary.ToJson
                sql.CommandText = $"UPDATE `DXCC` SET `geometry`='{geometry}' WHERE `DXCCnum`={DXCC}"
                sql.ExecuteNonQuery()         ' insert into database
            Next
        End Using
        AppendText(TextBox1, $"Done{vbCrLf}")
    End Sub
    Shared Function IsDBNullorEmpty(ByVal value As Object) As Boolean
        ' test if DB field is null or empty
        Return IsDBNull(value) OrElse value = ""
    End Function
    Public Async Function CreateGrids(connect As SqliteConnection, country As String) As Task(Of Integer)

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, sqlupd As SqliteCommand
        Dim responseString As String = "", geometry As String, dxcc As Integer, query As String

        AppendText(TextBox1, $"Creating grids for {country}{vbCrLf}")

        sql = connect.CreateCommand
        sqlupd = connect.CreateCommand
        ' Retrieve any existing geometry
        sql.CommandText = $"SELECT * FROM `DXCC` WHERE `Entity`='{SQLescape(country)}' AND `Deleted`=0"
        SQLdr = sql.ExecuteReader()
        If Not SQLdr.HasRows Then
            MsgBox($"No data found for {country}", vbCritical + vbOKOnly, "Data lookup error")
            End
        End If
        SQLdr.Read()
        dxcc = SQLdr("DXCCnum")
        If IsDBNullorEmpty(SQLdr("geometry")) Then
            ' We have no cached geometry, so create it
            AppendText(TextBox1, $"  No cached geometry - generating{vbCrLf}")
            Dim OSMTimer As New Stopwatch
            OSMTimer.Restart()
            AppendText(TextBox1, "Fetching way data from OSM ")
            Using httpClient As New System.Net.Http.HttpClient()
                httpClient.Timeout = New TimeSpan(0, 10, 0)        ' 10 min timeout
                Dim url = "https://overpass-api.de/api/interpreter"
                Dim criteria As String
                If Not IsDBNullorEmpty(SQLdr("query")) Then
                    criteria = SQLdr("query")
                Else
                    MsgBox($"No search criteria for {country}", vbCritical + vbOKOnly, "Not criteria")
                    Exit Function
                End If
                ' the name is sometimes not sufficient to specify the target of the query, e.g. "Christmas Island" which is not unique
                ' we can specify a bounding box in this case

                'Dim query = $"[out:json][timeout:100];
                '    {criteria};
                '(
                '    way(r)[maritime != yes];
                '    way(area)[natural=coastline];
                ');
                'out geom;"
                Dim bbox As String = ""
                If Not IsDBNull(SQLdr("bbox")) AndAlso SQLdr("bbox") <> "" Then
                    ' validate bounding box and add to query
                    Dim box = Split(SQLdr("bbox"), ",")
                    Debug.Assert(box.Length = 4, "Bounding box must contain 4 values")
                    Debug.Assert(CDbl(box(0)) < CDbl(box(2)), "Bottom value must be less than top")
                    ' Need to check left and right, but there could be a situation (e.g. Fiji) where the bounding box crosses the Antimeridian
                    Debug.Assert(CDbl(box(1)) < CDbl(box(3)) Or (CDbl(box(1)) > 0 And CDbl(box(3)) < 0 And Math.Abs(CDbl(box(1)) + CDbl(box(3))) < 10), "Left value must be less than right")
                    bbox = $"({SQLdr("bbox")})"
                    Dim bounds = ParseBox(SQLdr("bbox"))
                End If
                query = $"[out:json][timeout:100];
                        {criteria};
                    out geom{bbox};"
                Dim body = New StringContent(query)
                Try
                    Dim httpResult As System.Net.Http.HttpResponseMessage = Await httpClient.PostAsync(url, body)
                    httpResult.EnsureSuccessStatusCode()
                    responseString = Await httpResult.Content.ReadAsStringAsync()
                    AppendText(TextBox1, $"{responseString.Length} bytes retrieved {OSMTimer.ElapsedMilliseconds / 1000:f1}s{vbCrLf}")
                Catch ex As HttpRequestException
                    MsgBox($"{ex.Message}{vbCrLf}query={query}", vbCritical + vbOKOnly, "OSM request error")
                End Try
            End Using

            Dim response = JsonNode.Parse(responseString)
            Dim elements As JsonNode = response!elements        ' both way and relation queries have an "elements" node
            If elements.AsArray.Count = 0 Then
                MessageBox.Show($"No geometry found for {country} with query={query}")
                Exit Function
            End If

            Dim plb As New PolylineBuilder(SpatialReferences.Wgs84), relations As New List(Of Integer), rel

            For Each element In elements.AsArray
                Dim ElementType = element!type.ToString
                rel = element!id
                Select Case ElementType
                    Case "relation"
                        Dim members As JsonArray = element!members
                        ' populate polylinebuilder with ways as linestrings
                        For Each member As JsonObject In members
                            Dim ty As String = member!type
                            If ty = "way" Then      ' a linestring
                                Dim way = New Part(plb.SpatialReference)
                                For Each pnt In member!geometry.AsArray
                                    If pnt IsNot Nothing Then way.AddPoint(CDbl(pnt("lon")), CDbl(pnt("lat")))
                                Next
                                If Not way.IsEmpty Then
                                    plb.AddPart(way)
                                    relations.Add(CInt(rel.ToString))      ' the reference for this way
                                End If
                            End If
                        Next
                    Case "way"
                        ' it's a way
                        Dim way = New Part(plb.SpatialReference)
                        For Each pnt In element!geometry.AsArray
                            If pnt IsNot Nothing Then way.AddPoint(CDbl(pnt("lon")), CDbl(pnt("lat")))
                        Next
                        If Not way.IsEmpty Then
                            plb.AddPart(way)
                            relations.Add(CInt(rel.ToString))      ' the reference for this way
                        End If
                End Select
            Next

            'Connectivity(plb)
            'Repair(plb)
            Dim poly As Polygon = CreatePolygon(plb, relations, country)
            geometry = poly.ToJson           ' convert to json
            SQLdr.Close()
            sql.CommandText = $"UPDATE `DXCC` SET `geometry`='{geometry}' WHERE `DXCCnum`={dxcc}"
            sql.ExecuteNonQuery()         ' insert into database
        End If
        If Not SQLdr.IsClosed Then SQLdr.Close()
        Return 1
    End Function

    Function ParseBox(bbox As String) As Geometry
        ' Convert a Bounding Box specification to a multipoint
        ' It may be a bbox of form "a,b,c,d",  or a polygon of form "poly:a b c d e f g h i j"
        Dim mpb As New MultipointBuilder(SpatialReferences.Wgs84), result As Multipoint = Nothing
        If IsDBNullorEmpty(bbox) Then Return result        ' no bounds
        Dim groups = bbox.Split(",")
        If groups.Count = 4 Then
            mpb.Points.Add(New MapPoint(CDbl(groups(1)), CDbl(groups(0))))
            mpb.Points.Add(New MapPoint(CDbl(groups(1)), CDbl(groups(2))))
            mpb.Points.Add(New MapPoint(CDbl(groups(3)), CDbl(groups(2))))
            mpb.Points.Add(New MapPoint(CDbl(groups(3)), CDbl(groups(0))))
            mpb.Points.Add(New MapPoint(CDbl(groups(1)), CDbl(groups(0))))
            result = mpb.ToGeometry
        Else
            Dim matches = Regex.Match(bbox, "^poly:""([\d\.\- ]+)""$")
            If matches.success Then
                Dim data = Split(matches.Groups(1).Value, " ")      ' split space separated list of coordinates
                Dim X As Double, Y As Double
                Debug.Assert(data.Count Mod 2 = 0, "Odd number of coordinates")
                For i = 0 To data.Count - 1 Step 2
                    Debug.Assert(Double.TryParse(data(i), Y), "Badly formed double")
                    Debug.Assert(Double.TryParse(data(i + 1), X), "Badly formed double")
                    mpb.Points.Add(New MapPoint(CDbl(X), CDbl(Y)))         ' add xy pair
                Next
                result = mpb.ToGeometry
            End If
        End If
        Return result
    End Function
    Function CreatePolygon(plb As PolylineBuilder, relations As List(Of Integer), country As String) As Polygon
        ' Convert a polyline into a polygon
        ' the ways retrieved are disjoint fragments of polygons.
        ' We must connect up the disjoint fragments.
        ' Note: fragments may connect Head to Tail, Tail to Tail, Head to Head or Tail to Head

        ' It is a 2 pass process
        ' 1. Look for shared borders and remove them. For countries that are created by an amalgamation of states, e.g. the Russia's and India,
        '    adjacent states will each have an identical border. We don't want internal borders, so remove any ways that share a start/end point. Borders can only be shared if
        '    they in different relations.
        ' 2. Connect all ways that have a connection to only one other way. Ignore situation where three or more ways are joined (shouldn't be any).

        Dim finished As Boolean = False
        Dim pass As Integer = 1
        Dim PassWatch As New Stopwatch     ' measures total execution time
        Dim touches
        Dim head As New Part(plb.SpatialReference)      ' the front of the joined segment
        Dim tail As New Part(plb.SpatialReference)      ' the back of the joined segment
        Dim closed As Integer        ' number of ways closed
        Dim RemovedCount As Integer
        Dim original As New PolylineBuilder(plb.ToGeometry)        ' save if needed

        ' dump out ways fro debug purposes
        'Using dump As New StreamWriter("PLB dump.txt", False)
        '    For Each prt In plb.Parts
        '        If Not prt.IsEmpty Then
        '            With prt.Points
        '                dump.WriteLine($"{ .First.X:f7},{ .First.Y:f7} { .Last.X:f7},{ .Last.Y:f7} points={prt.PointCount} Part={plb.Parts.IndexOf(prt)}")
        '            End With
        '        Else
        '            dump.WriteLine($"Part {plb.Parts.IndexOf(prt)} is empty")
        '        End If
        '    Next
        'End Using
        Debug.Assert(plb.Parts.Any, "Geometry is empty")
        ' Measure size of starting data
        Dim beforeParts = plb.Parts.Count
        Dim beforePoints = 0
        For Each prt In plb.Parts
            beforePoints += prt.PointCount
        Next
        ' Delete internal, shared ways.
        With ProgressBar1
            .Minimum = 0
            .Maximum = plb.Parts.Count - 1
            .Step = 10
        End With
        For partindex = 0 To plb.Parts.Count - 1
            Dim pi = partindex
            Dim PartNdx = plb.Parts(pi)
            ProgressBar1.Value = partindex
            If Not PartNdx.IsEmpty Then
                If Not PartNdx.StartPoint.IsEqual(PartNdx.EndPoint) Then      ' ignore closed ways
                    touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso
                       ((CoIncident(PartNdx.StartPoint, item.value.StartPoint) And CoIncident(PartNdx.EndPoint, item.value.EndPoint)) Or (CoIncident(PartNdx.StartPoint, item.value.EndPoint) And CoIncident(PartNdx.EndPoint, item.value.StartPoint))) AndAlso
                        relations(pi) <> relations(item.index)).ToList
                    If touches.count > 0 Then
                        plb.Parts(partindex).Clear()
                        For Each touch In touches
                            plb.Parts(touch.index).Clear()
                        Next
                    End If
                End If
            End If
        Next
        ' remove deleted edges
        Dim RemovedEdges = 0
        For ndx = plb.Parts.Count - 1 To 0 Step -1
            If plb.Parts(ndx).IsEmpty Then
                plb.Parts.RemoveAt(ndx)
                RemovedEdges += 1
            End If
        Next
        AppendText(TextBox1, $"{RemovedEdges} internal edges removed{vbCrLf}")
        If plb.Parts.Count = 0 Then
            plb = original        ' restore original
            MsgBox("Polygon empty", vbCritical + vbOKOnly, "Polygon empty")
        End If

        While Not finished
            AppendText(TextBox1, $"Pass {pass}, ways {plb.Parts.Count} ")
            PassWatch.Restart()      ' time each pass
            closed = 0        ' count number of ways already closed
            RemovedCount = 0

            ' show progress in joining up ways. For big countires it's slow
            With ProgressBar1
                .Minimum = 0
                .Maximum = plb.Parts.Count - 1
                .Step = 10
            End With
            ' for each part, try to join it to another part
            If plb.Parts.Count > 1 Then     ' must have more than 1 part to join them
                Dim CaseCount = New Integer() {0, 0, 0, 0, 0}
                For PartIndex = 0 To plb.Parts.Count - 1
                    Dim pi As Integer = PartIndex       ' use local variable to fix lambda function issue
                    Dim PartNdx As New Part(plb.Parts(PartIndex))
                    Dim TouchCase As Integer = 0
                    If Not plb.Parts(PartIndex).IsEmpty Then
                        ' part is not already removed
                        ProgressBar1.Value = PartIndex
                        If CoIncident(plb.Parts(PartIndex).StartPoint, plb.Parts(PartIndex).EndPoint) Then
                            closed += 1      ' way is closed already
                        Else
                            Try
                                ' Case 1 - end1 -> begin2
                                touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso CoIncident(PartNdx.EndPoint, item.value.StartPoint)).ToList
                                If touches.count = 1 Then
                                    head = PartNdx
                                    tail = touches(0).value
                                    TouchCase = 1
                                Else
                                    If touches.count > 1 Then Continue For
                                    'Case 2 - begin1 -> end2
                                    touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso CoIncident(PartNdx.StartPoint, item.value.EndPoint)).ToList
                                    If touches.count = 1 Then
                                        head = touches(0).value
                                        tail = PartNdx
                                        TouchCase = 2
                                    Else
                                        If touches.count > 1 Then Continue For
                                        'Case 3 - begin1 -> begin2
                                        touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso CoIncident(PartNdx.StartPoint, item.value.StartPoint)).ToList
                                        If touches.count = 1 Then
                                            ' join both parts together
                                            head = ReversePart(touches(0).value)
                                            tail = PartNdx
                                            TouchCase = 3
                                        Else
                                            If touches.count > 1 Then Continue For
                                            ' Case 4 - end1 -> end2
                                            touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso CoIncident(PartNdx.EndPoint, item.value.EndPoint)).ToList
                                            If touches.count = 1 Then
                                                ' join both parts together
                                                ' reverse 2
                                                head = PartNdx
                                                tail = ReversePart(touches(0).value)
                                                TouchCase = 4
                                            Else Continue For
                                            End If
                                        End If
                                    End If
                                End If

                            Catch ex As Exception
                                MessageBox.Show($"Part {PartIndex} {ex.Message}{vbCrLf}{ex.StackTrace}")
                            End Try
                            If TouchCase > 0 Then
                                CaseCount(TouchCase) += 1
                                ' Replace the current part with the join of 2 touching parts
                                plb.Parts(PartIndex) = JoinParts(head, tail)
                                plb.Parts(touches(0).index).Clear()                     ' make an empty segment
                            End If
                        End If
                    End If
                    Application.DoEvents()        ' yield
                Next
                ' Remove empty segments
                For ndx = plb.Parts.Count - 1 To 0 Step -1
                    If plb.Parts(ndx).IsEmpty Then
                        plb.Parts.RemoveAt(ndx)
                        RemovedCount += 1
                    End If
                Next
                AppendText(TextBox1, "Count of cases ")
                For i = 1 To 4
                    AppendText(TextBox1, $"{i}:{CaseCount(i)} ")
                Next
                AppendText(TextBox1, $" ")
            End If
            AppendText(TextBox1, $"closed {closed} ")
            AppendText(TextBox1, $"{PassWatch.ElapsedMilliseconds / 1000:f1}s{vbCrLf}")
            pass += 1
            If RemovedCount = 0 Then finished = True
        End While

        Dim unclosed = 0
        For PartIndex As Integer = 0 To plb.Parts.Count - 1
            If Not CoIncident(plb.Parts(PartIndex).StartPoint, plb.Parts(PartIndex).EndPoint) Then
                unclosed += 1
                plb.Parts(PartIndex).AddPoint(plb.Parts(PartIndex).StartPoint)        ' close the open polygon
            End If
        Next
        If plb.IsEmpty Then plb = original      ' if plb is now empty, restore original
        AppendText(TextBox1, $"{unclosed} unclosed polygons closed{vbCrLf}")

        Dim poly = New Polygon(plb.Parts)       ' make the polygon
        poly = poly.Simplify    ' make sure all polygons have correct winding direction

        ' We frequently see a situation where we have an outer ring inside another outer ring. The outer ring is usually the administrative boundary
        ' whilst the inner ring is the land boundary. In this case we remove the outer ring.
        Dim OuterRemoval As New List(Of String) From {"Austral Is", "Martinique", "Bonaire", "Christmas Is", "Clipperton Is", "Lakshadweep Is", "Sable Is", "Scarborough Reef", "St Paul Is", "St. Pierre & Miquelon", "Marquesas Is", "Wake Is", "Willis Is"}     ' countires which need outer ring removed
        If OuterRemoval.Contains(country) Then
            '**********************************************
            ' Make a separate polygon for every part so we can use spatial comparisons
            ' Pre calculate all the windings to save time
            plb = New PolylineBuilder(poly.ToPolyline)          ' convert back to polyline so we can edit
            Dim polygons As New List(Of (poly As Polygon, winding As Winding))
            For ndx = 0 To plb.Parts.Count - 1
                Dim poly1 = New Polygon(plb.Parts(ndx))
                Dim wind As Winding
                If PolygonArea(poly1.Parts(0)) < 0 Then wind = Winding.Outer Else wind = Winding.Inner
                polygons.Add((poly1, wind))
            Next
            Dim OutersRemoved As Integer = 0
            For outer = 0 To plb.Parts.Count - 1
                If polygons(outer).winding = Winding.Outer Then        ' it's an outer ring
                    plb.Parts(outer).Clear()      ' erase outer polygon
                    OutersRemoved += 1
                End If
            Next
            If OutersRemoved > 0 Then
                poly = New Polygon(plb.Parts)           ' put polygon back together
                poly = poly.Simplify    ' make sure all polygons have correct winding direction
                AppendText(TextBox1, $"{OutersRemoved} outer boundaries removed{vbCrLf}")
            End If
        End If
        '**********************************************

        'poly = Dissolve(poly)
        If poly.IsEmpty Then
            AppendText(TextBox1, $"WARNING: Simplified polygon is empty. Using original{vbCrLf}")
            poly = New Polygon(original.Parts)
        End If
        Dim polyGeneralized As Polygon = GeneralizeByPart(poly)
        Dim afterParts = polyGeneralized.Parts.Count
        Dim afterPoints = 0
        For Each p In polyGeneralized.Parts
            afterPoints += p.PointCount
        Next
        AppendText(TextBox1, $"Polyline converted to polygon. Before parts={beforeParts} points={beforePoints}, After parts={afterParts} points={afterPoints}. Reduced to {afterPoints / beforePoints * 100:f1}%{vbCrLf}")
        Return polyGeneralized
    End Function
    Function Connectivity(plb As PolylineBuilder)
        ' Do any ways intersect ?
        Dim joins As New List(Of List(Of String)), SelfClosed As Integer = 0

        With ProgressBar1
            .Minimum = 0
            .Maximum = plb.Parts.Count - 1
            .Step = 2
        End With
        For outer = 0 To plb.Parts.Count - 1
            Dim joints As New List(Of String)
            ProgressBar1.Value = outer
            For inner = 0 To plb.Parts.Count - 1
                If outer <> inner Then
                    If CoIncident(plb.Parts(outer).EndPoint, plb.Parts(inner).StartPoint) Then joints.Add($"E->B{inner}")
                    If CoIncident(plb.Parts(outer).EndPoint, plb.Parts(inner).EndPoint) Then joints.Add($"E->E{inner}")
                    If CoIncident(plb.Parts(outer).StartPoint, plb.Parts(inner).StartPoint) Then joints.Add($"B->B{inner}")
                    If CoIncident(plb.Parts(outer).StartPoint, plb.Parts(inner).EndPoint) Then joints.Add($"B->E{inner}")
                End If
            Next
            joins.Add(joints)
            If CoIncident(plb.Parts(outer).StartPoint, plb.Parts(outer).EndPoint) Then SelfClosed += 1
        Next
        ' print results
        AppendText(TextBox1, $"Lines with more than 1 connection are surrounded by **{vbCrLf}")
        For i = 0 To plb.Parts.Count - 1
            If joins(i).Count = 2 Or CoIncident(plb.Parts(i).StartPoint, plb.Parts(i).EndPoint) Then
                AppendText(TextBox1, $"{i} {String.Join(",", joins(i).ToArray)} ")
            Else
                AppendText(TextBox1, $"**{i} {String.Join(",", joins(i).ToArray)}** ")
            End If
            If i Mod 10 = 9 Then AppendText(TextBox1, vbCrLf)
        Next
        AppendText(TextBox1, vbCrLf)

        Dim JoinCounts = New Integer() {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        For Each j In joins
            JoinCounts(j.Count) += 1
        Next
        AppendText(TextBox1, "Join counts ")
        For i = LBound(JoinCounts) To UBound(JoinCounts)
            AppendText(TextBox1, $"{i}={JoinCounts(i)} ")
        Next
        AppendText(TextBox1, vbCrLf)
        AppendText(TextBox1, $"{plb.Parts.Count} total parts{vbCrLf}")
        AppendText(TextBox1, $"{SelfClosed} ways were self closing{vbCrLf}")
    End Function

    Public Function Dissolve(poly As Polygon)
        ' Are there any overlapping polygons? If so, remove them
        Dim intersects
        Dim finished As Boolean = False, intersections As Integer
        Dim PartsBefore As Integer = poly.Parts.Count
        Dim plb As New PolylineBuilder(poly.Parts)      ' convert polygon to polyline so we can edit it

        With ProgressBar1
            .Minimum = 0
            .Maximum = plb.Parts.Count - 1
            .Step = 2
        End With
        While Not finished
            intersections = 0
            For outer = 0 To plb.Parts.Count - 1
                Dim PartNdx = plb.Parts(outer)
                Dim pi = outer
                ProgressBar1.Value = outer
                ' Is there an intersection of parts ?
                intersects = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).Where(Function(item) Not item.value.IsEmpty AndAlso item.index <> pi AndAlso PartNdx.Intersect(item.value) IsNot Nothing).ToList
                If intersects.count > 0 Then
                    ' touch to touch start. Is there also end to end?
                    For Each intersect In intersects
                        plb.Parts(outer) = plb.Parts(outer).Union(plb.Parts(intersect.index))   ' union the two together
                        plb.Parts(intersect.index).Clear()
                        intersections += 1
                    Next
                End If
                ' remove empty parts
                For ndx = plb.Parts.Count - 1 To 0 Step -1
                    If plb.Parts(ndx).IsEmpty Then plb.Parts.RemoveAt(ndx)
                Next
            Next
            If intersections = 0 Then finished = True
        End While
        Dim PartsAfter = plb.Parts.Count
        AppendText(TextBox1, $"Dissolve reduced parts from {PartsBefore} to {PartsAfter}{vbCrLf}")
        Return New Polygon(plb.Parts)
    End Function
    Private Function Repair(ByRef plb As PolylineBuilder)
        ' every Way must have a connection to its begin and one to its end
        ' if not, repair the way by connecting it to the closest other point

        Dim touches As Integer?
        With ProgressBar1
            .Minimum = 0
            .Maximum = plb.Parts.Count - 1
            .Step = 2
        End With
        Dim nodes As New Dictionary(Of Tuple(Of Double, Double), Integer)
        For Each prt In plb.Parts
            Dim startpt As New Tuple(Of Double, Double)(prt.StartPoint.X, prt.StartPoint.Y)
            Dim endpt As New Tuple(Of Double, Double)(prt.EndPoint.X, prt.EndPoint.Y)
            If Not nodes.TryAdd(startpt, 1) Then
                nodes(startpt) += 1
            End If
            If Not nodes.TryAdd(endpt, 1) Then
                nodes(endpt) += 1
            End If
        Next
        For Each n In nodes
            If n.Value <> 2 Then
                AppendText(TextBox1, $"node count= {n.Value}{vbCrLf}")
            End If
        Next
        Exit Function
        ' Remove any zero length ways
        Dim removed As Integer = 0, closed As Integer = 0
        For outer = plb.Parts.Count - 1 To 0 Step -1
            If CoIncident(plb.Parts(outer).StartPoint, plb.Parts(outer).EndPoint) Then closed += 1
            If plb.Parts(outer).PointCount <= 1 Then
                AppendText(TextBox1, $"way {outer} has 0 length{vbCrLf}")
                plb.Parts.RemoveAt(outer)           ' remove the way
                removed += 1
            ElseIf plb.Parts(outer).Count = 1 And CoIncident(plb.Parts(outer).StartPoint, plb.Parts(outer).EndPoint) Then
                AppendText(TextBox1, $"way {outer} has 0 length{vbCrLf}")
                plb.Parts.RemoveAt(outer)           ' remove the way
                removed += 1
            End If
        Next
        AppendText(TextBox1, $"{closed} closed ways found And {removed} empty ways removed{vbCrLf}")

        For outer = 0 To plb.Parts.Count - 1
            Try
                ProgressBar1.Value = outer
                If Not CoIncident(plb.Parts(outer).StartPoint, plb.Parts(outer).EndPoint) Then      ' only test non closed ways
                    Dim pi As Integer = outer           ' local variable needed for lambda function
                    Dim PartNdx As New Part(plb.Parts(outer))          ' local variable needed for lambda function

                    Dim StartJoined As Boolean, FinishJoined As Boolean
                    ' Does start of way touch another ?
                    touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).FirstOrDefault(Function(item) item.index <> pi AndAlso CoIncident(PartNdx.Points(0), item.value.Points(0)))?.index
                    If touches Is Nothing Then touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).FirstOrDefault(Function(item) item.index <> pi AndAlso CoIncident(PartNdx.Points(0), item.value.Points(item.value.Points.Count - 1)))?.index
                    StartJoined = touches IsNot Nothing
                    ' Does finish of way touch another ?
                    touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).FirstOrDefault(Function(item) item.index <> pi AndAlso CoIncident(PartNdx.Points(PartNdx.Points.Count - 1), item.value.Points(0)))?.index
                    If touches Is Nothing Then touches = plb.Parts.Select(Function(prt, i) New With {.value = prt, .index = i}).FirstOrDefault(Function(item) item.index <> pi AndAlso CoIncident(PartNdx.Points(PartNdx.Points.Count - 1), item.value.Points(item.value.Points.Count - 1)))?.index
                    FinishJoined = touches IsNot Nothing
                    If Not StartJoined Or Not FinishJoined Then
                        ' a way is missing a connection. Attempt to repair by replacing broken link point with closest point
                        AppendText(TextBox1, $"Repair required On way {outer}{vbCrLf}")
                        If Not StartJoined Then
                            Dim NearestEnd = plb.Parts.OrderBy(Of Double)(Function(i) Not PartNdx.Equals(i) And GeometryEngine.Distance(PartNdx.Points(0), i.Points(i.Points.Count - 1))).First
                            Dim index = plb.Parts.IndexOf(NearestEnd)
                            AppendText(TextBox1, $"Nearest To start Of {outer} Is End Of {index}, Distance = {GeometryEngine.Distance(PartNdx.Points(0), plb.Parts(index).Points(plb.Parts(index).Points.Count - 1))}{vbCrLf}")
                            index = index
                        End If
                    End If
                End If
            Catch ex As Exception
                MessageBox.Show($"Part {outer} {ex.Message}{vbCrLf}{ex.StackTrace}")
            End Try
        Next
    End Function
    Shared Function CoIncident(a As MapPoint, b As MapPoint) As Boolean
        ' test if points are coincident with small tolerance
        Const tol = 0.000001
        Return a.IsEqual(b)
        'Return Math.Abs(a.X - b.X) <= tol AndAlso Math.Abs(a.Y - b.Y) <= tol AndAlso a.SpatialReference.Wkid = b.SpatialReference.Wkid
    End Function
    Shared Function ReversePart(p As Part) As Part
        ' reverse the order of points in a part
        Dim result As New Part(p.SpatialReference)
        Debug.Assert(Not p.IsEmpty, "ReverseParts - part Is empty")
        For ndx = p.Points.Count - 1 To 0 Step -1
            result.AddPoint(p.Points(ndx))
        Next
        Debug.Assert(p.PointCount = result.PointCount, "Point count Error")
        Return result
    End Function

    Shared Function JoinParts(a As Part, b As Part) As Part
        ' Join parts a and b. Parts a and b face right, and are connected a -> b
        Debug.Assert(Not a.IsEmpty And Not b.IsEmpty, "part Is empty")
        Debug.Assert(CoIncident(a.EndPoint, b.StartPoint), "Parts Not contiguous")
        Dim result As New Part(a.SpatialReference)
        result.AddSegments(a)
        'b.RemoveAt(0)       ' remove duplicate segment at junction
        result.AddSegments(b)
        Debug.Assert(result.SegmentCount = a.SegmentCount + b.SegmentCount, "Error joining")
        Return result
    End Function
    Private Sub MakeKMLByCQZoneToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeKMLByCQZoneToolStripMenuItem.Click
        ' make KML file for specified CQ zone
        KMLbyCQ.ShowDialog()
    End Sub

    Private Sub MakeKMLToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeKMLToolStripMenuItem.Click
        ' Make a KML file of all Entity boundaries
        Dim DXCClist As New List(Of Integer), BoundingBoxes As New List(Of (name As String, box As String))   ' bounding boxes to add at end
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Dim BaseFilename As String = $"{Application.StartupPath}\KML\DXCC Map of the World"
        Using connect As New SqliteConnection(DXCC_DATA),
            kml As New StreamWriter($"{BaseFilename}.kml", False)

            connect.Open()
            ' Create list of DXCC to convert to KML (all)
            sql = connect.CreateCommand
            sql.CommandText = "Select * FROM DXCC WHERE Deleted=0 And geometry Is Not NULL ORDER BY Entity"     ' fetch all geometry
            SQLdr = sql.ExecuteReader
            DXCClist.Clear()
            BoundingBoxes.Clear()

            While SQLdr.Read
                DXCClist.Add(SQLdr("DXCCnum"))
                Dim BoundingBox = SQLdr("bbox")
                If Not IsDBNullorEmpty(BoundingBox) Then
                    BoundingBoxes.Add((name:=SQLdr("Entity"), box:=ParseBox(BoundingBox).ToJson))
                End If
            End While
            SQLdr.Close()
            kml.WriteLine(KMLheader)
            KMLlist(connect, kml, DXCClist)
            If DXCClist.Contains(15) Or DXCClist.Contains(54) Then EUASborder(kml)      ' kml contains EU or AS Russia, then include the boundary
            PrefixFolder(kml)
            GridSquareFolder(connect, kml, DXCClist)
            ZoneFolder(connect, kml)
            IARUFolder(connect, kml)
            TimeZoneFolder(connect, kml)
            BoundingBoxFolder(kml, BoundingBoxes)
            AntarcticFolder(connect, kml)
            kml.WriteLine(KMLfooter)
            kml.Close()
            ' compress to zip file
            System.IO.File.Delete(BaseFilename & ".kmz")
            Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
            zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
            zip.Dispose()
            Dim kmlSize As Long = FileLen(BaseFilename & ".kml")
            Dim kmzSize As Long = FileLen(BaseFilename & ".kmz")
            AppendText(TextBox1, $"KML file {BaseFilename} of {kmlSize / 1024:f0} Kb compressed to {kmzSize / 1024:f0} Kb, i.e. {kmzSize / kmlSize * 100:f0}%{vbCrLf}")
            AppendText(TextBox1, $"Done{vbCrLf}")
        End Using
    End Sub
    Shared Function SQLescape(st As String) As String
        ' escape special characters for SQL
        Return st.Replace("'", "''")
    End Function
    Shared Function DegtoRad(deg As Double) As Double
        ' Convert degrees to radians
        Return deg * Math.PI / 180
    End Function
    Shared Function RadtoDeg(rad As Double) As Double
        ' Convert radians to degrees
        Return rad * 180 / Math.PI
    End Function

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

    Private Async Sub MakeShapefileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeShapefileToolStripMenuItem.Click
        Dim Shapefile = $"{Application.StartupPath}\DXCC.shp"
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, countries As Integer = 0
        Dim sft As New ShapefileFeatureTable(Shapefile)
        Dim fqr As FeatureQueryResult, qp As New QueryParameters
        Await sft.LoadAsync()
        ' Delete all existing features
        With qp
            .WhereClause = "1=1"            ' get all features
            .OrderByFields.Add(New OrderBy("FID", SortOrder.Descending))    ' sort by descending order of FID
            .ReturnGeometry = False
        End With
        fqr = Await sft.QueryFeaturesAsync(qp)
        Dim Removed = fqr.Count             ' features removed
        Await sft.DeleteFeaturesAsync(fqr.ToList)
        sft.Close()
        sft = New ShapefileFeatureTable(Shapefile)
        Await sft.LoadAsync()
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM DXCC WHERE Deleted=0 AND geometry IS NOT NULL ORDER BY Entity"     ' fetch all geometry
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                countries += 1
                Dim feature = sft.CreateFeature
                With feature        ' add all attributes to feature
                    .SetAttributeValue("DXCCnum", CSng(SQLdr("DXCCnum")))
                    .SetAttributeValue("Entity", SQLdr("Entity"))
                    .SetAttributeValue("Continent", SQLdr("Continent"))
                    .SetAttributeValue("prefix", SQLdr("prefix"))
                    .SetAttributeValue("CQ", SQLdr("CQ"))
                    .SetAttributeValue("ITU", SQLdr("ITU"))
                    .SetAttributeValue("IARU", CSng(SQLdr("IARU")))
                    .SetAttributeValue("lat", CDbl(SQLdr("lat")))
                    .SetAttributeValue("lon", CDbl(SQLdr("lon")))
                    .SetAttributeValue("StartDate", CDate(SQLdr("StartDate")))
                    .Geometry = Polygon.FromJson(SQLdr("geometry"))
                End With
                Await sft.AddFeatureAsync(feature)
            End While
        End Using
        sft.Close()
        AppendText(TextBox1, $"Shapefile file created with {Removed} countries removed and {countries} added{vbCrLf}")
    End Sub

    Private Sub EntityReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EntityReportToolStripMenuItem.Click
        ' Create list of entities and state of geometry
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim partcount As Integer, pointcount As Integer, JSONcount As Integer, lines As Integer = 0
        Dim TotalEntity As Integer = 0, TotalParts As Integer = 0, TotalPoints As Integer = 0, TotalQueries As Integer = 0, TotalJSON As Integer = 0
        Dim TotalUnclosed As Integer = 0, TotalNotes As Integer = 0, TotalBbox As Integer = 0
        Dim outer As Integer, inner As Integer, unclosed As Integer

        Using connect As New SqliteConnection(DXCC_DATA),
              html As New StreamWriter($"{Application.StartupPath}\DXCC.html", False)
            connect.Open()
            sql = connect.CreateCommand
            ' Find total entries in report
            sql.CommandText = "select count(*) as Total from DXCC where Deleted=0 and DXCCnum != 999"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            With ProgressBar1
                .Minimum = 0
                .Maximum = SQLdr("Total")
                .Value = 0
            End With
            SQLdr.Close()
            sql.CommandText = "select * from `DXCC` where `Deleted`=0 and DXCCnum != 999 order by `Entity`"
            SQLdr = sql.ExecuteReader
            html.WriteLine("<!DOCTYPE html>")
            html.WriteLine("<style> table td:nth-child(2), td:nth-child(3), td:nth-child(4),td:nth-child(5),td:nth-child(6) {text-align:right;} .red td{color:red;font-weight: bold;}</style>")
            html.WriteLine("<table border=1>")
            html.WriteLine("<tr><th>Entity</th><th>Parts</th><th>Outer</th><th>Inner</th><th>Points</th><th>Unclosed</th><th>JSON</th><th>OSM Overpass QL</th><th>bbox</th><th>Notes</th></tr>")
            While SQLdr.Read
                ProgressBar1.Value += 1
                partcount = 0
                pointcount = 0
                JSONcount = 0
                outer = 0
                inner = 0
                unclosed = 0
                If Not IsDBNullorEmpty(SQLdr("query")) Then TotalQueries += 1
                Dim cls = " class=red"
                If Not IsDBNullorEmpty(SQLdr("geometry")) Then
                    Dim geometry As Polygon = Polygon.FromJson(SQLdr("geometry"))
                    If Not geometry.IsEmpty Then cls = ""
                    JSONcount = SQLdr("geometry").ToString.Length
                    TotalJSON += JSONcount
                    partcount = geometry.Parts.Count

                    For Each prt In geometry.Parts
                        pointcount += prt.PointCount
                        If PolygonArea(prt) < 0 Then outer += 1 Else inner += 1
                        If Not prt.Points.First.IsEqual(prt.Points.Last) Then unclosed += 1
                    Next
                    TotalUnclosed += unclosed
                End If
                TotalEntity += 1
                TotalParts += partcount
                TotalPoints += pointcount
                Dim bbox As String, notes As String
                If IsDBNullorEmpty(SQLdr("bbox")) Then
                    bbox = ""
                Else
                    bbox = $"({SQLdr("bbox")})"
                    TotalBbox += 1
                End If
                If IsDBNullorEmpty(SQLdr("notes")) Then
                    notes = ""
                Else
                    notes = hyperlink(SQLdr("notes"))
                    TotalNotes += 1
                End If
                html.WriteLine($"<tr{cls}><td>{Strings.Replace(SQLdr("Entity"), " ", "&nbsp;")}</td><td>{partcount:n0}</td><td>{outer}</td><td>{inner}</td><td>{pointcount:n0}</td><td>{unclosed}</td><td>{JSONcount:n0}</td><td>{SQLdr("query")}</td><td>{bbox}</td><td>{notes}</td></tr>")
                lines += 1
            End While
            html.WriteLine($"<tr><td>{TotalEntity}</td><td>{TotalParts:n0}</td><td></td><td></td><td>{TotalPoints:n0}</td><td>{TotalUnclosed:n0}</td><td>{TotalJSON:n0}</td><td>{TotalQueries}</td><td>{TotalBbox}</td><td>{TotalNotes}</td></tr>")
            html.WriteLine("</table>")
            AppendText(TextBox1, $"{lines} lines written To html file{vbCrLf}")
        End Using
    End Sub

    Function hyperlink(links As String) As String
        ' convert a list of hyperlinks, separated by semi-colon, and return html hyperlinks separated by <br>
        Dim hyperlinkList As New List(Of String)
        Dim hyperlinks = links.Split(";").ToList
        For ndx = 0 To hyperlinks.Count - 1
            hyperlinks(ndx) = Regex.Replace(hyperlinks(ndx), "[^0-9a-zA-Z\-\:\./]", "")     ' remove noise characters
        Next
        For Each link In hyperlinks
            Dim matches = Regex.Match(link, "^https.*?-(.*)$")
            hyperlinkList.Add($"<a href=""{link}"">{matches.Groups(1)}</a>")
            Dim i = 1
        Next
        Return String.Join(";<br>", hyperlinkList)
    End Function

    Shared Function PolygonArea(polygon As ReadOnlyPart) As Double
        ' Calculate the area of a polygon using the 'Shoelace' or Gauss's formula
        ' https://en.wikipedia.org/wiki/Shoelace_formula
        ' if result < 0 then CW winding (outer), else CCW (inner)
        Contract.Requires(polygon IsNot Nothing AndAlso Not polygon.IsEmpty, "Illegal polygon")
        Dim result As Double = 0
        If polygon.Count > 2 Then           ' ignore degenerate polygon
            For Each s In polygon
                result += s.StartPoint.X * s.EndPoint.Y - s.EndPoint.X * s.StartPoint.Y
            Next
        End If
        Return result / 2
    End Function
    Private Sub ReprocessCountryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReprocessCountryToolStripMenuItem.Click
        ' Reprocess GIS data for a country
        ' 1. Remove any existing geometry
        ' 2. Download and process OSM data
        ' 3. Remake country KML file
        ReprocessCountry.Show()
    End Sub

    Private Sub CheckISO3166ReferencesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CheckISO3166ReferencesToolStripMenuItem.Click
        Checkiso3166references
    End Sub

    Private Sub CountryCollisionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CountryCollisionsToolStripMenuItem.Click
        CountryCollisions
    End Sub

    Private Sub RemoveEmptyGeometryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveEmptyGeometryToolStripMenuItem.Click
        ' Remove any geometry that is empty. This can be the result of an overagressive Generalize operation
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, sqlR As SqliteCommand
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sqlR = connect.CreateCommand
            ' Find how many DXCC have geometry
            sql.CommandText = "SELECT COUNT(*) as Total FROM DXCC WHERE geometry IS NOT NULL"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            With ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = SQLdr("Total")
            End With
            SQLdr.Close()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                ProgressBar1.Value += 1
                Dim poly As Polygon = Polygon.FromJson(SQLdr("geometry"))
                If poly.IsEmpty Then
                    sqlR.CommandText = $"UPDATE DXCC SET geometry=NULL WHERE DXCCnum={SQLdr("DXCCnum")}"
                    sqlR.ExecuteNonQuery()
                    AppendText(TextBox1, $"Removed empty geometry for {SQLdr("Entity")}{vbCrLf}")
                End If
            End While
        End Using
    End Sub

    Private Async Sub UseOSMLandPolygonsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UseOSMLandPolygonsToolStripMenuItem.Click
        ' Retrieve a land polygon given a position
        Dim myQuery = New QueryParameters, sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim EntityList As New List(Of String) From {{"Antarctica"}}

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\Land Polygons\complete\land_polygons.shp")
            AppendText(TextBox1, $"{Features.NumberOfFeatures} land polygons loaded{vbCrLf}")

            For Each entity In EntityList
                sql.CommandText = $"SELECT * FROM DXCC WHERE Entity='{SQLescape(entity)}'"   ' update database
                SQLdr = sql.ExecuteReader()
                SQLdr.Read()
                Dim position As New MapPoint(CDbl(SQLdr("lon")), CDbl(SQLdr("lat")), SpatialReferences.Wgs84)
                SQLdr.Close()
                ' Query to find the polygon at the Entity location
                With myQuery
                    .ReturnGeometry = True
                    .Geometry = position
                    .SpatialRelationship = SpatialRelationship.Intersects
                    .OutSpatialReference = SpatialReferences.Wgs84
                    .MaxFeatures = 1
                End With
                ' Find the polygon, and update the database
                Dim feature = Await Features.QueryFeaturesAsync(myQuery)
                If Not feature.Any Then
                    AppendText(TextBox1, $"No data retrieved for {entity}{vbCrLf}")
                Else
                    AppendText(TextBox1, $"Extracted data for {entity}{vbCrLf}")
                    Dim geometry As Polygon = feature(0).Geometry
                    Dim poly = Generalize(geometry)                   ' convert to JSON
                    sql.CommandText = $"UPDATE DXCC SET geometry='{poly.ToJson}' WHERE Entity='{SQLescape(entity)}'"   ' update database
                    sql.ExecuteNonQuery()
                End If
            Next
        End Using
        AppendText(TextBox1, "Done")
    End Sub

    Shared Function Generalize(poly As Polygon) As Polygon
        ' Generalize a polygon. If generalized out of existence then return original polygon
        Dim distance = GeometryEngine.DistanceGeodetic(New MapPoint(poly.Extent.XMin, poly.Extent.YMin, poly.SpatialReference), New MapPoint(poly.Extent.XMax, poly.Extent.YMax, poly.SpatialReference), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic)
        Dim GeneralizeDistance = Math.Min(CLOSENESS, distance.Distance * 0.01)              ' Generalize to 1 percent of envelope size, or CLOSENESS, whichever is less 
        Dim GeneralizeAngle As Double = RadtoDeg(Math.Asin(GeneralizeDistance / EARTH_RADIUS))
        Dim polyGeneralized As Polygon = GeometryEngine.Generalize(poly, GeneralizeAngle, False)   ' Generalize a polygon to reduce it in size
        If polyGeneralized.IsEmpty Then polyGeneralized = poly       ' If the polygon is generalized To nothing, the original polygon is returned
        Return polyGeneralized
    End Function

    Shared Function GeneralizeByPart(poly As Polygon) As Polygon
        ' Generalizing a Polygon as a single entity often ends up with degenerate or mangled polygons.
        ' This function generalizes by individual parts. It will use a different generalize distance depending on the size of the part.
        ' Large parts will be generalized to CLOSENESS meters.
        ' Small parts will be generalized to 1% of the extent, where 1% of the extent is less than CLOSENESS
        Dim plb As New PolylineBuilder(SpatialReferences.Wgs84)     ' create a polyline builder to collect generalized parts
        Dim distance As GeodeticDistanceResult

        For Each prt In poly.Parts
            Dim PolyPart = New Polygon(prt)         ' create a new polygon from the part
            With PolyPart.Extent
                distance = GeometryEngine.DistanceGeodetic(New MapPoint(.XMin, .YMin, poly.SpatialReference), New MapPoint(.XMax, .YMax, poly.SpatialReference), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic)  ' calculate diagonal distance of the extent
            End With
            Dim GeneralizeDistance = Math.Min(CLOSENESS, distance.Distance * 0.01)              ' Generalize to 1 percent of envelope size, or CLOSENESS, whichever is less 
            Dim GeneralizeAngle As Double = RadtoDeg(Math.Asin(GeneralizeDistance / EARTH_RADIUS))
            Dim polyPartGeneralized As Polygon = GeometryEngine.Generalize(PolyPart, GeneralizeAngle, False)   ' Generalize a polygon to reduce it in size
            For Each p In polyPartGeneralized.Parts
                plb.AddPart(p)      ' add parts (there should only be one) of the generalized polygon
            Next
        Next
        Dim polyGeneralized As New Polygon(plb.Parts)
        If polyGeneralized.IsEmpty Then polyGeneralized = poly       ' If the polygon is generalized to nothing, the original polygon is returned
        Return polyGeneralized
    End Function

    Private Sub MakeKMLAllEntitiesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MakeKMLAllEntitiesToolStripMenuItem.Click
        ' Make individual KML files for all entities
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            With ProgressBar1
                .Minimum = 0
                .Maximum = 340
                .Value = 0
            End With
            sql = connect.CreateCommand
            sql.CommandText = ("SELECT * FROM `DXCC` WHERE `DELETED`=0 AND `DXCCnum`<>999 ORDER BY `Entity`")
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Using kml As New StreamWriter($"{Application.StartupPath}\KML\DXCC_{SQLdr("Entity")}.kml", False)
                    kml.WriteLine(KMLheader)
                    Placemark(connect, kml, SQLdr("DXCCnum"))
                    kml.WriteLine(KMLfooter)
                End Using
            End While
        End Using
    End Sub
    Private Sub InnerRingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InnerRingsToolStripMenuItem.Click
        ' Find countries with inner rings. Some are genuine
        Dim KnownInner As New List(Of String) From {"Argentina", "Australia", "Belarus", "Belgium", "China", "Fed Rep of Germany", "France", "French Polynesia", "Indonesia", "Italy", "Kyrgyzstan", "Mozambique", "Netherlands", "Oman", "Paraguay", "Republic of South Africa", "Serbia", "Spain", "Switzerland", "United Arab Emirates", "Uruguay", "Uzbekistan"}     ' list of entities known to contain genuine inner(s) (enclaves)
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, sqlR As SqliteCommand
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sqlR = connect.CreateCommand
            ' Find how many DXCC have geometry
            sql.CommandText = "SELECT COUNT(*) as Total FROM DXCC WHERE geometry IS NOT NULL"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            With ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = SQLdr("Total")
            End With
            SQLdr.Close()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                ProgressBar1.Value += 1
                If Not KnownInner.Contains(SQLdr("Entity")) Then
                    Dim poly As Polygon = Polygon.FromJson(SQLdr("geometry"))
                    For Each part In poly.Parts
                        If PolygonArea(part) > 0 Then
                            ' calculate centroid
                            Dim x As Double = 0, y As Double = 0, count As Integer = 0
                            For Each pnt In part.Points
                                x += pnt.X
                                y += pnt.Y
                                count += 1
                            Next
                            x /= count
                            y /= count
                            AppendText(TextBox1, $"Inner ring(s): {SQLdr("Entity")} at {y:f6},{x:f6}{vbCrLf}")
                        End If
                    Next

                End If
            End While
        End Using
    End Sub


    Private Sub ImportCQITUZonesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportCQITUZonesToolStripMenuItem.Click
        ImportCQITUZones()
    End Sub

    Private Async Sub ImportTimezonesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportTimezonesToolStripMenuItem.Click
        Await ImportTimeZones()
    End Sub

    Private Sub ImportEUASBorderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportEUASBorderToolStripMenuItem.Click
        ImportEUASBorder()
    End Sub
    Private Async Sub ImportIARURegionsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportIARURegionsToolStripMenuItem.Click
        Await ImportIARURegions()
    End Sub
    Private Async Sub ImportAntarcticBasesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportAntarcticBasesToolStripMenuItem.Click
        Await ImportAntarcticBases()
    End Sub
    Private Async Sub ImportIOTAGroupsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportIOTAGroupsToolStripMenuItem.Click
        Await ImportIOTAGroups()
    End Sub
    Private Async Sub ImportAntarcticaToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Await ImportAntarctica()
    End Sub
    Private Async Sub ImportIOTAIslandsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportIOTAIslandsToolStripMenuItem.Click
        Await ImportIOTAIslands()
    End Sub

    Private Async Sub ImportIOTADXCCMatchesOneIOTAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportIOTADXCCMatchesOneIOTAToolStripMenuItem.Click
        Await ImportIOTADXCCMatchesOneIOTA()
    End Sub
    Private Sub ImportISO3166ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportISO3166ToolStripMenuItem.Click
        ImportISO3166()
    End Sub

    Private Sub IOTACheckToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IOTACheckToolStripMenuItem.Click
        IOTACheck
    End Sub

    Private Sub ParseBoxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ParseBoxToolStripMenuItem.Click
        ' Test the Parse box function
        Dim testcases As New List(Of (input As String, result As Boolean)) From {{("-15,-171,-14,-167", True)}, {("-53,165,-48,179.9", True)},
            {("-53,165,-48", False)}, {("poly:""-12 -160 -12 -135 -27 -135 -17 -160 -12 -160""", True)}}
        For Each testcase In testcases
            Dim result = ParseBox(testcase.input) IsNot Nothing
            AppendText(TextBox1, $"test case {testcase.input}, Expected result {testcase.result}, Passed {testcase.result = result}{vbCrLf}")
        Next
    End Sub
End Class

