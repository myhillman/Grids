Imports System.IO
Imports System.Text.Json
Imports System.Text.RegularExpressions
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports Clipper2Lib
Imports DotNetDBF
Imports Esri.ArcGISRuntime
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Esri.ArcGISRuntime.Ogc
Imports NetTopologySuite
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Union
Imports OsmSharp.IO.PBF

Module Form1side
    Public Const GridFieldX = 20, GridFieldY = 10    ' size of gridfield in degrees in X,Y
    Public Const GridSquareX = 2, GridSquareY = 1    ' size of gridsquare in degrees in X,Y
    Public Const EARTH_RADIUS = 6371000    ' radius of earth in meters
    Const CLOSENESS = 1000     ' distance for Generalize in meters
    Public KMLheader As String
    Public Const KMLfooter = "</Document></kml>"       ' standard footer for kml file
    Public ColourMapping = {"", "red", "green", "blue", "yellow", "cyan", "magenta", "white"}   ' colours for polygons

    Public Sub MakeKMLheader()
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
<tr><td>Grid Squares</td><td>The boundary of every grid square that intersects with the land of an entity is displayed. The 4-character grid square code is displayed in the center of the grid square. This folder is searchable, so you can use it to locate any grid square.</td></tr>
<tr><td>IOTA</td><td>Island Groups for Islands On The Air (IOTA).</td></tr>
<tr><td>CQ Zones</td><td>CQ magazine (now defunct) zones used for Worked All Zones (WAZ) award now administered by ARRL.(Based on data extracted from http://zone-check.eu/.)</td></tr>
<tr><td>ITU Zones</td><td>International Telegraphic Union (ITU) zones. (Based on data extracted from http://zone-check.eu/.)</td></tr>
<tr><td>IARU regions</td><td>International Amateur Radio Union (IARU) regions. (Using data created by Tim Makins (EI8IC))</td></tr>
<tr><td>Timezones</td><td>World time zones</td></tr>
<tr><td>Antarctic bases</td><td>Location and basic details of Antarctic bases.</td></tr>
<tr><td>Bounding boxes</td><td>To extract the entity data it was sometimes necessary to use a bounding box to filter the returned geometry. Open this folder will display those boxes. They are not much use in normal operation, but are used for debugging.</td></tr>
</table>
<br>A huge thank you to Peter Forbes (VK3QI) who used his extensive knowledge and awesome research skills to validate all the data.
<br><br>Data created {Now:R}
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

    Public Delegate Sub SetProgressCallback(pb As System.Windows.Forms.ProgressBar, value As Integer)
    Public Sub UpdateProgressBar(pb As System.Windows.Forms.ProgressBar, value As Integer)
        If pb.InvokeRequired Then
            pb.Invoke(New SetProgressCallback(AddressOf UpdateProgressBar), New Object() {pb, value})
        Else
            pb.Value = value
        End If
        Application.DoEvents()
    End Sub

    Public Delegate Sub InitializeProgressCallback(pb As System.Windows.Forms.ProgressBar, min As Integer, value As Integer, max As Integer, stp As Integer)
    Public Sub InitializeProgressBar(pb As System.Windows.Forms.ProgressBar, min As Integer, value As Integer, max As Integer, stp As Integer)
        If pb.InvokeRequired Then
            pb.Invoke(New InitializeProgressCallback(AddressOf InitializeProgressBar), New Object() {pb, min, value, max, stp})
        Else
            With pb
                .Minimum = min
                .Value = value
                .Maximum = max
                .Step = stp
            End With
        End If
        Application.DoEvents()
    End Sub

    Public Function GridSquare(p As MapPoint) As String
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
    Public Async Function UseShapefile() As Task(Of Boolean)

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
        myQueryFilter.OrderByFields.Add(New OrderBy("country", SortOrder.Ascending))         ' sort the results by name
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
                SQLdr = sql.ExecuteReader
                If SQLdr.HasRows Then
                    SQLdr.Read()
                    prefix = SQLdr("prefix")
                    PrefixesFound.Add($"'{prefix}'")      ' remember prefixes found
                    ProductionGrids.Clear()
                    ' Make a list of grids in production system
                    DXCC = SQLdr("DXCCnum")
                    sql1.CommandText = $"SELECT * FROM DXCCtoGRID WHERE DXCC={DXCC} ORDER BY GRID"
                    SQLdr1 = sql1.ExecuteReader
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
                            Dim SquareBuffered = square.BufferGeodetic(-1000, LinearUnits.Meters)        ' reduce the square by 100 on all sides
                            If SquareBuffered.Intersects(country.Geometry) Then
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
                        Dim previous_field = squares(0).Substring(0, 2)      ' start with first field
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
                AppendText(Form1.TextBox1, $"{name}{vbCrLf}")
            Next
            ' Display prefixes not found
            Dim count = 0
            sql.CommandText = $"SELECT * FROM DXCC WHERE prefix NOT IN ({String.Join(",", PrefixesFound)}) AND Deleted=0 ORDER By Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                gridlist.WriteLine($"// Not found {SQLdr("Entity")} {SQLdr("prefix")}")
                count += 1
            End While
            gridlist.WriteLine($"// Total of {count} DXCC entities not found, {PrefixesFound.Count} found")
        End Using
        Return True
    End Function
    Public Async Function UseOSM() As Task(Of Boolean)
        ' Retrieve country boundaries from OpenStreetMap (OSM)
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim DXCC As Integer, Entity As String
        Dim ProcessList As New List(Of String)        ' list of countries to process

        Using connect As New SqliteConnection(DXCC_DATA)
            ' get all countries where the OSM parameters are known
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM `DXCC` WHERE `geometry` IS NULL AND `query` IS NOT NULL AND `Deleted`=0 ORDER BY `Entity`"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                DXCC = SQLdr("DXCCnum")
                Entity = SQLdr("Entity")
                ProcessList.Add(Entity)
            End While
            SQLdr.Close()     ' close the reader
            For Each item In ProcessList
                Await CreateGrids(connect, item)
            Next
        End Using
        AppendText(Form1.TextBox1, $"Done{vbCrLf}")
        Return True
    End Function
    Public Async Function CreateGrids(connect As SqliteConnection, country As String) As Task(Of Boolean)
        Dim sql As SqliteCommand, sqldr As SqliteDataReader, sqlupd As SqliteCommand

        ' ------------------------------------------------------------
        '  Validate inputs
        ' ------------------------------------------------------------
        If String.IsNullOrWhiteSpace(country) Then
            Throw New ArgumentException("Country name cannot be empty.", NameOf(country))
        End If

        sql = connect.CreateCommand()
        sqlupd = connect.CreateCommand()
        Dim cache = New OsmGeometryCacheService(connect)
        sql.CommandText = $"SELECT * FROM `DXCC` WHERE `Entity`='{SQLescape(country)}' AND `Deleted`=0"
        sqldr = sql.ExecuteReader()
        If sqldr.Read() Then
            AppendText(Form1.TextBox1, $"Creating grids for {sqldr("Entity")}{vbCrLf}")
            Dim DXCCnum = sqldr("DXCCnum")
            Dim Entity = sqldr("Entity")
            Dim relation = SafeStr(sqldr("relation"))
            Dim rule = SafeStr(sqldr("rule"))
            Dim bbox = SafeStr(sqldr("bbox"))
            Dim tolerance_m = SafeStr(sqldr("tolerance_m"))
            Dim geom = Await EvaluateExpressionAsync(relation, cache)     ' get geometry for the country using the specified query parameters
            Debug.WriteLine("Cache hits: " & cache.CacheHits)
            Debug.WriteLine("Cache misses: " & cache.CacheMisses)
            Dim esriGeom = NtsToEsri(geom)     ' convert to esri

            ' Convert to GeoJSON
            Dim GeoJson As String = GeometryToGeoJson(esriGeom)

            ' 4. Update database
            Dim hash = HashText($"{rule}|{bbox}|{tolerance_m}")   ' calculate hash of the parameters used to create the geometry, so we can detect if they change in future
            sqlupd.CommandText = "UPDATE DXCC SET geometry=@json, hash=@hash WHERE DXCCnum=@id"
            sqlupd.Parameters.AddWithValue("@json", GeoJson)
            sqlupd.Parameters.AddWithValue("@hash", hash)
            sqlupd.Parameters.AddWithValue("@id", DXCCnum)
            sqlupd.ExecuteNonQuery()
            AppendText(Form1.TextBox1, "Geometry stored successfully." & vbCrLf)

            ' Make standalone KML file for the country
            Using KML As New StreamWriter(Path.Combine(Application.StartupPath, "KML", $"DXCC_{Entity}.kml"), False)
                KML.WriteLine(KMLheader)
                Using conn As New SqliteConnection(DXCC_DATA)
                    conn.Open()
                    Placemark(conn, KML, DXCCnum)
                End Using
                KML.WriteLine(KMLfooter)
            End Using
        Else
            AppendText(Form1.TextBox1, $"Failed to retrieve geometry for {country} from Natural Earth." & vbCrLf)
            Return False
        End If
        Return True
    End Function
    
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

    Public Async Sub MakeShapefile()
        Dim Shapefile = $"{Application.StartupPath}\DXCC.shp"
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, countries = 0
        Dim sft As New ShapefileFeatureTable(Shapefile)
        Dim fqr As FeatureQueryResult, qp As New QueryParameters
        Await sft.LoadAsync
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
        Await sft.LoadAsync
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM `DXCC` WHERE `Deleted`=0 AND `geometry` IS NOT NULL ORDER BY `Entity`"     ' fetch all geometry
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
                    .Geometry = GeoJsonToGeometry(SQLdr("geometry"))
                End With
                Await sft.AddFeatureAsync(feature)
            End While
        End Using
        sft.Close()
        AppendText(Form1.TextBox1, $"Shapefile file created with {Removed} countries removed and {countries} added{vbCrLf}")
    End Sub
    Public Sub RemoveEmptyGeometry()
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

            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = SQLdr("Total")
            End With
            SQLdr.Close()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                Dim poly As Polygon = GeoJsonToGeometry(SQLdr("geometry"))
                If poly.IsEmpty Then
                    sqlR.CommandText = $"UPDATE DXCC SET geometry=NULL WHERE DXCCnum={SQLdr("DXCCnum")}"
                    sqlR.ExecuteNonQuery()
                    AppendText(Form1.TextBox1, $"Removed empty geometry for {SQLdr("Entity")}{vbCrLf}")
                End If
            End While
        End Using
    End Sub
    Public Async Sub UseOSMLandPolygons()
        ' Retrieve a land polygon given a position
        Dim myQuery = New QueryParameters, sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim EntityList As New List(Of String) From {{"Antarctica"}}

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            Dim Features = Await ShapefileFeatureTable.OpenAsync("D:\GIS Data\Land Polygons\complete\land_polygons.shp")
            AppendText(Form1.TextBox1, $"{Features.NumberOfFeatures} land polygons loaded{vbCrLf}")

            For Each entity In EntityList
                sql.CommandText = $"SELECT * FROM DXCC WHERE Entity='{SQLescape(entity)}'"   ' update database
                SQLdr = sql.ExecuteReader
                SQLdr.Read()
                Dim position As New MapPoint(SQLdr("lon"), SQLdr("lat"), SpatialReferences.Wgs84)
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
                    AppendText(Form1.TextBox1, $"No data retrieved for {entity}{vbCrLf}")
                Else
                    AppendText(Form1.TextBox1, $"Extracted data for {entity}{vbCrLf}")
                    Dim geometry As Polygon = feature(0).Geometry
                    Dim poly = GeneralizeByPart(geometry)                   ' reduce point count
                    poly = FixRingOrientation(poly)
                    sql.CommandText = "UPDATE DXCC SET geometry=@geom WHERE Entity=@entity"
                    sql.Parameters.Clear()
                    sql.Parameters.AddWithValue("@geom", GeometryToGeoJson(poly))
                    sql.Parameters.AddWithValue("@entity", entity)

                    sql.ExecuteNonQuery()

                End If
            Next
        End Using
        AppendText(Form1.TextBox1, "Done")
    End Sub
    Public Sub InnerRings()
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

            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = SQLdr("Total")
            End With
            SQLdr.Close()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                If Not KnownInner.Contains(SQLdr("Entity")) Then
                    Dim poly As Polygon = GeoJsonToGeometry(SQLdr("geometry"))
                    For Each prt In poly.Parts
                        If PolygonArea(prt.Points) > 0 Then
                            ' calculate centroid
                            Dim x As Double = 0, y As Double = 0, count = 0
                            For Each pnt In prt.Points
                                x += pnt.X
                                y += pnt.Y
                                count += 1
                            Next
                            x /= count
                            y /= count
                            AppendText(Form1.TextBox1, $"Inner ring(s): {SQLdr("Entity")} at {y:f6},{x:f6}{vbCrLf}")
                        End If
                    Next

                End If
            End While
        End Using
    End Sub

    Public Async Sub KMLArcgis()
        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Const BaseFilename = "DXCC Map"
        Dim Document = New KmlDocument With {
            .Name = "DXCC Map of the World"
        }
        Dim Folder = New KmlFolder
        Document.ChildNodes.Add(Folder)
        Folder.Name = "DXCC Entities"
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            ' Create list of DXCC to convert to KML (all)
            sql = connect.CreateCommand
            sql.CommandText = "Select * FROM `DXCC` WHERE `Deleted`=0 And `geometry` Is Not NULL ORDER BY `Entity`"     ' fetch all geometry
            sqldr = sql.ExecuteReader
            While sqldr.Read
                ' get the geometry which is in GeoJSON format, convert to Esri geometry, and then to KML geometry
                Dim geom = LoadPolygon(sqldr("geometry"))
                Dim labelPoint = geom.LabelPoint
                Dim kmlgeo As New KmlGeometry(geom, KmlAltitudeMode.ClampToGround, False, True)
                Dim placemark As New KmlPlacemark(kmlgeo)
                placemark.Name = sqldr("Entity")
                Folder.ChildNodes.Add(placemark)
            End While
            sqldr.Close()
            Await Document.SaveAsAsync($"{BaseFilename}.kml")
            ' compress to zip file
            'System.IO.File.Delete(BaseFilename & ".kmz")
            'Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
            'zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
            'zip.Dispose()
            'Dim kmlSize As Long = FileLen(BaseFilename & ".kml")
            'Dim kmzSize As Long = FileLen(BaseFilename & ".kmz")
            'AppendText(form1.textbox1, $"KML file {BaseFilename} of {kmlSize / 1024:f0} Kb compressed to {kmzSize / 1024:f0} Kb, i.e. {kmzSize / kmlSize * 100:f0}%{vbCrLf}")
            AppendText(Form1.TextBox1, $"Done{vbCrLf}")
        End Using
    End Sub
    Public Sub FindIARURegion()
        ' Find IARU region for entities missing one
        Dim IARUlines As New Dictionary(Of Integer, Polyline)       ' IARU boundary lines
        Dim IARU As New Dictionary(Of Integer, Integer()) From {
            {1, {1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 16, 17, 21}},
            {2, {1, 2, 3, 4, 5, 6, 10, 12, 13, 14, 16, 17, 21, 22, 23, 24, 26}},
            {3, {7, 8, 9, 10, 11, 12, 13, 14, 22, 23, 24, 26}}
            }
        Dim IARUPoly As New Dictionary(Of Integer, Polygon), sql As SqliteCommand, sqldr As SqliteDataReader, LineCounts As New Dictionary(Of Integer, Integer)
        Dim IARUregions As New Dictionary(Of Integer, Integer)

        ' Check that each line appears exactly twice in the IARU matrix
        For i = 1 To 26 : LineCounts.Add(i, 0) : Next
        For rgn = 1 To 3
            For i = LBound(IARU(rgn)) To UBound(IARU(rgn)) : LineCounts(IARU(rgn)(i)) += 1 : Next
        Next
        For i = 1 To LineCounts.Count
            If LineCounts(i) <> 2 Then AppendText(Form1.TextBox1, $"Invalid line count of {LineCounts(i)} for line {i}{vbCrLf}")
        Next

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            ' Retrieve all IARU line segments
            sql = connect.CreateCommand
            sql.CommandText = "Select * FROM `IARU`"     ' fetch all geometry
            sqldr = sql.ExecuteReader
            While sqldr.Read
                IARUlines.Add(sqldr("line"), Geometry.FromJson(sqldr("geometry")))
            End While
            sqldr.Close()
            ' Make polygons for all 3 regions
            For rgn = 1 To 3
                AppendText(Form1.TextBox1, $"Processing region {rgn}{vbCrLf}")
                ' make a polygon for this region
                Dim plb As New PolylineBuilder(SpatialReferences.Wgs84)
                For i = LBound(IARU(rgn)) To UBound(IARU(rgn))
                    ' add each line to a region
                    For Each prt In IARUlines(IARU(rgn)(i)).Parts
                        plb.AddPart(prt)
                    Next
                Next
                ' round all X,Y coordinates
                For prt = 0 To plb.Parts.Count - 1
                    For pnt = 0 To plb.Parts(prt).Points.Count - 1
                        plb.Parts(prt).SetPoint(pnt, New MapPoint(Math.Round(plb.Parts(prt).Points(pnt).X, 3), Math.Round(plb.Parts(prt).Points(pnt).Y, 3), SpatialReferences.Wgs84))
                    Next
                Next
                Repair(plb)
                ' TODO: fix this
                'Dim poly As Polygon = CreatePolygon(plb)
                'If rgn = 2 Then poly = ReversePolygon(poly)
                ''poly = poly.Simplify
                'poly = NormalizeCentralMeridian(poly)
                'IARUPoly.Add(rgn, poly)
            Next
            ' make a KML file for each region
            Using kml As New StreamWriter($"{Application.StartupPath}\IARU Regions.kml")
                kml.WriteLine(KMLheader)
                Dim styles = New String() {"", "#red", "#blue", "#yellow"}
                For r = 1 To 3
                    kml.WriteLine($"<Placemark><name>IARU region {r}</name><styleUrl>{styles(r)}</styleUrl>")
                    KMLPolygon(kml, IARUPoly(r), 1, False)
                    kml.WriteLine("</Placemark>")
                Next
                kml.WriteLine(KMLfooter)
            End Using
            ' Resolve IARU region where it is 0
            Dim count = 0, total = 0
            sql.CommandText = "Select * FROM `DXCC` WHERE IARU=0 AND Deleted=0 AND geometry IS NOT NULL ORDER By Entity"     ' fetch all geometry
            sqldr = sql.ExecuteReader
            While sqldr.Read
                total += 1
                Dim geom = GeoJsonToGeometry(sqldr("geometry"))
                Dim regions As New List(Of Integer)
                If IARUPoly(1).Intersects(geom) Then regions.Add(1)
                If Not IARUPoly(2).Intersects(geom) Then regions.Add(2)
                If IARUPoly(3).Intersects(geom) Then regions.Add(3)
                If regions.Count = 1 Then
                    AppendText(Form1.TextBox1, $"{sqldr("Entity")} is in IARU Region {regions(0)}{vbCrLf}")
                    IARUregions.Add(sqldr("DXCCnum"), regions(0))
                    count += 1
                Else
                    AppendText(Form1.TextBox1, $"{sqldr("Entity")} could not be isolated to a single region {regions}{vbCrLf}")
                End If
            End While
            sqldr.Close()
            ' Update data
            For Each entry In IARUregions
                sql.CommandText = $"UPDATE `DXCC` SET IARU={entry.Value} WHERE DXCCnum={entry.Key}"   ' update data
                sql.ExecuteNonQuery()
            Next
            AppendText(Form1.TextBox1, $"Done. {count} out of {total} regions resolved{vbCrLf}")
        End Using
    End Sub
    Sub Repair(ByRef plb As PolylineBuilder)
        ' Repair collection of polylines
        ' All lines must connect to another
        For outer = 0 To plb.Parts.Count - 1
            Dim StartTouch = False, EndTouch = False
            For inner = 0 To plb.Parts.Count - 1
                If outer <> inner Then
                    If CoIncident(plb.Parts(outer).StartPoint, plb.Parts(inner).StartPoint) Or
                   CoIncident(plb.Parts(outer).StartPoint, plb.Parts(inner).EndPoint) Then
                        StartTouch = True
                    End If
                    If CoIncident(plb.Parts(outer).EndPoint, plb.Parts(inner).EndPoint) Or
                       CoIncident(plb.Parts(outer).EndPoint, plb.Parts(inner).StartPoint) Then
                        EndTouch = True
                    End If
                    If StartTouch And EndTouch Then Exit For
                End If
            Next
            If Not StartTouch Then
                AppendText(Form1.TextBox1, $"line {outer} Start does not touch any other line{vbCrLf}Attempting repair ")
                ' Find the nearest matching start/end
                Dim ClosestDistance As Double = Double.MaxValue, ClosestIndex As Integer, ClosestEnd As String, OtherEnd As String, distance As Double
                For other = 0 To plb.Parts.Count - 1
                    If outer <> other Then
                        distance = GeometryEngine.Distance(plb.Parts(outer).StartPoint, plb.Parts(other).StartPoint)      ' calculate distance to all other starts
                        If distance < ClosestDistance Then
                            ClosestDistance = distance
                            ClosestIndex = other
                            OtherEnd = "Start"
                            ClosestEnd = "Start"
                        End If
                        distance = GeometryEngine.Distance(plb.Parts(outer).StartPoint, plb.Parts(other).EndPoint)      ' calculate distance to all other starts
                        If distance < ClosestDistance Then
                            ClosestDistance = distance
                            ClosestIndex = other
                            OtherEnd = "End"
                            ClosestEnd = "Start"
                        End If
                    End If
                Next
                AppendText(Form1.TextBox1, $"Closest to line {outer} {ClosestEnd} is line {ClosestIndex} {OtherEnd} at distance of {ClosestDistance:f3}deg{vbCrLf}")
                If ClosestDistance < 0.1 Then
                    ' Repair
                    Dim node As MapPoint
                    Select Case OtherEnd
                        Case "Start" : node = plb.Parts(ClosestIndex).StartPoint
                        Case "End" : node = plb.Parts(ClosestIndex).EndPoint
                    End Select
                    plb.Parts(outer).SetPoint(0, node)      ' reset start
                    AppendText(Form1.TextBox1, $"Repaired{vbCrLf}")
                End If
            End If

            If Not EndTouch Then
                AppendText(Form1.TextBox1, $"line {outer} End does not touch any other line{vbCrLf}Attempting repair ")
                ' Find the nearest matching start/end
                Dim ClosestDistance As Double = Double.MaxValue, ClosestIndex As Integer, ClosestEnd As String, OtherEnd As String, distance As Double
                For other = 0 To plb.Parts.Count - 1
                    If outer <> other Then
                        distance = GeometryEngine.Distance(plb.Parts(outer).EndPoint, plb.Parts(other).StartPoint)      ' calculate distance to all other starts
                        If distance < ClosestDistance Then
                            ClosestDistance = distance
                            ClosestIndex = other
                            OtherEnd = "Start"
                            ClosestEnd = "End"
                        End If
                        distance = GeometryEngine.Distance(plb.Parts(outer).EndPoint, plb.Parts(other).EndPoint)      ' calculate distance to all other starts
                        If distance < ClosestDistance Then
                            ClosestDistance = distance
                            ClosestIndex = other
                            OtherEnd = "End"
                            ClosestEnd = "End"
                        End If
                    End If
                Next
                AppendText(Form1.TextBox1, $"Closest to line {outer} {ClosestEnd} is line {ClosestIndex} {OtherEnd} at distance of {ClosestDistance:f3}deg{vbCrLf}")
                If ClosestDistance < 0.1 Then
                    ' Repair
                    Dim node As MapPoint
                    Select Case OtherEnd
                        Case "Start" : node = plb.Parts(ClosestIndex).StartPoint
                        Case "End" : node = plb.Parts(ClosestIndex).EndPoint
                    End Select
                    plb.Parts(outer).SetPoint(plb.Parts(outer).Points.Count - 1, node)      ' reset end
                    AppendText(Form1.TextBox1, $"Repaired{vbCrLf}")
                End If
            End If
        Next
    End Sub
    Function Generalize(poly As Polygon) As Polygon
        ' Generalize a polygon. If generalized out of existence then return original polygon
        Dim BeforeParts As Integer, BeforePoints As Integer = 0, AfterParts As Integer, AfterPoints As Integer = 0, timer As New Stopwatch

        timer.Start()
        BeforeParts = poly.Parts.Count
        For Each prt In poly.Parts
            BeforePoints += prt.PointCount
        Next
        Dim distance = GeometryEngine.DistanceGeodetic(New MapPoint(poly.Extent.XMin, poly.Extent.YMin, poly.SpatialReference), New MapPoint(poly.Extent.XMax, poly.Extent.YMax, poly.SpatialReference), LinearUnits.Meters, AngularUnits.Degrees, GeodeticCurveType.Geodesic)
        Dim GeneralizeDistance = Math.Min(CLOSENESS, distance.Distance * 0.01)              ' Generalize to 1 percent of envelope size, or CLOSENESS, whichever is less 
        Dim GeneralizeAngle As Double = RadtoDeg(Math.Asin(GeneralizeDistance / EARTH_RADIUS))
        Dim polyGeneralized As Polygon = GeometryEngine.Generalize(poly, GeneralizeAngle, False)   ' Generalize a polygon to reduce it in size
        If polyGeneralized.IsEmpty Then polyGeneralized = poly       ' If the polygon is generalized To nothing, the original polygon is returned
        AfterParts = polyGeneralized.Parts.Count
        For Each prt In polyGeneralized.Parts
            AfterPoints += prt.PointCount
        Next
        timer.Stop()
        AppendText(Form1.TextBox1, $"Generalize: before parts {BeforeParts} points {BeforePoints}, after parts {AfterParts} points {AfterPoints}, reduced to {AfterPoints / BeforePoints * 100:f1}% [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
        Return polyGeneralized
    End Function

    Function GeneralizeByPart(poly As Polygon) As Polygon
        ' Generalizing a Polygon as a single entity often ends up with degenerate or mangled polygons.
        ' This function generalizes by individual parts. It will use a different generalize distance depending on the size of the part.
        ' Large parts will be generalized to CLOSENESS meters.
        ' Small parts will be generalized to 1% of the extent, where 1% of the extent is less than CLOSENESS
        Dim plb As New PolylineBuilder(SpatialReferences.Wgs84)     ' create a polyline builder to collect generalized parts
        Dim distance As GeodeticDistanceResult, timer As New Stopwatch
        Dim BeforePoints As Integer = 0, AfterPoints As Integer = 0, BeforeParts As Integer, AfterParts As Integer

        timer.Start()
        BeforeParts = poly.Parts.Count
        For Each prt In poly.Parts
            BeforePoints += prt.PointCount
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
        ' Close all polygons
        For ndx = 0 To plb.Parts.Count - 1
            If Not CoIncident(plb.Parts(ndx).StartPoint, plb.Parts(ndx).EndPoint) Then plb.Parts(ndx).AddPoint(plb.Parts(ndx).StartPoint)
            AfterPoints += plb.Parts(ndx).PointCount
        Next
        Dim polyGeneralized As New Polygon(plb.Parts)
        If polyGeneralized.IsEmpty Then polyGeneralized = poly       ' If the polygon is generalized to nothing, the original polygon is returned
        AfterParts = polyGeneralized.Parts.Count
        timer.Stop()
        AppendText(Form1.TextBox1, $"GeneralizeByPart: before parts {BeforeParts} points {BeforePoints}, after parts {AfterParts} points {AfterPoints}, reduced to {AfterPoints / BeforePoints * 100:f1}% [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
        Return polyGeneralized
    End Function

    Public Sub ParseBox()
        ' Test the Parse box function
        Dim testcases As New List(Of (input As String, result As Boolean)) From {{("-15,-171,-14,-167", True)}, {("-53,165,-48,179.9", True)},
            {("-53,165,-48", False)}, {("poly:""-12 -160 -12 -135 -27 -135 -17 -160 -12 -160""", True)}}
        For Each testcase In testcases
            Dim result = ParseBboxorpoly(testcase.input) IsNot Nothing
            AppendText(Form1.TextBox1, $"test case {testcase.input}, Expected result {testcase.result}, Passed {testcase.result = result}{vbCrLf}")
        Next
    End Sub
    Public Function Between(value As Double, low As Double, high As Double) As Boolean
        If low > high Then
            Throw New ArgumentOutOfRangeException(NameOf(low), $"Lower bound ({low}) must be less than or equal to upper bound ({high}).")
        End If

        Return value >= low AndAlso value <= high
    End Function
End Module
