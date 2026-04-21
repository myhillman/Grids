Imports System.IO
Imports Esri.ArcGISRuntime
Imports Esri.ArcGISRuntime.Data
Imports Esri.ArcGISRuntime.Geometry
Imports Esri.ArcGISRuntime.Ogc
Imports NetTopologySuite
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Polygonize
Imports NetTopologySuite.Simplify
'Imports NetTopologySuite.IO.Shapefile
Imports Microsoft.Data.Sqlite

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

    Public Async Function UseShapefile() As Task(Of Boolean)

        Dim gf = NtsGeometryServices.Instance.CreateGeometryFactory(4326)
        Dim shpPath = "D:\GIS Data\World countries generalized\World_Countries_Generalized.shp"

        ' Load shapefile
        Dim countries As New List(Of (Name As String, Geom As NetTopologySuite.Geometries.Geometry))

        Using reader As New ShapefileDataReader(shpPath, gf)
            Dim fieldIndex = reader.GetOrdinal("country")
            While reader.Read()
                Dim geom As NetTopologySuite.Geometries.Geometry = reader.Geometry
                Dim name As String = reader.GetString(fieldIndex)
                countries.Add((name, geom))   ' or (name, geom) with your tuple/class
            End While
        End Using

        ' Name normalization table
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

        Dim PrefixesFound As New List(Of String)

        Using connect As New SqliteConnection(DXCC_DATA),
          gridlist As New StreamWriter($"{Application.StartupPath}\GridList.txt", False),
          missing As New StreamWriter($"{Application.StartupPath}\missing.txt", False)

            connect.Open()

            Dim sql = connect.CreateCommand()
            Dim sql1 = connect.CreateCommand()

            For Each c In countries

                Dim name = c.Name
                Dim geom = c.Geom

                ' Normalize name
                Dim newName As String = Nothing
                If NameChanges.TryGetValue(name, newName) Then name = newName

                ' Lookup DXCC entry
                sql.CommandText = $"SELECT * FROM DXCC WHERE Entity='{name.Replace("'", "''")}' AND Deleted=0"
                Dim SQLdr = sql.ExecuteReader()

                If SQLdr.Read() Then

                    Dim prefix = CStr(SQLdr("prefix"))
                    PrefixesFound.Add($"'{prefix}'")

                    Dim DXCC = CInt(SQLdr("DXCCnum"))
                    SQLdr.Close()

                    ' Load production grids
                    Dim ProductionGrids As New List(Of String)
                    sql1.CommandText = $"SELECT * FROM DXCCtoGRID WHERE DXCC={DXCC} ORDER BY GRID"
                    Dim SQLdr1 = sql1.ExecuteReader()
                    While SQLdr1.Read()
                        ProductionGrids.Add(CStr(SQLdr1("GRID")))
                    End While
                    SQLdr1.Close()

                    ' Compute grid squares
                    Dim squares As New List(Of String)

                    Dim env As NetTopologySuite.Geometries.Envelope = geom.EnvelopeInternal

                    Dim topLeftX = Math.Floor(env.MinX / GridSquareX) * GridSquareX
                    Dim topLeftY = Math.Floor(env.MinY / GridSquareY) * GridSquareY
                    Dim bottomRightX = Math.Ceiling(env.MaxX / GridSquareX) * GridSquareX
                    Dim bottomRightY = Math.Ceiling(env.MaxY / GridSquareY) * GridSquareY

                    Dim x = topLeftX
                    While x < bottomRightX
                        Dim y = topLeftY
                        While y < bottomRightY

                            Dim sqEnv As New NetTopologySuite.Geometries.Envelope(x, x + GridSquareX, y, y + GridSquareY)
                            Dim sqPoly = gf.ToGeometry(sqEnv)

                            If sqPoly.Intersects(geom) Then
                                squares.Add(GridSquare(x, y))
                            End If

                            y += GridSquareY
                        End While
                        x += GridSquareX
                    End While

                    squares.Sort()

                    ' Write gridlist entry
                    If squares.Count = 0 Then
                        gridlist.WriteLine($"// Found no grids for {name}")
                    Else
                        Dim csvfields As New List(Of String)
                        If name.Contains(" "c) Then name = $"""{name}"""
                        csvfields.Add(name)
                        csvfields.Add(prefix)

                        Dim locators As New List(Of String)
                        Dim square_list As New List(Of Integer)

                        Dim prevField = squares(0).Substring(0, 2)
                        square_list.Add(CInt(squares(0).Substring(2, 2)))

                        For i = 1 To squares.Count - 1
                            Dim field = squares(i).Substring(0, 2)
                            Dim sqr = CInt(squares(i).Substring(2, 2))

                            If field <> prevField Then
                                locators.Add($"{prevField} {RangeString(square_list)}")
                                square_list.Clear()
                            End If

                            prevField = field
                            square_list.Add(sqr)
                        Next

                        locators.Add($"{prevField} {RangeString(square_list)}")
                        csvfields.Add($"""{String.Join(", ", locators)}""")

                        gridlist.WriteLine(String.Join(", ", csvfields))
                    End If

                    ' Compare production vs world
                    If squares.Count > 0 Then
                        Dim Production = ProductionGrids.Except(squares)
                        Dim World = squares.Except(ProductionGrids)

                        missing.WriteLine(name)
                        If Not Production.Any() AndAlso Not World.Any() Then
                            missing.WriteLine("Lists match")
                        Else
                            missing.WriteLine($" Missing from production {String.Join(",", World)}")
                            missing.WriteLine($" Missing from world {String.Join(",", Production)}")
                        End If
                    End If

                End If

                AppendText(Form1.TextBox1, $"{name}{vbCrLf}")
            Next

            ' Show DXCC not found
            Dim count = 0
            sql.CommandText = $"SELECT * FROM DXCC WHERE prefix NOT IN ({String.Join(",", PrefixesFound)}) AND Deleted=0 ORDER BY Entity"
            Dim SQLdr2 = sql.ExecuteReader()
            While SQLdr2.Read()
                gridlist.WriteLine($"// Not found {SQLdr2("Entity")} {SQLdr2("prefix")}")
                count += 1
            End While
            SQLdr2.Close()
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
            sqldr.Close()
            Dim geom = Await EvaluateExpressionAsync(relation, cache)     ' get geometry for the country using the specified query parameters
            Debug.WriteLine("Cache hits: " & cache.CacheHits)
            Debug.WriteLine("Cache misses: " & cache.CacheMisses)

            ' Convert to GeoJSON
            Dim GeoJson As String = fromntsToGeoJson(geom)

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
                    .Geometry = fromGeoJsonToArcGis(SQLdr("geometry"))
                End With
                Await sft.AddFeatureAsync(feature)
            End While
            SQLdr.Close()
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
                Dim poly As NetTopologySuite.Geometries.Polygon = FromGeoJsonToNTS(SQLdr("geometry"))
                If poly.IsEmpty Then
                    sqlR.CommandText = $"UPDATE DXCC SET geometry=NULL WHERE DXCCnum={SQLdr("DXCCnum")}"
                    sqlR.ExecuteNonQuery()
                    AppendText(Form1.TextBox1, $"Removed empty geometry for {SQLdr("Entity")}{vbCrLf}")
                End If
            End While
            SQLdr.Close()
        End Using
    End Sub

    Public Sub UseOSMLandPolygons()

        Dim gf = NtsGeometryServices.Instance.CreateGeometryFactory(4326)
        Dim shpPath = "D:\GIS Data\Land Polygons\complete\land_polygons.shp"

        ' Load all land polygons into memory
        Dim landPolys As New List(Of NetTopologySuite.Geometries.Geometry)

        Using reader As New ShapefileDataReader(shpPath, gf)
            While reader.Read()
                Dim geom As NetTopologySuite.Geometries.Geometry = reader.Geometry
                landPolys.Add(geom)
            End While
        End Using

        AppendText(Form1.TextBox1, $"{landPolys.Count} land polygons loaded{vbCrLf}")

        Dim entityList As New List(Of String) From {"Antarctica"}

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            For Each entity In entityList

                ' Get DXCC coordinates
                Dim sql = connect.CreateCommand()
                sql.CommandText = $"SELECT * FROM DXCC WHERE Entity='{SQLescape(entity)}'"
                Dim dr = sql.ExecuteReader()

                If Not dr.Read() Then
                    AppendText(Form1.TextBox1, $"DXCC entry not found for {entity}{vbCrLf}")
                    Continue For
                End If

                Dim lon As Double = dr("lon")
                Dim lat As Double = dr("lat")
                dr.Close()

                Dim pt As New Point(lon, lat) With {.SRID = 4326}

                ' Find the land polygon containing the point
                Dim match As NetTopologySuite.Geometries.Geometry = Nothing

                For Each poly In landPolys
                    If poly.Contains(pt) OrElse poly.Intersects(pt) Then
                        match = poly
                        Exit For
                    End If
                Next

                If match Is Nothing Then
                    AppendText(Form1.TextBox1, $"No land polygon found for {entity}{vbCrLf}")
                    Continue For
                End If

                AppendText(Form1.TextBox1, $"Extracted land polygon for {entity}{vbCrLf}")

                ' Optional: generalize (your existing NTS version)
                Dim polyGen As NetTopologySuite.Geometries.Geometry = Generalize(match)

                ' Save to DB as GeoJSON
                Dim writer As New GeoJsonWriter()
                Dim json As String = writer.Write(polyGen)

                sql = connect.CreateCommand()
                sql.CommandText = "UPDATE DXCC SET geometry=@geom WHERE Entity=@entity"
                sql.Parameters.AddWithValue("@geom", json)
                sql.Parameters.AddWithValue("@entity", entity)
                sql.ExecuteNonQuery()

            Next
        End Using

    End Sub
    Public Sub InnerRings()

        Dim KnownInner As New List(Of String) From {
        "Argentina", "Australia", "Belarus", "Belgium", "China", "Fed Rep of Germany",
        "France", "French Polynesia", "Indonesia", "Italy", "Kyrgyzstan", "Mozambique",
        "Netherlands", "Oman", "Paraguay", "Republic of South Africa", "Serbia",
        "Spain", "Switzerland", "United Arab Emirates", "Uruguay", "Uzbekistan"
    }

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            sql = connect.CreateCommand()
            sql.CommandText = "SELECT COUNT(*) AS Total FROM DXCC WHERE geometry IS NOT NULL"
            SQLdr = sql.ExecuteReader()
            SQLdr.Read()

            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = SQLdr("Total")
            End With
            SQLdr.Close()

            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader()

            While SQLdr.Read()
                Form1.ProgressBar1.Value += 1

                Dim entity As String = SQLdr("Entity")
                If KnownInner.Contains(entity) Then Continue While

                ' Load geometry (Polygon or MultiPolygon)
                Dim geom As NetTopologySuite.Geometries.Geometry = FromGeoJsonToNTS(SQLdr("geometry"))

                ' Handle MultiPolygon by iterating each polygon
                Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)

                If TypeOf geom Is NetTopologySuite.Geometries.Polygon Then
                    polys.Add(DirectCast(geom, NetTopologySuite.Geometries.Polygon))
                ElseIf TypeOf geom Is NetTopologySuite.Geometries.MultiPolygon Then
                    Dim mp = DirectCast(geom, NetTopologySuite.Geometries.MultiPolygon)
                    For i = 0 To mp.NumGeometries - 1
                        polys.Add(DirectCast(mp.GetGeometryN(i), NetTopologySuite.Geometries.Polygon))
                    Next
                End If

                ' Check each polygon for interior rings
                For Each poly In polys
                    Dim holeCount = poly.NumInteriorRings

                    If holeCount > 0 Then
                        For i = 0 To holeCount - 1
                            Dim hole As LinearRing = poly.GetInteriorRingN(i)
                            Dim coords = hole.Coordinates

                            ' Compute centroid of the hole
                            Dim cx As Double = 0, cy As Double = 0
                            For Each c In coords
                                cx += c.X
                                cy += c.Y
                            Next
                            cx /= coords.Length
                            cy /= coords.Length

                            AppendText(Form1.TextBox1,
                            $"Inner ring(s): {entity} at {cy:F6},{cx:F6}{vbCrLf}")
                        Next
                    End If
                Next

            End While
            SQLdr.Close()
        End Using
    End Sub

    Public Sub KMLnts()

        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Const BaseFilename = "DXCC Map"

        Using connect As New SqliteConnection(DXCC_DATA),
            kml As New StreamWriter($"{BaseFilename}.kml", False)

            connect.Open()

            ' Write KML header
            kml.WriteLine("<?xml version=""1.0"" encoding=""UTF-8""?>")
            kml.WriteLine("<kml xmlns=""http://www.opengis.net/kml/2.2"">")
            kml.WriteLine("<Document>")
            kml.WriteLine("<name>DXCC Map of the World</name>")
            kml.WriteLine("<Folder>")
            kml.WriteLine("<name>DXCC Entities</name>")

            ' Query DXCC geometries
            sql = connect.CreateCommand()
            sql.CommandText = "SELECT * FROM DXCC WHERE Deleted=0 AND geometry IS NOT NULL ORDER BY Entity"
            sqldr = sql.ExecuteReader()

            While sqldr.Read()

                Dim entity As String = sqldr("Entity")
                Dim geomJson As String = sqldr("geometry")

                ' Load NTS geometry
                Dim geom As NetTopologySuite.Geometries.Geometry = FromGeoJsonToNTS(geomJson)

                ' Write placemark
                kml.WriteLine("<Placemark>")
                kml.WriteLine($"<name>{entity}</name>")

                ' Write polygon(s)
                WriteKmlSingleOrMultiPolygon(geom, kml, 1)

                kml.WriteLine("</Placemark>")
            End While
            sqldr.Close()

            ' Close folder + document
            kml.WriteLine("</Folder>")
            kml.WriteLine("</Document>")
            kml.WriteLine("</kml>")

        End Using

        AppendText(Form1.TextBox1, "Done" & vbCrLf)

    End Sub

    Function Generalize(poly As NetTopologySuite.Geometries.Polygon) As NetTopologySuite.Geometries.Polygon

        Dim timer As New Stopwatch()
        timer.Start()

        ' Count original points
        Dim beforePoints As Integer = poly.NumPoints

        ' Compute envelope diagonal length (approximate size)
        Dim env As NetTopologySuite.Geometries.Envelope = poly.EnvelopeInternal
        Dim dx = env.MaxX - env.MinX
        Dim dy = env.MaxY - env.MinY
        Dim diag = Math.Sqrt(dx * dx + dy * dy)

        ' Generalize to 1% of envelope diagonal, capped by CLOSENESS
        Dim tol = Math.Min(CLOSENESS, diag * 0.01)

        ' Simplify using Douglas-Peucker
        Dim simplified As NetTopologySuite.Geometries.Geometry = DouglasPeuckerSimplifier.Simplify(poly, tol)

        ' If simplification collapses geometry → return original
        If simplified Is Nothing OrElse simplified.IsEmpty OrElse Not TypeOf simplified Is NetTopologySuite.Geometries.Polygon Then
            simplified = poly
        End If

        Dim polyGen As NetTopologySuite.Geometries.Polygon = DirectCast(simplified, NetTopologySuite.Geometries.Polygon)

        ' Count points after simplification
        Dim afterPoints As Integer = polyGen.NumPoints

        timer.Stop()

        AppendText(Form1.TextBox1,
                   $"Generalize: before {beforePoints} pts, after {afterPoints} pts, " &
                   $"reduced to {afterPoints / beforePoints * 100:F1}% [{timer.Elapsed.TotalSeconds:F1}s]{vbCrLf}")

        Return polyGen

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
