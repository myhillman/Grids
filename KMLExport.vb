Imports System.IO
Imports System.IO.Compression
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Xml
Imports System.Xml.XPath
Imports Esri.ArcGISRuntime.Geometry
Imports Esri.ArcGISRuntime.Ogc
Imports Microsoft.Data.Sqlite

Module KMLExport

    Const CoordsPerLine = 10          ' coordinates per line
    Private prefixes As New List(Of (p As MapPoint, pfx As String))         ' prefixes for prefix folder
    Sub MakeKMLAllEntities()
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            With Form1.ProgressBar1
                .Minimum = 0
                .Maximum = 340
                .Value = 0
            End With
            sql = connect.CreateCommand
            sql.CommandText = ("SELECT * FROM `DXCC` WHERE `DELETED`=0 AND `DXCCnum`<>999 ORDER BY `Entity`")
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Using kml As New StreamWriter($"{Application.StartupPath}\KML\DXCC_{SQLdr("Entity")}.kml", False)
                    kml.WriteLine(Form1.KMLheader)
                    Placemark(connect, kml, SQLdr("DXCCnum"))
                    kml.WriteLine(Form1.KMLfooter)
                End Using
            End While
        End Using
    End Sub

    Sub MakeKML()
        ' Make a KML file of all Entity boundaries
        Dim DXCClist As New List(Of Integer), BoundingBoxes As New List(Of (name As String, box As String))   ' bounding boxes to add at end
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Dim BaseFilename As String = $"{Application.StartupPath}\KML\DXCC Map of the World"
        Using connect As New SqliteConnection(Form1.DXCC_DATA),
            kml As New StreamWriter($"{BaseFilename}.kml", False)

            connect.Open()
            ' Create list of DXCC to convert to KML (all)
            sql = connect.CreateCommand
            sql.CommandText = "Select * FROM `DXCC` WHERE `Deleted`=0 And `geometry` Is Not NULL ORDER BY `Entity`"     ' fetch all geometry
            SQLdr = sql.ExecuteReader
            DXCClist.Clear()
            BoundingBoxes.Clear()

            While SQLdr.Read
                DXCClist.Add(SQLdr("DXCCnum"))
                Dim BoundingBox = SQLdr("bbox")
                If Not Form1.IsDBNullorEmpty(BoundingBox) Then
                    Dim box = Form1.ParseBox(BoundingBox)
                    If box.Parts.Count > 1 Then
                        Form1.AppendText(Form1.TextBox1, $"{SQLdr("Entity")} crosses anti-meridian{vbCrLf}")
                    End If
                    BoundingBoxes.Add((name:=SQLdr("Entity"), box:=box.ToJson))
                End If
            End While
            SQLdr.Close()
            kml.WriteLine(Form1.KMLheader)
            KMLlist(connect, kml, DXCClist)
            If DXCClist.Contains(15) Or DXCClist.Contains(54) Then EUASborder(kml)      ' kml contains EU or AS Russia, then include the boundary
            PrefixFolder(kml)
            GridSquareFolder(connect, kml, DXCClist)
            IOTAFolder(connect, kml)
            ZoneFolder(connect, kml)
            IARUFolder(connect, kml)
            TimeZoneFolder(connect, kml)
            BoundingBoxFolder(kml, BoundingBoxes)
            AntarcticFolder(connect, kml)
            kml.WriteLine(Form1.KMLfooter)
            kml.Close()
            ' compress to zip file
            System.IO.File.Delete(BaseFilename & ".kmz")
            Dim zip As ZipArchive = ZipFile.Open(BaseFilename & ".kmz", ZipArchiveMode.Create)    ' create new archive file
            zip.CreateEntryFromFile(BaseFilename & ".kml", "doc.kml", CompressionLevel.Optimal)   ' compress output file
            zip.Dispose()
            Dim kmlSize As Long = FileLen(BaseFilename & ".kml")
            Dim kmzSize As Long = FileLen(BaseFilename & ".kmz")
            Form1.AppendText(Form1.TextBox1, $"KML file {BaseFilename} of {kmlSize / 1024:f0} Kb compressed to {kmzSize / 1024:f0} Kb, i.e. {kmzSize / kmlSize * 100:f0}%{vbCrLf}")
            Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
        End Using
    End Sub
    Sub KMLlist(connect As SqliteConnection, kml As StreamWriter, DXCClist As List(Of Integer))
        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = DXCClist.Count
        End With
        prefixes.Clear()
        Form1.AppendText(Form1.TextBox1, $"Making KML for {DXCClist.Count} Entities{vbCrLf}")
        kml.WriteLine("<Folder><name>DXCC Entities</name><open>0</open><description>Boundaries of DXCC entities</description>")
        For Each dxcc In DXCClist
            Placemark(connect, kml, dxcc)
        Next
        kml.WriteLine("</Folder>")
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
    Sub Placemark(connect As SqliteConnection, kml As StreamWriter, dxcc As Integer)
        ' create KML for one placemark
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim tagStack As New Stack(Of String)    ' FIFO for end tags

        sql = connect.CreateCommand
        sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={dxcc}"     ' fetch all geometry
        SQLdr = sql.ExecuteReader
        SQLdr.Read()
        Form1.ProgressBar1.Value += 1
        If IsDBNull(SQLdr("geometry")) Then
            ' There is no geometry to convert
            Dim errMsg = $"There is no geometry for {SQLdr("Entity")} to convert to KML"
            Form1.AppendText(Form1.TextBox1, $"{errMsg}{vbCrLf}")
            kml.WriteLine(errMsg)
            Exit Sub
        End If
        Dim boundaries As Polygon = Polygon.FromJson(SQLdr("geometry"))         ' retrieve geometry
        Form1.AppendText(Form1.TextBox1, $"Making KML for {SQLdr("Entity")}{vbCrLf}")
        kml.WriteLine($"<Placemark><styleUrl>#boundary_{SQLdr("colour")}</styleUrl>")
        kml.WriteLine($"<name>{KMLescape(SQLdr("Entity"))} ({SQLdr("prefix")})</name>")
        ' put the label at the centroid of the geometry
        Dim labelpoint = boundaries.LabelPoint
        kml.Write($"<Point>")
        KMLcoordinates(kml, labelpoint, 0)
        kml.WriteLine("</Point>")
        prefixes.Add((labelpoint, SQLdr("prefix")))     ' for the prefix folder
        kml.WriteLine("<ExtendedData>")
        kml.WriteLine($"<Data name=""Entity""><value>{KMLescape(SQLdr("Entity"))}</value></Data>")
        kml.WriteLine($"<Data name= ""DXCC number""><value>{SQLdr("DXCCnum")}</value></Data>")
        kml.WriteLine($"<Data name=""Prefix""><value>{SQLdr("prefix")}</value></Data>")
        kml.WriteLine($"<Data name=""CQ Zone""><value>{SQLdr("CQ")}</value></Data>")
        kml.WriteLine($"<Data name=""ITU Zone""><value>{SQLdr("ITU")}</value></Data>")
        kml.WriteLine($"<Data name=""IARU Region""><value>{SQLdr("IARU")}</value></Data>")
        kml.WriteLine($"<Data name=""Continent""><value>{SQLdr("Continent")}</value></Data>")
        kml.WriteLine($"<Data name=""Start Date""><value>{SQLdr("StartDate")}</value></Data>")
        kml.WriteLine($"<Data name=""lat""><value>{SQLdr("lat"):f3}</value></Data>")
        kml.WriteLine($"<Data name=""lon""><value>{SQLdr("lon"):f3}</value></Data>")
        If Not IsDBNull(SQLdr("query")) Then
            Dim QL = SQLdr("query")
            If Not IsDBNull(SQLdr("bbox")) Then QL &= $"({SQLdr("bbox")})"      ' add bounding box if any
            kml.WriteLine($"<Data name=""OSM query""><value>{KMLescape(QL)}</value></Data>")
        End If
        If Not IsDBNull(SQLdr("notes")) Then
            kml.WriteLine($"<Data name=""Notes""><value><![CDATA[{Form1.Hyperlink(SQLdr("notes"))}]]></value></Data>")
        End If
        kml.WriteLine("</ExtendedData>")
        KMLPolygon(kml, boundaries, 3, False)
        kml.WriteLine("</Placemark>")
    End Sub
    Function KMLescape(st As String) As String
        ' escape special characters for KML
        Dim escapes As New Dictionary(Of String, String) From {
            {"<", "&lt;"},
            {">", "&gt;"},
            {"&", "&amp;"},
            {"""", "&quot;"},
            {"\'", "&apos;"}
            }

        Dim result As String
        result = st
        For Each s In escapes
            result = result.Replace(s.Key, s.Value)
        Next
        Return result
    End Function
    Sub PrefixFolder(kml As StreamWriter)
        ' Create the Prefix folder
        Dim timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, $"Making KML for {prefixes.Count} prefixes.")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Prefixes</name>")
        kml.WriteLine("<description>A folder containing all prefix labels</description>")
        kml.WriteLine("<Style id=""prefix""><LabelStyle><color>ff00ffff</color><scale>5</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle></Style>")    ' empty IconStyle required to display labels correctly (don't know why)
        prefixes = prefixes.OrderBy(Function(a) a.pfx).ToList        ' sort prefix
        For Each prefix In prefixes
            kml.Write($"<Placemark><name>{prefix.pfx}</name><styleUrl>#prefix</styleUrl><Point>")
            KMLcoordinates(kml, prefix.p, 0)
            kml.WriteLine("</Point></Placemark>")
        Next
        kml.WriteLine("</Folder>")
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Sub BoundingBoxFolder(kml As StreamWriter, BoundingBoxes As List(Of (name As String, box As String)))
        ' Create the Bounding Box folder
        Const DensifyDegrees = 10        ' ensure lines have resolution so the follow lat/lon lines
        Dim timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, $"Making KML for {BoundingBoxes.Count} bounding boxes.")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Bounding Boxes</name>")
        kml.WriteLine("<description>Bounding boxes sometimes necessary to filter query. For debugging only.</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""bbox""><PolyStyle><color>7f7f7f7f</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>ffffffff</color><width>2</width></LineStyle></Style>")
        For Each BoundingBox In BoundingBoxes
            ' Display bounding box 
            kml.WriteLine($"<Placemark><name>{KMLescape(BoundingBox.name)}</name><visibility>0</visibility><styleUrl>#bbox</styleUrl>")
            Dim Polybox As Polygon = Geometry.FromJson(BoundingBox.box)        ' convert the json to geometry
            Polybox = Polybox.Densify(DensifyDegrees)       ' Densify the bounding box so that it clearly follows lat/lon lines better. 
            KMLPolygon(kml, Polybox, 1, False)
            kml.WriteLine("</Placemark>")
        Next
        kml.WriteLine("</Folder>")
        timer.Stop()
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub KMLLineString(kml As StreamWriter, linestring As Polyline, digits As Integer)
        ' Write a polyline to the kml file
        If linestring.Parts.Count > 1 Then kml.WriteLine("<MultiGeometry>")
        For Each ls In linestring.Parts
            kml.WriteLine("<LineString><tessellate>1</tessellate>")
            KMLcoordinates(kml, ls.Points.ToList, digits)
            kml.WriteLine("</LineString>")
        Next
        If linestring.Parts.Count > 1 Then kml.WriteLine("</MultiGeometry>")
    End Sub

    Sub KMLcoordinates(kml As StreamWriter, point As MapPoint, digits As Integer)
        Dim list As New List(Of MapPoint) From {point}
        KMLcoordinates(kml, list, digits)
    End Sub
    Sub KMLcoordinates(kml As StreamWriter, points As List(Of MapPoint), digits As Integer)
        ' produce the <coordinates>..</coordinates> KML for a list of points, using digits level of precision, and 10 coordinates per line
        Debug.Assert(points.Count > 0, "part is empty")
        Debug.Assert(points(0).SpatialReference.Equals(SpatialReferences.Wgs84), "Spatial reference not WGS84")
        Debug.Assert(digits >= 0 And digits <= 8, $"bad digits value {digits}")
        Const BlockSize = 10      ' max coordinates per line
        kml.Write($"<coordinates>")
        If points.Count > BlockSize Then kml.WriteLine()        ' small blocks fit on same line as <coordinates>
        Dim index As Integer = 0
        Dim last As Integer
        Do
            last = index + BlockSize - 1
            last = Math.Min(last, points.Count - 1)      ' last point in this block
            For ndx = index To last
                Dim X = points(ndx).X.ToString($"f{digits}")
                Dim Y = points(ndx).Y.ToString($"f{digits}")
                kml.Write($"{X},{Y}")
                If ndx < last Then kml.Write(" ")
            Next
            If points.Count > BlockSize Then kml.WriteLine()    ' end of block
            index += BlockSize       ' next block
        Loop Until index >= points.Count    ' all points done
        kml.WriteLine("</coordinates>")
    End Sub
    Sub KMLMultipoint(kml As StreamWriter, linestring As Multipoint, digits As Integer)
        ' Write a multipoint to the kml file
        kml.WriteLine("<LineString><tessellate>1</tessellate>")
        KMLcoordinates(kml, linestring.Points.ToList, digits)
        kml.WriteLine("</LineString>")
    End Sub
    Sub KMLPolygon(kml As StreamWriter, poly As Polygon, digits As Integer, MultiGeometryOpen As Boolean)
        ' Write a polygon to the kml file
        Dim tagstack As New Stack(Of String)
        If poly.Parts.Count > 1 And Not MultiGeometryOpen Then kml.WriteLine("<MultiGeometry>")
        For Each Prt In poly.Parts
            Dim area = Form1.PolygonArea(Prt)     ' determine inner or outer by testing polygon area
            If area > 0 Then
                ' It's an inner boundary
                kml.WriteLine("<innerBoundaryIs>")
                tagstack.Push("</innerBoundaryIs>")
            Else
                ' It's an outer boundary. Finish previous polygon, if any, and start a new one
                While (tagstack.Count > 0)
                    kml.WriteLine(tagstack.Pop)     ' empty the stack
                End While
                kml.WriteLine("<Polygon><tessellate>1</tessellate><outerBoundaryIs>")    ' open new one
                tagstack.Push("</Polygon>")         ' push end tag
                tagstack.Push("</outerBoundaryIs>")
            End If
            ' Now write the boundary (inner or outer)
            kml.WriteLine("<LinearRing>")
            ' Polygons are not normally closed. If you render them as is, you will get a filled polygon, but there will be 1 boundary segment missing.
            ' This looks ugly, so close the polygon so the boundary looks complete
            Dim pnts = Prt.Points.ToList
            If Not pnts(0).IsEqual(pnts(pnts.Count - 1)) Then pnts.Add(pnts(0))     ' if not closed, then close it
            KMLcoordinates(kml, pnts, digits)
            kml.WriteLine("</LinearRing>")
            kml.WriteLine(tagstack.Pop)    ' close of boundaryIs
        Next
        ' Close any open polygon
        While (tagstack.Count > 0)
            kml.WriteLine(tagstack.Pop)
        End While
        If poly.Parts.Count > 1 And Not MultiGeometryOpen Then kml.WriteLine("</MultiGeometry>")
    End Sub

    Function CrossesAntiMeridian(g As Geometry) As Boolean
        ' returns true if geometry crosses anti-meridian
        Return Math.Abs(g.Extent.XMax - g.Extent.XMin) > 300
    End Function
    Function NormalizeCentralMeridian(PolyBox As Polygon) As Polygon
        ' Normalize a polygon if it crosses the anti-meridian
        If CrossesAntiMeridian(PolyBox) Then       '  crosses anti-meridian
            Dim plb As New PolygonBuilder(PolyBox)        ' deconstruct polygon to polygon builder
            ' Add 360 to any negative X value in any part
            For prt = 0 To plb.Parts.Count - 1
                For pnt = 0 To plb.Parts(prt).Points.Count - 1
                    If plb.Parts(prt).Points(pnt).X < 0 Then
                        plb.Parts(prt).SetPoint(pnt, New MapPoint(plb.Parts(prt).Points(pnt).X + 360, plb.Parts(prt).Points(pnt).Y, plb.SpatialReference))   ' add 360 to X
                    End If
                Next
            Next
            Dim poly = plb.ToGeometry               ' reconstruct geometry
            Dim Normalized As Polygon = poly.NormalizeCentralMeridian   ' split geometry into parts that don't cross the antimeridian
            Normalized = Normalized.Simplify        ' Normalizing seems to produce a non simple result
            Return Normalized
        Else
            Return PolyBox       ' doesn't cross anti-meridian
        End If
    End Function
    Function NormalizeCentralMeridian(PolyBox As Polyline) As Polyline
        ' Normalize a polyline if it crosses the anti-meridian.
        Dim count As Integer = 0

        If CrossesAntiMeridian(PolyBox) Then       '  crosses anti-meridian
            With PolyBox.Extent
                Dim plb = New PolylineBuilder(PolyBox)      ' deconstruct polyline to polyline builder
                ' Add 360 to any negative X value in any part
                For prt = 0 To plb.Parts.Count - 1
                    For pnt = 0 To plb.Parts(prt).Points.Count - 1
                        If plb.Parts(prt).Points(pnt).X < 0 Then
                            count += 1
                            plb.Parts(prt).SetPoint(pnt, New MapPoint(plb.Parts(prt).Points(pnt).X + 360, plb.Parts(prt).Points(pnt).Y, plb.SpatialReference))   ' add 360 to X
                        End If
                    Next
                Next
                Dim Normalized = plb.ToGeometry       ' reconstruct geometry
                Normalized = Normalized.NormalizeCentralMeridian
                Return Normalized
            End With
        Else
            Return PolyBox       ' doesn't cross anti-meridian
        End If

    End Function

    Const EPS = 0.000001      ' a very small number (epsilon)
    Sub GridSquareFolder(connect As SqliteConnection, kml As StreamWriter, DXCClist As List(Of Integer))
        ' Create the Grid Square folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch, Ocean As Integer = 0
        Dim GridSquares As New SortedDictionary(Of String, Envelope)        ' list of gridsquares and their envelope
        Dim Land As New Dictionary(Of String, Integer?)        ' list of gridsquares on land. Don't draw GS if not on land
        Dim Drawn As New Dictionary(Of String, Integer?)       ' list of squares already drawn. Don't draw GS more than once

        timer.Start()
        sql = connect.CreateCommand
        ' Load list of gridsquare known to be land
        sql.CommandText = "SELECT * FROM LAND"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            Land.Add(SQLdr("gridsquare"), Nothing)
        End While
        SQLdr.Close()

        ' make the folder
        Form1.AppendText(Form1.TextBox1, $"Making KML for Grid Square folder ")
        Application.DoEvents()
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Grid Squares</name>")
        kml.WriteLine("<description>Gridsquare overlay</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""grid""><LabelStyle><color>9Fffff00</color><scale>2</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><PolyStyle><color>9Fffff00</color><fill>0</fill><outline>1</outline></PolyStyle><LineStyle><color>9Fffff00</color><width>2</width></LineStyle></Style>")

        For Each DXCC In DXCClist
            GridSquares.Clear()
            sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={DXCC}"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            Dim geometry = Polygon.FromJson(SQLdr("geometry"))          ' get the geometry
            Dim Entity = SQLdr("Entity")
            SQLdr.Close()
            Dim extent = geometry.Extent
            ' Create an envelope which exactly covers all the possible grid squares
            Dim GridExtent As New Envelope(Math.Floor(extent.XMin / Form1.GridSquareX) * Form1.GridSquareX, Math.Floor(extent.YMin / Form1.GridSquareY) * Form1.GridSquareY, Math.Ceiling(extent.XMax / Form1.GridSquareX) * Form1.GridSquareX, Math.Ceiling(extent.YMax / Form1.GridSquareY) * Form1.GridSquareY, SpatialReferences.Wgs84)
            ' test every gridsquare with the grid extent
            Dim X = GridExtent.XMin
            While X < GridExtent.XMax
                Dim Y = GridExtent.YMin
                While Y < GridExtent.YMax
                    Dim GridSq As New Envelope(X, Y, X + Form1.GridSquareX - EPS, Y + Form1.GridSquareY - EPS, SpatialReferences.Wgs84)
                    If GridSq.Intersects(geometry) Then
                        Dim gs = Form1.GridSquare(New MapPoint(X, Y, SpatialReferences.Wgs84))
                        If Land.ContainsKey(gs) Then
                            GridSquares.Add(gs, GridSq)
                        Else
                            Ocean += 1      ' gridsquare is not on land
                        End If
                    End If
                    Y += Form1.GridSquareY
                End While
                X += Form1.GridSquareX
            End While
            kml.WriteLine($"<Folder><name>{KMLescape(Entity)}</name>")
            For Each sq In GridSquares
                With sq.Value
                    Dim lb As New MapPoint(.XMin + .Width / 2, .YMin + .Height / 2, SpatialReferences.Wgs84)    ' label point is center of square
                    kml.WriteLine($"<Placemark><name>{sq.Key}</name><styleUrl>#grid</styleUrl><visibility>0</visibility>")
                    If Not Drawn.ContainsKey(sq.Key) Then
                        kml.WriteLine("<MultiGeometry>")
                        kml.Write($"<Point>")
                        KMLcoordinates(kml, lb, 1)
                        kml.WriteLine("</Point>")
                        Dim mpb As New MultipointBuilder(SpatialReferences.Wgs84)        ' linestring containing 2,3 or 4 sides
                        ' Avoid drawing each side of a gridsquare twice. This will reduce a gridsquare from 5 to 4 or 3 points
                        ' If there is a grid square to the left, don't draw the left boundary
                        ' If there is a grid square above, don't draw the top boundary
                        Dim AboveGS = Form1.GridSquare(New MapPoint(.XMin, .YMin + Form1.GridSquareY, SpatialReferences.Wgs84))       ' grid square above
                        If Not GridSquares.ContainsKey(AboveGS) Then mpb.Points.Add(New MapPoint(.XMin, .YMax))       ' first point
                        mpb.Points.Add(New MapPoint(.XMax, .YMax))
                        mpb.Points.Add(New MapPoint(.XMax, .YMin))
                        mpb.Points.Add(New MapPoint(.XMin, .YMin))
                        Dim LeftGS = Form1.GridSquare(New MapPoint((.XMin - Form1.GridSquareX) Mod 180, .YMin, SpatialReferences.Wgs84))       ' grid square to the left
                        If Not GridSquares.ContainsKey(LeftGS) Then mpb.Points.Add(New MapPoint(.XMin, .YMax))       ' last point
                        Dim mp = mpb.ToGeometry            ' convert to multipoint
                        KMLMultipoint(kml, mp, 0)
                        kml.WriteLine("</MultiGeometry>")
                        Drawn.Add(sq.Key, Nothing)    ' remember we have drawn this square already
                    End If
                End With
                kml.WriteLine($"</Placemark>")
            Next
            kml.WriteLine("</Folder>")
        Next
        kml.WriteLine("</Folder>")
        timer.Stop()
        Form1.AppendText(Form1.TextBox1, $" {Ocean} GS eliminated, {Drawn.Count:n0} Unique grid squares [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub IOTAFolder(connect As SqliteConnection, kml As StreamWriter)
        ' Create the Grid Square folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch
        Dim sql1 As SqliteCommand, SQLdr1 As SqliteDataReader, count As Integer = 0

        timer.Start()
        sql = connect.CreateCommand
        sql1 = connect.CreateCommand
        ' make the folder
        Form1.AppendText(Form1.TextBox1, $"Making KML for IOTA folder ")
        Application.DoEvents()
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>IOTA</name>")
        kml.WriteLine("<description>IOTA overlay</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""iota_normal""><LabelStyle><color>9F00ffff</color><scale>1</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><PolyStyle><color>3f00ffff</color><fill>1</fill><outline>1</outline></PolyStyle><LineStyle><color>fF00ffff</color><width>2</width></LineStyle></Style>")
        kml.WriteLine("<Style id=""iota_highlight""><LabelStyle><color>9F00ffff</color><scale>1</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><PolyStyle><color>3f00ffff</color><fill>0</fill><outline>1</outline></PolyStyle><LineStyle><color>fF00ffff</color><width>2</width></LineStyle></Style>")
        kml.WriteLine("<StyleMap id='iota'>")
        kml.WriteLine("<Pair><key>normal</key><styleUrl>#iota_normal</styleUrl></Pair>")
        kml.WriteLine("<Pair><key>highlight</key><styleUrl>#iota_highlight</styleUrl></Pair>")
        kml.WriteLine("</StyleMap>")
        sql.CommandText = "SELECT * FROM IOTA_Groups"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            count += 1
            kml.WriteLine($"<Placemark><name>{SQLdr("refno")}</name><styleUrl>#iota</styleUrl><visibility>0</visibility>")
            ' Retrieve the islands of this group
            sql1.CommandText = $"SELECT GROUP_CONCAT(name||(iif(comment != '',' ('||comment||')','')),', ') AS Islands FROM IOTA_Islands WHERE refno='{SQLdr("refno")}'"
            SQLdr1 = sql1.ExecuteReader
            SQLdr1.Read()
            Dim Islands = SQLdr1("Islands")
            SQLdr1.Close()
            Dim comment As String = IIf(Form1.IsDBNullorEmpty(SQLdr("comment")), "", $"{SQLdr("comment")}{vbCrLf}{vbCrLf}")      ' add optional comment
            kml.WriteLine($"<description>{KMLescape(comment)}{KMLescape(Islands)}</description>")
            ' create the bounding box polygon
            Dim env = New Envelope(New MapPoint(CDbl(SQLdr("longitude_min")), CDbl(SQLdr("latitude_min")), SpatialReferences.Wgs84), New MapPoint(CDbl(SQLdr("longitude_max")), CDbl(SQLdr("latitude_max")), SpatialReferences.Wgs84))
            Dim pgb As New PolygonBuilder(SpatialReferences.Wgs84)       ' drawn bounding box
            With env
                pgb.AddPoint(New MapPoint(.XMin, .YMax))       ' first point
                pgb.AddPoint(New MapPoint(.XMax, .YMax))
                pgb.AddPoint(New MapPoint(.XMax, .YMin))
                pgb.AddPoint(New MapPoint(.XMin, .YMin))
            End With
            Dim poly = pgb.ToGeometry
            If CrossesAntiMeridian(poly) And SQLdr("refno") <> "AN-016" Then poly = NormalizeCentralMeridian(poly)
            Dim labelpoint As MapPoint = poly.LabelPoint
            With env
                kml.WriteLine("<MultiGeometry>")
                kml.Write($"<Point>")
                KMLcoordinates(kml, labelPoint, 1)
                kml.WriteLine("</Point>")
                poly = poly.Densify(5)
                KMLPolygon(kml, poly, 2, True)
                kml.WriteLine("</MultiGeometry>")
            End With
            kml.WriteLine($"</Placemark>")
        End While
        kml.WriteLine("</Folder>")
        timer.Stop()
        Form1.AppendText(Form1.TextBox1, $"IOTA folder created with {count} island groups [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Public Sub ZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for CQ & ITU Zones. Use data created by Francesco Crosilla IV3TMM (SK)
        Dim sql As SqliteCommand, timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, "Making KML for CQ/ITU folders.")
        Application.DoEvents()
        sql = connect.CreateCommand
        ' Create CQ zones
        CreateZoneKML(connect, kml, "CQ")

        'Now do ITU zones
        CreateZoneKML(connect, kml, "ITU")
        timer.Stop()
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub CreateZoneKML(connect As SqliteConnection, kml As StreamWriter, zone As String)
        ' Create linestrings and lables for CQ or ITU zones
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        sql = connect.CreateCommand
        ' Create the folder
        kml.WriteLine($"<Folder><name>{zone} Zones</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<description>Polygons describing {zone} Zones. Based on data created by Elwood Downey (WBØOEW), extracted from http://zone-check.eu/.</description>")
        Select Case zone
            Case "CQ"
                ' add styles for hover over polygon
                kml.WriteLine($"<Style id='CQnormal'><PolyStyle><fill>0</fill><outline>1</outline></PolyStyle><LabelStyle><color>ff0000ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ff0000ff</color><width>5</width></LineStyle></Style>")
                kml.WriteLine($"<Style id='CQhighlight'><PolyStyle><color>400000ff</color><fill>1</fill><outline>1</outline></PolyStyle><LabelStyle><color>fF0000ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ff0000ff</color><width>5</width></LineStyle></Style>")
                kml.WriteLine($"<StyleMap id='CQ'>
                    <Pair><key>normal</key><styleUrl>#CQnormal</styleUrl></Pair>
                    <Pair><key>highlight</key><styleUrl>#CQhighlight</styleUrl></Pair>
                    </StyleMap>")
            Case "ITU"
                ' add styles for hover over polygon
                kml.WriteLine($"<Style id='ITUnormal'><PolyStyle><fill>0</fill><outline>1</outline></PolyStyle><LabelStyle><color>ffff00ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ffff00ff</color><width>5</width></LineStyle></Style>")
                kml.WriteLine($"<Style id='ITUhighlight'><PolyStyle><color>40ff00ff</color><fill>1</fill><outline>1</outline></PolyStyle><LabelStyle><color>ffff00ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ffff00ff</color><width>5</width></LineStyle></Style>")
                kml.WriteLine($"<StyleMap id='ITU'>
                    <Pair><key>normal</key><styleUrl>#ITUnormal</styleUrl></Pair>
                    <Pair><key>highlight</key><styleUrl>#ITUhighlight</styleUrl></Pair>
                    </StyleMap>")
            Case Else
                Throw New System.Exception("zone type was not CQ or ITU")
        End Select

        sql.CommandText = $"SELECT * FROM ZoneLines WHERE Type='{zone}' ORDER BY zone"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#{zone}</styleUrl><name>{zone} {SQLdr("zone")}</name>")
            Dim poly As Polygon = Geometry.FromJson(SQLdr("geometry"))            ' retrieve geometry
            Dim labelpoint As MapPoint = poly.LabelPoint    ' calculate labelpoint
            kml.WriteLine("<MultiGeometry>")
            kml.Write($"<Point>")
            KMLcoordinates(kml, labelpoint, 1)
            kml.WriteLine("</Point>")
            poly = poly.Densify(5)          ' ensure point every 5 degrees
            KMLPolygon(kml, poly, 2, True)
            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine("</Placemark>")
        End While
        SQLdr.Close()
        kml.WriteLine("</Folder>")
    End Sub
    Public Sub TimeZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for Timezones. 
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, $"Making KML for Timezone folder.")
        Application.DoEvents()
        sql = connect.CreateCommand
        kml.WriteLine("<Folder><name>Timezones</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<description>Polygons describing timezones. Data obtained from https://www.naturalearthdata.com/</description>")

        ' create styles for the 6 different colour polygons
        For i = 1 To 6
            kml.WriteLine($"<StyleMap id='TZ_{i}'>")
            kml.WriteLine($"<Pair><key>normal</key><styleUrl>#{Form1.ColourMapping(i)}</styleUrl></Pair>")
            kml.WriteLine("<Pair><key>highlight</key><styleUrl>#white</styleUrl></Pair>")
            kml.WriteLine("</StyleMap>")
        Next
        ' get list of all unique timezones
        Dim TZlist As New List(Of (name As String, color As Integer, places As String))
        sql.CommandText = "SELECT name,color,GROUP_CONCAT(places,"", "") AS places FROM Timezones GROUP BY name"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            TZlist.Add((SQLdr("name"), SQLdr("color"), SQLdr("places")))
        End While
        SQLdr.Close()
        TZlist.Sort(Function(x, y) CInt(x.name).CompareTo(CInt(y.name)))      ' sort in numerical order
        ' Now retrieve each timezone
        For Each tz In TZlist
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#TZ_{tz.color}</styleUrl><name>{tz.name}</name><description>{tz.places}</description>")
            kml.WriteLine("<MultiGeometry>")
            sql.CommandText = $"SELECT * FROM Timezones WHERE name='{tz.name}'"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Dim geom As Polygon = Geometry.FromJson(SQLdr("geometry"))
                KMLPolygon(kml, geom, 1, False)
            End While
            SQLdr.Close()
            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine("</Placemark>")
        Next
        kml.WriteLine("</Folder>")
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub IARUFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make the IARU regions folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, "Making KML for IARU folder.")
        kml.WriteLine("<Folder><name>IARU regions</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<Style id=""IARU""><LineStyle><color>ffffffff</color><width>3</width></LineStyle><LabelStyle><color>ffffffff</color><scale>15</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle></Style>")
        kml.WriteLine("<description>Lines defining IARU boundaries. Data provided by Tim Makins (EI8IC)/</description>")
        sql = connect.CreateCommand
        ' first do the lines
        kml.WriteLine("<Folder><name>Boundaries</name><visibility>0</visibility>")
        sql.CommandText = "SELECT * FROM IARU"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            Dim geom As Polyline = Geometry.FromJson(SQLdr("geometry"))
            kml.WriteLine($"<Placemark><name>Line {SQLdr("line")}</name><visibility>0</visibility><styleUrl>#IARU</styleUrl>")
            KMLLineString(kml, geom, 1)
            kml.WriteLine("</Placemark>")
        End While
        kml.WriteLine("</Folder>")
        ' now do the labels
        kml.WriteLine("<Folder><name>labels</name><visibility>0</visibility>")
        Dim labels As New List(Of (region As Integer, lat As Double, lon As Double)) From {(1, 0, -90), (2, 0, 20), (3, 5, 145)}      ' positions of the labels for regions
        For Each label In labels
            kml.Write($"<Placemark><visibility>0</visibility><styleUrl>#IARU</styleUrl><name>IARU {label.region}</name><Point>")
            Dim point = New MapPoint(label.lon, label.lat, SpatialReferences.Wgs84)
            KMLcoordinates(kml, point, 1)
            kml.WriteLine("</Point></Placemark>")
        Next
        kml.WriteLine("</Folder>")
        kml.WriteLine("</Folder>")
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Sub AntarcticFolder(connect As SqliteConnection, kml As StreamWriter)
        ' Create the Antarctic bases folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        Form1.AppendText(Form1.TextBox1, "Making Antarctic bases folder.")
        Application.DoEvents()
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Antarctic bases</name>")
        kml.WriteLine("<description>Details of all active Antartic bases. Data harvested from https://www.coolantarctica.com/Community/antarctic_bases.php</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""antarctic""><LabelStyle><color>ffff7f7f</color></LabelStyle><IconStyle><Icon><href>https://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><scale>3</scale></IconStyle></Style>")

        sql = connect.CreateCommand
        sql.CommandText = "SELECT * FROM Antarctic ORDER BY name"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            kml.WriteLine($"<Placemark><name>{KMLescape(SQLdr("name"))}</name><visibility>0</visibility><styleUrl>#antarctic</styleUrl>")
            Dim point As MapPoint = Geometry.FromJson(SQLdr("coordinates"))
            kml.WriteLine($"<Point><coordinates>{point.X:f2},{point.Y:f2}</coordinates></Point>")
            kml.WriteLine("<ExtendedData>")
            kml.WriteLine($"<Data name=""name""><value>{KMLescape(SQLdr("name"))}</value></Data>")
            kml.WriteLine($"<Data name=""nation""><value>{KMLescape(SQLdr("nation"))}</value></Data>")
            kml.WriteLine($"<Data name=""lat""><value>{point.Y:f3}</value></Data>")
            kml.WriteLine($"<Data name=""lon""><value>{point.X:f3}</value></Data>")
            kml.WriteLine($"<Data name=""situation""><value>{KMLescape(SQLdr("situation"))}</value></Data>")
            kml.WriteLine($"<Data name=""altitude""><value>{SQLdr("altitude")}</value></Data>")
            kml.WriteLine($"<Data name=""open""><value>{KMLescape(SQLdr("open"))}</value></Data>")
            kml.WriteLine("</ExtendedData>")
            kml.WriteLine($"</Placemark>")
        End While
        kml.WriteLine("</Folder>")
        Form1.AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Sub EUASborder(kml As StreamWriter)
        ' add the EU/AS Russia border
        Dim doc As XDocument = XDocument.Load($"{Application.StartupPath}\KML\BorderEUAS.kml")    ' read the XML
        Dim ns As XNamespace = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
        Dim nsmgr As New XmlNamespaceManager(New NameTable())
        nsmgr.AddNamespace("x", ns.NamespaceName)
        Dim placemark = doc.XPathSelectElement("//x:Placemark[2]", nsmgr).ToString       ' find second placemark
        placemark = Strings.Replace(placemark, $" xmlns=""{ns}""", "")         ' remove pesky namespace
        kml.WriteLine(placemark)                ' write border to kml file
    End Sub

End Module
