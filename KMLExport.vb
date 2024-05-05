Imports System.IO
Imports System.Xml
Imports System.Xml.XPath
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite

Module KMLExport

    Const CoordsPerLine = 10          ' coordinates per line
    Private prefixes As New List(Of (p As MapPoint, pfx As String))         ' prefixes for prefix folder
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
        Dim LabelPoint = boundaries.LabelPoint
        kml.Write($"<Point>")
        KMLcoordinates(kml, LabelPoint, 4)
        kml.WriteLine("</Point>")
        prefixes.Add((LabelPoint, SQLdr("prefix")))     ' for the prefix folder
        kml.WriteLine("<ExtendedData>")
        kml.WriteLine($"<Data name=""Entity""><value>{KMLescape(SQLdr("Entity"))}</value></Data>")
        kml.WriteLine($"<Data name= ""DXCC number""><value>{SQLdr("DXCCnum")}</value></Data>")
        kml.WriteLine($"<Data name=""Prefix""><value>{SQLdr("prefix")}</value></Data>")
        kml.WriteLine($"<Data name=""CQ Zone""><value>{SQLdr("CQ")}</value></Data>")
        kml.WriteLine($"<Data name=""ITU Zone""><value>{SQLdr("ITU")}</value></Data>")
        kml.WriteLine($"<Data name=""IARU Region""><value>{SQLdr("IARU")}</value></Data>")
        kml.WriteLine($"<Data name=""Continent""><value>{SQLdr("Continent")}</value></Data>")
        kml.WriteLine($"<Data name=""Start Date""><value>{SQLdr("StartDate")}</value></Data>")
        kml.WriteLine($"<Data name=""lat""><value>{SQLdr("lat")}</value></Data>")
        kml.WriteLine($"<Data name=""lon""><value>{SQLdr("lon")}</value></Data>")
        If Not IsDBNull(SQLdr("query")) Then
            Dim QL = SQLdr("query")
            If Not IsDBNull(SQLdr("bbox")) Then QL &= $"({SQLdr("bbox")})"      ' add bounding box if any
            kml.WriteLine($"<Data name=""OSM query""><value>{KMLescape(QL)}</value></Data>")
        End If
        If Not IsDBNull(SQLdr("notes")) Then
            kml.WriteLine($"<Data name=""Notes""><value><![CDATA[{Form1.hyperlink(SQLdr("notes"))}]]></value></Data>")
        End If
        kml.WriteLine("</ExtendedData>")
        If boundaries.Parts.Count > 1 Then kml.WriteLine("<MultiGeometry>")
        tagStack.Clear()
        For Each Prt In boundaries.Parts
            Dim area = Form1.PolygonArea(Prt)     ' determine inner or outer by testing polygon area
            If area > 0 Then
                ' It's an inner boundary
                kml.WriteLine("<innerBoundaryIs>")
                tagStack.Push("</innerBoundaryIs>")
            Else
                ' It's an outer boundary. Finish previous polygon, if any, and start a new one
                While (tagStack.Count > 0)
                    kml.WriteLine(tagStack.Pop)     ' empty the stack
                End While
                kml.WriteLine("<Polygon>")    ' open new one
                tagStack.Push("</Polygon>")         ' push end tag
                kml.WriteLine("<tessellate>1</tessellate>")
                kml.WriteLine("<outerBoundaryIs>")
                tagStack.Push("</outerBoundaryIs>")
            End If
            ' Now write the boundary (inner or outer)
            kml.WriteLine("<LinearRing>")
            KMLcoordinates(kml, Prt.Points.ToList, 5)
            kml.WriteLine("</LinearRing>")
            kml.WriteLine(tagStack.Pop)    ' close of boundaryIs
        Next
        ' Close any open polygon
        While (tagStack.Count > 0)
            kml.WriteLine(tagStack.Pop)
        End While
        If boundaries.Parts.Count > 1 Then kml.WriteLine("</MultiGeometry>")
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
        Form1.AppendText(Form1.TextBox1, $"Making KML for {prefixes.Count} prefixes{vbCrLf}")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Prefixes</name>")
        kml.WriteLine("<description>A folder containing all prefix labels</description>")
        kml.WriteLine("<Style id=""prefix""><LabelStyle><color>ff00ffff</color><scale>5</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle></Style>")    ' empty IconStyle required to display labels correctly (don't know why)
        prefixes = prefixes.OrderBy(Function(a) a.pfx).ToList        ' sort prefix
        For Each prefix In prefixes
            kml.Write($"<Placemark><name>{prefix.pfx}</name><styleUrl>#prefix</styleUrl><Point>")
            KMLcoordinates(kml, prefix.p, 2)
            kml.WriteLine("</Point></Placemark>")
        Next
        kml.WriteLine("</Folder>")
    End Sub
    Sub BoundingBoxFolder(kml As StreamWriter, BoundingBoxes As List(Of (name As String, box As String)))
        ' Create the Bounding Box folder
        Const DensifyDegrees = 5        ' ensure lines have resolution so the follow lat/lon lines

        Form1.AppendText(Form1.TextBox1, $"Making KML for {BoundingBoxes.Count} bounding boxes{vbCrLf}")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Bounding Boxes</name>")
        kml.WriteLine("<description>Bounding boxes sometimes necessary to filter query. For debugging only.</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""bbox""><PolyStyle><fill>0</fill><outline>1</outline></PolyStyle><LineStyle><color>ffffffff</color><width>2</width></LineStyle></Style>")
        For Each BoundingBox In BoundingBoxes
            ' Display bounding box as a closed linestring
            kml.WriteLine($"<Placemark><name>{KMLescape(BoundingBox.name)}</name><visibility>0</visibility><styleUrl>#bbox</styleUrl>")
            ' Densify the bounding box so that it clearly follows lat/lon lines better. Exclude Fiji as Densify takes long path
            ' You can't seem to Densify a  Multipoint, so convert to Polygon, Densify, and convert back to Multipoint
            Dim box As Multipoint = Geometry.FromJson(BoundingBox.box)        ' convert the json to geometry
            If BoundingBox.name <> "Fiji" Then
                Dim PolyBox As New Polygon(box.Points)          ' Convert to Polygon
                PolyBox = PolyBox.Densify(DensifyDegrees)       ' Densify
                box = New Multipoint(PolyBox.Parts(0).Points)   ' Convert back to Multipoint
            End If
            ' Write out the linestring
            kml.WriteLine("<LineString><tessellate>1</tessellate>")
            KMLcoordinates(kml, box.Points.ToList, 2)
            kml.WriteLine("</LineString>")
            kml.WriteLine($"</Placemark>")
        Next
        kml.WriteLine("</Folder>")
    End Sub
    Sub GridSquareFolder(connect As SqliteConnection, kml As StreamWriter, DXCClist As List(Of Integer))
        ' Create the Grid Square folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Form1.AppendText(Form1.TextBox1, $"Making KML for Grid Square folder{vbCrLf}")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Grid Squares</name>")
        kml.WriteLine("<description>Gridsquare overlay</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""grid""><LabelStyle><color>ff000000</color><scale>2</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><PolyStyle><color>3fffffff</color><fill>0</fill><outline>1</outline></PolyStyle><LineStyle><color>ff000000</color><width>2</width></LineStyle></Style>")

        sql = connect.CreateCommand
        For Each DXCC In DXCClist
            sql.CommandText = $"SELECT * FROM DXCC WHERE DXCCnum={DXCC}"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            Dim geometry = Polygon.FromJson(SQLdr("geometry"))          ' get the geometry
            Dim Entity = SQLdr("Entity")
            SQLdr.Close()
            Dim extent = geometry.Extent
            ' Create an envelope which exactly covers all the possible grid squares
            Dim GridExtent As New Envelope(Math.Floor(extent.XMin / Form1.GridSquareX) * Form1.GridSquareX, Math.Floor(extent.YMin / Form1.GridSquareY) * Form1.GridSquareY, Math.Ceiling(extent.XMax / Form1.GridSquareX) * Form1.GridSquareX, Math.Ceiling(extent.YMax / Form1.GridSquareY) * Form1.GridSquareY, SpatialReferences.Wgs84)
            ' test every gridsquare with the grid extent
            Dim GridSquares As New List(Of (grid As String, square As Envelope))
            Dim X = GridExtent.XMin
            While X < GridExtent.XMax
                Dim Y = GridExtent.YMin
                While Y < GridExtent.YMax
                    Dim GridSq As New Envelope(X, Y, X + Form1.GridSquareX, Y + Form1.GridSquareY, SpatialReferences.Wgs84)
                    If GridSq.Intersects(geometry) Then
                        GridSquares.Add((Form1.GridSquare(New MapPoint(X, Y, SpatialReferences.Wgs84)), GridSq))
                    End If
                    Y += Form1.GridSquareY
                End While
                X += Form1.GridSquareX
            End While
            Dim sorted = From g In GridSquares Order By g.grid Select g    ' sort the list in gridsquare order
            GridSquares = sorted.ToList
            kml.WriteLine($"<Folder><name>{KMLescape(Entity)}</name>")
            For Each sq In GridSquares
                With sq
                    Dim lb As New MapPoint(.square.Extent.XMin + .square.Width / 2, .square.Extent.YMin + .square.Height / 2, SpatialReferences.Wgs84)    ' label point is center of square
                    kml.WriteLine($"<Placemark><name>{Form1.GridSquare(lb)}</name><styleUrl>#grid</styleUrl><visibility>0</visibility>")
                    kml.WriteLine("<MultiGeometry>")
                    kml.Write($"<Point>")
                    KMLcoordinates(kml, lb, 1)
                    kml.WriteLine("</Point>")
                    kml.WriteLine("<LineString>")
                    Dim ls As New List(Of MapPoint)
                    With .square
                        ls.Add(New MapPoint(.XMin, .YMax, SpatialReferences.Wgs84))
                        ls.Add(New MapPoint(.XMax, .YMax, SpatialReferences.Wgs84))
                        ls.Add(New MapPoint(.XMax, .YMin, SpatialReferences.Wgs84))
                        ls.Add(New MapPoint(.XMin, .YMin, SpatialReferences.Wgs84))
                        ls.Add(New MapPoint(.XMin, .YMax, SpatialReferences.Wgs84))
                        KMLcoordinates(kml, ls, 0)
                    End With
                    kml.WriteLine($"</LineString>")
                    kml.WriteLine("</MultiGeometry>")
                    kml.WriteLine($"</Placemark>")
                End With
            Next
            kml.WriteLine("</Folder>")
        Next
        kml.WriteLine("</Folder>")
    End Sub

    Public Sub ZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for CQ & ITU Zones. Use data created by Francesco Crosilla IV3TMM (SK)
        Dim sql As SqliteCommand

        Form1.AppendText(Form1.TextBox1, $"Making KML for CQ/ITU folders{vbCrLf}")
        sql = connect.CreateCommand
        kml.WriteLine("<Folder><name>CQ Zones</name><visibility>0</visibility>")
        kml.WriteLine("<description>Lines describing CQ Zones. Based on data created by Francesco Crosilla IV3TMM (SK).</description>")
        kml.WriteLine("<Style id=""CQ""><LabelStyle><color>fF0000ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>fF0000Ff</color><width>5</width></LineStyle></Style>")
        CreateZoneKML(connect, kml, "CQ")
        kml.WriteLine("</Folder>")

        'Now do ITU zones
        kml.WriteLine("<Folder><name>ITU Zones</name><visibility>0</visibility>")
        kml.WriteLine("<description>Lines describing ITU Zones. Based on data created by Francesco Crosilla IV3TMM (SK).</description>")
        kml.WriteLine("<Style id=""ITU""><LabelStyle><color>ffff02fc</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ffff02fc</color><width>5</width></LineStyle></Style>")
        CreateZoneKML(connect, kml, "ITU")
        kml.WriteLine("</Folder>")
    End Sub
    Sub CreateZoneKML(connect As SqliteConnection, kml As StreamWriter, zone As String)
        ' Create linestrings and lables for CQ or ITU zones
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        sql = connect.CreateCommand
        ' First do the linestrings
        kml.WriteLine("<Folder><name>Boundaries</name><open>0</open>")
        sql.CommandText = $"SELECT * FROM ZoneLines WHERE Type='{zone}' ORDER BY line"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#{zone}</styleUrl><name>{zone} line {SQLdr("line")}</name>")
            Dim linestring As Multipoint = Geometry.FromJson(SQLdr("geometry"))
            kml.WriteLine("<LineString><tessellate>1</tessellate>")
            KMLcoordinates(kml, linestring.Points.ToList, 1)
            kml.WriteLine("</LineString>")
            kml.WriteLine("</Placemark>")
        End While
        SQLdr.Close()
        kml.WriteLine("</Folder>")

        ' Now do the zone labels
        kml.WriteLine("<Folder><name>Labels</name><open>0</open>")
        sql.CommandText = $"SELECT * FROM ZoneLabels WHERE Type='{zone}' ORDER BY zone"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            Dim pnt As MapPoint = Geometry.FromJson(SQLdr("geometry"))        ' retrieve the geometry
            kml.Write($"<Placemark><visibility>0</visibility><styleUrl>#{zone}</styleUrl><name>{zone} {SQLdr("zone")}</name><Point>")
            KMLcoordinates(kml, pnt, 1)
            kml.WriteLine("</Point></Placemark>")
        End While
        SQLdr.Close()
        kml.WriteLine("</Folder>")
    End Sub
    Public Sub TimeZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for Timezones. 
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Form1.AppendText(Form1.TextBox1, $"Making KML for Timezone folder{vbCrLf}")
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
                kml.WriteLine("<Polygon><tessellate>1</tessellate><outerBoundaryIs><LinearRing>")
                KMLcoordinates(kml, geom.Parts(0).Points.ToList, 1)
                kml.WriteLine("</LinearRing></outerBoundaryIs></Polygon>")
            End While
            SQLdr.Close()
            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine("</Placemark>")
        Next
        kml.WriteLine("</Folder>")
    End Sub

    Sub IARUFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make the IARU regions folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

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
            kml.WriteLine($"<Placemark><name>Line {SQLdr("line")}</name><visibility>0</visibility><styleUrl>#IARU</styleUrl><LineString><tessellate>1</tessellate>")
            Dim count As Integer = 0
            For Each prt In geom.Parts
                KMLcoordinates(kml, prt.Points.ToList, 1)
            Next
            kml.WriteLine("</LineString></Placemark>")
        End While
        kml.WriteLine("</Folder>")
        ' now do the labels
        kml.WriteLine("<Folder><name>labels</name><visibility>0</visibility>")
        Dim labels As New List(Of (region As Integer, lat As Double, lon As Double)) From {(1, -70, 15), (2, 20, 0), (3, 145, 5)}      ' positions of the labels for regions
        For Each label In labels
            kml.Write($"<Placemark><visibility>0</visibility><styleUrl>#IARU</styleUrl><name>IARU {label.region}</name><Point>")
            Dim point = New MapPoint(label.lon, label.lat, SpatialReferences.Wgs84)
            KMLcoordinates(kml, point, 1)
            kml.WriteLine("</Point></Placemark>")
        Next
        kml.WriteLine("</Folder>")
        kml.WriteLine("</Folder>")
    End Sub
    Sub AntarcticFolder(connect As SqliteConnection, kml As StreamWriter)
        ' Create the Antarctic bases folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

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
            kml.WriteLine($"<Point><coordinates>{point.X:f4},{point.Y:f4}</coordinates></Point>")
            kml.WriteLine("<ExtendedData>")
            kml.WriteLine($"<Data name=""name""><value>{KMLescape(SQLdr("name"))}</value></Data>")
            kml.WriteLine($"<Data name=""nation""><value>{KMLescape(SQLdr("nation"))}</value></Data>")
            kml.WriteLine($"<Data name=""lat""><value>{point.Y:f4}</value></Data>")
            kml.WriteLine($"<Data name=""lon""><value>{point.X:f4}</value></Data>")
            kml.WriteLine($"<Data name=""situation""><value>{KMLescape(SQLdr("situation"))}</value></Data>")
            kml.WriteLine($"<Data name=""altitude""><value>{SQLdr("altitude")}</value></Data>")
            kml.WriteLine($"<Data name=""open""><value>{KMLescape(SQLdr("open"))}</value></Data>")
            kml.WriteLine("</ExtendedData>")
            kml.WriteLine($"</Placemark>")
        End While
        kml.WriteLine("</Folder>")
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
        If points.Count > 1 Then kml.WriteLine()
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
            If points.Count > 1 Then kml.WriteLine()    ' end of block
            index += BlockSize       ' next block
        Loop Until index >= points.Count    ' all points done
        kml.Write("</coordinates>")
        If points.Count > 1 Then kml.WriteLine()
    End Sub
End Module
