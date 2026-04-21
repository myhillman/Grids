Imports System.Diagnostics.Eventing.Reader
Imports System.Globalization
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml
Imports System.Xml.XPath
Imports NetTopologySuite
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO
Imports NetTopologySuite.IO.Handlers
Imports NetTopologySuite.Operation.Polygonize
Imports NetTopologySuite.Precision
Module KMLExport

    Const CoordsPerLine = 10          ' coordinates per line
    Private prefixes As New List(Of (p As Point, pfx As String))
    ' prefixes for prefix folder
    Sub MakeKMLAllEntities()
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            With Form1.ProgressBar1
                .Minimum = 0
                .Maximum = 340
                .Value = 0
            End With
            sql = connect.CreateCommand
            sql.CommandText = ("SELECT * FROM `DXCC` WHERE `DELETED`=0 AND geometry is not null AND `DXCCnum`<>999 ORDER BY `Entity`")
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Dim entity = SafeStr(SQLdr("Entity"))
                Using kml As New StreamWriter(Path.Combine(Application.StartupPath, "KML", $"DXCC_{entity}.kml"), False)
                    kml.WriteLine(KMLheader)
                    Placemark(connect, kml, SQLdr("DXCCnum"))
                    kml.WriteLine(KMLfooter)
                End Using
                Form1.ProgressBar1.Value += 1
            End While
        End Using
    End Sub

    Sub MakeKML()
        ' Make a KML file of all Entity boundaries
        Dim DXCClist As New List(Of Integer)
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        Dim BaseFilename As String = Path.Combine(Application.StartupPath, "KML", "DXCC Map of the World")
        Using connect As New SqliteConnection(DXCC_DATA),
            kml As New StreamWriter($"{BaseFilename}.kml", False)

            connect.Open()
            ' Create list of DXCC to convert to KML (all)
            sql = connect.CreateCommand
            sql.CommandText = "Select * FROM `DXCC` WHERE `Deleted`=0 And `geometry` Is Not NULL ORDER BY `Entity`"     ' fetch all geometry
            SQLdr = sql.ExecuteReader
            DXCClist.Clear()

            While SQLdr.Read
                DXCClist.Add(SQLdr("DXCCnum"))
            End While
            SQLdr.Close()
            kml.WriteLine(KMLheader)
            KMLlist(connect, kml, DXCClist)
            If DXCClist.Contains(15) Or DXCClist.Contains(54) Then EUASborder(kml)      ' kml contains EU or AS Russia, then include the boundary
            PrefixFolder(kml)
            GridSquareFolder(connect, kml, DXCClist)
            IOTAFolder(connect, kml)
            ZoneFolder(connect, kml)       ' data is lost
            IARUFolder(connect, kml)
            TimeZoneFolder(connect, kml)
            BoundingBoxFolder(connect, kml)
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
            AppendText(Form1.TextBox1, $"KML file {BaseFilename} of {kmlSize / 1024:f0} Kb compressed to {kmzSize / 1024:f0} Kb, i.e. {kmzSize / kmlSize * 100:f0}%{vbCrLf}")
            AppendText(Form1.TextBox1, $"Done{vbCrLf}")
        End Using
    End Sub
    Sub KMLlist(connect As SqliteConnection, kml As StreamWriter, DXCClist As List(Of Integer))
        With Form1.ProgressBar1
            .Minimum = 0
            .Value = 0
            .Maximum = DXCClist.Count
        End With
        prefixes.Clear()
        AppendText(Form1.TextBox1, $"Making KML for {DXCClist.Count} Entities{vbCrLf}")
        kml.WriteLine("<Folder><name>DXCC Entities</name><open>0</open><description>Boundaries of DXCC entities</description>")
        For Each dxcc In DXCClist
            Placemark(connect, kml, dxcc)
        Next
        kml.WriteLine("</Folder>")
        AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub

    Sub Placemark(connect As SqliteConnection, kml As StreamWriter, dxcc As Integer)

        ' --- Load DXCC row ---
        Dim sql = connect.CreateCommand()
        sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={dxcc}"
        Dim SQLdr = sql.ExecuteReader()
        SQLdr.Read()

        Dim entity = SafeStr(SQLdr("Entity"))
        Dim geoJson = SafeStr(SQLdr("geometry"))

        If String.IsNullOrWhiteSpace(geoJson) Then
            Dim errMsg = $"There is no geometry for {entity} to convert to KML"
            AppendText(Form1.TextBox1, errMsg & vbCrLf)
            kml.WriteLine(errMsg)
            Exit Sub
        End If

        ' --- Convert GeoJSON → NTS geometry ---
        Dim ntsGeom As NetTopologySuite.Geometries.Geometry = FromGeoJsonToNTS(geoJson)

        If ntsGeom Is Nothing OrElse ntsGeom.IsEmpty Then
            Dim errMsg = $"Geometry for {entity} is empty or invalid"
            AppendText(Form1.TextBox1, errMsg & vbCrLf)
            kml.WriteLine(errMsg)
            Exit Sub
        End If

        AppendText(Form1.TextBox1, $"Making KML for {entity}{vbCrLf}")

        ' --- Start Placemark ---
        kml.WriteLine($"<Placemark><styleUrl>#boundary_{SQLdr("colour")}</styleUrl>")
        kml.WriteLine($"<name>{KMLescape(entity)} ({SQLdr("prefix")})</name>")

        ' --- Label point using NTS centroid ---
        Dim centroid As Coordinate = ntsGeom.Centroid.Coordinate
        Dim centroidPnt = factory.CreatePoint(centroid)

        WriteKmlPoint(centroidPnt, kml, 0)

        ' Store prefix label for PrefixFolder (now using NTS coordinate)
        prefixes.Add((centroidPnt, SQLdr("prefix").ToString()))

        ' --- ExtendedData ---
        kml.WriteLine("<ExtendedData>")
        kml.WriteLine($"<Data name=""Entity""><value>{KMLescape(entity)}</value></Data>")
        kml.WriteLine($"<Data name=""DXCC number""><value>{SQLdr("DXCCnum")}</value></Data>")
        kml.WriteLine($"<Data name=""Prefix""><value>{SQLdr("prefix")}</value></Data>")
        kml.WriteLine($"<Data name=""CQ Zone""><value>{SQLdr("CQ")}</value></Data>")
        kml.WriteLine($"<Data name=""ITU Zone""><value>{SQLdr("ITU")}</value></Data>")
        kml.WriteLine($"<Data name=""IARU Region""><value>{SQLdr("IARU")}</value></Data>")
        kml.WriteLine($"<Data name=""Continent""><value>{SQLdr("Continent")}</value></Data>")
        kml.WriteLine($"<Data name=""Start Date""><value>{SQLdr("StartDate")}</value></Data>")
        kml.WriteLine($"<Data name=""lat""><value>{SQLdr("lat"):f3}</value></Data>")
        kml.WriteLine($"<Data name=""lon""><value>{SQLdr("lon"):f3}</value></Data>")

        Dim query As New StringBuilder()
        If Not IsDBNullorEmpty(SQLdr("source")) Then query.Append($"{SQLdr("source").ToString().Trim()}: ")
        If Not IsDBNullorEmpty(SQLdr("rule")) Then query.Append(SQLdr("rule").ToString().Trim())
        If Not IsDBNullorEmpty(SQLdr("bbox")) Then query.Append($" ({SQLdr("bbox")})")

        If query.Length > 0 Then
            kml.WriteLine($"<Data name=""GIS query""><value>{KMLescape(query.ToString())}</value></Data>")
        End If

        If Not IsDBNullorEmpty(SQLdr("notes")) Then
            kml.WriteLine($"<Data name=""Notes""><value><![CDATA[{Hyperlink(SQLdr("notes"))}]]></value></Data>")
        End If

        kml.WriteLine("</ExtendedData>")

        ' --- Write polygon(s) ---
        WriteKmlSingleOrMultiPolygon(ntsGeom, kml, 3)

        kml.WriteLine("</Placemark>")
    End Sub

    Sub PrefixFolder(kml As StreamWriter)

        Dim timer As New Stopwatch
        timer.Start()

        AppendText(Form1.TextBox1, $"Making KML for {prefixes.Count} prefixes.")

        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Prefixes</name>")
        kml.WriteLine("<description>A folder containing all prefix labels</description>")
        kml.WriteLine("<Style id=""prefix""><LabelStyle><color>ff00ffff</color><scale>5</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle></Style>")

        ' prefixes is now List(Of (p As Coordinate, pfx As String))
        prefixes = prefixes.OrderBy(Function(a) a.pfx).ToList()

        For Each prefix In prefixes
            kml.Write($"<Placemark><name>{prefix.pfx}</name><styleUrl>#prefix</styleUrl>")
            ' prefix.p is now an NTS Coordinate
            WriteKmlPoint(prefix.p, kml, 0)
            kml.WriteLine("</Placemark>")
        Next
        kml.WriteLine("</Folder>")

        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub

    Sub BoundingBoxFolder(connect As SqliteConnection, kml As StreamWriter)

        Const DensifyDegrees = 10
        Dim timer As New Stopwatch
        timer.Start()

        AppendText(Form1.TextBox1, "Making KML for bounding boxes.")
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Bounding Boxes</name>")
        kml.WriteLine("<description>Bounding boxes sometimes necessary to filter query. For debugging only.</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<open>0</open>")
        kml.WriteLine("<Style id=""bbox"">
              <IconStyle><scale>0</scale></IconStyle>
              <LabelStyle><color>ffffffff</color><scale>1.2</scale></LabelStyle>
              <LineStyle><color>ffffffff</color><width>2</width></LineStyle>
              <PolyStyle><color>00ffffff</color><fill>0</fill><outline>1</outline></PolyStyle>
            </Style>")

        Dim sql = connect.CreateCommand()
        sql.CommandText = "SELECT * FROM `DXCC` WHERE `Deleted`=0 AND `bbox` IS NOT NULL AND bbox <> '' ORDER BY `Entity`"
        Dim SQLdr = sql.ExecuteReader()

        While SQLdr.Read()

            Dim entity = SafeStr(SQLdr("Entity"))
            DebugOutput = (entity = "Fiji")

            ' ParseBBoxOrPoly still returns Runtime geometry → convert to GeoJSON → NTS
            Dim ntsGeom = ParseBBoxOrPoly(SQLdr("bbox"))

            ' --- Label point using NTS centroid ---
            Dim centroid As Coordinate = ntsGeom.Centroid.Coordinate
            kml.WriteLine("<Placemark>")
            kml.WriteLine($"  <name>{KMLescape(entity)}</name>")
            kml.WriteLine("  <styleUrl>#bbox</styleUrl>")
            kml.WriteLine($"  <Point><coordinates>{centroid.X:f2},{centroid.Y:f2}</coordinates></Point>")
            kml.WriteLine("  <MultiGeometry>")

            ' --- Convert bbox geometry into polygons ---
            Dim polys As New List(Of NetTopologySuite.Geometries.Polygon)

            If TypeOf ntsGeom Is NetTopologySuite.Geometries.Polygon Then
                polys.Add(DirectCast(ntsGeom, NetTopologySuite.Geometries.Polygon))

            ElseIf TypeOf ntsGeom Is MultiPolygon Then
                Dim mp = DirectCast(ntsGeom, MultiPolygon)
                For i = 0 To mp.NumGeometries - 1
                    polys.Add(DirectCast(mp.GetGeometryN(i), NetTopologySuite.Geometries.Polygon))
                Next

            ElseIf TypeOf ntsGeom Is GeometryCollection Then
                Dim gc = DirectCast(ntsGeom, GeometryCollection)
                For i = 0 To gc.NumGeometries - 1
                    If TypeOf gc.GetGeometryN(i) Is NetTopologySuite.Geometries.Polygon Then
                        polys.Add(DirectCast(gc.GetGeometryN(i), NetTopologySuite.Geometries.Polygon))
                    End If
                Next

            ElseIf TypeOf ntsGeom Is LineString OrElse TypeOf ntsGeom Is Point Then
                ' Should not happen for bbox, but ignore gracefully
                Continue While
            End If

            If DebugOutput Then Debug.WriteLine($"Fiji polys from bbox: {polys.Count}")

            ' --- Densify + wrap + export ---
            For Each poly In polys

                Dim dense = NtsDensify(poly, DensifyDegrees)
                Dim wrapped = WrapPolygonTo180(dense)

                Dim env = wrapped.EnvelopeInternal
                Dim dec = DecimalsForBbox(env.MinX, env.MinY, env.MaxX, env.MaxY)

                WriteKmlPolygon(wrapped, kml, dec)
            Next

            kml.WriteLine("  </MultiGeometry>")
            kml.WriteLine("</Placemark>")
        End While

        SQLdr.Close()
        kml.WriteLine("</Folder>")
        timer.Stop()
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub
    Public Sub FindIARURegion()

        ' Deleted

    End Sub


    Sub KMLLineString(kml As StreamWriter,
                      geom As Geometry,
                      digits As Integer)

        If geom Is Nothing OrElse geom.IsEmpty Then Exit Sub

        ' ----------------------------
        ' Single LineString
        ' ----------------------------
        If TypeOf geom Is LineString Then
            Dim ls = DirectCast(geom, LineString)

            kml.WriteLine("<LineString><tessellate>1</tessellate>")
            WriteCoordinates(kml, ls.Coordinates, digits)
            kml.WriteLine("</LineString>")
            Exit Sub
        End If

        ' ----------------------------
        ' MultiLineString
        ' ----------------------------
        If TypeOf geom Is MultiLineString Then
            Dim mls = DirectCast(geom, MultiLineString)

            kml.WriteLine("<MultiGeometry>")

            For i = 0 To mls.NumGeometries - 1
                Dim ls = DirectCast(mls.GetGeometryN(i), LineString)

                kml.WriteLine("<LineString><tessellate>1</tessellate>")
                WriteCoordinates(kml, ls.Coordinates, digits)
                kml.WriteLine("</LineString>")
            Next

            kml.WriteLine("</MultiGeometry>")
            Exit Sub
        End If

        ' ----------------------------
        ' Unexpected geometry type
        ' ----------------------------
        Throw New InvalidOperationException(
            $"KMLLineString expected LineString or MultiLineString, got {geom.GeometryType}")
    End Sub

    Public Sub KMLMultipoint(kml As StreamWriter,
                             mp As NetTopologySuite.Geometries.MultiPoint,
                             digits As Integer)

        If mp Is Nothing OrElse mp.IsEmpty Then Exit Sub

        kml.WriteLine("<LineString><tessellate>1</tessellate>")

        ' Extract coordinates from all points
        Dim coords As New List(Of Coordinate)
        For i = 0 To mp.NumGeometries - 1
            Dim pt = DirectCast(mp.GetGeometryN(i), NetTopologySuite.Geometries.Point)
            coords.Add(pt.Coordinate)
        Next

        WriteCoordinates(kml, coords.ToArray(), digits)

        kml.WriteLine("</LineString>")
    End Sub


    ' New signature: NTS geometry only
    Sub KMLPolygon(kml As StreamWriter,
                   ntsGeom As NetTopologySuite.Geometries.Geometry,
                   digits As Integer,
                   multiGeometryOpen As Boolean)

        If ntsGeom Is Nothing OrElse ntsGeom.IsEmpty Then Exit Sub

        ' --- Case 1: MultiPolygon ---
        If TypeOf ntsGeom Is MultiPolygon Then
            Dim mp = DirectCast(ntsGeom, MultiPolygon)

            Dim useMulti As Boolean = (mp.NumGeometries > 1 AndAlso Not multiGeometryOpen)
            If useMulti Then kml.WriteLine("<MultiGeometry>")

            For i = 0 To mp.NumGeometries - 1
                Dim poly = DirectCast(mp.GetGeometryN(i), NetTopologySuite.Geometries.Polygon)
                WriteSinglePolygon(kml, poly, digits)
            Next

            If useMulti Then kml.WriteLine("</MultiGeometry>")
            Exit Sub
        End If

        ' --- Case 2: Polygon ---
        If TypeOf ntsGeom Is NetTopologySuite.Geometries.Polygon Then
            WriteSinglePolygon(kml, DirectCast(ntsGeom, NetTopologySuite.Geometries.Polygon), digits)
            Exit Sub
        End If

        ' --- Case 3: GeometryCollection ---
        If TypeOf ntsGeom Is GeometryCollection Then
            Dim gc = DirectCast(ntsGeom, GeometryCollection)

            kml.WriteLine("<MultiGeometry>")
            For i = 0 To gc.NumGeometries - 1
                WriteKmlSingleOrMultiPolygon(gc.GetGeometryN(i), kml, digits)
            Next
            kml.WriteLine("</MultiGeometry>")
            Exit Sub
        End If

        ' Other geometry types (LineString, Point) are ignored for polygon export
    End Sub
    Private Sub WriteSinglePolygon(kml As StreamWriter,
                                   poly As NetTopologySuite.Geometries.Polygon,
                                   digits As Integer)

        If poly Is Nothing OrElse poly.IsEmpty Then Exit Sub

        kml.WriteLine("<Polygon><tessellate>1</tessellate>")

        ' --- Outer ring ---
        kml.WriteLine("<outerBoundaryIs>")
        kml.WriteLine("<LinearRing>")
        WriteCoordinates(kml, poly.ExteriorRing.Coordinates, digits)
        kml.WriteLine("</LinearRing>")
        kml.WriteLine("</outerBoundaryIs>")

        ' --- Holes ---
        For i = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(i)
            kml.WriteLine("<innerBoundaryIs>")
            kml.WriteLine("<LinearRing>")
            WriteCoordinates(kml, hole.Coordinates, digits)
            kml.WriteLine("</LinearRing>")
            kml.WriteLine("</innerBoundaryIs>")
        Next

        kml.WriteLine("</Polygon>")
    End Sub
    Public Sub WriteCoordinates(kml As StreamWriter,
                                 coords As Coordinate(),
                                 digits As Integer)

        Const BlockSize = 10

        kml.Write("<coordinates>")
        If coords.Length > BlockSize Then kml.WriteLine()

        Dim index As Integer = 0
        While index < coords.Length
            Dim last = Math.Min(index + BlockSize - 1, coords.Length - 1)

            For i = index To last
                Dim x = coords(i).X.ToString($"f{digits}")
                Dim y = coords(i).Y.ToString($"f{digits}")
                kml.Write($"{x},{y}")
                If i < last Then kml.Write(" ")
            Next

            If coords.Length > BlockSize Then kml.WriteLine()
            index += BlockSize
        End While

        kml.WriteLine("</coordinates>")
    End Sub

    Function CrossesAntiMeridian(g As NetTopologySuite.Geometries.Geometry) As Boolean
        ' Returns true if geometry spans a longitude width large enough to imply AM crossing
        Dim env = g.EnvelopeInternal
        Return Math.Abs(env.MaxX - env.MinX) >= 300
    End Function

    Const EPS = 0.000001      ' a very small number (epsilon)

    Sub GridSquareFolder(connect As SqliteConnection, kml As StreamWriter, DXCClist As List(Of Integer))

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch, Ocean As Integer = 0
        Dim GridSquares As New SortedDictionary(Of String, Envelope)   ' gridsquare → envelope
        Dim Land As New Dictionary(Of String, Integer?)                ' gridsquares known to be land
        Dim Drawn As New Dictionary(Of String, Integer?)               ' gridsquares already drawn

        timer.Start()
        sql = connect.CreateCommand()

        ' Load list of gridsquares known to be land
        sql.CommandText = "SELECT * FROM LAND"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            Land(SafeStr(SQLdr("gridsquare"))) = Nothing
        End While
        SQLdr.Close()

        ' Folder header
        AppendText(Form1.TextBox1, $"Making KML for Grid Square folder ")
        Application.DoEvents()
        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Grid Squares</name>")
        kml.WriteLine("<description>Gridsquare overlay</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""grid""><LabelStyle><color>9Fffff00</color><scale>2</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><PolyStyle><color>9Fffff00</color><fill>0</fill><outline>1</outline></PolyStyle><LineStyle><color>9Fffff00</color><width>2</width></LineStyle></Style>")

        For Each DXCC In DXCClist

            GridSquares.Clear()

            ' Get DXCC geometry (must now be NTS Geometry)
            sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={DXCC}"
            SQLdr = sql.ExecuteReader()
            SQLdr.Read()
            Dim entity = SafeStr(SQLdr("Entity"))
            Dim geometry As Geometry = FromGeoJsonToNTS(SafeStr(SQLdr("geometry")))   ' NTS geometry
            If Not geometry.IsValid Then
                Debug.WriteLine($"Geometry for {entity} is invalid")
            End If
            SQLdr.Close()

            Dim extent = geometry.EnvelopeInternal
            Dim factory = geometry.Factory

            ' Envelope covering all possible grid squares for this entity
            Dim gridExtent As New Envelope(
            Math.Floor(extent.MinX / GridSquareX) * GridSquareX,
            Math.Ceiling(extent.MaxX / GridSquareX) * GridSquareX,
            Math.Floor(extent.MinY / GridSquareY) * GridSquareY,
            Math.Ceiling(extent.MaxY / GridSquareY) * GridSquareY
        )

            ' Test every gridsquare in that envelope
            Dim X = gridExtent.MinX
            While X < gridExtent.MaxX
                Dim Y = gridExtent.MinY
                While Y < gridExtent.MaxY

                    Dim sqEnv As New Envelope(
                    X,
                    X + GridSquareX - EPS,
                    Y,
                    Y + GridSquareY - EPS
                )

                    ' Build polygon for this grid square and test intersection with DXCC geometry
                    Dim sqPoly As Geometry = factory.ToGeometry(sqEnv)
                    If sqPoly.Intersects(geometry) Then

                        Dim gs As String = GridSquare(X, Y)   ' 4‑char Maidenhead

                        If Land.ContainsKey(gs) Then
                            If Not GridSquares.ContainsKey(gs) Then
                                GridSquares.Add(gs, sqEnv)
                            End If
                        Else
                            Ocean += 1
                        End If
                    End If

                    Y += GridSquareY
                End While
                X += GridSquareX
            End While

            ' Folder for this DXCC
            kml.WriteLine($"<Folder><name>{KMLescape(entity)}</name>")

            For Each sq In GridSquares

                Dim env = sq.Value
                Dim center As New Coordinate(env.MinX + env.Width / 2, env.MinY + env.Height / 2)

                kml.WriteLine($"<Placemark><name>{sq.Key}</name><styleUrl>#grid</styleUrl><visibility>0</visibility>")

                If Not Drawn.ContainsKey(sq.Key) Then

                    kml.WriteLine("<MultiGeometry>")

                    ' Label point
                    kml.Write("<Point>")
                    WriteCoordinates(kml, {center}, 1)
                    kml.WriteLine("</Point>")

                    ' Build boundary segments as multipoint (like original logic)
                    Dim coords As New List(Of Coordinate)

                    ' Grid square above
                    Dim aboveGS As String = GridSquare(env.MinX, env.MinY + GridSquareY)
                    If Not GridSquares.ContainsKey(aboveGS) Then
                        coords.Add(New Coordinate(env.MinX, env.MaxY))   ' first point
                    End If

                    coords.Add(New Coordinate(env.MaxX, env.MaxY))
                    coords.Add(New Coordinate(env.MaxX, env.MinY))
                    coords.Add(New Coordinate(env.MinX, env.MinY))

                    ' Grid square to the left
                    Dim leftGS As String = GridSquare(env.MinX - GridSquareX, env.MinY)
                    If Not GridSquares.ContainsKey(leftGS) Then
                        coords.Add(New Coordinate(env.MinX, env.MaxY))   ' last point
                    End If

                    Dim mp As MultiPoint = factory.CreateMultiPointFromCoords(coords.ToArray())
                    KMLMultipoint(kml, mp, 0)

                    kml.WriteLine("</MultiGeometry>")

                    Drawn(sq.Key) = Nothing
                End If

                kml.WriteLine("</Placemark>")
            Next

            kml.WriteLine("</Folder>")
        Next

        kml.WriteLine("</Folder>")
        timer.Stop()
        AppendText(Form1.TextBox1, $" {Ocean} GS eliminated, {Drawn.Count:n0} Unique grid squares [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub

    Sub IOTAFolder(connect As SqliteConnection, kml As StreamWriter)

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch
        Dim sql1 As SqliteCommand, SQLdr1 As SqliteDataReader, count As Integer = 0

        timer.Start()
        sql = connect.CreateCommand()
        sql1 = connect.CreateCommand()

        AppendText(Form1.TextBox1, $"Making KML for IOTA folder ")
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
        SQLdr = sql.ExecuteReader()

        While SQLdr.Read()

            count += 1
            Dim refno = SafeStr(SQLdr("refno"))

            kml.WriteLine($"<Placemark><name>{refno}</name><styleUrl>#iota</styleUrl><visibility>0</visibility>")

            ' Islands list
            sql1.CommandText = $"SELECT GROUP_CONCAT(name||(iif(comment != '',' ('||comment||')','')),', ') AS Islands FROM IOTA_Islands WHERE refno='{refno}'"
            SQLdr1 = sql1.ExecuteReader()
            SQLdr1.Read()
            Dim Islands = SafeStr(SQLdr1("Islands"))
            SQLdr1.Close()

            Dim comment As String =
            If(IsDBNullorEmpty(SQLdr("comment")), "", $"{SQLdr("comment")}{vbCrLf}{vbCrLf}")

            kml.WriteLine($"<description>{KMLescape(comment)}{KMLescape(Islands)}</description>")

            Dim polys As New List(Of Polygon)

            ' Bounding box
            Dim LonMin = SafeDbl(SQLdr("longitude_min"))
            Dim LatMin = SafeDbl(SQLdr("latitude_min"))
            Dim LonMax = SafeDbl(SQLdr("longitude_max"))
            Dim LatMax = SafeDbl(SQLdr("latitude_max"))

            If LonMin < LonMax Then
                ' Normal (no AM crossing)
                Dim coords = {
                New Coordinate(LonMin, LatMax),   ' top-left
                New Coordinate(LonMin, LatMin),   ' bottom-left
                New Coordinate(LonMax, LatMin),   ' bottom-right
                New Coordinate(LonMax, LatMax),   ' top-right
                New Coordinate(LonMin, LatMax)    ' close
            }
                polys.Add(factory.CreatePolygon(coords))
            Else
                ' Crosses anti‑meridian: split into west and east parts

                ' West part: LonMin → 180
                Dim westCoords = {
                New Coordinate(LonMin, LatMax),
                New Coordinate(LonMin, LatMin),
                New Coordinate(179.9999, LatMin),
                New Coordinate(179.9999, LatMax),
                New Coordinate(LonMin, LatMax)
            }
                Dim w = factory.CreatePolygon(westCoords)
                If Not w.IsEmpty Then polys.Add(w)

                ' East part: -180 → LonMax
                Dim eastCoords = {
                New Coordinate(-179.9999, LatMax),
                New Coordinate(-179.9999, LatMin),
                New Coordinate(LonMax, LatMin),
                New Coordinate(LonMax, LatMax),
                New Coordinate(-179.9999, LatMax)
            }
                Dim e = factory.CreatePolygon(eastCoords)
                If Not e.IsEmpty Then polys.Add(e)
            End If

            ' Label point: use centroid of first polygon
            Dim labelCoord As Coordinate = polys(0).Centroid.Coordinate
            Dim labelPoint As Point = factory.CreatePoint(labelCoord)

            kml.WriteLine("<MultiGeometry>")

            ' Label point
            WriteKmlPoint(labelPoint, kml, 1)

            ' Polygons
            For Each poly In polys
                WriteKmlPolygon(poly, kml, 2)
            Next

            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine("</Placemark>")

        End While

        kml.WriteLine("</Folder>")
        timer.Stop()
        AppendText(Form1.TextBox1, $"IOTA folder created with {count} island groups [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub

    Public Sub ZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for CQ & ITU Zones. Use data created by Francesco Crosilla IV3TMM (SK)
        Dim sql As SqliteCommand, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, "Making KML for CQ/ITU folders.")
        Application.DoEvents()
        sql = connect.CreateCommand
        ' Create CQ zones
        CreateZoneKML(connect, kml, "CQ")

        'Now do ITU zones
        CreateZoneKML(connect, kml, "ITU")
        timer.Stop()
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Sub CreateZoneKML(connect As SqliteConnection, kml As StreamWriter, zone As String)
        Dim description As String
        Dim sql As SqliteCommand = connect.CreateCommand()
        Dim SQLdr As SqliteDataReader
        Dim factory As GeometryFactory = NtsGeometryServices.Instance.CreateGeometryFactory(4326)

        ' Folder header
        kml.WriteLine($"<Folder><name>{zone} Zones</name><visibility>0</visibility><open>0</open>")

        Select Case zone
            Case "CQ"
                ' add styles for lines
                description = "Lines describing {zone} Zones. Based on data extracted from WAZ (@ IV3TMM) v1.2.kml (deleted)"
                kml.WriteLine($"<Style id='style'><LabelStyle><color>ff0000ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ff0000ff</color><width>5</width></LineStyle></Style>")
            Case "ITU"
                description = "Lines describing {zone} Zones. Based on data extracted from ITU (@ IV3TMM) v1.1.kml (deleted)"
                ' add styles for lines
                kml.WriteLine($"<Style id='style'><LabelStyle><color>ffff00ff</color><scale>6</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle><LineStyle><color>ffff00ff</color><width>5</width></LineStyle></Style>")
            Case Else
                Throw New System.Exception("zone type was not CQ or ITU")
        End Select
        kml.WriteLine($"<description>{description}</description>")

        ' Load zone lines
        sql.CommandText = $"SELECT * FROM ZoneLines WHERE Type='{zone}' ORDER BY zone"
        SQLdr = sql.ExecuteReader()

        kml.WriteLine($"<Folder><name>{zone} zone lines</name>")
        While SQLdr.Read()
            Dim zoneNum = SafeStr(SQLdr("zone"))
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#style</styleUrl><name>Line {zoneNum}</name>")
            GeoJsonToKML(SQLdr("geometry"), kml)    ' Display zone lines
            kml.WriteLine("</Placemark>")
        End While

        SQLdr.Close()
        kml.WriteLine("</Folder>")

        ' load zone labels
        kml.WriteLine($"<Folder><name>{zone} zone labels</name>")
        sql.CommandText = $"SELECT * FROM ZoneLabels WHERE Type='{zone}' ORDER BY zone"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            Dim zoneNum = SafeStr(SQLdr("zone"))
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#style</styleUrl><name>{zone} {zoneNum}</name>")
            GeoJsonToKML(SQLdr("geometry"), kml)    ' Display zone lines
            kml.WriteLine("</Placemark>")
        End While

        SQLdr.Close()
        kml.WriteLine("</Folder>")
        kml.WriteLine("</Folder>")

    End Sub

    Public Sub TimeZoneFolder(connect As SqliteConnection, kml As StreamWriter)

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, $"Making KML for Timezone folder.")
        Application.DoEvents()

        sql = connect.CreateCommand()

        kml.WriteLine("<Folder><name>Timezones</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<description>Polygons describing timezones. Data obtained from https://www.naturalearthdata.com/</description>")

        ' Style maps for 6 colours
        For i = 1 To 6
            kml.WriteLine($"<StyleMap id='TZ_{i}'>")
            kml.WriteLine($"<Pair><key>normal</key><styleUrl>#{ColourMapping(i)}</styleUrl></Pair>")
            kml.WriteLine("<Pair><key>highlight</key><styleUrl>#white</styleUrl></Pair>")
            kml.WriteLine("</StyleMap>")
        Next

        ' List of unique timezones
        Dim TZlist As New List(Of (name As String, color As Integer, places As String))
        sql.CommandText = "SELECT name,color,GROUP_CONCAT(places,', ') AS places FROM Timezones GROUP BY name"
        SQLdr = sql.ExecuteReader()
        While SQLdr.Read()
            TZlist.Add((SafeStr(SQLdr("name")), CInt(SQLdr("color")), SafeStr(SQLdr("places"))))
        End While
        SQLdr.Close()

        ' Sort numerically by name (e.g. -12, -11, …, 0, 1, …, 14)
        TZlist.Sort(Function(x, y) CSng(x.name).CompareTo(CSng(y.name)))

        ' Emit each timezone
        For Each tz In TZlist

            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#TZ_{tz.color}</styleUrl><name>{tz.name}</name><description>{tz.places}</description>")

            sql.CommandText = $"SELECT * FROM Timezones WHERE name='{tz.name}'"
            SQLdr = sql.ExecuteReader()

            While SQLdr.Read()
                GeoJsonToKML(SQLdr("geometry"), kml)
            End While

            SQLdr.Close()
            kml.WriteLine("</Placemark>")

        Next

        kml.WriteLine("</Folder>")
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub


    Sub IARUFolder(connect As SqliteConnection, kml As StreamWriter)

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, "Making KML for IARU folder.")

        kml.WriteLine("<Folder><name>IARU regions</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<Style id=""IARU""><LineStyle><color>ffffffff</color><width>3</width></LineStyle><LabelStyle><color>ffffffff</color><scale>15</scale></LabelStyle><IconStyle><Icon></Icon></IconStyle></Style>")
        kml.WriteLine("<description>Lines defining IARU boundaries. Data provided by Tim Makins (EI8IC)</description>")

        sql = connect.CreateCommand()

        ' ----------------------------
        ' 1. Boundary lines
        ' ----------------------------
        kml.WriteLine("<Folder><name>Boundaries</name><visibility>0</visibility>")
        sql.CommandText = "SELECT * FROM IARU"
        SQLdr = sql.ExecuteReader()

        While SQLdr.Read()

            Dim geoJson As String = SafeStr(SQLdr("geometry"))
            Dim ntsGeom As Geometry = FromGeoJsonToNTS(geoJson)

            ' Expecting LineString or MultiLineString
            kml.WriteLine($"<Placemark><name>Line {SQLdr("line")}</name><visibility>0</visibility><styleUrl>#IARU</styleUrl>")

            If TypeOf ntsGeom Is LineString Then
                KMLLineString(kml, DirectCast(ntsGeom, LineString), 1)

            ElseIf TypeOf ntsGeom Is MultiLineString Then
                Dim mls = DirectCast(ntsGeom, MultiLineString)
                For i = 0 To mls.NumGeometries - 1
                    KMLLineString(kml, DirectCast(mls.GetGeometryN(i), LineString), 1)
                Next
            End If

            kml.WriteLine("</Placemark>")
        End While

        kml.WriteLine("</Folder>")

        ' ----------------------------
        ' 2. Labels
        ' ----------------------------
        kml.WriteLine("<Folder><name>labels</name><visibility>0</visibility>")

        Dim labels As New List(Of (region As Integer, lat As Double, lon As Double)) From {
        (1, 0, -90),
        (2, 0, 20),
        (3, 5, 145)
    }

        For Each label In labels
            kml.Write($"<Placemark><visibility>0</visibility><styleUrl>#IARU</styleUrl><name>IARU {label.region}</name><Point>")

            Dim coord As New Coordinate(label.lon, label.lat)
            WriteCoordinates(kml, {coord}, 1)

            kml.WriteLine("</Point></Placemark>")
        Next

        kml.WriteLine("</Folder>")
        kml.WriteLine("</Folder>")

        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub

    Sub AntarcticFolder(connect As SqliteConnection, kml As StreamWriter)

        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, "Making Antarctic bases folder.")
        Application.DoEvents()

        kml.WriteLine("<Folder>")
        kml.WriteLine("<name>Antarctic bases</name>")
        kml.WriteLine("<description>Details of all active Antarctic bases. Data harvested from https://www.coolantarctica.com/Community/antarctic_bases.php</description>")
        kml.WriteLine("<visibility>0</visibility>")
        kml.WriteLine("<Style id=""antarctic""><LabelStyle><color>ffff7f7f</color></LabelStyle><IconStyle><Icon><href>https://maps.google.com/mapfiles/kml/paddle/blu-blank.png</href></Icon><scale>3</scale></IconStyle></Style>")

        sql = connect.CreateCommand()
        sql.CommandText = "SELECT * FROM Antarctic ORDER BY name"
        SQLdr = sql.ExecuteReader()

        While SQLdr.Read()

            Dim name = KMLescape(SQLdr("name"))
            Dim geoJson = SafeStr(SQLdr("coordinates"))

            ' Parse GeoJSON → NTS Point
            Dim ntsGeom As Geometry = FromGeoJsonToNTS(geoJson)
            Dim pt As Point = DirectCast(ntsGeom, Point)

            kml.WriteLine($"<Placemark><name>{name}</name><visibility>0</visibility><styleUrl>#antarctic</styleUrl>")

            ' Write point
            WriteKmlPoint(pt, kml, 2)

            ' Extended data
            kml.WriteLine("<ExtendedData>")
            kml.WriteLine($"<Data name=""name""><value>{name}</value></Data>")
            kml.WriteLine($"<Data name=""nation""><value>{KMLescape(SQLdr("nation"))}</value></Data>")
            kml.WriteLine($"<Data name=""lat""><value>{pt.Coordinate.Y:f3}</value></Data>")
            kml.WriteLine($"<Data name=""lon""><value>{pt.Coordinate.X:f3}</value></Data>")
            kml.WriteLine($"<Data name=""situation""><value>{KMLescape(SQLdr("situation"))}</value></Data>")
            kml.WriteLine($"<Data name=""altitude""><value>{SQLdr("altitude")}</value></Data>")
            kml.WriteLine($"<Data name=""open""><value>{KMLescape(SQLdr("open"))}</value></Data>")
            kml.WriteLine("</ExtendedData>")

            kml.WriteLine("</Placemark>")
        End While

        kml.WriteLine("</Folder>")
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")

    End Sub

    Sub EUASborder(kml As StreamWriter)
        ' add the EU/AS Russia border
        Dim doc As XDocument = XDocument.Load($"{Application.StartupPath}\BorderEUAS.kml")    ' read the XML
        Dim ns As XNamespace = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
        Dim nsmgr As New XmlNamespaceManager(New NameTable())
        nsmgr.AddNamespace("x", ns.NamespaceName)
        Dim placemark = doc.XPathSelectElement("//x:Placemark[2]", nsmgr).ToString       ' find second placemark
        placemark = Strings.Replace(placemark, $" xmlns=""{ns}""", "")         ' remove pesky namespace
        kml.WriteLine(placemark)                ' write border to kml file
    End Sub
    Private Function NtsDensify(poly As NetTopologySuite.Geometries.Polygon, degrees As Double) As NetTopologySuite.Geometries.Polygon
        If degrees <= 0 Then Return poly

        Dim tol = degrees * (Math.PI / 180.0) ' convert degrees to radians-ish spacing
        Dim densifier = NetTopologySuite.Densify.Densifier.Densify(poly, tol)
        Return DirectCast(densifier, NetTopologySuite.Geometries.Polygon)
    End Function
    Private Function WrapPolygonTo180(poly As NetTopologySuite.Geometries.Polygon) As NetTopologySuite.Geometries.Polygon
        Dim gf As New GeometryFactory()

        Dim wrap = Function(x As Double)
                       While x > 180 : x -= 360 : End While
                       While x < -180 : x += 360 : End While
                       Return x
                   End Function

        Dim ext = poly.ExteriorRing.Coordinates
        Dim newExt(ext.Length - 1) As Coordinate
        For i = 0 To ext.Length - 1
            newExt(i) = New Coordinate(wrap(ext(i).X), ext(i).Y)
        Next

        Dim holes As New List(Of LinearRing)
        For h = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(h).Coordinates
            Dim newHole(hole.Length - 1) As Coordinate
            For i = 0 To hole.Length - 1
                newHole(i) = New Coordinate(wrap(hole(i).X), hole(i).Y)
            Next
            holes.Add(gf.CreateLinearRing(newHole))
        Next

        Return gf.CreatePolygon(gf.CreateLinearRing(newExt), holes.ToArray())
    End Function

    Function NtsDensify(geom As Geometry, maxDeg As Double) As Geometry
        If geom Is Nothing OrElse geom.IsEmpty Then Return geom

        Dim factory = geom.Factory

        Select Case geom.GeometryType

            Case "LineString"
                Return DensifyLineString(DirectCast(geom, LineString), maxDeg, factory)

            Case "Polygon"
                Dim poly = DirectCast(geom, Polygon)

                Dim shell = DensifyLineString(poly.Shell, maxDeg, factory)
                Dim holes(poly.NumInteriorRings - 1) As LinearRing

                For i = 0 To poly.NumInteriorRings - 1
                    Dim hole As LinearRing = poly.GetInteriorRingN(i)
                Next

                Return factory.CreatePolygon(DirectCast(shell, LinearRing), holes)

            Case "MultiPolygon"
                Dim polys(geom.NumGeometries - 1) As Geometry
                For i = 0 To geom.NumGeometries - 1
                    polys(i) = NtsDensify(geom.GetGeometryN(i), maxDeg)
                Next
                Return factory.CreateMultiPolygon(polys.Cast(Of Polygon).ToArray())

            Case Else
                Return geom
        End Select
    End Function


    Private Function DensifyLineString(ls As LineString, maxDeg As Double, factory As GeometryFactory) As LineString
        Dim coords As New List(Of Coordinate)
        Dim pts = ls.Coordinates

        For i = 0 To pts.Length - 2
            Dim a = pts(i)
            Dim b = pts(i + 1)

            coords.Add(New Coordinate(a.X, a.Y))

            Dim dx = b.X - a.X
            Dim dy = b.Y - a.Y
            Dim dist = Math.Sqrt(dx * dx + dy * dy)

            If dist > maxDeg Then
                Dim steps = CInt(Math.Floor(dist / maxDeg))
                For s = 1 To steps
                    Dim t = s / (steps + 1)
                    coords.Add(New Coordinate(a.X + dx * t, a.Y + dy * t))
                Next
            End If
        Next

        coords.Add(New Coordinate(pts.Last.X, pts.Last.Y))

        Return factory.CreateLineString(coords.ToArray())
    End Function

    ' Primitives to convert GeoJson to KML
    Public Sub GeoJsonToKML(geoJson As String, kml As StreamWriter)

        Dim reader As New GeoJsonReader()
        Dim geom As Geometry = reader.Read(Of Geometry)(geoJson)

        Dim decimals As Integer
        If TypeOf geom Is Point Then
            decimals = 1
        Else
            decimals = GetPrecisionFromEnvelope(geom)
        End If

        Select Case geom.OgcGeometryType
            Case OgcGeometryType.Point
                WriteKmlPoint(DirectCast(geom, Point), kml, decimals)

            Case OgcGeometryType.LineString
                WriteKmlLineString(DirectCast(geom, LineString), kml, decimals)

            Case OgcGeometryType.Polygon
                WriteKmlPolygon(DirectCast(geom, Polygon), kml, decimals)

            Case OgcGeometryType.MultiPolygon
                WriteKmlMultiPolygon(DirectCast(geom, MultiPolygon), kml, decimals)

            Case Else
                Throw New NotSupportedException($"Geometry type {geom.OgcGeometryType} not supported.")
        End Select

    End Sub

    ''' <summary>
    ''' Determines the number of decimal places to use when exporting coordinates
    ''' to KML, based on the spatial extent of the geometry.  
    ''' <para>
    ''' Larger geometries require fewer decimal places, while smaller geometries
    ''' require more precision. This helps reduce file size without sacrificing
    ''' positional accuracy.
    ''' </para>
    ''' </summary>
    ''' <param name="geom">
    ''' The geometry whose bounding envelope is used to determine precision.
    ''' </param>
    ''' <returns>
    ''' An integer representing the number of decimal places to use when formatting
    ''' coordinate values.
    ''' </returns>
    Private Function GetPrecisionFromEnvelope(geom As Geometry) As Integer
        Dim env = geom.EnvelopeInternal
        Dim span = Math.Max(env.Width, env.Height)

        ' Simple heuristic – tune if you like
        If span > 1000 Then
            Return 1
        ElseIf span > 100 Then
            Return 2
        ElseIf span > 10 Then
            Return 3
        Else
            Return 4
        End If
    End Function
    ''' <summary>
    ''' Writes a KML <Point> element to the provided stream, using the specified
    ''' coordinate precision.  
    ''' <para>
    ''' Coordinates are written in <c>longitude,latitude</c> order as required by KML.
    ''' </para>
    ''' </summary>
    ''' <param name="pt">The NTS <see cref="Point"/> to export.</param>
    ''' <param name="kml">The <see cref="StreamWriter"/> receiving the KML output.</param>
    ''' <param name="decimals">The number of decimal places to use for coordinates.</param>

    Private Sub WriteKmlPoint(pt As Point, kml As StreamWriter, decimals As Integer)
        Dim lon = FormatCoord(pt.X, decimals)
        Dim lat = FormatCoord(pt.Y, decimals)

        kml.WriteLine($"<Point><coordinates>{lon},{lat}</coordinates></Point>")
    End Sub
    ''' <summary>
    ''' Writes a KML <LineString> element representing the supplied NTS geometry.
    ''' </summary>
    ''' <param name="ls">The <see cref="LineString"/> to export.</param>
    ''' <param name="kml">The output stream for the generated KML.</param>
    ''' <param name="decimals">The number of decimal places to use when formatting coordinates.</param>

    Private Sub WriteKmlLineString(ls As LineString, kml As StreamWriter, decimals As Integer)
        kml.WriteLine("<LineString>")
        kml.Write("  <coordinates>")
        WriteKmlCoordSequence(ls.CoordinateSequence, kml, decimals)
        kml.WriteLine("</coordinates>")
        kml.WriteLine("</LineString>")
    End Sub
    ''' <summary>
    ''' Writes a complete KML <Polygon> element, including outer and inner rings,
    ''' to the specified stream.  
    ''' <para>
    ''' The method handles both the exterior boundary and any interior holes,
    ''' formatting each ring as a <LinearRing> with the requested precision.
    ''' </para>
    ''' </summary>
    ''' <param name="poly">The polygon to export.</param>
    ''' <param name="kml">The stream writer receiving the KML output.</param>
    ''' <param name="decimals">The number of decimal places to use for coordinates.</param>

    Private Sub WriteKmlPolygon(poly As Polygon, kml As StreamWriter, decimals As Integer)
        kml.WriteLine("<Polygon><tessellate>1</tessellate>")

        ' Outer ring
        kml.WriteLine("  <outerBoundaryIs>")
        kml.WriteLine("    <LinearRing>")
        kml.Write("      <coordinates>")
        WriteKmlCoordSequence(poly.ExteriorRing.CoordinateSequence, kml, decimals)
        kml.WriteLine("</coordinates>")
        kml.WriteLine("    </LinearRing>")
        kml.WriteLine("  </outerBoundaryIs>")

        ' Holes
        For i = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(i)
            kml.WriteLine("  <innerBoundaryIs>")
            kml.WriteLine("    <LinearRing>")
            kml.Write("      <coordinates>")
            WriteKmlCoordSequence(hole.CoordinateSequence, kml, decimals)
            kml.WriteLine("</coordinates>")
            kml.WriteLine("    </LinearRing>")
            kml.WriteLine("  </innerBoundaryIs>")
        Next

        kml.WriteLine("</Polygon>")
    End Sub
    ''' <summary>
    ''' Writes a KML <MultiGeometry> element containing one or more <Polygon>
    ''' elements derived from the supplied <see cref="MultiPolygon"/>.  
    ''' <para>
    ''' Each polygon is exported using <see cref="WriteKmlPolygon"/>.
    ''' </para>
    ''' </summary>
    ''' <param name="mp">The multipolygon to export.</param>
    ''' <param name="kml">The output stream for the generated KML.</param>
    ''' <param name="decimals">The number of decimal places to use for coordinates.</param>

    Private Sub WriteKmlMultiPolygon(mp As MultiPolygon, kml As StreamWriter, decimals As Integer)
        kml.WriteLine("<MultiGeometry>")
        For i = 0 To mp.NumGeometries - 1
            Dim poly = DirectCast(mp.GetGeometryN(i), Polygon)
            WriteKmlPolygon(poly, kml, decimals)
        Next
        kml.WriteLine("</MultiGeometry>")
    End Sub
    Sub WriteKmlSingleOrMultiPolygon(g As Geometry, kml As StreamWriter, decimals As Integer)
        If TypeOf g Is Polygon Then
            Dim pg = DirectCast(g, Polygon)
            WriteKmlPolygon(pg, kml, 3)
        ElseIf TypeOf g Is MultiPolygon Then
            Dim mp = DirectCast(g, MultiPolygon)
            WriteKmlMultiPolygon(mp, kml, 3)
        Else
            Throw New Exception($"Unexpected geometry type {g.GeometryType}")
        End If
    End Sub
    ''' <summary>
    ''' Writes a sequence of coordinates in KML format using <c>longitude,latitude</c>
    ''' pairs, inserting a line break after every 10 coordinate pairs to improve
    ''' readability and reduce excessively long lines in the output.
    ''' <para>
    ''' Coordinates are formatted using the specified number of decimal places and
    ''' written in a space‑separated list, as required by the KML specification.
    ''' </para>
    ''' </summary>
    ''' <param name="seq">
    ''' The coordinate sequence to export. Each entry is written as a
    ''' <c>lon,lat</c> pair.
    ''' </param>
    ''' <param name="kml">
    ''' The <see cref="StreamWriter"/> receiving the generated KML coordinate text.
    ''' </param>
    ''' <param name="decimals">
    ''' The number of decimal places to apply when formatting each coordinate value.
    ''' </param>
    ''' <remarks>
    ''' This method does not write enclosing KML elements such as
    ''' <c>&lt;coordinates&gt;</c>; callers are responsible for wrapping the output
    ''' appropriately.
    ''' </remarks>
    Private Sub WriteKmlCoordSequence(seq As CoordinateSequence, kml As StreamWriter, decimals As Integer)

        Const MaxPerLine As Integer = 10
        Dim countOnLine As Integer = 0

        For i = 0 To seq.Count - 1

            ' Insert space between pairs (except at start of line)
            If countOnLine > 0 Then
                kml.Write(" ")
            End If

            ' Write lon,lat
            Dim x = FormatCoord(seq.GetX(i), decimals)
            Dim y = FormatCoord(seq.GetY(i), decimals)
            kml.Write($"{x},{y}")

            countOnLine += 1

            ' If we hit 10 pairs, break line
            If countOnLine = MaxPerLine AndAlso i < seq.Count - 1 Then
                kml.WriteLine()
                countOnLine = 0
            End If

        Next

    End Sub

    ''' <summary>
    ''' Formats a numeric coordinate value using invariant culture and the specified
    ''' number of decimal places.  
    ''' <para>
    ''' Ensures consistent output for KML and GeoJSON regardless of system locale.
    ''' </para>
    ''' </summary>
    ''' <param name="value">The coordinate value to format.</param>
    ''' <param name="decimals">The number of decimal places to include.</param>
    ''' <returns>
    ''' A string containing the formatted coordinate.
    ''' </returns>

    Private Function FormatCoord(value As Double, decimals As Integer) As String
        Dim fmt = "F" & decimals.ToString()
        Return value.ToString(fmt, CultureInfo.InvariantCulture)
    End Function
    Public Function DecimalsForBbox(minX As Double, minY As Double,
                                maxX As Double, maxY As Double) As Integer

        ' Handle antimeridian wrap
        Dim width As Double
        If minX > maxX Then
            ' Example: minX = 170, maxX = -170 → actual width = 20°
            width = (maxX + 360.0) - minX
        Else
            width = maxX - minX
        End If

        Dim height As Double = maxY - minY
        Dim w = Math.Max(Math.Abs(width), Math.Abs(height))

        If w >= 10 Then Return 1
        If w >= 1 Then Return 2
        If w >= 0.1 Then Return 3
        If w >= 0.01 Then Return 4
        If w >= 0.001 Then Return 5
        Return 6

    End Function

End Module
