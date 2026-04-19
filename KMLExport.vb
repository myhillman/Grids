Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml
Imports System.Xml.XPath
Imports System.Linq
Imports DotSpatial.Projections.Transforms
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite

Module KMLExport

    Const CoordsPerLine = 10          ' coordinates per line
    Private prefixes As New List(Of (p As MapPoint, pfx As String))         ' prefixes for prefix folder
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
                Using kml As New StreamWriter($"{Application.StartupPath}\KML\DXCC_{entity}.kml", False)
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

        Dim BaseFilename As String = $"{Application.StartupPath}\KML\DXCC Map of the World"
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
            'ZoneFolder(connect, kml)       ' data is lost
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
        ' create KML for one placemark
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim tagStack As New Stack(Of String)    ' FIFO for end tags

        sql = connect.CreateCommand
        sql.CommandText = $"SELECT * FROM `DXCC` WHERE `DXCCnum`={dxcc}"     ' fetch all geometry
        SQLdr = sql.ExecuteReader
        SQLdr.Read()
        Dim Entity = SQLdr("Entity")
        Dim boundaries As Polygon = GeoJsonToGeometry(SafeStr(SQLdr("geometry")))
        If boundaries.IsEmpty Then  ' There is no geometry to convert
            Dim errMsg = $"There is no geometry for {SQLdr("Entity")} to convert to KML"
            AppendText(Form1.TextBox1, $"{errMsg}{vbCrLf}")
            kml.WriteLine(errMsg)
            Exit Sub
        End If
        AppendText(Form1.TextBox1, $"Making KML for {Entity}{vbCrLf}")
        kml.WriteLine($"<Placemark><styleUrl>#boundary_{SQLdr("colour")}</styleUrl>")
        kml.WriteLine($"<name>{KMLescape(Entity)} ({SQLdr("prefix")})</name>")
        ' Get a label point for the entity 
        Dim labelPoint As MapPoint = GeometryEngine.LabelPoint(boundaries)
        kml.Write($"<Point>")
        KMLcoordinates(kml, labelPoint, 0)
        kml.WriteLine("</Point>")
        prefixes.Add((labelPoint, SQLdr("prefix")))     ' for the prefix folder
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
        Dim query As New StringBuilder()
        If Not IsDBNullorEmpty(SQLdr("source")) Then query.Append($"{SQLdr("source").ToString().Trim()}: ")
        If Not IsDBNullorEmpty(SQLdr("rule")) Then query.Append(SQLdr("rule").ToString().Trim())
        If Not IsDBNullorEmpty(SQLdr("bbox")) Then query.Append($" ({SQLdr("bbox")})")
        If query.Length > 0 Then
            kml.WriteLine($"<Data name=""GIS query""><value>{KMLescape(query.ToString)}</value></Data>")
        End If
        If Not IsDBNullorEmpty(SQLdr("notes")) Then
            kml.WriteLine($"<Data name=""Notes""><value><![CDATA[{Hyperlink(SQLdr("notes"))}]]></value></Data>")
        End If
        kml.WriteLine("</ExtendedData>")
        KMLPolygon(kml, boundaries, 3, False)
        kml.WriteLine("</Placemark>")
    End Sub

    Sub PrefixFolder(kml As StreamWriter)
        ' Create the Prefix folder
        Dim timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, $"Making KML for {prefixes.Count} prefixes.")
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
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

Sub BoundingBoxFolder(connect As SqliteConnection, kml As StreamWriter)
    Const DensifyDegrees = 10
    Dim labelpoint As MapPoint
    Dim timer As New Stopwatch

    timer.Start()
    AppendText(Form1.TextBox1, $"Making KML for bounding boxes.")
    kml.WriteLine("<Folder>")
    kml.WriteLine("<name>Bounding Boxes</name>")
    kml.WriteLine("<description>Bounding boxes sometimes necessary to filter query. For debugging only.</description>")
    kml.WriteLine("<visibility>0</visibility>")
    kml.WriteLine("<Style id=""bbox"">
              <IconStyle>
                <scale>0</scale>
              </IconStyle>
              <LabelStyle>
                <color>ffffffff</color>
                <scale>1.2</scale>
              </LabelStyle>
              <LineStyle>
                <color>ffffffff</color>
                <width>2</width>
              </LineStyle>
              <PolyStyle>
                <color>00ffffff</color>
                <fill>0</fill>
                <outline>1</outline>
              </PolyStyle>
            </Style>")

    Dim sql = connect.CreateCommand
    sql.CommandText = "Select * FROM `DXCC` WHERE `Deleted`=0 And `bbox` Is Not NULL and bbox <> '' ORDER BY `Entity`"
    Dim SQLdr = sql.ExecuteReader

    While SQLdr.Read
        Dim entity = SafeStr(SQLdr("Entity"))
        DebugOutput = (entity = "Fiji")

        Dim bbox = ParseBBoxOrPoly(SQLdr("bbox"))

        kml.WriteLine("<Placemark>")
        kml.WriteLine($"  <name>{KMLescape(entity)}</name>")
        kml.WriteLine("  <styleUrl>#bbox</styleUrl>")

        ' Label point
        If TypeOf bbox Is Envelope Then
            With bbox.Extent
                Dim cx = (.XMin + .XMax) / 2.0
                Dim cy = (.YMin + .YMax) / 2.0
                labelpoint = New MapPoint(cx, cy, SpatialReferences.Wgs84)
            End With
        Else
            labelpoint = GeometryEngine.LabelPoint(bbox)
        End If
        kml.WriteLine($"  <Point><coordinates>{labelpoint.X:f2},{labelpoint.Y:f2}</coordinates></Point>")
        kml.WriteLine("  <MultiGeometry>")

        ' Collect polygons exactly as ParseBBoxOrPoly produced them
        Dim polys As New List(Of Polygon)

        If TypeOf bbox Is Envelope Then
            Dim ring As New List(Of MapPoint) From {
                New MapPoint(bbox.Extent.XMin, bbox.Extent.YMin),
                New MapPoint(bbox.Extent.XMax, bbox.Extent.YMin),
                New MapPoint(bbox.Extent.XMax, bbox.Extent.YMax),
                New MapPoint(bbox.Extent.XMin, bbox.Extent.YMax),
                New MapPoint(bbox.Extent.XMin, bbox.Extent.YMin)
            }
            Dim pb As New PolygonBuilder(bbox.SpatialReference)
            pb.AddPart(ring)
            polys.Add(pb.ToGeometry())

        ElseIf TypeOf bbox Is Polygon Then
            polys.Add(CType(bbox, Polygon))

        ElseIf TypeOf bbox Is Multipart Then
            Dim mp = CType(bbox, Multipart)
            For Each part In mp.Parts
                Dim pb As New PolygonBuilder(mp.SpatialReference)
                pb.AddPart(part.Points)
                polys.Add(pb.ToGeometry())
            Next
        End If

        If DebugOutput Then Debug.WriteLine($"Fiji polys from bbox: {polys.Count}")

        ' Densify + wrap + write each polygon as‑is
        For Each p In polys
            Dim finalPoly = DensifyAndWrapPolygon(p, DensifyDegrees)

            If DebugOutput Then
                Debug.WriteLine("---- RING ----")
                For Each part In finalPoly.Parts
                    Debug.WriteLine("---- RING 0 ----")
                    For Each mp In part.Points
                        Debug.WriteLine($"{mp.X}, {mp.Y}")
                    Next
                Next
            End If

            Dim ext = finalPoly.Extent
            Dim dec = DecimalsForBbox(ext.XMin, ext.YMin, ext.XMax, ext.YMax)
            KMLPolygon(kml, finalPoly, dec, True)
        Next

        kml.WriteLine("  </MultiGeometry>")
        kml.WriteLine("</Placemark>")
    End While

    SQLdr.Close()
    kml.WriteLine("</Folder>")
    timer.Stop()
    AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
End Sub
    Private Function DensifyAndWrapPolygon(src As Polygon,
                                           densifyDegrees As Double) As Polygon

        Dim sr = src.SpatialReference

        ' Densify as polyline
        Dim lb As New PolylineBuilder(sr)
        For Each part In src.Parts
            lb.AddPart(part.Points)
        Next
        Dim denseLine = GeometryEngine.Densify(lb.ToGeometry(), densifyDegrees)

        ' Collect densified points (single ring per polygon in our bbox case)
        Dim densePts As New List(Of MapPoint)
        For Each p In CType(denseLine, Polyline).Parts(0).Points
            densePts.Add(p)
        Next

        ' Wrap to [-180,180]
        Dim wrapped As New List(Of MapPoint)
        For Each p In densePts
            Dim x = p.X
            While x > 180 : x -= 360 : End While
            While x < -180 : x += 360 : End While
            wrapped.Add(New MapPoint(x, p.Y, sr))
        Next

        ' Build final polygon
        Dim pb As New PolygonBuilder(SpatialReferences.Wgs84)
        pb.AddPart(wrapped)
        Return pb.ToGeometry()
    End Function
    Private Function SplitWrappedForKml(ring As List(Of MapPoint)) As List(Of List(Of MapPoint))
        Const AM As Double = 180.0
        Const EPS As Double = 1e-9

        Dim parts As New List(Of List(Of MapPoint))
        Dim current As New List(Of MapPoint)
        current.Add(ring(0))

        For i = 1 To ring.Count - 1
            Dim p1 = current.Last()
            Dim p2 = ring(i)

            Dim x1 = p1.X
            Dim x2 = p2.X

            Dim touchesOrCrosses = _
                    ((x1 <= AM + EPS AndAlso x2 >= AM - EPS) OrElse _
                     (x1 >= AM - EPS AndAlso x2 <= AM + EPS)) _
                    AndAlso (Math.Abs(x1 - x2) > EPS)

            If touchesOrCrosses Then
                ' intersection at x = 180
                Dim t = (AM - x1) / (x2 - x1)
                Dim y = p1.Y + t * (p2.Y - p1.Y)

                current.Add(New MapPoint(AM, y, p1.SpatialReference))
                parts.Add(New List(Of MapPoint)(current))

                current = New List(Of MapPoint)
                current.Add(New MapPoint(AM, y, p1.SpatialReference))
            End If

            current.Add(p2)
        Next

        parts.Add(current)

        ' close rings
        For Each prt In parts
            Dim f = prt(0)
            Dim l = prt(prt.Count - 1)
            If f.X <> l.X OrElse f.Y <> l.Y Then prt.Add(f)
        Next

        Return parts
    End Function


    Private Function PreparePolygonForKml(poly As Polygon,
                                          densifyDegrees As Double) As List(Of Polygon)

        Dim sr = poly.SpatialReference
        Dim result As New List(Of Polygon)

        For Each part In poly.Parts
            Dim rawRing = part.Points.ToList()

            ' 1. unwrap for safe geometry ops
            Dim unwrapped = UnwrapRing(rawRing)
            Dim ccw = NormalizeRingCCW(unwrapped)

            ' 2. densify in unwrapped space (optional)
            Dim workRing = ccw
            If densifyDegrees > 0 Then
                Dim lb As New PolylineBuilder(sr)
                lb.AddPart(workRing)
                Dim denseLine = GeometryEngine.Densify(lb.ToGeometry(), densifyDegrees)

                Dim densePts As New List(Of MapPoint)
                For Each p In CType(denseLine, Polyline).Parts(0).Points
                    densePts.Add(p)
                Next
                workRing = densePts
            End If

            ' 3. wrap for export
            Dim wrapped = WrapBackTo180(workRing)

            ' 4. split in WRAPPED space for KML
            Dim splitParts = SplitWrappedForKml(wrapped)

            ' 5. build polygons
            For Each r In splitParts
                Dim pb As New PolygonBuilder(sr)
                pb.AddPart(r)
                result.Add(pb.ToGeometry())
            Next
        Next

        Return result
    End Function

    ' Unwrap a ring so no edge has |Δlon| > 180
Private Function UnwrapRing(ring As List(Of MapPoint)) As List(Of MapPoint)
    Dim out As New List(Of MapPoint)
    If ring.Count = 0 Then Return out

    Dim prev = ring(0).X
    Dim offset As Double = 0

    out.Add(New MapPoint(prev, ring(0).Y, ring(0).SpatialReference))

    For i = 1 To ring.Count - 1
        Dim lon = ring(i).X
        Dim lat = ring(i).Y
        Dim d = lon - prev

        If d > 180 Then offset -= 360
        If d < -180 Then offset += 360

        out.Add(New MapPoint(lon + offset, lat, ring(i).SpatialReference))
        prev = lon
    Next

    Return out
End Function

' Enforce CCW orientation on an unwrapped ring
Private Function NormalizeRingCCW(ring As List(Of MapPoint)) As List(Of MapPoint)
    Dim pts = New List(Of MapPoint)(ring)

    ' ensure closed
    If pts.Count > 0 Then
        Dim f = pts(0)
        Dim l = pts(pts.Count - 1)
        If f.X <> l.X OrElse f.Y <> l.Y Then pts.Add(f)
    End If

    ' signed area
    Dim area As Double = 0
    For i = 0 To pts.Count - 2
        Dim p1 = pts(i)
        Dim p2 = pts(i + 1)
        area += (p1.X * p2.Y) - (p2.X * p1.Y)
    Next
    area /= 2.0

    If area < 0 Then pts.Reverse()
    Return pts
End Function

' Split an UNWRAPPED ring at x = 180 (antimeridian)
Private Function SplitRingAt180Unwrapped(ring As List(Of MapPoint)) As List(Of List(Of MapPoint))
    Const AM As Double = 180.0
    Const EPS As Double = 1e-9

    Dim parts As New List(Of List(Of MapPoint))
    If ring Is Nothing OrElse ring.Count = 0 Then Return parts

    Dim current As New List(Of MapPoint)
    current.Add(ring(0))

    For i = 1 To ring.Count - 1
        Dim p1 = current.Last()
        Dim p2 = ring(i)

        Dim x1 = p1.X
        Dim y1 = p1.Y
        Dim x2 = p2.X
        Dim y2 = p2.Y

        Dim crosses = (x1 < AM - EPS AndAlso x2 > AM + EPS) OrElse
                      (x1 > AM + EPS AndAlso x2 < AM - EPS)

        If crosses Then
            Dim t = (AM - x1) / (x2 - x1)
            Dim y = y1 + t * (y2 - y1)

            current.Add(New MapPoint(AM, y, p1.SpatialReference))
            parts.Add(New List(Of MapPoint)(current))

            current = New List(Of MapPoint)
            current.Add(New MapPoint(AM, y, p1.SpatialReference))
        End If

        current.Add(p2)
    Next

    parts.Add(current)

    ' close each part
    Dim finalParts As New List(Of List(Of MapPoint))
    For Each prt In parts
        If prt.Count < 3 Then Continue For
        Dim f = prt(0)
        Dim l = prt(prt.Count - 1)
        If f.X <> l.X OrElse f.Y <> l.Y Then prt.Add(f)
        finalParts.Add(prt)
    Next

    Return finalParts
End Function

' Wrap longitudes back to [-180,180] for export
Private Function WrapBackTo180(ring As List(Of MapPoint)) As List(Of MapPoint)
    Dim out As New List(Of MapPoint)
    For Each p In ring
        Dim x = p.X
        While x > 180 : x -= 360 : End While
        While x < -180 : x += 360 : End While
        out.Add(New MapPoint(x, p.Y, p.SpatialReference))
    Next
    Return out
End Function

    Public Sub DumpPolygonParts(poly As Polygon, Optional label As String = "")
        If Not String.IsNullOrEmpty(label) Then
            Debug.WriteLine($"--- {label} ---")
        End If

        If poly Is Nothing Then
            Debug.WriteLine("Polygon = NULL")
            Return
        End If

        Debug.WriteLine($"Parts: {poly.Parts.Count}")

        Dim partIndex As Integer = 0

        For Each part In poly.Parts
            partIndex += 1

            Dim pts = part.Points
            Debug.WriteLine($" Part {partIndex}: {pts.Count} points")

            If pts.Count > 0 Then
                Dim p0 = pts(0)
                Dim p1 = pts(Math.Max(0, pts.Count - 1))

                Debug.WriteLine($"   First: {p0.X}, {p0.Y}")
                Debug.WriteLine($"   Last : {p1.X}, {p1.Y}")
            End If
        Next

        Debug.WriteLine("-----------------------------")
    End Sub

    Function DecimalsForBbox(minX As Double, minY As Double,
                             maxX As Double, maxY As Double) As Integer

        Dim dx As Double

        ' Detect wrap BEFORE rounding
        If minX > maxX Then
            dx = (180 - minX) + (maxX + 180)
        Else
            dx = maxX - minX
        End If

        Dim dy = Math.Abs(maxY - minY)
        Dim w = Math.Max(dx, dy)

        If w >= 10 Then Return 1
        If w >= 1 Then Return 2
        If w >= 0.1 Then Return 3
        If w >= 0.01 Then Return 4
        If w >= 0.001 Then Return 5
        Return 6
    End Function

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
        If points Is Nothing OrElse points.Count = 0 Then
            Throw New InvalidOperationException("Ring part is empty.")
        End If

        If points(0).SpatialReference Is Nothing OrElse
   Not points(0).SpatialReference.Equals(SpatialReferences.Wgs84) Then

            Throw New InvalidOperationException($"Spatial reference must be WGS84 is {points(0).SpatialReference.ToString}")
        End If

        If digits < 0 OrElse digits > 8 Then
            Throw New ArgumentOutOfRangeException(NameOf(digits), $"Digits value {digits} is outside the valid range 0–8.")
        End If

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

        Dim parts = poly.Parts.ToList()

        ' Count outer rings (clockwise)
        Dim outerCount As Integer = 0
        For Each prt In parts
            If PolygonArea(prt.Points) > 0 Then outerCount += 1
        Next

        ' If more than one outer ring → MultiGeometry
        Dim useMulti As Boolean = (outerCount > 1 AndAlso Not MultiGeometryOpen)
        If useMulti Then kml.WriteLine("<MultiGeometry>")

        Dim polygonOpen As Boolean = False

        For Each prt In parts

            Dim area = PolygonArea(prt.Points)
            Dim isOuter As Boolean = (area > 0)

            If isOuter Then

                ' Close previous polygon if open
                If polygonOpen Then
                    kml.WriteLine("</Polygon>")
                End If

                ' Start new polygon
                kml.WriteLine("<Polygon><tessellate>1</tessellate>")
                kml.WriteLine("<outerBoundaryIs>")
                polygonOpen = True

                ' Write outer ring
                kml.WriteLine("<LinearRing>")
                Dim pts = prt.Points.ToList()
                If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))
                KMLcoordinates(kml, pts, digits)
                kml.WriteLine("</LinearRing>")
                kml.WriteLine("</outerBoundaryIs>")

            Else
                ' Inner ring (hole)
                If Not polygonOpen Then
                    ' Defensive: inner ring without outer → ignore
                    Continue For
                End If

                kml.WriteLine("<innerBoundaryIs>")
                kml.WriteLine("<LinearRing>")
                Dim pts = prt.Points.ToList()
                If Not pts(0).IsEqual(pts.Last()) Then pts.Add(pts(0))
                KMLcoordinates(kml, pts, digits)
                kml.WriteLine("</LinearRing>")
                kml.WriteLine("</innerBoundaryIs>")

            End If

        Next

        ' Close last polygon
        If polygonOpen Then
            kml.WriteLine("</Polygon>")
        End If

        If useMulti Then kml.WriteLine("</MultiGeometry>")

    End Sub


    Function CrossesAntiMeridian(g As Geometry) As Boolean
        ' returns true if geometry crosses anti-meridian
        Return Math.Abs(g.Extent.XMax - g.Extent.XMin) >= 300
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
    Function NormalizeAntiMeridian(g As Geometry) As Geometry
        If g Is Nothing Then Return Nothing

        ' --- Case 1: Envelope (fast path) ---
        If TypeOf g Is Envelope Then
            Dim env = DirectCast(g, Envelope)

            Dim xmin = NormalizeLon(env.XMin)
            Dim xmax = NormalizeLon(env.XMax)

            ' Rebuild envelope with normalized longitudes
            Return New Envelope(xmin, env.YMin, xmax, env.YMax, env.SpatialReference)
        End If

        ' --- Case 2: Polygon (your existing logic) ---
        If TypeOf g Is Polygon Then
            Dim poly = DirectCast(g, Polygon)

            If poly.Parts.Count = 0 Then Return poly
            If Not CrossesAntiMeridian(poly) Then Return poly

            Dim pb As New PolygonBuilder(poly.SpatialReference)

            For Each Prt In poly.Parts
                Dim pts As New List(Of MapPoint)

                For Each p In Prt.Points
                    Dim lng = NormalizeLon(p.X)
                    pts.Add(New MapPoint(lng, p.Y, p.SpatialReference))
                Next

                pb.AddPart(pts)
            Next

            Dim fixedPoly = pb.ToGeometry
            Return GeometryEngine.Simplify(fixedPoly)
        End If

        ' --- Case 3: Other geometry types (Polyline, Point, etc.) ---
        Return g
    End Function
    Public Function NormalizeLon(lon As Double) As Double
        ' Bring longitude into the -180..180 range
        While lon > 180
            lon -= 360
        End While

        While lon < -180
            lon += 360
        End While

        Return lon
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
        AppendText(Form1.TextBox1, $"Making KML for Grid Square folder ")
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
            Dim geometry = GeoJsonToGeometry(SQLdr("geometry"))          ' get the geometry
            Dim Entity = SQLdr("Entity")
            SQLdr.Close()
            Dim extent = geometry.Extent
            ' Create an envelope which exactly covers all the possible grid squares
            Dim GridExtent As New Envelope(Math.Floor(extent.XMin / GridSquareX) * GridSquareX, Math.Floor(extent.YMin / GridSquareY) * GridSquareY, Math.Ceiling(extent.XMax / GridSquareX) * GridSquareX, Math.Ceiling(extent.YMax / GridSquareY) * GridSquareY, SpatialReferences.Wgs84)
            ' test every gridsquare with the grid extent
            Dim X = GridExtent.XMin
            While X < GridExtent.XMax
                Dim Y = GridExtent.YMin
                While Y < GridExtent.YMax
                    Dim GridSq As New Envelope(X, Y, X + GridSquareX - EPS, Y + GridSquareY - EPS, SpatialReferences.Wgs84)
                    If GridSq.Intersects(geometry) Then
                        Dim gs = GridSquare(New MapPoint(X, Y, SpatialReferences.Wgs84))
                        If Land.ContainsKey(gs) Then
                            GridSquares.Add(gs, GridSq)
                        Else
                            Ocean += 1      ' gridsquare is not on land
                        End If
                    End If
                    Y += GridSquareY
                End While
                X += GridSquareX
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
                        Dim AboveGS = GridSquare(New MapPoint(.XMin, .YMin + GridSquareY, SpatialReferences.Wgs84))       ' grid square above
                        If Not GridSquares.ContainsKey(AboveGS) Then mpb.Points.Add(New MapPoint(.XMin, .YMax))       ' first point
                        mpb.Points.Add(New MapPoint(.XMax, .YMax))
                        mpb.Points.Add(New MapPoint(.XMax, .YMin))
                        mpb.Points.Add(New MapPoint(.XMin, .YMin))
                        Dim LeftGS = GridSquare(New MapPoint((.XMin - GridSquareX) Mod 180, .YMin, SpatialReferences.Wgs84))       ' grid square to the left
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
        AppendText(Form1.TextBox1, $" {Ocean} GS eliminated, {Drawn.Count:n0} Unique grid squares [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub IOTAFolder(connect As SqliteConnection, kml As StreamWriter)
        ' Create the Grid Square folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch
        Dim sql1 As SqliteCommand, SQLdr1 As SqliteDataReader, count As Integer = 0

        timer.Start()
        sql = connect.CreateCommand
        sql1 = connect.CreateCommand
        ' make the folder
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
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            count += 1
            Dim refno = SafeStr(SQLdr("refno"))
            kml.WriteLine($"<Placemark><name>{refno}</name><styleUrl>#iota</styleUrl><visibility>0</visibility>")
            ' Retrieve the islands of this group
            sql1.CommandText = $"SELECT GROUP_CONCAT(name||(iif(comment != '',' ('||comment||')','')),', ') AS Islands FROM IOTA_Islands WHERE refno='{refno}'"
            SQLdr1 = sql1.ExecuteReader
            SQLdr1.Read()
            Dim Islands = SQLdr1("Islands")
            SQLdr1.Close()
            Dim comment As String = IIf(IsDBNullorEmpty(SQLdr("comment")), "", $"{SQLdr("comment")}{vbCrLf}{vbCrLf}")      ' add optional comment
            kml.WriteLine($"<description>{KMLescape(comment)}{KMLescape(Islands)}</description>")
            Dim polys As New List(Of Polygon)
            ' create the bounding box polygon
            Dim LonMin = SafeDbl(SQLdr("longitude_min"))
            Dim LatMin = SafeDbl(SQLdr("latitude_min"))
            Dim LonMax = SafeDbl(SQLdr("longitude_max"))
            Dim LatMax = SafeDbl(SQLdr("latitude_max"))
            If LonMin < LonMax Then           ' Normal
                Dim pgb As New Esri.ArcGISRuntime.Geometry.PolygonBuilder(SpatialReferences.Wgs84)       ' drawn bounding box CCW
                pgb.AddPoint(New MapPoint(LonMin, LatMax))   ' top-left
                pgb.AddPoint(New MapPoint(LonMin, LatMin))   ' bottom-left
                pgb.AddPoint(New MapPoint(LonMax, LatMin))   ' bottom-right
                pgb.AddPoint(New MapPoint(LonMax, LatMax))   ' top-right
                pgb.AddPoint(New MapPoint(LonMin, LatMax))   ' close
                polys.Add(pgb.ToGeometry)
            Else
                ' crosses anti-meridian. Make two rings
                ' West part: LonMin → 180
                Dim west As New PolygonBuilder(SpatialReferences.Wgs84)
                west.AddPart(New List(Of MapPoint) From {
                    New MapPoint(LonMin, LatMax),   ' top‑left
                    New MapPoint(LonMin, LatMin),   ' bottom‑left
                    New MapPoint(179.9999, LatMin),   ' bottom‑right
                    New MapPoint(179.9999, LatMax),    ' top‑right
                    New MapPoint(LonMin, LatMax)   ' close
                })
                Dim w = west.ToGeometry()
                If w IsNot Nothing AndAlso Not w.IsEmpty Then polys.Add(w)

                Dim east As New PolygonBuilder(SpatialReferences.Wgs84)
                ' East part: -180 → LonMax
                east.AddPart(New List(Of MapPoint) From {
                    New MapPoint(-179.9999, LatMax),    ' top‑left
                    New MapPoint(-179.9999, LatMin),    ' bottom‑left
                    New MapPoint(LonMax, LatMin),   ' bottom‑right
                    New MapPoint(LonMax, LatMax),   ' top‑right
                    New MapPoint(-179.9999, LatMax)   ' close
                })
                Dim e = east.ToGeometry()
                If e IsNot Nothing AndAlso Not e.IsEmpty Then polys.Add(e)
            End If
            Dim labelpoint As MapPoint = polys(0).LabelPoint
            kml.WriteLine("<MultiGeometry>")
            kml.Write($"<Point>")
            KMLcoordinates(kml, labelpoint, 1)
            kml.WriteLine("</Point>")
            For Each poly In polys
                poly = DensifyCleanMergeOrient(poly, 5)
                KMLPolygon(kml, poly, 2, True)
            Next
            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine($"</Placemark>")
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
        ' Create linestrings and lables for CQ or ITU zones
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader

        sql = connect.CreateCommand
        ' Create the folder
        kml.WriteLine($"<Folder><name>{zone} Zones</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<description>Polygons describing {zone} Zones. Based on data extracted from http://zone-check.eu/.</description>")
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
            Dim poly As Polygon = GeoJsonToGeometry(SQLdr("geometry"))            ' retrieve geometry
            If poly.Parts.Count > 0 Then
                Dim labelpoint As MapPoint = poly.LabelPoint    ' calculate labelpoint
                kml.WriteLine("<MultiGeometry>")
                kml.Write($"<Point>")
                KMLcoordinates(kml, labelpoint, 1)
                kml.WriteLine("</Point>")
                poly = NormalizeAntiMeridian(poly)
                poly = DensifyCleanMergeOrient(poly, 5)
                KMLPolygon(kml, poly, 2, True)
                kml.WriteLine("</MultiGeometry>")
            End If
            kml.WriteLine("</Placemark>")
        End While
        SQLdr.Close()
        kml.WriteLine("</Folder>")
    End Sub
    Public Sub TimeZoneFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make a folder for Timezones. 
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, $"Making KML for Timezone folder.")
        Application.DoEvents()
        sql = connect.CreateCommand
        kml.WriteLine("<Folder><name>Timezones</name><visibility>0</visibility><open>0</open>")
        kml.WriteLine("<description>Polygons describing timezones. Data obtained from https://www.naturalearthdata.com/</description>")

        ' create styles for the 6 different colour polygons
        For i = 1 To 6
            kml.WriteLine($"<StyleMap id='TZ_{i}'>")
            kml.WriteLine($"<Pair><key>normal</key><styleUrl>#{ColourMapping(i)}</styleUrl></Pair>")
            kml.WriteLine("<Pair><key>highlight</key><styleUrl>#white</styleUrl></Pair>")
            kml.WriteLine("</StyleMap>")
        Next
        ' get list of all unique timezones
        Dim TZlist As New List(Of (name As String, color As Integer, places As String))
        sql.CommandText = "SELECT name,color,GROUP_CONCAT(places,', ') AS places FROM Timezones GROUP BY name"
        SQLdr = sql.ExecuteReader
        While SQLdr.Read
            TZlist.Add((SQLdr("name"), SQLdr("color"), SQLdr("places")))
        End While
        SQLdr.Close()
        TZlist.Sort(Function(x, y) CSng(x.name).CompareTo(CSng(y.name)))      ' sort in numerical order
        ' Now retrieve each timezone
        For Each tz In TZlist
            kml.WriteLine($"<Placemark><visibility>0</visibility><styleUrl>#TZ_{tz.color}</styleUrl><name>{tz.name}</name><description>{tz.places}</description>")

            kml.WriteLine("<MultiGeometry>")
            sql.CommandText = $"SELECT * FROM Timezones WHERE name='{tz.name}'"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Dim geom As Polygon = GeoJsonToGeometry(SQLdr("geometry"))
                'Debug.WriteLine($"""{tz.name}"",""{tz.places}"",{geom.Parts.Count}")
                KMLPolygon(kml, geom, 1, True)
            End While
            SQLdr.Close()
            kml.WriteLine("</MultiGeometry>")
            kml.WriteLine("</Placemark>")
        Next
        kml.WriteLine("</Folder>")
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub

    Sub IARUFolder(connect As SqliteConnection, kml As StreamWriter)
        ' make the IARU regions folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, "Making KML for IARU folder.")
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
        AppendText(Form1.TextBox1, $" [{timer.Elapsed.Seconds:f1}s]{vbCrLf}")
    End Sub
    Sub AntarcticFolder(connect As SqliteConnection, kml As StreamWriter)
        ' Create the Antarctic bases folder
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, timer As New Stopwatch

        timer.Start()
        AppendText(Form1.TextBox1, "Making Antarctic bases folder.")
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

End Module
