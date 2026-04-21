Imports NetTopologySuite.Geometries
Imports Microsoft.Data.Sqlite
Imports System.IO

Module Reports
    Sub EntityReport()
        ' Create list of entities and state of geometry
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim ringCount As Integer, pointCount As Integer, JSONcount As Integer, lines As Integer = 0
        Dim TotalEntity As Integer = 0, TotalRings As Integer = 0, TotalPoints As Integer = 0, TotalQueries As Integer = 0, TotalJSON As Integer = 0
        Dim TotalUnclosed As Integer = 0, TotalNotes As Integer = 0, TotalBbox As Integer = 0
        Dim outer As Integer, inner As Integer, unclosed As Integer

        Dim htmlFile = Path.Combine(Application.StartupPath, "DXCC_Entity_report.html")

        Using connect As New SqliteConnection(DXCC_DATA),
          html As New StreamWriter(htmlFile, False)

            connect.Open()
            sql = connect.CreateCommand

            ' Find total entries in report
            sql.CommandText = "select count(*) as Total from DXCC where Deleted=0 and DXCCnum != 999"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()

            With Form1.ProgressBar1
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
            html.WriteLine("<tr><th>Entity</th><th>Rings</th><th>Outer</th><th>Inner</th><th>Points</th><th>Unclosed</th><th>JSON</th><th>OSM Overpass QL</th><th>bbox</th><th>Notes</th></tr>")

            While SQLdr.Read
                Form1.ProgressBar1.Value += 1

                ringCount = 0
                pointCount = 0
                JSONcount = 0
                outer = 0
                inner = 0
                unclosed = 0

                If Not IsDBNullorEmpty(SQLdr("query")) Then TotalQueries += 1

                Dim cls = " class=red"

                ' ------------------------------
                ' GEOMETRY PROCESSING (NTS)
                ' ------------------------------
                If Not IsDBNullorEmpty(SQLdr("geometry")) Then
                    Dim geom As Geometry = FromGeoJsonToNTS(SQLdr("geometry"))

                    If Not geom.IsEmpty Then cls = ""

                    JSONcount = SQLdr("geometry").ToString.Length
                    TotalJSON += JSONcount

                    ' Handle Polygon or MultiPolygon
                    If TypeOf geom Is Polygon Then
                        ProcessPolygon(DirectCast(geom, Polygon), ringCount, pointCount, outer, inner, unclosed)

                    ElseIf TypeOf geom Is MultiPolygon Then
                        Dim mp = DirectCast(geom, MultiPolygon)
                        For i = 0 To mp.NumGeometries - 1
                            ProcessPolygon(DirectCast(mp.GetGeometryN(i), Polygon), ringCount, pointCount, outer, inner, unclosed)
                        Next
                    End If

                    TotalUnclosed += unclosed
                End If

                TotalEntity += 1
                TotalRings += ringCount
                TotalPoints += pointCount

                ' ------------------------------
                ' BBOX + NOTES
                ' ------------------------------
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
                    notes = Hyperlink(SQLdr("notes"))
                    TotalNotes += 1
                End If

                ' ------------------------------
                ' WRITE ROW
                ' ------------------------------
                html.WriteLine(
                $"<tr{cls}><td>{Strings.Replace(SQLdr("Entity"), " ", "&nbsp;")}</td>" &
                $"<td>{ringCount:n0}</td><td>{outer}</td><td>{inner}</td>" &
                $"<td>{pointCount:n0}</td><td>{unclosed}</td><td>{JSONcount:n0}</td>" &
                $"<td>{SQLdr("query")}</td><td>{bbox}</td><td>{notes}</td></tr>"
            )

                lines += 1
            End While

            ' ------------------------------
            ' TOTALS ROW
            ' ------------------------------
            html.WriteLine(
            $"<tr><td>{TotalEntity}</td><td>{TotalRings:n0}</td><td></td><td></td>" &
            $"<td>{TotalPoints:n0}</td><td>{TotalUnclosed:n0}</td><td>{TotalJSON:n0}</td>" &
            $"<td>{TotalQueries}</td><td>{TotalBbox}</td><td>{TotalNotes}</td></tr>"
        )

            html.WriteLine("</table>")

            AppendText(Form1.TextBox1, $"{lines} lines written to HTML file{vbCrLf}")
            OpenHtml(htmlFile)
        End Using
    End Sub
    Private Sub ProcessPolygon(poly As Polygon,
                               ByRef ringCount As Integer,
                               ByRef pointCount As Integer,
                               ByRef outer As Integer,
                               ByRef inner As Integer,
                               ByRef unclosed As Integer)

        ' ------------------------------
        ' Outer ring (shell)
        ' ------------------------------
        Dim shell = poly.Shell
        ringCount += 1
        outer += 1
        pointCount += shell.NumPoints
        If Not shell.IsClosed Then unclosed += 1

        ' ------------------------------
        ' Inner rings (holes)
        ' ------------------------------
        For i = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(i)
            ringCount += 1
            inner += 1
            pointCount += hole.NumPoints
            If Not hole.IsClosed Then unclosed += 1
        Next
    End Sub

    Sub GeometrySizeTable()
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim ringCount As Integer, lines As Integer = 0
        Dim TotalEntity As Integer = 0, TotalRings As Integer = 0, TotalSize As Integer = 0

        Dim htmlFile = Path.Combine(Application.StartupPath, "Geometry Size.html")

        Using connect As New SqliteConnection(DXCC_DATA),
          html As New StreamWriter(htmlFile, False)

            connect.Open()
            sql = connect.CreateCommand

            ' ---------------------------------------------------------
            ' Total size of all geometry JSON
            ' ---------------------------------------------------------
            sql.CommandText = "select sum(length(geometry)) as TotalSize from DXCC where Deleted=0 and DXCCnum != 999"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            TotalSize = SafeInt(SQLdr("TotalSize"))
            SQLdr.Close()

            ' ---------------------------------------------------------
            ' Total number of entities
            ' ---------------------------------------------------------
            sql.CommandText = "select count(*) as Total from DXCC where Deleted=0 and DXCCnum != 999"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()

            With Form1.ProgressBar1
                .Minimum = 0
                .Maximum = SQLdr("Total")
                .Value = 0
            End With

            SQLdr.Close()

            ' ---------------------------------------------------------
            ' Main query
            ' ---------------------------------------------------------
            sql.CommandText =
            "select *, LENGTH(geometry) AS size " &
            "from DXCC where Deleted=0 and DXCCnum != 999 " &
            "order by size DESC"

            SQLdr = sql.ExecuteReader

            ' ---------------------------------------------------------
            ' HTML header
            ' ---------------------------------------------------------
            html.WriteLine("<!DOCTYPE html>")
            html.WriteLine("
            <style>
                table {
                    border-collapse: collapse;
                    border: 1px solid #444;
                }
                table th, table td {
                    border: 1px solid #444;
                    padding: 4px 6px;
                }
                table td:nth-child(2),
                table td:nth-child(3),
                table td:nth-child(4),
                table td:nth-child(5),
                table td:nth-child(6) {
                    text-align: right;
                }
                table thead th {
                    background-color: #f0f0f0;
                    font-weight: bold;
                }
                .red td {
                    color: red;
                    font-weight: bold;
                }
            </style>")

            html.WriteLine("<table>")
            html.WriteLine("<tr><th>Entity</th><th>Rings</th><th>Geometry Size</th><th>Percent</th></tr>")

            ' ---------------------------------------------------------
            ' Process each DXCC row
            ' ---------------------------------------------------------
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                TotalEntity += 1

                Dim Entity = SafeStr(SQLdr("Entity"))
                Dim geometryJson = SafeStr(SQLdr("geometry"))
                Dim size = SafeInt(SQLdr("size"))

                ' ------------------------------
                ' Count rings using NTS
                ' ------------------------------
                If IsDBNullorEmpty(geometryJson) Then
                    ringCount = 0
                Else
                    Dim geom As Geometry = FromGeoJsonToNTS(geometryJson)
                    ringCount = CountRings(geom)
                    TotalRings += ringCount
                End If

                ' ------------------------------
                ' Write row
                ' ------------------------------
                html.WriteLine(
                $"<tr><td>{Strings.Replace(Entity, " ", "&nbsp;")}</td>" &
                $"<td>{ringCount:n0}</td><td>{size}</td>" &
                $"<td>{If(TotalSize > 0, size / TotalSize * 100, 0):f1}</td></tr>"
            )

                lines += 1
            End While

            ' ---------------------------------------------------------
            ' Totals row
            ' ---------------------------------------------------------
            html.WriteLine(
            $"<tr><td>{TotalEntity}</td><td>{TotalRings:n0}</td>" &
            $"<td>{TotalSize:n0}</td><td></td></tr>"
        )

            html.WriteLine("</table>")

            AppendText(Form1.TextBox1, $"{lines} lines written to HTML file{vbCrLf}")
            OpenHtml(htmlFile)
        End Using
    End Sub
    Private Function CountRings(g As Geometry) As Integer
        If g Is Nothing OrElse g.IsEmpty Then Return 0

        Dim count As Integer = 0

        If TypeOf g Is Polygon Then
            Dim poly = DirectCast(g, Polygon)
            count += 1 + poly.NumInteriorRings   ' shell + holes

        ElseIf TypeOf g Is MultiPolygon Then
            Dim mp = DirectCast(g, MultiPolygon)
            For i = 0 To mp.NumGeometries - 1
                Dim poly = DirectCast(mp.GetGeometryN(i), Polygon)
                count += 1 + poly.NumInteriorRings
            Next
        End If

        Return count
    End Function

    Sub KMLFileSize()
        ' Create summary of KML file sizes
        Dim lines As Integer = 0, TotalFolders As Integer = 0, FolderSize As Integer = 0, TotalSubFolders As Integer = 0, TotalPlacemarks As Integer = 0
        Dim TotalPolygons As Integer = 0, TotalLineStrings As Integer = 0

        Dim htmlFile = Path.Combine(Application.StartupPath, "KML File Size.html")
        Using html As New StreamWriter(htmlFile, False)
            html.WriteLine("<!DOCTYPE html>")
            html.WriteLine("<style> table td:nth-child(2), td:nth-child(3), td:nth-child(4),td:nth-child(5),td:nth-child(6) {text-align:right;} .red td{color:red;font-weight: bold;}</style>")
            html.WriteLine("<table border=1>")
            html.WriteLine("<tr><th>Folder</th><th>Sub-folders</th><th>Placemarks</th><th>Polygons</th><th>LineStrings</th><th>Size</th></tr>")
            Dim doc = XDocument.Load($"{Application.StartupPath}\KML\DXCC Map of the World.kml")    ' read the XML
            Dim ns = doc.Root.Name.Namespace      ' get namespace name so we can qualify everything
            Dim document = doc.Root.Element(ns + "Document")        ' root of kml
            Dim folders = document.Elements(ns + "Folder")       ' find all the top level folders
            For Each folder In folders
                TotalFolders += 1
                Dim size = folder.ToString.Length
                FolderSize += size
                Dim name = folder.Element(ns + "name").Value
                Dim subfolders = folder.Elements(ns + "Folder").Count
                TotalSubFolders += subfolders
                Dim placemarks = folder.Descendants(ns + "Placemark").Count
                TotalPlacemarks += placemarks
                Dim polygons = folder.Descendants(ns + "Polygon").Count
                TotalPolygons += polygons
                Dim linestrings = folder.Descendants(ns + "LineString").Count
                TotalLineStrings += linestrings
                html.WriteLine($"<tr><td>{name}</td><td>{subfolders:n0}</td><td>{placemarks:n0}</td><td>{polygons:n0}</td><td>{linestrings:n0}</td><td>{size:n0}</td></tr>")
                lines += 1
            Next
            html.WriteLine($"<tr><td>{TotalFolders}</td><td>{TotalSubFolders:n0}</td><td>{TotalPlacemarks:n0}</td><td>{TotalPolygons:n0}</td><td>{TotalLineStrings:n0}</td><td>{FolderSize:n0}</td></tr>")
            html.WriteLine("</table>")
            AppendText(Form1.TextBox1, $"{lines} lines written To html file{vbCrLf}")
            OpenHtml(htmlFile)
        End Using
    End Sub
End Module
