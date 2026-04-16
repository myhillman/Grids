Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Xml
Imports System.Xml.XPath

Module Reports
    Sub EntityReport()
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
            html.WriteLine("<tr><th>Entity</th><th>Parts</th><th>Outer</th><th>Inner</th><th>Points</th><th>Unclosed</th><th>JSON</th><th>OSM Overpass QL</th><th>bbox</th><th>Notes</th></tr>")
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                partcount = 0
                pointcount = 0
                JSONcount = 0
                outer = 0
                inner = 0
                unclosed = 0
                If Not IsDBNullorEmpty(SQLdr("query")) Then TotalQueries += 1
                Dim cls = " class=red"
                If Not IsDBNullorEmpty(SQLdr("geometry")) Then
                    Dim geometry As Polygon = GeoJsonToGeometry(SQLdr("geometry"))
                    If Not geometry.IsEmpty Then cls = ""
                    JSONcount = SQLdr("geometry").ToString.Length
                    TotalJSON += JSONcount
                    partcount = geometry.Parts.Count

                    For Each prt In geometry.Parts
                        pointcount += prt.PointCount
                        If PolygonArea(prt.Points) < 0 Then outer += 1 Else inner += 1
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
                    notes = Hyperlink(SQLdr("notes"))
                    TotalNotes += 1
                End If
                html.WriteLine($"<tr{cls}><td>{Strings.Replace(SQLdr("Entity"), " ", "&nbsp;")}</td><td>{partcount:n0}</td><td>{outer}</td><td>{inner}</td><td>{pointcount:n0}</td><td>{unclosed}</td><td>{JSONcount:n0}</td><td>{SQLdr("query")}</td><td>{bbox}</td><td>{notes}</td></tr>")
                lines += 1
            End While
            html.WriteLine($"<tr><td>{TotalEntity}</td><td>{TotalParts:n0}</td><td></td><td></td><td>{TotalPoints:n0}</td><td>{TotalUnclosed:n0}</td><td>{TotalJSON:n0}</td><td>{TotalQueries}</td><td>{TotalBbox}</td><td>{TotalNotes}</td></tr>")
            html.WriteLine("</table>")
            AppendText(Form1.TextBox1, $"{lines} lines written To html file{vbCrLf}")
        End Using
    End Sub
    Sub GeometrySizeTable()
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader
        Dim partcount As Integer, lines As Integer = 0
        Dim TotalEntity As Integer = 0, TotalParts As Integer = 0, TotalSize As Integer = 0

        Using connect As New SqliteConnection(DXCC_DATA),
              html As New StreamWriter($"{Application.StartupPath}\Geometry Size.html", False)
            connect.Open()
            sql = connect.CreateCommand
            ' Find total size of geometry
            sql.CommandText = "select sum(length(geometry)) as TotalSize from DXCC where Deleted=0 and DXCCnum != 999"
            SQLdr = sql.ExecuteReader
            SQLdr.Read()
            TotalSize = SQLdr("TotalSize")
            SQLdr.Close()
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
            sql.CommandText = "select *,LENGTH(geometry) AS size from `DXCC` where `Deleted`=0 and DXCCnum != 999 ORDER BY size DESC"
            SQLdr = sql.ExecuteReader
            html.WriteLine("<!DOCTYPE html>")
            html.WriteLine("<style> table td:nth-child(2), td:nth-child(3), td:nth-child(4),td:nth-child(5),td:nth-child(6) {text-align:right;} .red td{color:red;font-weight: bold;}</style>")
            html.WriteLine("<table border=1>")
            html.WriteLine("<tr><th>Entity</th><th>Parts</th><th>Geometry Size</th><th>Percent</th></tr>")
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                TotalEntity += 1
                Dim poly As Polygon = GeoJsonToGeometry(SQLdr("geometry"))
                partcount = poly.Parts.Count
                TotalParts += partcount
                html.WriteLine($"<tr><td>{Strings.Replace(SQLdr("Entity"), " ", "&nbsp;")}</td><td>{partcount:n0}</td><td>{SQLdr("size")}</td><td>{SQLdr("size") / TotalSize * 100:f1}</td></tr>")
                lines += 1
            End While
            html.WriteLine($"<tr><td>{TotalEntity}</td><td>{TotalParts:n0}</td><td>{TotalSize:n0}</td><td></td></tr>")
            html.WriteLine("</table>")
            AppendText(Form1.TextBox1, $"{lines} lines written To html file{vbCrLf}")
        End Using
    End Sub
    Sub KMLFileSize()
        ' Create summary of KML file sizes
        Dim lines As Integer = 0, TotalFolders As Integer = 0, FolderSize As Integer = 0, TotalSubFolders As Integer = 0, TotalPlacemarks As Integer = 0
        Dim TotalPolygons As Integer = 0, TotalLineStrings As Integer = 0

        Using html As New StreamWriter($"{Application.StartupPath}\KML File Size.html", False)
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
        End Using
    End Sub
End Module
