Imports System.IO
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Module Validate
    ' module contains all validate functions
    Sub CountryCollisions()
        ' Search for DXCC that intersect
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, theWorld As New List(Of (Entity As String, geom As Geometry)), overlaps As Integer = 0
        ' List of country collisions that are acceptable, i.e. there isn't much we can do about it, and they don;t compromise the utility of the data
        Dim AcceptableCollisions As New List(Of (A As String, B As String)) From {
            {("Finland", "Aland Is")},
            {("Alaska", "Canada")},
            {("Canada", "United States")},
            {("Cuba", "Guantanamo Bay")},
            {("Italy", "Vatican")},
            {("United Nations HQ", "United States")},
            {("Argentina", "Chile")},
            {("Morocco", "Western Sahara")},
            {("Sudan", "South Sudan")}
}
        Form1.AppendText(Form1.TextBox1, $"The following collisions are deemed acceptable/unavoidable{vbCrLf}")
        For Each accept In AcceptableCollisions
            Form1.AppendText(Form1.TextBox1, $"{accept.A} - {accept.B}{vbCrLf}")
        Next
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
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
            ' Retrieve all the geometry
            Form1.AppendText(Form1.TextBox1, $"Loading geometry for {Form1.ProgressBar1.Maximum} countries{vbCrLf}")
            theWorld.Clear()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1
                Dim poly As Polygon = Polygon.FromJson(SQLdr("geometry"))
                Dim polyBuffered = GeometryEngine.BufferGeodetic(poly, -10000, LinearUnits.Meters)       ' shrink the polygons a little
                ' if the buffered polygon disappears, replace it with the original
                If Not polyBuffered.IsEmpty Then
                    poly = polyBuffered
                End If
                theWorld.Add((Entity:=SQLdr("Entity"), geom:=poly))
            End While
            Form1.AppendText(Form1.TextBox1, $"Geometry for {theWorld.Count} entities loaded{vbCrLf}")
            ' test for collisions
            For outer = 0 To theWorld.Count - 1
                Form1.ProgressBar1.Value = outer
                For inner = outer + 1 To theWorld.Count - 1
                    Dim Acceptable As Boolean = AcceptableCollisions.Contains((theWorld.Item(outer).Entity, theWorld.Item(inner).Entity)) Or AcceptableCollisions.Contains((theWorld.Item(inner).Entity, theWorld.Item(outer).Entity))
                    If Not Acceptable Then
                        If GeometryEngine.Intersects(theWorld.Item(outer).geom, theWorld.Item(inner).geom) Then
                            Form1.AppendText(Form1.TextBox1, $"The entities {theWorld.Item(outer).Entity} and {theWorld.Item(inner).Entity} overlap{vbCrLf}")
                            overlaps += 1
                        End If
                    End If
                Next
            Next
            Form1.AppendText(Form1.TextBox1, $"{overlaps} overlaps found{vbCrLf}")
        End Using
    End Sub
    Sub CheckISO3166References()
        ' Create a table of ISO 3166-1 references. Any value != 1 requires a look
        Dim sqlISO As SqliteCommand, ISOdr As SqliteDataReader, sqlDXCC As SqliteCommand, DXCCdr As SqliteDataReader, updated As Integer = 0, count As Integer = 0
        Using connect As New SqliteConnection(Form1.DXCC_DATA),
              ISO As New StreamWriter($"{Application.StartupPath}\ISOreport.html", False)

            connect.Open()
            sqlISO = connect.CreateCommand
            sqlDXCC = connect.CreateCommand
            sqlISO.CommandText = "SELECT count(*) AS Total FROM ISO31661"
            ISOdr = sqlISO.ExecuteReader
            ISOdr.Read()
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = CInt(ISOdr("Total"))
            End With
            ISOdr.Close()
            ISO.WriteLine("<table border=1")
            ISO.WriteLine("<tr><th>Entity</th><th>Code</th><th>References</th></tr>")
            sqlISO.CommandText = "SELECT * FROM ISO31661 ORDER by Entity"
            ISOdr = sqlISO.ExecuteReader
            While ISOdr.Read
                count += 1
                Form1.ProgressBar1.Value = count
                sqlDXCC.CommandText = $"SELECT COUNT(*) as Total FROM DXCC WHERE query LIKE '%={ISOdr("Code")}]%'"
                Dim refs = 0
                DXCCdr = sqlDXCC.ExecuteReader
                DXCCdr.Read()
                refs += DXCCdr("Total")
                DXCCdr.Close()
                ISO.WriteLine($"<tr><td>{ISOdr("Entity")}</td><td>{ISOdr("Code")}</td><td>{refs}</td></tr>")
            End While
            ISO.WriteLine("</table>")
        End Using
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
    Sub IOTACheck()
        ' Check that IOTA which belong to a single DXCC, intersect with that DXCC
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, sql1 As SqliteCommand, SQLdr1 As SqliteDataReader
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql1 = connect.CreateCommand
            sql.CommandText = "SELECT *,IOTA_DXCC_IOTA.dxcc_num AS DXCC FROM IOTA_DXCC_IOTA JOIN IOTA_Groups USING (refno)"     ' pick the correct DXCC
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                ' create an envelope which is the IOTA group boundary
                Dim iota As New Envelope(New MapPoint(SQLdr("longitude_min"), SQLdr("latitude_min"), SpatialReferences.Wgs84), New MapPoint(SQLdr("longitude_max"), SQLdr("latitude_max"), SpatialReferences.Wgs84))
                ' now get the DXCC geometry
                sql1.CommandText = $"SELECT * FROM DXCC WHERE dxccnum={SQLdr("DXCC")}"
                SQLdr1 = sql1.ExecuteReader
                SQLdr1.Read()
                Form1.AppendText(Form1.TextBox1, $"Checking IOTA {SQLdr("refno")} and {SQLdr1("Entity")}{vbCrLf}")
                Dim geometry As Polygon = Polygon.FromJson(SQLdr1("geometry"))      ' retrieve the DXCC geometry
                ' Now check for intersection of IOTA and DXCC
                If Not geometry.Intersects(iota) Then
                    Form1.AppendText(Form1.TextBox1, $"************** IOTA ref {SQLdr("refno")} - {SQLdr("name")} does not intersect with DXCC {SQLdr1("Entity")} **************{vbCrLf}")
                End If
                SQLdr1.Close()
            End While
        End Using
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub

    Public Sub AdjacentColourCheck()
        ' Check that all adjacent entities have a different colour
        ' The colours are 1=red, 2=green, 3=Blue, 4=Yellow, 5=Magenta
        Dim countries As New List(Of (prefix As String, geom As Geometry, colour As Integer)), adjacent As New List(Of (prefix As String, colour As Integer)), sql As SqliteCommand, sqldr As SqliteDataReader
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            ' Read in geometry for all entities
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry is not NULL"
            sqldr = sql.ExecuteReader
            While sqldr.Read
                countries.Add((prefix:=sqldr("prefix"), geom:=Geometry.FromJson(sqldr("geometry")), colour:=sqldr("colour")))
            End While
        End Using
        Form1.AppendText(Form1.TextBox1, $"{countries.Count} countries loaded{vbCrLf}")
        ' Test every country
        With Form1.ProgressBar1
            .Minimum = 0
            .Maximum = countries.Count
            .Value = 0
        End With
        For outer = 0 To countries.Count - 2
            Form1.ProgressBar1.Value += 1
            adjacent.Clear()
            For inner = outer + 1 To countries.Count - 1
                If outer <> inner Then
                    If GeometryEngine.Touches(countries(outer).geom, countries(inner).geom) Then
                        adjacent.Add((countries(inner).prefix, countries(inner).colour))            ' record adjacent country and colour
                    End If
                End If
            Next
            For Each item In adjacent
                If item.colour = countries(outer).colour Then Form1.AppendText(Form1.TextBox1, $"Adjacent countries {countries(outer).prefix} and {item.prefix} share a colour{vbCrLf}")
            Next
        Next
        Form1.AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
End Module

