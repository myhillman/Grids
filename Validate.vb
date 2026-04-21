Imports System.IO
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite
Module Validate
    ' module contains all validate functions
    Sub CountryCollisions()
        ' Search for DXCC that intersect
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, theWorld As New List(Of (Entity As String, geom As NetTopologySuite.Geometries.Geometry)), overlaps As Integer = 0
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
        AppendText(Form1.TextBox1, $"The following collisions are deemed acceptable/unavoidable{vbCrLf}")
        For Each accept In AcceptableCollisions
            AppendText(Form1.TextBox1, $"{accept.A} - {accept.B}{vbCrLf}")
        Next
        Using connect As New SqliteConnection(DXCC_DATA)
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
            AppendText(Form1.TextBox1, $"Loading geometry for {Form1.ProgressBar1.Maximum} countries{vbCrLf}")
            theWorld.Clear()
            sql.CommandText = "SELECT * FROM DXCC WHERE geometry IS NOT NULL ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Form1.ProgressBar1.Value += 1

                Dim poly As NetTopologySuite.Geometries.Geometry =
                        FromGeoJsonToNTS(SQLdr("geometry"))

                ' Convert meters → degrees
                Const meters As Double = -10000
                Dim deg As Double = meters / 111320.0

                ' NTS buffer
                Dim polyBuffered As NetTopologySuite.Geometries.Geometry = poly.Buffer(deg)

                ' If the buffered polygon disappears, keep the original
                If Not polyBuffered.IsEmpty Then
                    poly = polyBuffered
                End If

                theWorld.Add((Entity:=SQLdr("Entity"), geom:=poly))
            End While
            AppendText(Form1.TextBox1, $"Geometry for {theWorld.Count} entities loaded{vbCrLf}")
            ' test for collisions
            For outer = 0 To theWorld.Count - 1
                Form1.ProgressBar1.Value = outer

                For inner = outer + 1 To theWorld.Count - 1

                    Dim Acceptable As Boolean =
                            AcceptableCollisions.Contains((theWorld(outer).Entity, theWorld(inner).Entity)) OrElse
                            AcceptableCollisions.Contains((theWorld(inner).Entity, theWorld(outer).Entity))

                    If Not Acceptable Then
                        ' NTS intersects
                        If theWorld(outer).geom.Intersects(theWorld(inner).geom) Then
                            AppendText(Form1.TextBox1,
                                       $"The entities {theWorld(outer).Entity} and {theWorld(inner).Entity} overlap{vbCrLf}")
                            overlaps += 1
                        End If
                    End If

                Next
            Next
            AppendText(Form1.TextBox1, $"{overlaps} overlaps found{vbCrLf}")
        End Using
    End Sub
    Sub CheckISO3166References()
        ' Create a table of ISO 3166-1 references. Any value != 1 requires a look
        Dim sqlISO As SqliteCommand, ISOdr As SqliteDataReader, sqlDXCC As SqliteCommand, DXCCdr As SqliteDataReader, updated As Integer = 0, count As Integer = 0
        Using connect As New SqliteConnection(DXCC_DATA),
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
                .Maximum = SafeInt(ISOdr("Total"))
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
        AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub
    Sub IOTACheck()

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            Dim sql = connect.CreateCommand()
            Dim sql1 = connect.CreateCommand()

            sql.CommandText =
                "SELECT *, IOTA_DXCC_IOTA.dxcc_num AS DXCC " &
                "FROM IOTA_DXCC_IOTA JOIN IOTA_Groups USING (refno)"

            Dim rdr = sql.ExecuteReader()

            While rdr.Read()

                ' Build NTS envelope
                Dim minLon As Double = rdr("longitude_min")
                Dim minLat As Double = rdr("latitude_min")
                Dim maxLon As Double = rdr("longitude_max")
                Dim maxLat As Double = rdr("latitude_max")

                Dim env As New NetTopologySuite.Geometries.Envelope(minLon, maxLon, minLat, maxLat)

                ' Convert envelope → polygon
                Dim iotaPoly As NetTopologySuite.Geometries.Geometry = factory.ToGeometry(env)

                ' Load DXCC geometry (already NTS)
                sql1.CommandText = $"SELECT * FROM DXCC WHERE dxccnum={rdr("DXCC")}"
                Dim rdr1 = sql1.ExecuteReader()
                rdr1.Read()

                AppendText(Form1.TextBox1,
                           $"Checking IOTA {rdr("refno")} and {rdr1("Entity")}{vbCrLf}")

                Dim dxccGeom As NetTopologySuite.Geometries.Geometry = FromGeoJsonToNTS(rdr1("geometry"))

                rdr1.Close()

                ' NTS intersection test
                If Not dxccGeom.Intersects(iotaPoly) Then
                    AppendText(Form1.TextBox1,
                               $"************** IOTA ref {rdr("refno")} - {rdr("name")} does not intersect with DXCC {rdr1("Entity")} **************{vbCrLf}")
                End If

            End While
        End Using

        AppendText(Form1.TextBox1, $"Done{vbCrLf}")

    End Sub


    Public Sub AdjacentColourCheck()
        ' Check that all adjacent entities have a different colour
        ' Colours: 1=Red, 2=Green, 3=Blue, 4=Yellow, 5=Magenta

        Dim countries As New List(Of (prefix As String,
                                      geom As NetTopologySuite.Geometries.Geometry,
                                      colour As Integer))

        Dim adjacent As New List(Of (prefix As String, colour As Integer))

        Dim sql As SqliteCommand, sqldr As SqliteDataReader

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT prefix, geometry, colour FROM DXCC WHERE geometry IS NOT NULL"
            sqldr = sql.ExecuteReader

            While sqldr.Read
                countries.Add((
                    prefix:=SafeStr(sqldr("prefix")),
                    geom:=FromGeoJsonToNTS(SafeStr(sqldr("geometry"))),
                    colour:=SafeInt(sqldr("colour"))
                ))
            End While
        End Using

        AppendText(Form1.TextBox1, $"{countries.Count} countries loaded{vbCrLf}")

        ' Progress bar
        With Form1.ProgressBar1
            .Minimum = 0
            .Maximum = countries.Count
            .Value = 0
        End With

        ' ---------------------------------------------------------
        ' MAIN LOOP — check adjacency using NTS .Touches()
        ' ---------------------------------------------------------
        For outer = 0 To countries.Count - 2
            Form1.ProgressBar1.Value += 1
            adjacent.Clear()

            For inner = outer + 1 To countries.Count - 1

                ' NTS adjacency test
                If countries(outer).geom.Touches(countries(inner).geom) Then
                    adjacent.Add((countries(inner).prefix, countries(inner).colour))
                End If

            Next

            ' Check for colour conflicts
            For Each item In adjacent
                If item.colour = countries(outer).colour Then
                    AppendText(Form1.TextBox1,
                        $"Adjacent countries {countries(outer).prefix} and {item.prefix} share a colour{vbCrLf}")
                End If
            Next
        Next

        AppendText(Form1.TextBox1, $"Done{vbCrLf}")
    End Sub

End Module

