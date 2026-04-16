Imports System.IO
Imports System.Text.RegularExpressions
Imports Microsoft.Data.Sqlite

Public Class KMLbyCQ
    Private Sub KMLbyCQ_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' fetch list of all CQ zones used to populate listbox
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, CQzones As New List(Of Integer)
        Using connect As New SqliteConnection(DXCC_DATA),
            kml As New StreamWriter($"{Application.StartupPath}\DXCC.kml", False)
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM DXCC where Deleted=0 and CQ IS NOT NULL"
            SQLdr = sql.ExecuteReader
            ' Read all CQ zones, and then build list of all zones
            While SQLdr.Read
                If Not IsDBNullorEmpty(SQLdr("CQ")) Then
                    Dim zones = SQLdr("CQ")       ' a list of CQ zones
                    Dim zone = Split(zones, ",")
                    ' zones could be listed as n-m or just n in CSV list
                    For i = LBound(zone) To UBound(zone)
                        If IsNumeric(zone(i)) Then
                            ' Single value
                            If Not CQzones.Contains(zone(i)) Then CQzones.Add(zone(i))
                        Else
                            ' might be range
                            Dim matches As MatchCollection = Regex.Matches(zone(i), "(\d+)-(\d+)")
                            If matches.Count = 1 Then
                                For j = CInt(matches(0).Groups(1).Value) To CInt(matches(0).Groups(2).Value)
                                    If Not CQzones.Contains(j) Then CQzones.Add(j)
                                Next
                            Else
                                MsgBox($"Could not decipher CQ zone {zone(i)}")
                            End If
                        End If
                    Next
                End If
            End While
            SQLdr.Close()
            CQzones.Sort()
            ' Populate the listbox
            ListBox1.Items.Add("All")
            For Each cqzone In CQzones
                ListBox1.Items.Add(cqzone)
            Next
        End Using
    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        ' Create KML for all entities in selected zone
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, DXCClist As New List(Of Integer), CQzoneList As New List(Of Integer)

        If ListBox1.SelectedItem IsNot Nothing Then
            Using connect As New SqliteConnection(DXCC_DATA)
                connect.Open()
                connect.CreateFunction("within", Function(A As Integer, B As String) Within(A, B))           ' find CQ zone within list
                sql = connect.CreateCommand
                CQzoneList.Clear()
                If ListBox1.SelectedItem.ToString = "All" Then
                    ' Proccess all zones
                    For i = 1 To ListBox1.Items.Count - 1
                        CQzoneList.Add(i)
                    Next
                Else
                    CQzoneList.Add(ListBox1.SelectedItem)   ' just the selected one
                End If

                With Form1.ProgressBar1
                    .Minimum = 0
                    .Maximum = CQzoneList.Count
                    .Value = 0
                End With

                For Each CQ In CQzoneList
                    'Form1.ProgressBar1.Value += 1
                    AppendText(Form1.TextBox1, $"Making CQ zone {CQ}{vbCrLf}")
                    DXCClist.Clear()
                    Application.DoEvents()
                    Dim kml As New StreamWriter($"{Application.StartupPath}\KML\CQ_Zone_{CQ}.kml", False)
                    sql.CommandText = $"SELECT * FROM DXCC WHERE within({CQ},CQ) AND Deleted=0 AND geometry IS NOT NULL ORDER BY Entity"      ' fetch all DXCC in this CQ zone
                    SQLdr = sql.ExecuteReader
                    While SQLdr.Read
                        DXCClist.Add(SQLdr("DXCCnum"))
                    End While
                    SQLdr.Close()
                    kml.WriteLine(KMLheader)
                    KMLlist(connect, kml, DXCClist)
                    kml.WriteLine(KMLfooter)
                    kml.Close()
                Next
                AppendText(Form1.TextBox1, $"Done{vbCrLf}")
            End Using
        End If
    End Sub

    Shared Function Within(needle As Integer, haystack As String) As Boolean
        ' find needle in haystack, where haystack contains a CSV list of n or n-m entries
        Dim words = Split(haystack, ",")
        ' words could be listed as n-m or just n in CSV list
        For i = LBound(words) To UBound(words)
            If IsNumeric(words(i)) Then
                ' Single value
                If words(i) = needle Then Return True       ' found it
            Else
                ' might be range
                Dim matches As MatchCollection = Regex.Matches(words(i), "(\d+)-(\d+)")
                If matches.Count = 1 Then
                    For j = CInt(matches(0).Groups(1).Value) To CInt(matches(0).Groups(2).Value)
                        If j = needle Then Return True       ' found it
                    Next
                End If
            End If
        Next
        Return False
    End Function
End Class