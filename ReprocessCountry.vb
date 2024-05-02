Imports System.IO
Imports System.Windows.Forms
Imports Esri.ArcGISRuntime.UI.GeoAnalysis
Imports Microsoft.Data.Sqlite

Public Class ReprocessCountry
    ' Reprocess GIS data for a country
    ' 1. Remove any existing geometry
    ' 2. Download and process OSM data
    ' 3. Remake country KML file

    Dim datasource As New DataTable
    Private Async Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Dim sql As SqliteCommand, DXCClist As New Dictionary(Of String, Integer)
        Dim value As Integer, display As String

        Me.Hide()

        If ListBox1.SelectedIndex <> -1 Then
            If ListBox1.SelectedValue = 0 Then
                ' add all DXCC to list
                For ndx = 1 To ListBox1.Items.Count - 1
                    value = datasource.Rows(ndx).Item("value")
                    display = datasource.Rows(ndx).Item("display")
                    DXCClist.Add(display, value)
                Next
            Else
                ' add just the selected one
                value = datasource.Rows(ListBox1.SelectedIndex).Item("value")
                display = datasource.Rows(ListBox1.SelectedIndex).Item("display")
                DXCClist.Add(display, value)        ' add selected value to list
            End If
        End If
        ' Now reprocess the list of DXCC
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            For Each n In DXCClist
                ' Remove the geometry
                sql.CommandText = $"UPDATE `DXCC` SET `geometry`=NULL WHERE `DXCCnum`={n.Value}"
                sql.ExecuteNonQuery()
                ' get the OSM data
                Await Form1.CreateGrids(connect, n.Key)
                ' remake KML file
                Dim kml As New StreamWriter($"{Application.StartupPath}\KML\DXCC_{n.Key}.kml", False)
                kml.WriteLine(Form1.KMLheader)
                KMLlist(connect, kml, New List(Of Integer) From {n.Value})
                kml.WriteLine(Form1.KMLfooter)
                kml.Close()
            Next
        End Using
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ReprocessCountry_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sql As SqliteCommand, SQLdr As SqliteDataReader, DXCClist As New List(Of Integer)
        Using connect As New SqliteConnection(Form1.DXCC_DATA)
            datasource.Clear()
            datasource.Columns.Add("display", GetType(String))
            datasource.Columns.Add("value", GetType(Integer))
            datasource.Rows.Add("All", 0)
            Dim lastLetter As String = ""       ' start letter of last item added
            connect.Open()
            sql = connect.CreateCommand
            sql.CommandText = "SELECT * FROM DXCC WHERE Deleted=0 AND DXCCnum != 999 ORDER BY Entity"
            SQLdr = sql.ExecuteReader
            While SQLdr.Read
                Dim firstLetter As String = SQLdr("Entity")(0)
                If firstLetter <> lastLetter Then datasource.Rows.Add($"{firstLetter}-----------------------", -1)
                datasource.Rows.Add(SQLdr("Entity"), SQLdr("DXCCnum"))
                LastLetter = firstLetter
            End While
            ListBox1.DisplayMember = "display"
            ListBox1.ValueMember = "value"
            ListBox1.DataSource = datasource
        End Using
    End Sub
End Class
