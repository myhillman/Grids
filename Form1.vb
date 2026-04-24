Imports System.ComponentModel

Partial Public Class Form1
    Inherits System.Windows.Forms.Form

    Public Sub New()
        If LicenseManager.UsageMode = LicenseUsageMode.Designtime Then
            ' Skip runtime-only code, but DO call InitializeComponent
            InitializeComponent()
            Return
        End If
        InitializeComponent()
        InitHttp()
        SQLitePCL.Batteries.Init()
        InitializeOSMCache()
        MakeKMLheader()
        AddHandler AppDomain.CurrentDomain.FirstChanceException,
            Sub(sender, e)
                If TypeOf e.Exception Is InvalidCastException Then
                    Debug.WriteLine("FIRST CHANCE CAST FAIL: " & e.Exception.Message)
                    Debug.WriteLine(e.Exception.StackTrace)
                End If
            End Sub
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Dim m As New ToolStripMenuItem("Convert legacy")
        AddHandler m.Click, AddressOf ConvertLegacyToolStripMenuItem_Click
        ImportToolStripMenuItem.DropDownItems.Add(m)
        Dim n As New ToolStripMenuItem("OSM Alternatives")
        AddHandler n.Click, AddressOf OSMAlternativesToolStripMenuItem_Click
        ImportToolStripMenuItem.DropDownItems.Add(n)
    End Sub
    Public Async Sub UseShapefileToolStripMenuItem1_Click(sender As Object, e As EventArgs) _
        Handles UseShapefileToolStripMenuItem1.Click
        Await UseShapefile()
    End Sub

    Private Async Sub UseOSMToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles UseOSMToolStripMenuItem.Click
        Await UseOSM()
    End Sub

    Private Sub MakeKMLByCQZoneToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles MakeKMLByCQZoneToolStripMenuItem.Click
        KMLbyCQ.ShowDialog()
    End Sub

    Private Sub MakeKMLToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles MakeKMLToolStripMenuItem.Click
        MakeKML()
    End Sub

    Private Sub MakeShapefileToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles MakeShapefileToolStripMenuItem.Click
        MakeShapefile()
    End Sub

    Private Sub EntityReportToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles EntityReportToolStripMenuItem.Click
        EntityReport()
    End Sub

    Private Sub ReprocessCountryToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ReprocessCountryToolStripMenuItem.Click
        ReprocessCountry.Show()
    End Sub

    Private Sub CheckISO3166ReferencesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles CheckISO3166ReferencesToolStripMenuItem.Click
        CheckISO3166References()
    End Sub

    Private Sub CountryCollisionsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles CountryCollisionsToolStripMenuItem.Click
        CountryCollisions()
    End Sub

    Private Sub RemoveEmptyGeometryToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles RemoveEmptyGeometryToolStripMenuItem.Click
        RemoveEmptyGeometry()
    End Sub

    Private Sub UseOSMLandPolygonsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles UseOSMLandPolygonsToolStripMenuItem.Click
        UseOSMLandPolygons()
    End Sub

    Private Sub MakeKMLAllEntitiesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles MakeKMLAllEntitiesToolStripMenuItem.Click
        MakeKMLAllEntities()
    End Sub

    Private Sub InnerRingsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles InnerRingsToolStripMenuItem.Click
        InnerRings()
    End Sub

    Private Sub ImportCQITUZonesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportCQITUZonesToolStripMenuItem.Click
        ImportCQITUZones()
    End Sub

    Private Async Sub ImportTimezonesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportTimezonesToolStripMenuItem.Click
        Await ImportTimeZones()
    End Sub

    Private Sub ImportEUASBorderToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportEUASBorderToolStripMenuItem.Click
        ImportEUASBorder()
    End Sub

    Private Async Sub ImportIARURegionsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportIARURegionsToolStripMenuItem.Click
        Await ImportIARURegions()
    End Sub

    Private Async Sub ImportAntarcticBasesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportAntarcticBasesToolStripMenuItem.Click
        Await ImportAntarcticBases()
    End Sub

    Private Async Sub ImportIOTAGroupsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportIOTAGroupsToolStripMenuItem.Click
        Await ImportIOTAGroups()
    End Sub

    Private Async Sub ImportAntarcticaToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportAntarcticaToolStripMenuItem.Click
        Await ImportAntarctica()
    End Sub

    Private Async Sub ImportIOTAIslandsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportIOTAIslandsToolStripMenuItem.Click
        Await ImportIOTAIslands()
    End Sub

    Private Async Sub ImportIOTADXCCMatchesOneIOTAToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportIOTADXCCMatchesOneIOTAToolStripMenuItem.Click
        Await ImportIOTADXCCMatchesOneIOTA()
    End Sub

    Private Sub ImportISO3166ToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportISO3166ToolStripMenuItem.Click
        ImportISO3166()
    End Sub

    Private Sub IOTACheckToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles IOTACheckToolStripMenuItem.Click
        IOTACheck()
    End Sub

    Private Sub ParseBoxToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ParseBoxToolStripMenuItem.Click
        ParseBox()
    End Sub

    Private Sub AdjacentColorCheckToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles AdjacentColorCheckToolStripMenuItem.Click
        AdjacentColourCheck()
    End Sub

    Private Async Sub ImportAntarcticaToolStripMenuItem_Click_1(sender As Object, e As EventArgs) _
        Handles ImportAntarcticaToolStripMenuItem.Click
        Await ImportAntarctica()
    End Sub

    Private Sub ImportPolyFromKMLToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportPolyFromKMLToolStripMenuItem.Click
        ImportPolyFromKML()
    End Sub

    Private Sub GeometrySizeTableToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles GeometrySizeTableToolStripMenuItem.Click
        GeometrySizeTable()
    End Sub

    Private Sub KMLFileSizeToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles KMLFileSizeToolStripMenuItem.Click
        KMLFileSize()
    End Sub

    Private Async Sub ImportLandSquareListToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ImportLandSquareListToolStripMenuItem.Click
        Await LandSquareList()
    End Sub

    Private Sub KMLntsToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles KMLntsToolStripMenuItem.Click
        KMLnts()
    End Sub

    Private Sub FindIARURegionToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles FindIARURegionToolStripMenuItem.Click
        FindIARURegion()
    End Sub

    Private Async Sub ConvertOPToToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ConvertOPToToolStripMenuItem.Click

        Dim sql As SqliteCommand, sqldr As SqliteDataReader, count = 0, sqlupd As SqliteCommand

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            sql = connect.CreateCommand
            sqlupd = connect.CreateCommand

            sql.CommandText = "SELECT count(*) FROM DXCC WHERE Entity <> 'Antarctica' AND (query <> '' OR bbox <> '') AND (ESRIjson IS NULL OR ESRIjson='')"
            count = SafeInt(sql.ExecuteScalar)

            ProgressBar1.Minimum = 0
            ProgressBar1.Value = 0
            ProgressBar1.Maximum = count

            sql.CommandText = "SELECT * FROM DXCC WHERE Entity <> 'Antarctica' AND (query <> '' OR bbox <> '') AND (ESRIjson IS NULL OR ESRIjson='')"
            sqldr = sql.ExecuteReader
            count = 0

            While sqldr.Read
                Await CreateGrids(connect, sqldr("Entity"))
                count += 1
                ProgressBar1.Value = count
            End While
        End Using
    End Sub

    Private Sub ValidateOSMQueriesToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles ValidateOSMQueriesToolStripMenuItem.Click
        ValidateDXCC()
    End Sub

    Private Sub TestAATToolStripMenuItem_Click(sender As Object, e As EventArgs) _
        Handles TestAATToolStripMenuItem.Click
        BuildAndSaveAllDxccGeometries()
    End Sub

    Private Async Sub ConvertLegacyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Await ConvertLegacy.ConvertLegacyQueriesAsync()
    End Sub

    Private Async Sub OSMAlternativesToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Await FindOSMAlternatives()
    End Sub
End Class
