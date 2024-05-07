<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        MenuStrip1 = New MenuStrip()
        UseShapefileToolStripMenuItem = New ToolStripMenuItem()
        UseShapefileToolStripMenuItem1 = New ToolStripMenuItem()
        UseOSMToolStripMenuItem = New ToolStripMenuItem()
        MakeKMLToolStripMenuItem = New ToolStripMenuItem()
        MakeKMLByCQZoneToolStripMenuItem = New ToolStripMenuItem()
        MakeKMLAllEntitiesToolStripMenuItem = New ToolStripMenuItem()
        MakeShapefileToolStripMenuItem = New ToolStripMenuItem()
        EntityReportToolStripMenuItem = New ToolStripMenuItem()
        ReprocessCountryToolStripMenuItem = New ToolStripMenuItem()
        ExitToolStripMenuItem = New ToolStripMenuItem()
        ImportToolStripMenuItem = New ToolStripMenuItem()
        ImportISO3166ToolStripMenuItem = New ToolStripMenuItem()
        ImportEUASBorderToolStripMenuItem = New ToolStripMenuItem()
        ImportCQITUZonesToolStripMenuItem = New ToolStripMenuItem()
        ImportTimezonesToolStripMenuItem = New ToolStripMenuItem()
        ImportIARURegionsToolStripMenuItem = New ToolStripMenuItem()
        UseOSMLandPolygonsToolStripMenuItem = New ToolStripMenuItem()
        ImportAntarcticBasesToolStripMenuItem = New ToolStripMenuItem()
        ImportIOTAToolStripMenuItem = New ToolStripMenuItem()
        ImportIOTAGroupsToolStripMenuItem = New ToolStripMenuItem()
        ImportIOTAIslandsToolStripMenuItem = New ToolStripMenuItem()
        ImportIOTADXCCMatchesOneIOTAToolStripMenuItem = New ToolStripMenuItem()
        ImportAntarcticaToolStripMenuItem = New ToolStripMenuItem()
        ImportPolyFromKMLToolStripMenuItem = New ToolStripMenuItem()
        ChecksToolStripMenuItem = New ToolStripMenuItem()
        CheckISO3166ReferencesToolStripMenuItem = New ToolStripMenuItem()
        CountryCollisionsToolStripMenuItem = New ToolStripMenuItem()
        RemoveEmptyGeometryToolStripMenuItem = New ToolStripMenuItem()
        IOTACheckToolStripMenuItem = New ToolStripMenuItem()
        InnerRingsToolStripMenuItem = New ToolStripMenuItem()
        AdjacentColorCheckToolStripMenuItem = New ToolStripMenuItem()
        GeometrySizeTableToolStripMenuItem = New ToolStripMenuItem()
        TestToolStripMenuItem = New ToolStripMenuItem()
        ParseBoxToolStripMenuItem = New ToolStripMenuItem()
        TextBox1 = New TextBox()
        ProgressBar1 = New ProgressBar()
        DummyToolStripMenuItem = New ToolStripMenuItem()
        OpenFileDialog1 = New OpenFileDialog()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {UseShapefileToolStripMenuItem, ImportToolStripMenuItem, ChecksToolStripMenuItem, TestToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(1286, 24)
        MenuStrip1.TabIndex = 0
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' UseShapefileToolStripMenuItem
        ' 
        UseShapefileToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {UseShapefileToolStripMenuItem1, UseOSMToolStripMenuItem, MakeKMLToolStripMenuItem, MakeKMLByCQZoneToolStripMenuItem, MakeKMLAllEntitiesToolStripMenuItem, MakeShapefileToolStripMenuItem, EntityReportToolStripMenuItem, ReprocessCountryToolStripMenuItem, ExitToolStripMenuItem})
        UseShapefileToolStripMenuItem.Name = "UseShapefileToolStripMenuItem"
        UseShapefileToolStripMenuItem.Size = New Size(37, 20)
        UseShapefileToolStripMenuItem.Text = "File"
        ' 
        ' UseShapefileToolStripMenuItem1
        ' 
        UseShapefileToolStripMenuItem1.Name = "UseShapefileToolStripMenuItem1"
        UseShapefileToolStripMenuItem1.Size = New Size(194, 22)
        UseShapefileToolStripMenuItem1.Text = "Use shapefile"
        ' 
        ' UseOSMToolStripMenuItem
        ' 
        UseOSMToolStripMenuItem.Name = "UseOSMToolStripMenuItem"
        UseOSMToolStripMenuItem.Size = New Size(194, 22)
        UseOSMToolStripMenuItem.Text = "Use OSM"
        ' 
        ' MakeKMLToolStripMenuItem
        ' 
        MakeKMLToolStripMenuItem.Name = "MakeKMLToolStripMenuItem"
        MakeKMLToolStripMenuItem.Size = New Size(194, 22)
        MakeKMLToolStripMenuItem.Text = "Make KML"
        ' 
        ' MakeKMLByCQZoneToolStripMenuItem
        ' 
        MakeKMLByCQZoneToolStripMenuItem.Name = "MakeKMLByCQZoneToolStripMenuItem"
        MakeKMLByCQZoneToolStripMenuItem.Size = New Size(194, 22)
        MakeKMLByCQZoneToolStripMenuItem.Text = "Make KML by CQ zone"
        ' 
        ' MakeKMLAllEntitiesToolStripMenuItem
        ' 
        MakeKMLAllEntitiesToolStripMenuItem.Name = "MakeKMLAllEntitiesToolStripMenuItem"
        MakeKMLAllEntitiesToolStripMenuItem.Size = New Size(194, 22)
        MakeKMLAllEntitiesToolStripMenuItem.Text = "Make KML all entities"
        ' 
        ' MakeShapefileToolStripMenuItem
        ' 
        MakeShapefileToolStripMenuItem.Name = "MakeShapefileToolStripMenuItem"
        MakeShapefileToolStripMenuItem.Size = New Size(194, 22)
        MakeShapefileToolStripMenuItem.Text = "Make Shapefile"
        ' 
        ' EntityReportToolStripMenuItem
        ' 
        EntityReportToolStripMenuItem.Name = "EntityReportToolStripMenuItem"
        EntityReportToolStripMenuItem.Size = New Size(194, 22)
        EntityReportToolStripMenuItem.Text = "Entity Report"
        ' 
        ' ReprocessCountryToolStripMenuItem
        ' 
        ReprocessCountryToolStripMenuItem.Name = "ReprocessCountryToolStripMenuItem"
        ReprocessCountryToolStripMenuItem.Size = New Size(194, 22)
        ReprocessCountryToolStripMenuItem.Text = "Reprocess country"
        ' 
        ' ExitToolStripMenuItem
        ' 
        ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        ExitToolStripMenuItem.Size = New Size(194, 22)
        ExitToolStripMenuItem.Text = "Exit"
        ' 
        ' ImportToolStripMenuItem
        ' 
        ImportToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ImportISO3166ToolStripMenuItem, ImportEUASBorderToolStripMenuItem, ImportCQITUZonesToolStripMenuItem, ImportTimezonesToolStripMenuItem, ImportIARURegionsToolStripMenuItem, UseOSMLandPolygonsToolStripMenuItem, ImportAntarcticBasesToolStripMenuItem, ImportIOTAToolStripMenuItem, ImportAntarcticaToolStripMenuItem, ImportPolyFromKMLToolStripMenuItem})
        ImportToolStripMenuItem.Name = "ImportToolStripMenuItem"
        ImportToolStripMenuItem.Size = New Size(55, 20)
        ImportToolStripMenuItem.Text = "Import"
        ' 
        ' ImportISO3166ToolStripMenuItem
        ' 
        ImportISO3166ToolStripMenuItem.Name = "ImportISO3166ToolStripMenuItem"
        ImportISO3166ToolStripMenuItem.Size = New Size(217, 22)
        ImportISO3166ToolStripMenuItem.Text = "Import ISO3166"
        ' 
        ' ImportEUASBorderToolStripMenuItem
        ' 
        ImportEUASBorderToolStripMenuItem.Name = "ImportEUASBorderToolStripMenuItem"
        ImportEUASBorderToolStripMenuItem.Size = New Size(217, 22)
        ImportEUASBorderToolStripMenuItem.Text = "Import EU/AS border"
        ' 
        ' ImportCQITUZonesToolStripMenuItem
        ' 
        ImportCQITUZonesToolStripMenuItem.Name = "ImportCQITUZonesToolStripMenuItem"
        ImportCQITUZonesToolStripMenuItem.Size = New Size(217, 22)
        ImportCQITUZonesToolStripMenuItem.Text = "Import CQ/ITU zones"
        ' 
        ' ImportTimezonesToolStripMenuItem
        ' 
        ImportTimezonesToolStripMenuItem.Name = "ImportTimezonesToolStripMenuItem"
        ImportTimezonesToolStripMenuItem.Size = New Size(217, 22)
        ImportTimezonesToolStripMenuItem.Text = "Import Timezones"
        ' 
        ' ImportIARURegionsToolStripMenuItem
        ' 
        ImportIARURegionsToolStripMenuItem.Name = "ImportIARURegionsToolStripMenuItem"
        ImportIARURegionsToolStripMenuItem.Size = New Size(217, 22)
        ImportIARURegionsToolStripMenuItem.Text = "Import IARU regions"
        ' 
        ' UseOSMLandPolygonsToolStripMenuItem
        ' 
        UseOSMLandPolygonsToolStripMenuItem.Name = "UseOSMLandPolygonsToolStripMenuItem"
        UseOSMLandPolygonsToolStripMenuItem.Size = New Size(217, 22)
        UseOSMLandPolygonsToolStripMenuItem.Text = "Import OSM land polygons"
        ' 
        ' ImportAntarcticBasesToolStripMenuItem
        ' 
        ImportAntarcticBasesToolStripMenuItem.Name = "ImportAntarcticBasesToolStripMenuItem"
        ImportAntarcticBasesToolStripMenuItem.Size = New Size(217, 22)
        ImportAntarcticBasesToolStripMenuItem.Text = "Import Antarctic Bases"
        ' 
        ' ImportIOTAToolStripMenuItem
        ' 
        ImportIOTAToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ImportIOTAGroupsToolStripMenuItem, ImportIOTAIslandsToolStripMenuItem, ImportIOTADXCCMatchesOneIOTAToolStripMenuItem})
        ImportIOTAToolStripMenuItem.Name = "ImportIOTAToolStripMenuItem"
        ImportIOTAToolStripMenuItem.Size = New Size(217, 22)
        ImportIOTAToolStripMenuItem.Text = "Import IOTA"
        ' 
        ' ImportIOTAGroupsToolStripMenuItem
        ' 
        ImportIOTAGroupsToolStripMenuItem.Name = "ImportIOTAGroupsToolStripMenuItem"
        ImportIOTAGroupsToolStripMenuItem.Size = New Size(269, 22)
        ImportIOTAGroupsToolStripMenuItem.Text = "Import IOTA Groups"
        ' 
        ' ImportIOTAIslandsToolStripMenuItem
        ' 
        ImportIOTAIslandsToolStripMenuItem.Name = "ImportIOTAIslandsToolStripMenuItem"
        ImportIOTAIslandsToolStripMenuItem.Size = New Size(269, 22)
        ImportIOTAIslandsToolStripMenuItem.Text = "Import IOTA Islands"
        ' 
        ' ImportIOTADXCCMatchesOneIOTAToolStripMenuItem
        ' 
        ImportIOTADXCCMatchesOneIOTAToolStripMenuItem.Name = "ImportIOTADXCCMatchesOneIOTAToolStripMenuItem"
        ImportIOTADXCCMatchesOneIOTAToolStripMenuItem.Size = New Size(269, 22)
        ImportIOTADXCCMatchesOneIOTAToolStripMenuItem.Text = "Import IOTA DXCC matches one IOTA"
        ' 
        ' ImportAntarcticaToolStripMenuItem
        ' 
        ImportAntarcticaToolStripMenuItem.Name = "ImportAntarcticaToolStripMenuItem"
        ImportAntarcticaToolStripMenuItem.Size = New Size(217, 22)
        ImportAntarcticaToolStripMenuItem.Text = "Import Antarctica"
        ' 
        ' ImportPolyFromKMLToolStripMenuItem
        ' 
        ImportPolyFromKMLToolStripMenuItem.Name = "ImportPolyFromKMLToolStripMenuItem"
        ImportPolyFromKMLToolStripMenuItem.Size = New Size(217, 22)
        ImportPolyFromKMLToolStripMenuItem.Text = "Import poly from KML"
        ' 
        ' ChecksToolStripMenuItem
        ' 
        ChecksToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {CheckISO3166ReferencesToolStripMenuItem, CountryCollisionsToolStripMenuItem, RemoveEmptyGeometryToolStripMenuItem, IOTACheckToolStripMenuItem, InnerRingsToolStripMenuItem, AdjacentColorCheckToolStripMenuItem, GeometrySizeTableToolStripMenuItem})
        ChecksToolStripMenuItem.Name = "ChecksToolStripMenuItem"
        ChecksToolStripMenuItem.Size = New Size(57, 20)
        ChecksToolStripMenuItem.Text = "Checks"
        ' 
        ' CheckISO3166ReferencesToolStripMenuItem
        ' 
        CheckISO3166ReferencesToolStripMenuItem.Name = "CheckISO3166ReferencesToolStripMenuItem"
        CheckISO3166ReferencesToolStripMenuItem.Size = New Size(209, 22)
        CheckISO3166ReferencesToolStripMenuItem.Text = "Check ISO3166 references"
        ' 
        ' CountryCollisionsToolStripMenuItem
        ' 
        CountryCollisionsToolStripMenuItem.Name = "CountryCollisionsToolStripMenuItem"
        CountryCollisionsToolStripMenuItem.Size = New Size(209, 22)
        CountryCollisionsToolStripMenuItem.Text = "Country Collisions"
        ' 
        ' RemoveEmptyGeometryToolStripMenuItem
        ' 
        RemoveEmptyGeometryToolStripMenuItem.Name = "RemoveEmptyGeometryToolStripMenuItem"
        RemoveEmptyGeometryToolStripMenuItem.Size = New Size(209, 22)
        RemoveEmptyGeometryToolStripMenuItem.Text = "Remove empty geometry"
        ' 
        ' IOTACheckToolStripMenuItem
        ' 
        IOTACheckToolStripMenuItem.Name = "IOTACheckToolStripMenuItem"
        IOTACheckToolStripMenuItem.Size = New Size(209, 22)
        IOTACheckToolStripMenuItem.Text = "IOTA check"
        ' 
        ' InnerRingsToolStripMenuItem
        ' 
        InnerRingsToolStripMenuItem.Name = "InnerRingsToolStripMenuItem"
        InnerRingsToolStripMenuItem.Size = New Size(209, 22)
        InnerRingsToolStripMenuItem.Text = "Inner rings"
        ' 
        ' AdjacentColorCheckToolStripMenuItem
        ' 
        AdjacentColorCheckToolStripMenuItem.Name = "AdjacentColorCheckToolStripMenuItem"
        AdjacentColorCheckToolStripMenuItem.Size = New Size(209, 22)
        AdjacentColorCheckToolStripMenuItem.Text = "Adjacent color check"
        ' 
        ' GeometrySizeTableToolStripMenuItem
        ' 
        GeometrySizeTableToolStripMenuItem.Name = "GeometrySizeTableToolStripMenuItem"
        GeometrySizeTableToolStripMenuItem.Size = New Size(209, 22)
        GeometrySizeTableToolStripMenuItem.Text = "Geometry size table"
        ' 
        ' TestToolStripMenuItem
        ' 
        TestToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {ParseBoxToolStripMenuItem})
        TestToolStripMenuItem.Name = "TestToolStripMenuItem"
        TestToolStripMenuItem.Size = New Size(39, 20)
        TestToolStripMenuItem.Text = "Test"
        ' 
        ' ParseBoxToolStripMenuItem
        ' 
        ParseBoxToolStripMenuItem.Name = "ParseBoxToolStripMenuItem"
        ParseBoxToolStripMenuItem.Size = New Size(122, 22)
        ParseBoxToolStripMenuItem.Text = "ParseBox"
        ' 
        ' TextBox1
        ' 
        TextBox1.Location = New Point(12, 27)
        TextBox1.Multiline = True
        TextBox1.Name = "TextBox1"
        TextBox1.ScrollBars = ScrollBars.Both
        TextBox1.Size = New Size(1242, 422)
        TextBox1.TabIndex = 1
        ' 
        ' ProgressBar1
        ' 
        ProgressBar1.Location = New Point(12, 465)
        ProgressBar1.Name = "ProgressBar1"
        ProgressBar1.Size = New Size(776, 21)
        ProgressBar1.TabIndex = 2
        ' 
        ' DummyToolStripMenuItem
        ' 
        DummyToolStripMenuItem.Name = "DummyToolStripMenuItem"
        DummyToolStripMenuItem.Size = New Size(32, 19)
        ' 
        ' OpenFileDialog1
        ' 
        OpenFileDialog1.FileName = "OpenFileDialog1"
        ' 
        ' Form1
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(1286, 502)
        Controls.Add(ProgressBar1)
        Controls.Add(TextBox1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "Form1"
        Text = "Grids"
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents UseShapefileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UseShapefileToolStripMenuItem1 As ToolStripMenuItem
    Friend WithEvents UseOSMToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MakeKMLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents MakeShapefileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents EntityReportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportISO3166ToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MakeKMLByCQZoneToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ReprocessCountryToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CheckISO3166ReferencesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents CountryCollisionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents RemoveEmptyGeometryToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportEUASBorderToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents UseOSMLandPolygonsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportCQITUZonesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportTimezonesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents DummyToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportIARURegionsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents MakeKMLAllEntitiesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents InnerRingsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportAntarcticBasesToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportIOTAToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportIOTAGroupsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportIOTAIslandsToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportIOTADXCCMatchesOneIOTAToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents IOTACheckToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents TestToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ParseBoxToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents AdjacentColorCheckToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ChecksToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportAntarcticaToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents ImportPolyFromKMLToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents GeometrySizeTableToolStripMenuItem As ToolStripMenuItem

End Class
