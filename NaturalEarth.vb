Imports Esri.ArcGISRuntime.Geometry
Imports Esri.ArcGISRuntime.Data
Imports System.Collections.Concurrent

Public Module NaturalEarth

    Private Const NE_COUNTRIES As String = "D:\GIS DATA\Natural Earth\Countries\ne_10m_admin_0_countries.shp"
    Private Const NE_SUBUNITS As String = "D:\GIS DATA\Natural Earth\MapSubUnits\ne_10m_admin_0_map_subunits.shp"
    Private Const NE_MAPUNITS As String = "D:\GIS DATA\Natural Earth\MapUnits\ne_10m_admin_0_map_units.shp"
    Private Const NE_REGIONPOLYS As String = "D:\GIS DATA\Natural Earth\RegionPolys\ne_10m_geography_regions_polys.shp"
    Private Const NE_STATESPROVINCES As String = "D:\GIS DATA\Natural Earth\StatesProvinces\ne_10m_admin_1_states_provinces.shp"

    Private ReadOnly _cache As New ConcurrentDictionary(Of String, Geometry)(StringComparer.OrdinalIgnoreCase)
    Private _loaded As Boolean = False
    Private ReadOnly _lock As New Object()


    ' ------------------------------------------------------------
    ' PUBLIC LOOKUP
    ' ------------------------------------------------------------
    Public Async Function Lookup(name As String) As Task(Of Geometry)

        Await EnsureLoadedAsync()
        Dim key = Normalize(name)

        ' 1. Fast path: direct key hit
        Dim g As Geometry = Nothing
        If _byKey.TryGetValue(key, g) Then
            Return g
        End If

        ' 2. Slow path: search across multiple attributes
        Dim fields() As String = {
        "SUBUNIT",
        "ADMIN",
        "NAME_EN",
        "NAME",
        "BRK_NAME",
        "SOVEREIGNT"
    }

        For Each f In _features
            Dim attrs = f.Attributes

            For Each fld In fields
                If attrs.ContainsKey(fld) AndAlso attrs(fld) IsNot Nothing Then
                    Dim v = Normalize(attrs(fld).ToString())
                    If v = key Then
                        g = f.Geometry
                        ' cache for next time
                        _byKey(key) = g
                        Return g
                    End If
                End If
            Next
        Next

        Return Nothing
    End Function

    ' ------------------------------------------------------------
    ' LOAD BOTH SHAPEFILES
    ' ------------------------------------------------------------
    Private Async Function EnsureLoadedAsync() As Task
        If _loaded Then Return

        SyncLock _lock
            If _loaded Then Return
            _loaded = True
        End SyncLock

        ' Load countries
        Await LoadShapefileAsync(NE_COUNTRIES)

        ' Load mapunits
        Await LoadShapefileAsync(NE_MAPUNITS)

        ' Load subunits
        Await LoadShapefileAsync(NE_SUBUNITS)

        ' Load region polys
        Await LoadShapefileAsync(NE_REGIONPOLYS)

        ' Load states & provinces
        Await LoadShapefileAsync(NE_STATESPROVINCES)

    End Function


    ' ------------------------------------------------------------
    ' SHAPEFILE LOADER
    ' ------------------------------------------------------------
    Private ReadOnly _byKey As New SortedDictionary(Of String, Geometry)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _features As New List(Of Feature)

    Private Async Function LoadShapefileAsync(path As String) As Task
        Dim table As New ShapefileFeatureTable(path)
        Await table.LoadAsync()

        Dim features = Await table.QueryFeaturesAsync(New QueryParameters())

        For Each f In features
            _features.Add(f)

            Dim name = ExtractName(f)
            If String.IsNullOrWhiteSpace(name) Then Continue For

            Dim key = Normalize(name)
            If Not _byKey.ContainsKey(key) Then
                _byKey.Add(key, f.Geometry)
            End If
        Next
        Debug.WriteLine($"Loaded {features.Count} features from {path}")
    End Function

    ' ------------------------------------------------------------
    ' ATTRIBUTE NAME EXTRACTION
    ' ------------------------------------------------------------
    Private Function ExtractName(f As Feature) As String

        ' Skip parent territories that hide child islands
        Dim skipParents As String() = {
        "Saint Helena",
        "France",
        "United Kingdom",
        "Portugal",
        "Spain",
        "Ecuador"
    }

        Dim attrs = f.Attributes
        Dim admin = TryGet(attrs, "ADMIN")
        Dim subunit = TryGet(attrs, "SUBUNIT")

        ' 1. SUBUNIT (if not a parent)
        If Not String.IsNullOrEmpty(subunit) AndAlso Not skipParents.Contains(subunit) Then
            Return subunit
        End If

        ' 2. ADMIN (if not a parent)
        If Not String.IsNullOrEmpty(admin) AndAlso Not skipParents.Contains(admin) Then
            Return admin
        End If

        ' 3. NAME_EN
        Dim nameEn = TryGet(attrs, "NAME_EN")
        If Not String.IsNullOrEmpty(nameEn) Then Return nameEn

        ' 4. NAME
        Dim name = TryGet(attrs, "NAME")
        If Not String.IsNullOrEmpty(name) Then Return name

        ' 5. BRK_NAME
        Dim brk = TryGet(attrs, "BRK_NAME")
        If Not String.IsNullOrEmpty(brk) Then Return brk

        ' 6. SOVEREIGNT
        Dim sov = TryGet(attrs, "SOVEREIGNT")
        If Not String.IsNullOrEmpty(sov) Then Return sov

        Return Nothing
    End Function
    Private Function TryGet(attrs As IDictionary(Of String, Object), key As String) As String
        If attrs.ContainsKey(key) AndAlso attrs(key) IsNot Nothing Then
            Return attrs(key).ToString().Trim()
        End If
        Return Nothing
    End Function

    ' ------------------------------------------------------------
    ' NORMALIZATION
    ' ------------------------------------------------------------
    Private Function Normalize(s As String) As String
        If s Is Nothing Then Return ""
        Return Clean(s).ToUpperInvariant()
    End Function
    Private Function Clean(s As String) As String
        If s Is Nothing Then Return Nothing
        Dim sb As New System.Text.StringBuilder(s.Length)
        For Each ch In s
            If ch >= " "c AndAlso ch <= "~"c Then
                sb.Append(ch)
            End If
        Next
        Return sb.ToString().Trim()
    End Function

End Module
