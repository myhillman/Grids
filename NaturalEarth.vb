Imports System.Collections.Concurrent
Imports System.IO
Imports NetTopologySuite.Features
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.IO

Public Module NaturalEarth

    Private ReadOnly _cache As New ConcurrentDictionary(Of String, Geometry)(StringComparer.OrdinalIgnoreCase)
    Private _loaded As Boolean = False
    Private ReadOnly _lock As New Object()


    ' ------------------------------------------------------------
    ' PUBLIC LOOKUP
    ' ------------------------------------------------------------
    Private ReadOnly _wktReader As New NetTopologySuite.IO.WKTReader()

    Public Async Function Lookup(name As String) As Task(Of Geometry)

        Await EnsureLoadedAsync()
        Dim key = Normalize(name)

        ' 1. Fast path: direct key hit
        Dim GeoJson As Geometry = Nothing
        If _byKey.TryGetValue(key, GeoJson) Then
            Return GeoJson
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
                If attrs.Exists(fld) AndAlso attrs(fld) IsNot Nothing Then
                    Dim v = Normalize(attrs(fld).ToString())
                    If v = key Then

                        GeoJson = f.Geometry

                        ' Cache for next time
                        _byKey(key) = GeoJson

                        Return GeoJson
                    End If
                End If
            Next
        Next

        Return Nothing
    End Function


    ' ------------------------------------------------------------
    ' LOAD GeoJson data
    ' ------------------------------------------------------------
    Private Const NE_COUNTRIES As String = "D:\GIS DATA\Natural Earth\Countries\ne_10m_admin_0_countries.json"
    Private Const NE_SUBUNITS As String = "D:\GIS DATA\Natural Earth\MapSubUnits\ne_10m_admin_0_map_subunits.json"
    Private Const NE_MAPUNITS As String = "D:\GIS DATA\Natural Earth\MapUnits\ne_10m_admin_0_map_units.json"
    Private Const NE_REGIONPOLYS As String = "D:\GIS DATA\Natural Earth\RegionPolys\ne_10m_geography_regions_polys.json"
    Private Const NE_STATESPROVINCES As String = "D:\GIS DATA\Natural Earth\StatesProvinces\ne_10m_admin_1_states_provinces.json"
    Private Const NE_MINORISLANDS As String = "D:\GIS DATA\Natural Earth\MinorIslands\ne_10m_admin_0_scale_rank_minor_islands.json"
    Private Async Function EnsureLoadedAsync() As Task
        If _loaded Then Return

        SyncLock _lock
            If _loaded Then Return
            _loaded = True
        End SyncLock

        ' the features collection is first in best dressed, so add the smaller geometries first and work way up

        ' Load region polys
        'Await LoadGeoJsonAsync(NE_REGIONPOLYS)

        ' Load subunits
        Await LoadGeoJsonAsync(NE_SUBUNITS)

        ' Load mapunits
        Await LoadGeoJsonAsync(NE_MAPUNITS)

        ' Load countries
        Await LoadGeoJsonAsync(NE_COUNTRIES)

        ' Load countries
        Await LoadGeoJsonAsync(NE_MINORISLANDS)

        ' Load states & provinces
        Await LoadGeoJsonAsync(NE_STATESPROVINCES)

        ' After everything is loaded, save to SQLite
        SaveFeaturesToSqlite()

    End Function
    Private Function SaveFeaturesToSqlite()
        Debug.WriteLine("Saving features to SQLite database...")
        Using conn As New SqliteConnection(DXCC_DATA)
            conn.Open()

            ' Create table if not exists
            Dim createCmd = conn.CreateCommand()
            createCmd.CommandText =
                "
                CREATE TABLE IF NOT EXISTS ne_features (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    layer TEXT,
                    name TEXT,
                    attributes TEXT
                );
                "
            createCmd.ExecuteNonQuery()

            ' ⭐ BEGIN TRANSACTION ⭐
            Using tx = conn.BeginTransaction()

                ' Clear existing data
                Dim deleteCmd = conn.CreateCommand()
                deleteCmd.CommandText = "DELETE FROM ne_features;"
                deleteCmd.Transaction = tx
                Dim deleted = deleteCmd.ExecuteNonQuery()
                Debug.WriteLine($"{deleted} existing features deleted from database")

                ' Prepare insert command ONCE
                Dim insertCmd = conn.CreateCommand()
                insertCmd.CommandText = "INSERT INTO ne_features (layer, name, attributes) VALUES ($layer, $name, $attributes);"
                insertCmd.Parameters.Add("$layer", SqliteType.Text)
                insertCmd.Parameters.Add("$name", SqliteType.Text)
                insertCmd.Parameters.Add("$attributes", SqliteType.Text)
                insertCmd.Transaction = tx

                Dim writer As New NetTopologySuite.IO.GeoJsonWriter()

                ' Insert each feature
                For Each f In _features
                    Dim layer = TryGet(f.Attributes, "LAYER")
                    Dim name = ExtractName(f)
                    If String.IsNullOrWhiteSpace(name) Then
                        Debug.WriteLine("NULL NAME: " & Newtonsoft.Json.JsonConvert.SerializeObject(f.Attributes))
                    End If
                    Dim attrJson = Newtonsoft.Json.JsonConvert.SerializeObject(f.Attributes)

                    insertCmd.Parameters("$layer").Value = layer
                    insertCmd.Parameters("$name").Value = name
                    insertCmd.Parameters("$attributes").Value = attrJson

                    insertCmd.ExecuteNonQuery()
                Next

                ' ⭐ COMMIT ⭐
                tx.Commit()
                Debug.WriteLine($"{_features.Count} features saved to database")
            End Using
        End Using
    End Function

    ' ------------------------------------------------------------
    ' SHAPEFILE LOADER
    ' ------------------------------------------------------------
    Private ReadOnly _byKey As New SortedDictionary(Of String, Geometry)(StringComparer.OrdinalIgnoreCase)
    Private ReadOnly _features As New List(Of NetTopologySuite.Features.Feature)
    Private _featureLayer As New Dictionary(Of String, String)
    Private ReadOnly _layerOrder As New List(Of String)


    Private Async Function LoadGeoJsonAsync(path As String) As Task
        If Not File.Exists(path) Then
            Throw New FileNotFoundException($"Natural Earth GeoJSON file not found: {path}")
        End If

        Dim json As String = Await File.ReadAllTextAsync(path)
        Dim reader As New NetTopologySuite.IO.GeoJsonReader()
        Dim fc = reader.Read(Of NetTopologySuite.Features.FeatureCollection)(json)

        Dim layer = System.IO.Path.GetFileNameWithoutExtension(path)

        Debug.WriteLine($"{fc.Count} features loaded for layer {layer}")

        ' Register layer in priority list if not already present
        If Not _layerOrder.Contains(layer.ToLower()) Then
            _layerOrder.Add(layer.ToLower())
        End If

        For Each f In fc

            ' Extract name
            Dim name = ExtractName(f)
            If String.IsNullOrWhiteSpace(name) Then Continue For

            Dim key = Normalize(name)

            ' Tag layer
            f.Attributes.Add("LAYER", layer)
            _features.Add(f)

            ' Repair geometry if needed
            Dim geom = f.Geometry
            If geom IsNot Nothing AndAlso Not geom.IsValid Then
                geom = geom.Buffer(0)
            End If

            ' If we have no entry yet → add it
            If Not _byKey.ContainsKey(key) Then
                _byKey(key) = geom
                _featurelayer(key) = layer
                Continue For
            End If

            ' Compare layer priority
            Dim existingLayer = _featureLayer(key)
            If LayerPriority(layer) > LayerPriority(existingLayer) Then
                ' Replace with higher‑priority geometry
                _byKey(key) = geom
                _featureLayer(key) = layer
                Continue For
            End If

            ' Same name, same or lower priority → merge geometry
            Dim merged = NetTopologySuite.Operation.Union.UnaryUnionOp.Union(
            New List(Of Geometry) From {_byKey(key), geom}
        )

            _byKey(key) = merged
        Next
    End Function
    Private Function LayerPriority(layer As String) As Integer
        Dim idx = _layerOrder.IndexOf(layer.ToLower())
        If idx = -1 Then Return Integer.MaxValue   ' unknown layer = lowest priority
        Return idx
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

        Dim admin = StripSuffix(TryGet(attrs, "ADMIN"))
        Dim subunit = StripSuffix(TryGet(attrs, "SUBUNIT"))

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


        ' ⭐ NEW FALLBACKS ⭐

        Dim geonunit = TryGet(attrs, "GEONUNIT")
        If Not String.IsNullOrEmpty(geonunit) Then Return geonunit

        Dim mapunit = TryGet(attrs, "MAPUNIT")
        If Not String.IsNullOrEmpty(mapunit) Then Return mapunit

        Dim label = TryGet(attrs, "LABEL")
        If Not String.IsNullOrEmpty(label) Then Return label

        Dim labelEn = TryGet(attrs, "LABEL_EN")
        If Not String.IsNullOrEmpty(labelEn) Then Return labelEn

        Dim nameLong = TryGet(attrs, "NAME_LONG")
        If Not String.IsNullOrEmpty(nameLong) Then Return nameLong

        Dim abbrev = TryGet(attrs, "ABBREV")
        If Not String.IsNullOrEmpty(abbrev) Then Return abbrev

        Dim isoA2 = TryGet(attrs, "ISO_A2")
        If Not String.IsNullOrEmpty(isoA2) Then Return isoA2

        Dim isoA3 = TryGet(attrs, "ISO_A3")
        If Not String.IsNullOrEmpty(isoA3) Then Return isoA3

        Return Nothing
    End Function

    Private Function TryGet(attrs As AttributesTable, field As String) As String
        If attrs.Exists(field) Then
            Dim v = attrs(field)
            If v IsNot Nothing Then Return v.ToString()
        End If
        Return Nothing
    End Function
    Private Function StripSuffix(value As String) As String
        If value Is Nothing Then Return Nothing

        Dim i = value.IndexOf(" (")
        If i > 0 Then
            Return value.Substring(0, i)
        End If

        Return value
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
