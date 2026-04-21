Imports esriArcGIS.runtime.Geometry
Imports NetTopologySuite.Geometries
Module esriArcGIS
    ' Module holds all functions requiring Esri ArcGIS Runtime SDK, to isolate it from the rest of the codebase and avoid unnecessary dependencies

    ''' <summary>
    ''' Converts an ArcGIS Runtime geometry (in ArcGIS JSON format) into an
    ''' equivalent GeoJSON geometry string.  
    ''' <para>
    ''' Supports Point (<c>x,y</c>), LineString and MultiLineString
    ''' (<c>paths</c>), Polygon (<c>rings</c>), and MultiPolygon
    ''' (<c>parts → rings</c>).  
    ''' The function inspects the ArcGIS JSON structure and emits the
    ''' corresponding GeoJSON <c>type</c> and <c>coordinates</c> array.
    ''' </para>
    ''' </summary>
    ''' <param name="geom">
    ''' The ArcGIS Runtime geometry to convert. Its <c>ToJson()</c> output is
    ''' parsed to determine the geometry type and coordinate structure.
    ''' </param>
    ''' <returns>
    ''' A valid GeoJSON geometry string representing the input geometry.
    ''' </returns>
    ''' <exception cref="Exception">
    ''' Thrown when the ArcGIS geometry type is not recognised or does not
    ''' contain any of the supported ArcGIS JSON properties (<c>x,y</c>,
    ''' <c>paths</c>, <c>rings</c>, <c>parts</c>).
    ''' </exception>
    Public Function FromArcGisToGeoJson(geom As Esri.ArcGISRuntime.Geometry.Geometry) As String
        Dim jsonarcgeom As String = geom.ToJson()
        Dim jo = System.Text.Json.JsonDocument.Parse(jsonarcgeom).RootElement

        ' ---------------------------------------------------------
        ' POINT  (ArcGIS: x,y)
        ' ---------------------------------------------------------
        If jo.TryGetProperty("x", Nothing) AndAlso jo.TryGetProperty("y", Nothing) Then
            Dim x = jo.GetProperty("x").GetDouble()
            Dim y = jo.GetProperty("y").GetDouble()
            Return $"{{ ""type"": ""Point"", ""coordinates"": [{x}, {y}] }}"
        End If

        ' ---------------------------------------------------------
        ' LINESTRING or MULTILINESTRING (ArcGIS: "paths")
        ' ---------------------------------------------------------
        If jo.TryGetProperty("paths", Nothing) Then
            Dim paths = jo.GetProperty("paths")

            If paths.GetArrayLength() = 1 Then
                ' Single path → LineString
                Return $"{{ ""type"": ""LineString"", ""coordinates"": {paths(0).ToString()} }}"
            Else
                ' Multiple paths → MultiLineString
                Return $"{{ ""type"": ""MultiLineString"", ""coordinates"": {paths.ToString()} }}"
            End If
        End If

        ' ---------------------------------------------------------
        ' POLYGON (ArcGIS: "rings")
        ' ---------------------------------------------------------
        If jo.TryGetProperty("rings", Nothing) Then
            Dim rings = jo.GetProperty("rings")
            Return $"{{ ""type"": ""Polygon"", ""coordinates"": {rings.ToString()} }}"
        End If

        ' ---------------------------------------------------------
        ' MULTIPOLYGON (ArcGIS: "parts")
        ' ---------------------------------------------------------
        If jo.TryGetProperty("parts", Nothing) Then
            Dim parts = jo.GetProperty("parts")

            Dim polys As New List(Of String)
            For Each part In parts.EnumerateArray()
                Dim rings = part.GetProperty("rings")
                polys.Add(rings.ToString())
            Next

            Dim joined = String.Join(",", polys.Select(Function(p) $"[{p}]"))
            Return $"{{ ""type"": ""MultiPolygon"", ""coordinates"": [{joined}] }}"
        End If

        Throw New Exception("Unsupported ArcGIS geometry type.")
    End Function

    Public Function FromGeoJsonToArcGis(geoJson As String) _
        As Esri.ArcGISRuntime.Geometry.Geometry

        ' 1. GeoJSON → NTS
        Dim reader As New NetTopologySuite.IO.GeoJsonReader()
        Dim ntsGeom As NetTopologySuite.Geometries.Geometry =
                reader.Read(Of NetTopologySuite.Geometries.Geometry)(geoJson)

        ' 2. NTS → ArcGIS JSON
        Dim arcJson As String = NtsToArcGisJson(ntsGeom)

        ' 3. ArcGIS JSON → ArcGIS Runtime geometry
        Return Esri.ArcGISRuntime.Geometry.Geometry.FromJson(arcJson)
    End Function
    Private Function NtsToArcGisJson(geom As NetTopologySuite.Geometries.Geometry) As String

        If TypeOf geom Is NetTopologySuite.Geometries.Polygon Then
            Dim poly = DirectCast(geom, NetTopologySuite.Geometries.Polygon)

            Dim rings = New List(Of List(Of Double()))()

            ' Shell
            Dim shell = poly.Shell.Coordinates.Select(
                Function(c) New Double() {c.X, c.Y}
                ).ToList()
            rings.Add(shell)

            ' Holes
            For i = 0 To poly.NumInteriorRings - 1
                Dim hole = poly.GetInteriorRingN(i).Coordinates.Select(
                    Function(c) New Double() {c.X, c.Y}
                    ).ToList()
                rings.Add(hole)
            Next

            Return $"{{ ""rings"": {System.Text.Json.JsonSerializer.Serialize(rings)}, ""spatialReference"": {{ ""wkid"": 4326 }} }}"
        End If

        If TypeOf geom Is NetTopologySuite.Geometries.MultiPolygon Then
            Dim mp = DirectCast(geom, NetTopologySuite.Geometries.MultiPolygon)

            Dim allRings = New List(Of List(Of Double()))()

            For i = 0 To mp.NumGeometries - 1
                Dim poly = DirectCast(mp.GetGeometryN(i), NetTopologySuite.Geometries.Polygon)

                ' Shell
                Dim shell = poly.Shell.Coordinates.Select(
                    Function(c) New Double() {c.X, c.Y}
                    ).ToList()
                allRings.Add(shell)

                ' Holes
                For h = 0 To poly.NumInteriorRings - 1
                    Dim hole = poly.GetInteriorRingN(h).Coordinates.Select(
                        Function(c) New Double() {c.X, c.Y}
                        ).ToList()
                    allRings.Add(hole)
                Next
            Next

            Return $"{{ ""rings"": {System.Text.Json.JsonSerializer.Serialize(allRings)}, ""spatialReference"": {{ ""wkid"": 4326 }} }}"
        End If

        Throw New Exception("Unsupported geometry type for ArcGIS JSON conversion.")
    End Function

    Function FromArcGisToNTS(geom As Esri.ArcGISRuntime.Geometry.Geometry) As NetTopologySuite.Geometries.Geometry
        Dim geoJson = FromArcGisToGeoJson(geom)
        Return FromGeoJsonToNTS(geoJson)
    End Function
End Module
