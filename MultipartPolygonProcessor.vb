Imports Esri.ArcGISRuntime.Geometry

Public Module MultipartPolygonProcessor

    Private Const DistanceThresholdMeters As Double = 350000 ' 350 km

    ''' <summary>
    ''' Takes a list of Esri polygons (each representing a separate island/part),
    ''' identifies the mainland, inserts holes, removes remote islands,
    ''' and returns a single multipart Esri Polygon.
    ''' </summary>
    Public Function Process(polygons As List(Of Polygon)) As Polygon

        If polygons Is Nothing OrElse polygons.Count = 0 Then
            Return Nothing
        End If

        Dim sr = polygons(0).SpatialReference

        ' --- 1. Identify mainland (largest area) ---
        Dim mainland As Polygon =
            polygons.OrderByDescending(Function(p) Math.Abs(GeometryEngine.Area(p))).First()

        Dim otherPolys = polygons.Where(Function(p) Not p.Equals(mainland)).ToList()

        Dim holes As New List(Of Polygon)
        Dim nearIslands As New List(Of Polygon)

        ' --- 2. Classify each polygon ---
        For Each poly In otherPolys

            ' 2A. Check if polygon is inside mainland → hole
            If GeometryEngine.Contains(mainland, poly) Then
                holes.Add(poly)
                Continue For
            End If

            ' 2B. Compute distance to mainland
            Dim dist As Double = GeometryEngine.Distance(mainland, poly)

            If dist < DistanceThresholdMeters Then
                ' Near-shore island → keep
                nearIslands.Add(poly)
            Else
                ' Remote island → discard (Paracel, Spratly, etc.)
            End If

        Next

        ' --- 3. Build final multipart polygon ---
        Dim builder As New Esri.ArcGISRuntime.Geometry.PolygonBuilder(sr)

        ' 3A. Mainland first
        For Each part In mainland.Parts
            builder.AddPart(part)
        Next

        ' 3B. Add holes (interior rings)
        For Each hole In holes
            For Each part In hole.Parts
                builder.AddPart(part)
            Next
        Next

        ' 3C. Add near-shore islands
        For Each island In nearIslands
            For Each part In island.Parts
                builder.AddPart(part)
            Next
        Next

        Return builder.ToGeometry()

    End Function

End Module
