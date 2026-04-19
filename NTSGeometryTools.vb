Imports NetTopologySuite.Densify
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Geometries.Utilities
Imports NetTopologySuite.IO
Imports NetTopologySuite.Operation.Densify
Imports NetTopologySuite.Operation.Valid

Public Module NtsGeometryTools

    Private ReadOnly _reader As New GeoJsonReader()
    Private ReadOnly _writer As New GeoJsonWriter()

    ' ------------------------------------------------------------
    ' Normalize a single longitude into [-180, 180]
    ' ------------------------------------------------------------
    Private Function NormalizeLongitude(x As Double) As Double
        Dim v = x
        While v > 180
            v -= 360
        End While
        While v < -180
            v += 360
        End While
        Return v
    End Function

    ' ------------------------------------------------------------
    ' Recursively normalize coordinates of any geometry
    ' ------------------------------------------------------------
    Private Function NormalizeCoords(g As Geometry) As Geometry
        If g Is Nothing Then Return Nothing

        Dim f = g.Factory

        Select Case True

            Case TypeOf g Is Point
                Dim p = DirectCast(g, Point)
                Return f.CreatePoint(New Coordinate(NormalizeLongitude(p.X), p.Y))

            Case TypeOf g Is LineString
                Dim ls = DirectCast(g, LineString)
                Dim coords = ls.Coordinates _
                    .Select(Function(c) New Coordinate(NormalizeLongitude(c.X), c.Y)) _
                    .ToArray()
                Return f.CreateLineString(coords)

            Case TypeOf g Is LinearRing
                Dim lr = DirectCast(g, LinearRing)
                Dim coords = lr.Coordinates _
                    .Select(Function(c) New Coordinate(NormalizeLongitude(c.X), c.Y)) _
                    .ToArray()
                Return f.CreateLinearRing(coords)

            Case TypeOf g Is Polygon
                Dim poly = DirectCast(g, Polygon)

                ' Exterior ring
                Dim shell = DirectCast(NormalizeCoords(poly.ExteriorRing), LinearRing)

                ' Interior rings (holes)
                Dim holeCount = poly.NumInteriorRings
                Dim holes(holeCount - 1) As LinearRing

                For i = 0 To holeCount - 1
                    holes(i) = DirectCast(NormalizeCoords(poly.GetInteriorRingN(i)), LinearRing)
                Next

                Return f.CreatePolygon(shell, holes)

            Case TypeOf g Is MultiPolygon
                Dim mp = DirectCast(g, MultiPolygon)
                Dim polys(mp.NumGeometries - 1) As Polygon

                For i = 0 To mp.NumGeometries - 1
                    polys(i) = DirectCast(NormalizeCoords(mp.GetGeometryN(i)), Polygon)
                Next

                Return f.CreateMultiPolygon(polys)

            Case Else
                ' Generic geometry collection fallback
                Dim parts(g.NumGeometries - 1) As Geometry
                For i = 0 To g.NumGeometries - 1
                    parts(i) = NormalizeCoords(g.GetGeometryN(i))
                Next
                Return f.BuildGeometry(parts)

        End Select
    End Function

    ' ------------------------------------------------------------
    ' Full pipeline: MakeValid → Normalize → Densify → GeoJSON
    ' ------------------------------------------------------------
    Public Function NormalizeAndDensifyGeoJson(
        geoJson As String,
        maxSegmentLength As Double
    ) As String

        If String.IsNullOrWhiteSpace(geoJson) Then Return geoJson

        Dim g = _reader.Read(Of Geometry)(geoJson)
        If g Is Nothing OrElse g.IsEmpty Then Return geoJson

        ' Fix topology
        If Not g.IsValid Then
            g = GeometryFixer.Fix(g)
        End If

        ' Normalize longitudes
        g = NormalizeCoords(g)

        ' Densify
        If maxSegmentLength > 0 Then
            g = Densifier.Densify(g, maxSegmentLength)
        End If

        Return _writer.Write(g)
    End Function

    ' ------------------------------------------------------------
    ' Normalize only
    ' ------------------------------------------------------------
    Public Function NormalizeGeoJson(geoJson As String) As String
        If String.IsNullOrWhiteSpace(geoJson) Then Return geoJson

        Dim g = _reader.Read(Of Geometry)(geoJson)
        If g Is Nothing OrElse g.IsEmpty Then Return geoJson

        If Not g.IsValid Then
            g = GeometryFixer.Fix(g)
        End If

        g = NormalizeCoords(g)

        Return _writer.Write(g)
    End Function

    ' ------------------------------------------------------------
    ' Densify only
    ' ------------------------------------------------------------
    Public Function DensifyGeoJson(geoJson As String, maxSegmentLength As Double) As String
        If String.IsNullOrWhiteSpace(geoJson) OrElse maxSegmentLength <= 0 Then Return geoJson

        Dim g = _reader.Read(Of Geometry)(geoJson)
        If g Is Nothing OrElse g.IsEmpty Then Return geoJson

        If Not g.IsValid Then
            g = GeometryFixer.Fix(g)
        End If

        g = Densifier.Densify(g, maxSegmentLength)

        Return _writer.Write(g)
    End Function

End Module

