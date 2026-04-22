Imports System.Xml
Imports NetTopologySuite.Geometries

Public Class KmlDocument
    Private ReadOnly _placemarks As New List(Of KmlPlacemark)

    Private ReadOnly _rawKMLBlocks As New List(Of String)
    Public Sub AddRawKML(xml As String)
        _rawKMLBlocks.Add(xml)
    End Sub

    Public Sub AddPlacemark(pm As KmlPlacemark)
        _placemarks.Add(pm)
    End Sub

    Public Sub WriteToFile(path As String)
        Dim settings As New XmlWriterSettings With {
        .Indent = True,
        .IndentChars = "  ",
        .NewLineOnAttributes = False,
        .Encoding = System.Text.Encoding.UTF8
    }

        Using writer = XmlWriter.Create(path, settings)
            writer.WriteStartDocument()
            writer.WriteStartElement("kml", "http://www.opengis.net/kml/2.2")
            writer.WriteStartElement("Document")
            ' Write raw KML blocks
            For Each raw In _rawKMLBlocks
                writer.WriteRaw(raw)
            Next
            For Each pm In _placemarks
                pm.Write(writer)
            Next

            writer.WriteEndElement() ' Document
            writer.WriteEndElement() ' kml
            writer.WriteEndDocument()
        End Using
    End Sub
End Class


' ------------------------------------------------------------
'   Placemark
' ------------------------------------------------------------
Public Class KmlPlacemark
        Public Property Name As String
        Public Property Description As String
        Public Property StyleUrl As String
        Public Property Geometry As Geometry
    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("Placemark")

        If Not String.IsNullOrWhiteSpace(Name) Then
            writer.WriteElementString("name", Name)
        End If

        If Not String.IsNullOrWhiteSpace(Description) Then
            writer.WriteStartElement("description")
            writer.WriteRaw(Description)   ' <-- critical to allow HTML content in description
            writer.WriteEndElement()
        End If

        If Not String.IsNullOrWhiteSpace(StyleUrl) Then
            writer.WriteElementString("styleUrl", StyleUrl)
        End If

        KmlGeometryWriter.WriteGeometry(writer, Geometry)

        writer.WriteEndElement() ' Placemark
    End Sub
End Class


    ' ------------------------------------------------------------
    '   Geometry Writer
    ' ------------------------------------------------------------
    Public Module KmlGeometryWriter

        Public Sub WriteGeometry(writer As XmlWriter, geom As Geometry)
            If geom Is Nothing OrElse geom.IsEmpty Then Return

            Select Case geom.OgcGeometryType
                Case OgcGeometryType.Point
                    WritePoint(writer, DirectCast(geom, Point))

                Case OgcGeometryType.LineString
                    WriteLineString(writer, DirectCast(geom, LineString))

                Case OgcGeometryType.Polygon
                    WritePolygon(writer, DirectCast(geom, Polygon))

                Case OgcGeometryType.MultiPolygon
                    WriteMultiPolygon(writer, DirectCast(geom, MultiPolygon))

                Case OgcGeometryType.MultiLineString
                    WriteMultiLineString(writer, DirectCast(geom, MultiLineString))

                Case Else
                    ' fallback: write WKT as description
                    writer.WriteElementString("description", geom.AsText())
            End Select
        End Sub


        ' ---------------- POINT ----------------
        Private Sub WritePoint(writer As XmlWriter, pt As Point)
            writer.WriteStartElement("Point")
            writer.WriteElementString("coordinates", FormatCoord(pt.Coordinate))
            writer.WriteEndElement()
        End Sub


        ' ---------------- LINESTRING ----------------
        Private Sub WriteLineString(writer As XmlWriter, ls As LineString)
            writer.WriteStartElement("LineString")
            writer.WriteElementString("coordinates", FormatCoords(ls.Coordinates))
            writer.WriteEndElement()
        End Sub


        ' ---------------- MULTILINESTRING ----------------
        Private Sub WriteMultiLineString(writer As XmlWriter, mls As MultiLineString)
            writer.WriteStartElement("MultiGeometry")

            For i = 0 To mls.NumGeometries - 1
                WriteLineString(writer, DirectCast(mls.GetGeometryN(i), LineString))
            Next

            writer.WriteEndElement()
        End Sub


        ' ---------------- POLYGON ----------------
        Private Sub WritePolygon(writer As XmlWriter, poly As Polygon)
            writer.WriteStartElement("Polygon")

            ' Outer ring
            writer.WriteStartElement("outerBoundaryIs")
            writer.WriteStartElement("LinearRing")
            writer.WriteElementString("coordinates", FormatCoords(poly.ExteriorRing.Coordinates))
            writer.WriteEndElement() ' LinearRing
            writer.WriteEndElement() ' outerBoundaryIs

            ' Holes
            For i = 0 To poly.NumInteriorRings - 1
                Dim hole = poly.GetInteriorRingN(i)
                writer.WriteStartElement("innerBoundaryIs")
                writer.WriteStartElement("LinearRing")
                writer.WriteElementString("coordinates", FormatCoords(hole.Coordinates))
                writer.WriteEndElement() ' LinearRing
                writer.WriteEndElement() ' innerBoundaryIs
            Next

            writer.WriteEndElement() ' Polygon
        End Sub


        ' ---------------- MULTIPOLYGON ----------------
        Private Sub WriteMultiPolygon(writer As XmlWriter, mp As MultiPolygon)
            writer.WriteStartElement("MultiGeometry")

            For i = 0 To mp.NumGeometries - 1
                WritePolygon(writer, DirectCast(mp.GetGeometryN(i), Polygon))
            Next

            writer.WriteEndElement()
        End Sub


        ' ------------------------------------------------------------
        '   Coordinate Formatting
        ' ------------------------------------------------------------
        Private Function FormatCoord(c As Coordinate) As String
            Return $"{c.X:0.########},{c.Y:0.########}"
        End Function

        Private Function FormatCoords(coords As Coordinate()) As String
            Dim sb As New System.Text.StringBuilder()

            For i = 0 To coords.Length - 1
                Dim c = coords(i)
                sb.Append($"{c.X:0.########},{c.Y:0.########} ")
            Next

            Return sb.ToString().Trim()
        End Function

    End Module
