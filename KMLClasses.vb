Imports System.Globalization
Imports System.Xml
Imports NetTopologySuite.Geometries

''' <summary>
''' Represents a complete KML file, containing a single <Document> element
''' and responsible for writing the XML header and root <kml> element.
''' 
''' The <see cref="KmlDocument"/> stored in the <c>Document</c> property
''' contains all styles, folders, placemarks, and extended data that make up
''' the body of the KML file. This class handles only the outer KML structure.
''' </summary>
''' <remarks>
''' The <c>Write</c> method outputs:
''' - The XML declaration
''' - The <kml> root element with the KML 2.2 namespace
''' - The serialized <Document> content
''' </remarks>
''' <param name="writer">
''' The <see cref="XmlWriter"/> used to generate the KML output.
''' </param>
Public Class KmlFile
    Public Property Document As KmlDocument
    Public Property Description As String

    Public Sub New()
        Document = New KmlDocument
    End Sub

    Public Sub Write(writer As XmlWriter)
        ' XML header
        writer.WriteStartDocument()

        ' <kml>
        writer.WriteStartElement("kml", "http://www.opengis.net/kml/2.2")

        Document.Write(writer)      ' write the document

        writer.WriteEndElement() ' </kml>
        writer.WriteEndDocument()
    End Sub
End Class
''' <summary>
''' Represents the root <Document> element inside a KML file.
''' 
''' A <Document> contains all user‑defined KML content, including:
''' - The document name and description
''' - Styles and StyleMaps
''' - Folders
''' - Placemarks
''' - Raw KML blocks inserted verbatim
''' 
''' This class provides collection properties for each of these elements and
''' serializes them in the correct KML‑compliant order when written.
''' </summary>
Public Class KmlDocument
    Public Property Name As String
    Public Property Description As String
    Public Property Open As Boolean = False
    Public Property StyleUrl As String
    Public Property Styles As List(Of KmlStyle)
    Public Property StyleMaps As List(Of KmlStyleMap)
    Public Property Folders As List(Of KmlFolder)
    Public Property Placemarks As List(Of KmlPlacemark)

    Private ReadOnly _rawKMLBlocks As List(Of String)

    Public Sub New()
        Styles = New List(Of KmlStyle)
        Folders = New List(Of KmlFolder)
        Placemarks = New List(Of KmlPlacemark)
        StyleMaps = New List(Of KmlStyleMap)
        _rawKMLBlocks = New List(Of String)
    End Sub
    Public Sub AddRawKML(xml As String)
        _rawKMLBlocks.Add(xml)
    End Sub

    Public Sub AddPlacemark(pm As KmlPlacemark)
        Placemarks.Add(pm)
    End Sub

    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("Document")

        If Not String.IsNullOrWhiteSpace(Name) Then
            writer.WriteElementString("name", Name)
        End If

        If Not String.IsNullOrWhiteSpace(Description) Then
            writer.WriteStartElement("description")
            writer.WriteRaw(Description)
            writer.WriteEndElement()
        End If

        writer.WriteElementString("open", If(Open, "1", "0"))

        For Each block In _rawKMLBlocks
            writer.WriteRaw(block)
        Next

        If Not String.IsNullOrWhiteSpace(StyleUrl) Then
            writer.WriteElementString("styleUrl", StyleUrl)
        End If

        ' StyleMaps must be written before Styles, because StyleMaps reference Styles by URL
        For Each sm In StyleMaps
            sm.Write(writer)
        Next

        ' Styles must be written before Folders and Placemarks, because they may reference Styles by URL
        For Each st In Styles
            st.Write(writer)
        Next

        ' Folders must be written before Placemarks, because Placemarks may be nested inside Folders
        For Each f In Folders
            f.Write(writer)
        Next

        'Placemarks at the Document level (not inside any Folder)
        For Each pm In Placemarks
            pm.Write(writer)
        Next

        writer.WriteEndElement() ' </Document>
    End Sub
End Class
''' <summary>
''' Represents a KML <Folder> element, used to group placemarks and nested folders
''' within a document. Folders provide hierarchical organization in Google Earth,
''' allowing related geographic features to be grouped together.
''' 
''' A folder may contain:
''' - A name and description
''' - A style reference applied to all child elements
''' - Nested folders
''' - Placemarks
''' 
''' The <see cref="Write"/> method outputs the folder and all of its contents
''' in proper KML structure and order.
''' </summary>
Public Class KmlFolder
    Public Property Name As String
    Public Property Description As String
    Public Property Open As Boolean = False     ' <open>0</open> by default
    Public Property StyleUrl As String
    Public Property Placemarks As List(Of KmlPlacemark)
    Public Property Folders As List(Of KmlFolder)

    Public Sub New()
        Placemarks = New List(Of KmlPlacemark)
        Folders = New List(Of KmlFolder)
    End Sub

    ' ------------------------------------------------------------
    '   Folder
    ' ------------------------------------------------------------
    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("Folder")

        If Not String.IsNullOrWhiteSpace(Name) Then
            writer.WriteElementString("name", Name)
        End If

        If Not String.IsNullOrWhiteSpace(Description) Then
            writer.WriteStartElement("description")
            writer.WriteRaw(Description)
            writer.WriteEndElement()
        End If

        ' <open>0</open> or <open>1</open>
        writer.WriteElementString("open", If(Open, "1", "0"))

        If Not String.IsNullOrWhiteSpace(StyleUrl) Then
            writer.WriteElementString("styleUrl", StyleUrl)
        End If

        ' Write nested folders
        Dim f As KmlFolder
        For Each f In Folders
            f.Write(writer)
        Next

        ' Write placemarks
        Dim pm As KmlPlacemark
        For Each pm In Placemarks
            pm.Write(writer)
        Next

        writer.WriteEndElement() ' Folder
    End Sub
End Class


' ------------------------------------------------------------
'   Placemark
' ------------------------------------------------------------
''' <summary>
''' Represents a KML <Placemark> element, which defines a geographic feature
''' such as a point, line, polygon, or multi‑geometry object. A placemark may
''' also contain descriptive text, styling information, and arbitrary metadata
''' stored in an <ExtendedData> block.
''' 
''' This class supports:
''' - Multiple geometry types (Point, LineString, Polygon, MultiPolygon, etc.)
''' - Optional name and description (HTML allowed)
''' - A style reference via <styleUrl>
''' - Precision‑controlled coordinate formatting
''' - Arbitrary key/value metadata written as <ExtendedData><Data> elements
''' 
''' The <see cref="Write"/> method serializes the placemark in full KML‑compliant
''' structure, including automatic wrapping of multiple geometries inside a
''' <MultiGeometry> container when required.
''' </summary>
Public Class KmlPlacemark

    ''' <summary>
    ''' The human‑readable name of the placemark, written as a <name> element.
    ''' </summary>
    Public Property Name As String

    ''' <summary>
    ''' Optional descriptive text for the placemark, written as a <description> element.
    ''' Raw XML/HTML is allowed and written without escaping.
    ''' </summary>
    Public Property Description As String

    ''' <summary>
    ''' Optional style reference applied to the placemark, written as a <styleUrl> element.
    ''' </summary>
    Public Property StyleUrl As String

    ''' <summary>
    ''' A list of geometry objects associated with the placemark. Supports multiple
    ''' geometries; if more than one is present, they are wrapped in a <MultiGeometry>.
    ''' </summary>
    Public Property Geometry As List(Of Geometry)

    ''' <summary>
    ''' Controls the number of decimal places used when writing coordinate values.
    ''' Default is 3 (approx. 100m precision).
    ''' </summary>
    Public Property CoordinateDigits As Integer = 3

    ''' <summary>
    ''' Arbitrary metadata stored in the placemark's <ExtendedData> block.
    ''' Keys become Data/@name attributes; values become <value> elements.
    ''' </summary>
    Public Property ExtendedData As Dictionary(Of String, String)

    ''' <summary>
    ''' Initializes a new instance of the <see cref="KmlPlacemark"/> class,
    ''' creating empty geometry and metadata collections.
    ''' </summary>
    Public Sub New()
        Geometry = New List(Of Geometry)
        ExtendedData = New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
    End Sub

    ''' <summary>
    ''' Writes the <Placemark> element and all associated content using the provided
    ''' <see cref="XmlWriter"/>. The output includes:
    ''' - Name and description
    ''' - styleUrl
    ''' - ExtendedData (if present)
    ''' - Geometry or MultiGeometry
    ''' 
    ''' Geometry is written using <see cref="KmlGeometryWriter.WriteGeometry"/>,
    ''' ensuring consistent coordinate formatting and KML‑compliant structure.
    ''' </summary>
    ''' <param name="writer">The <see cref="XmlWriter"/> used to output KML.</param>
    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("Placemark")

        If Not String.IsNullOrWhiteSpace(Name) Then
            writer.WriteElementString("name", Name)
        End If

        If Not String.IsNullOrWhiteSpace(Description) Then
            writer.WriteStartElement("description")
            writer.WriteRaw(Description)
            writer.WriteEndElement()
        End If

        If Not String.IsNullOrWhiteSpace(StyleUrl) Then
            writer.WriteElementString("styleUrl", StyleUrl)
        End If

        If ExtendedData IsNot Nothing AndAlso ExtendedData.Count > 0 Then
            writer.WriteWhitespace(vbLf)
            writer.WriteStartElement("ExtendedData")
            writer.WriteWhitespace(vbLf)

            For Each kvp In ExtendedData
                writer.WriteStartElement("Data")
                writer.WriteAttributeString("name", KMLescape(kvp.Key))

                writer.WriteStartElement("value")
                writer.WriteString(KMLescape(kvp.Value))
                writer.WriteEndElement()

                writer.WriteEndElement()
                writer.WriteWhitespace(vbLf)
            Next

            writer.WriteEndElement()
        End If

        If Geometry IsNot Nothing AndAlso Geometry.Count > 0 Then
            If Geometry.Count = 1 Then
                KmlGeometryWriter.WriteGeometry(writer, Geometry(0), CoordinateDigits)
            Else
                writer.WriteWhitespace(vbLf)
                writer.WriteStartElement("MultiGeometry")
                For Each geom In Geometry
                    KmlGeometryWriter.WriteGeometry(writer, geom, CoordinateDigits)
                Next
                writer.WriteEndElement()
            End If
        End If

        writer.WriteEndElement()
    End Sub

End Class


''' <summary>
''' Represents a KML <Style> element, defining visual appearance for placemarks,
''' polygons, and linework. A style may include a <LineStyle> block, a <PolyStyle>
''' block, or both, depending on which properties are populated.
''' 
''' Supported features include:
''' - Line color and width (for points, lines, and polygon outlines)
''' - Polygon fill color
''' - Fill and outline enable/disable flags
''' 
''' Colors must be provided in KML AABBGGRR format. The <see cref="Write"/> method
''' automatically omits empty style sections, producing only the elements required
''' by the populated properties.
''' </summary>
Public Class KmlStyle

    ''' <summary>
    ''' The unique identifier for the style, written as the "id" attribute
    ''' of the <Style> element. Referenced by placemarks via <styleUrl>.
    ''' </summary>
    Public Property Id As String

    ''' <summary>
    ''' The line color in KML AABBGGRR format. If omitted, no <color> element
    ''' is written inside <LineStyle>.
    ''' </summary>
    Public Property LineColor As String

    ''' <summary>
    ''' The width of the line in pixels. If zero or negative, the width element
    ''' is omitted from the <LineStyle> block.
    ''' </summary>
    Public Property LineWidth As Double

    ''' <summary>
    ''' The polygon fill color in KML AABBGGRR format. If omitted, no <color>
    ''' element is written inside <PolyStyle>.
    ''' </summary>
    Public Property PolyColor As String

    ''' <summary>
    ''' Indicates whether the polygon interior should be filled. Written as
    ''' <fill>1</fill> or <fill>0</fill> inside <PolyStyle>.
    ''' </summary>
    Public Property Fill As Boolean = True

    ''' <summary>
    ''' Indicates whether the polygon outline should be drawn. Written as
    ''' <outline>1</outline> or <outline>0</outline> inside <PolyStyle>.
    ''' </summary>
    Public Property Outline As Boolean = True

    ''' <summary>
    ''' Writes the <Style> element and any applicable <LineStyle> and <PolyStyle>
    ''' blocks using the provided <see cref="XmlWriter"/>. Only populated style
    ''' components are emitted, ensuring compact and valid KML output.
    ''' </summary>
    ''' <param name="writer">The <see cref="XmlWriter"/> used to output KML.</param>
    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("Style")
        writer.WriteAttributeString("id", Id)

        ' LineStyle
        If Not String.IsNullOrEmpty(LineColor) OrElse LineWidth > 0 Then
            writer.WriteStartElement("LineStyle")
            If Not String.IsNullOrEmpty(LineColor) Then writer.WriteElementString("color", LineColor)
            If LineWidth > 0 Then writer.WriteElementString("width", LineWidth.ToString(CultureInfo.InvariantCulture))
            writer.WriteEndElement()
        End If

        ' PolyStyle
        If Not String.IsNullOrEmpty(PolyColor) OrElse Not Fill OrElse Not Outline Then
            writer.WriteStartElement("PolyStyle")
            If Not String.IsNullOrEmpty(PolyColor) Then writer.WriteElementString("color", PolyColor)
            writer.WriteElementString("fill", If(Fill, "1", "0"))
            writer.WriteElementString("outline", If(Outline, "1", "0"))
            writer.WriteEndElement()
        End If

        writer.WriteEndElement()
    End Sub

End Class

''' <summary>
''' Represents a KML <StyleMap> element, which defines a pair of styles used for
''' interactive rendering in Google Earth. A StyleMap associates:
''' - A "normal" style (default appearance)
''' - A "highlight" style (used when the user hovers over the feature)
''' 
''' StyleMaps are typically used for polygons, boundaries, and other features
''' where visual feedback on mouse‑over improves usability. Each StyleMap is
''' identified by an "id" attribute and referenced by placemarks via <styleUrl>.
''' </summary>
Public Class KmlStyleMap

    ''' <summary>
    ''' The unique identifier for the StyleMap, written as the "id" attribute
    ''' of the <StyleMap> element.
    ''' </summary>
    Public Property Id As String

    ''' <summary>
    ''' The URL of the normal style, written inside the <Pair> element with
    ''' <key>normal</key>. Typically references a <Style> by its ID.
    ''' </summary>
    Public Property NormalStyleUrl As String

    ''' <summary>
    ''' The URL of the highlight style, written inside the <Pair> element with
    ''' <key>highlight</key>. Used when the user hovers over the feature.
    ''' </summary>
    Public Property HighlightStyleUrl As String

    ''' <summary>
    ''' Writes the <StyleMap> element and its two <Pair> entries (normal and
    ''' highlight) using the provided <see cref="XmlWriter"/>. The output is
    ''' fully KML‑compliant and suitable for use with interactive features in
    ''' Google Earth.
    ''' </summary>
    ''' <param name="writer">The <see cref="XmlWriter"/> used to output KML.</param>
    Public Sub Write(writer As XmlWriter)
        writer.WriteStartElement("StyleMap")
        writer.WriteAttributeString("id", Id)

        ' normal pair
        writer.WriteStartElement("Pair")
        writer.WriteElementString("key", "normal")
        writer.WriteElementString("styleUrl", NormalStyleUrl)
        writer.WriteEndElement()

        ' highlight pair
        writer.WriteStartElement("Pair")
        writer.WriteElementString("key", "highlight")
        writer.WriteElementString("styleUrl", HighlightStyleUrl)
        writer.WriteEndElement()

        writer.WriteEndElement() ' </StyleMap>
    End Sub

End Class


' ------------------------------------------------------------
'   Geometry Writer
' ------------------------------------------------------------
Public Module KmlGeometryWriter

    Public Sub WriteGeometry(writer As XmlWriter, geom As Geometry, digits As Integer)
        If geom Is Nothing OrElse geom.IsEmpty Then Return

        Select Case geom.OgcGeometryType
            Case OgcGeometryType.Point
                WritePoint(writer, DirectCast(geom, Point), digits)

            Case OgcGeometryType.LineString
                WriteLineString(writer, DirectCast(geom, LineString), digits)

            Case OgcGeometryType.Polygon
                WritePolygon(writer, DirectCast(geom, Polygon), digits)

            Case OgcGeometryType.MultiPolygon
                WriteMultiPolygon(writer, DirectCast(geom, MultiPolygon), digits)

            Case OgcGeometryType.MultiLineString
                WriteMultiLineString(writer, DirectCast(geom, MultiLineString), digits)

            Case Else
                ' fallback: write WKT as description
                writer.WriteElementString("description", geom.AsText())
        End Select
    End Sub


    ' ---------------- POINT ----------------
    Private Sub WritePoint(writer As XmlWriter, pt As Point, digits As Integer)
        writer.WriteWhitespace(vbLf)
        writer.WriteStartElement("Point")
        Dim lon As String = FormatCoord(pt.Coordinate.X, digits)
        Dim lat As String = FormatCoord(pt.Coordinate.Y, digits)
        Dim coordText As String = lon & "," & lat
        writer.WriteElementString("coordinates", coordText)
        writer.WriteEndElement()
    End Sub


    ' ---------------- LINESTRING ----------------
    Private Sub WriteLineString(writer As XmlWriter, ls As LineString, digits As Integer)
        writer.WriteWhitespace(vbLf)
        writer.WriteStartElement("LineString")
        writer.WriteElementString("coordinates", FormatCoords(ls.Coordinates, digits))
        writer.WriteEndElement()
    End Sub


    ' ---------------- MULTILINESTRING ----------------
    Private Sub WriteMultiLineString(writer As XmlWriter, mls As MultiLineString, digits As Integer)
        writer.WriteWhitespace(vbLf)
        writer.WriteStartElement("MultiGeometry")

        For i = 0 To mls.NumGeometries - 1
            WriteLineString(writer, DirectCast(mls.GetGeometryN(i), LineString), digits)
        Next

        writer.WriteEndElement()
    End Sub


    ' ---------------- POLYGON ----------------
    Private Sub WritePolygon(writer As XmlWriter, poly As Polygon, digits As Integer)
        writer.WriteWhitespace(vbLf)
        writer.WriteStartElement("Polygon")

        ' Outer ring
        writer.WriteStartElement("outerBoundaryIs")
        writer.WriteStartElement("LinearRing")
        writer.WriteElementString("coordinates", FormatCoords(poly.ExteriorRing.Coordinates, digits))
        writer.WriteEndElement() ' LinearRing
        writer.WriteEndElement() ' outerBoundaryIs

        ' Holes
        For i = 0 To poly.NumInteriorRings - 1
            Dim hole = poly.GetInteriorRingN(i)
            writer.WriteStartElement("innerBoundaryIs")
            writer.WriteStartElement("LinearRing")
            writer.WriteElementString("coordinates", FormatCoords(hole.Coordinates, digits))
            writer.WriteEndElement() ' LinearRing
            writer.WriteEndElement() ' innerBoundaryIs
        Next

        writer.WriteEndElement() ' Polygon
    End Sub


    ' ---------------- MULTIPOLYGON ----------------
    Private Sub WriteMultiPolygon(writer As XmlWriter, mp As MultiPolygon, digits As Integer)
        writer.WriteWhitespace(vbLf)
        writer.WriteStartElement("MultiGeometry")

        For i = 0 To mp.NumGeometries - 1
            WritePolygon(writer, DirectCast(mp.GetGeometryN(i), Polygon), digits)
        Next

        writer.WriteEndElement()
    End Sub


    ' ------------------------------------------------------------
    '   Coordinate Formatting
    ' ------------------------------------------------------------
    Public Function FormatCoord(value As Double, digits As Integer) As String
        ' Ensures consistent rounding and avoids scientific notation
        Return Math.Round(value, digits).ToString("F" & digits, CultureInfo.InvariantCulture)
    End Function


    Private Function FormatCoords(coords As Coordinate(), digits As Integer) As String
        Dim sb As New System.Text.StringBuilder()

        For i = 0 To coords.Length - 1
            Dim c = coords(i)
            sb.Append($"{FormatCoord(c.X, digits)},{FormatCoord(c.Y, digits)} ")
        Next

        Return sb.ToString().Trim()
    End Function

End Module
