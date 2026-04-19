Imports System.Text.RegularExpressions
Imports Esri.ArcGISRuntime.Geometry
Imports Microsoft.Data.Sqlite

Module dxccSources
    Public Class DxccSource
        Public Property DXCCnum As Integer
        Public Property name As String
        Public Property source As String
        Public Property rule As String
        Public Property tolerance_m As Integer
        Public Property bbox As String
        Public Property notes As String
    End Class
    Public Async Sub BuildAndSaveAllDxccGeometries()
        Dim count As Integer = 0, success As Integer = 0
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            connect.CreateFunction(name:="hashV", function:=Function(values As Object()) HashVariadic(values))

            Dim sql = "SELECT DXCCnum, Entity FROM DXCC WHERE (`hash` <> hashV(`rule`,`bbox`,`tolerance_m`,`geometry`) OR geometry is NULL) AND `Deleted`=0 AND `DXCCnum`<>999 ORDER BY `Entity`"

            Using cmd As New SqliteCommand(sql, connect)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        count += 1
                        Debug.WriteLine($"Getting geometry for {r("Entity")}")
                        Dim id = r("DXCCnum")
                        Try
                            Dim result = Await BuildAndSaveDxccGeometry(id, connect)
                            If result Then
                                Debug.WriteLine($"Built geometry for DXCC ID: {id}")
                                success += 1
                            End If
                        Catch ex As Exception
                            Debug.WriteLine($"Error building geometry for DXCC ID {id}: {ex.Message}")
                        End Try
                    End While
                End Using
            End Using
        End Using
        Debug.WriteLine($"Done. Processed {count} DXCC entries. Successfully built {success} geometries.")
    End Sub
    Private Function HashVariadic(values As Object()) As String
        ' SQLite may pass Nothing or an empty array
        If values Is Nothing OrElse values.Length = 0 Then
            Return HashText("")   ' always return a valid hash
        End If

        Dim parts As New List(Of String)(values.Length)

        For Each v In values
            If v Is Nothing OrElse v Is DBNull.Value Then
                parts.Add("")   ' treat NULL as empty string
            Else
                ' ToString() can throw on some types → guard it
                Try
                    parts.Add(v.ToString())
                Catch
                    parts.Add("")   ' fallback
                End Try
            End If
        Next

        Dim combined = String.Join("|", parts)
        Return HashText(combined)
    End Function


    Public Async Function BuildAndSaveDxccGeometry(dxccId As Integer, connect As SqliteConnection) As Task(Of Boolean)
        Try
            Dim src = LoadDxccSource(dxccId, connect)
            Dim geom = Await BuildDXCCGeometryAsync(src)
            If geom Is Nothing OrElse geom.IsEmpty Then
                Debug.WriteLine($"No geometry built for DXCC ID {dxccId}. Skipping save.")
                Return False
            End If
            SaveDxccGeometry(dxccId, geom, connect)
        Catch ex As Exception
            Debug.WriteLine($"Error building geometry for DXCC ID {dxccId}: {ex.Message}")
            Return False
        End Try
        Return True
    End Function

    Private Function LoadDxccSource(dxccId As Integer, conn As SqliteConnection) As DxccSource

        Dim sql = "SELECT DXCCnum, Entity, source, rule, bbox, tolerance_m, notes FROM DXCC WHERE DXCCnum=@id"

        Using cmd As New SqliteCommand(sql, conn)
            cmd.Parameters.AddWithValue("@id", dxccId)

            Using r = cmd.ExecuteReader()
                If Not r.Read() Then Throw New Exception("DXCC not found: " & dxccId)

                Return New DxccSource With {
                    .DXCCnum = r("DXCCnum"),
                    .name = SafeStr(r("Entity")),
                    .source = SafeStr(r("source")),
                    .rule = SafeStr(r("rule")),
                    .tolerance_m = If(r.IsDBNull(r.GetOrdinal("tolerance_m")), 0, r.GetInt32(r.GetOrdinal("tolerance_m"))),
                    .bbox = SafeStr(r("bbox")),
                    .notes = SafeStr(r("notes"))
                }
            End Using
        End Using
    End Function

    Private Async Function BuildFromOSM(osmName As String) As Task(Of Geometry)
        ' simple parser to start. ways start with W.
        Dim osmId = osmName
        Dim osmType As String = "relation"  ' default to relation, which is the most common type for countries. This means that if a user just enters "France" it will resolve to the relation, which is what we want in 99% of cases. If they want a specific way, they can enter "W.waysid"
        If osmName.StartsWith("W", StringComparison.OrdinalIgnoreCase) Then
            osmType = "way"
            osmId = osmName.Substring(1)
        End If

        Dim g = Await OsmLand.ResolveOSM(osmId, osmType)

        If g Is Nothing Then
            Throw New Exception("OSM geometry not found: " & osmName)
        End If

        'Return SimplifyGeometry(g, tol)
        Return g
    End Function
    Private Function BuildFromFake(fakeName As String) As Geometry
        ' Some geometries just dont exist anywhere, especially small rocks, reefs and atolls
        ' This makes fake geometry for these
        Dim p As MapPoint, g As Geometry

        Select Case fakeName
            Case "Willis Is"
                p = New MapPoint(150.2, -16.29, SpatialReferences.Wgs84)    ' center of island
                g = GeometryEngine.Buffer(p, MetersToDegrees(p, 200))       ' its a small island

            Case "Mellish Reef"
                p = New MapPoint(155.8644, -17.4126, SpatialReferences.Wgs84)    ' center of island
                g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

            Case "St Peter & St Paul Rocks"
                p = New MapPoint(-29.335, 0.9158, SpatialReferences.Wgs84)    ' center of island
                g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

            Case "Malpelo Is"
                p = New MapPoint(-81.5833, 3.9833, SpatialReferences.Wgs84)    ' center of island
                g = GeometryEngine.Buffer(p, MetersToDegrees(p, 1000))       ' its a small island

            Case Else
                Throw New Exception($"FAKE parameter {fakeName} not recognised")
        End Select

        Return g
    End Function

    Public Function MetersToDegrees(g As Geometry, meters As Double) As Double
        ' Use the latitude of the geometry's center
        Dim env = g.Extent
        Dim lat = (env.YMin + env.YMax) / 2.0
        Dim metersPerDegree = 111320 * Math.Cos(lat * Math.PI / 180)
        Return meters / metersPerDegree
    End Function
    Private Sub SaveDxccGeometry(dxccId As Integer, geom As Geometry, conn As SqliteConnection)
        If geom Is Nothing OrElse geom.IsEmpty Then
            Debug.WriteLine($"Geometry for DXCC {dxccId} is empty. Not saved")
            Return      ' don't save empty geometry
        End If
        Dim GeoJson As String = GeometryToGeoJson(geom)   ' GeoJSON
        Try
            conn.CreateFunction(name:="hashV", function:=Function(values As Object()) HashVariadic(values))
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            ' ignore as probably already registered
        End Try
        ' Update geometry and hash in one statement
        Dim sql = "UPDATE DXCC SET geometry = $g, hash = hashV(rule, bbox, tolerance_m, $g) WHERE DXCCnum = $id"

        Using cmd As New SqliteCommand(sql, conn)
            cmd.Parameters.AddWithValue("$g", GeoJson)
            cmd.Parameters.AddWithValue("$id", dxccId)
            cmd.ExecuteNonQuery()
        End Using
    End Sub
    Public Async Function BuildDXCCGeometryAsync(src As DxccSource) As Task(Of Geometry)

        ' STEP 1 — Resolve base geometry
        Dim baseGeom As Geometry = Await ResolveRuleExpressionAsync(src)
        ' STEP 2 — Aggressive pre-simplification (critical)
        Dim preTol = 0.00002 ' ~2m
        baseGeom = GeometryEngine.Generalize(baseGeom, preTol, True)
        Debug.WriteLine("BASE JSON:")
        Debug.WriteLine(baseGeom.ToJson())
        If baseGeom Is Nothing Then
            Debug.WriteLine($"Failed to resolve base geometry for DXCC '{src.name}' using rule: {src.rule}")
            Return Nothing
        End If
        baseGeom = GeometryEngine.Simplify(baseGeom)

        ' STEP 2 — Apply bbox/poly if present
        If Not String.IsNullOrWhiteSpace(src.bbox) Then

            Dim clipGeom As Geometry = ParseBBoxOrPoly(src.bbox)
            If clipGeom Is Nothing Then Return Nothing

            ' Normalize clipper BEFORE clipping
            clipGeom = GeometryEngine.Simplify(clipGeom)

            ' Clip
            baseGeom = GeometryEngine.Intersection(baseGeom, clipGeom)

            ' CLEANUP STEP — critical for political boundaries
            baseGeom = GeometryEngine.Buffer(baseGeom, 0)

            Debug.WriteLine("CLIPPED JSON:")
            Debug.WriteLine(baseGeom.ToJson())
        End If

        If baseGeom Is Nothing OrElse baseGeom.IsEmpty Then Return Nothing

        ' STEP 3 — Generalize AFTER clipping
        If src.tolerance_m > 0 Then
            Dim degTol = src.tolerance_m / 110540
            baseGeom = GeometryEngine.Generalize(baseGeom, degTol, True)
        End If

        ' STEP 4 — Round AFTER generalization
        baseGeom = RoundGeometry(baseGeom, 6)

        Return baseGeom
    End Function



    Public Function SimplifyGeometry(geom As Geometry, tolerance As Double) As Geometry
        ' Null or trivial geometry → return as-is
        If geom Is Nothing Then Return Nothing
        If TypeOf geom Is MapPoint Then Return geom
        If TypeOf geom Is Polyline AndAlso DirectCast(geom, Polyline).Parts.Count = 0 Then Return geom
        If TypeOf geom Is Polygon AndAlso DirectCast(geom, Polygon).Parts.Count = 0 Then Return geom

        ' ArcGIS Runtime 200.x Simplify() is too aggressive — do NOT use it
        ' Instead use Generalize() with preserveTopology:=True
        Dim simplified As Geometry = GeometryEngine.Generalize(geom, tolerance, True)

        ' Generalize can introduce floating-point drift → round coordinates
        simplified = RoundGeometry(simplified, 6)

        Return simplified
    End Function
    Public Function RoundGeometry(geom As Geometry, decimals As Integer) As Geometry
        If geom Is Nothing Then Return Nothing

        If TypeOf geom Is MapPoint Then
            Dim p = DirectCast(geom, MapPoint)
            Return New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals))
        End If

        If TypeOf geom Is Polygon Then
            Dim poly = DirectCast(geom, Polygon)
            Dim builder As New PolygonBuilder(poly.SpatialReference)

            For Each part In poly.Parts
                Dim newPart As New List(Of MapPoint)
                For Each p In part.Points
                    newPart.Add(New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals)))
                Next
                builder.AddPart(newPart)
            Next

            Return builder.ToGeometry()
        End If

        If TypeOf geom Is Polyline Then
            Dim line = DirectCast(geom, Polyline)
            Dim builder As New PolylineBuilder(line.SpatialReference)

            For Each part In line.Parts
                Dim newPart As New List(Of MapPoint)
                For Each p In part.Points
                    newPart.Add(New MapPoint(Math.Round(p.X, decimals), Math.Round(p.Y, decimals)))
                Next
                builder.AddPart(newPart)
            Next

            Return builder.ToGeometry()
        End If

        Return geom
    End Function


    ' ---------------------------------------------------------
    '  RULE PARSER (UNION SUPPORT) — NOW TAKES DxccSource
    ' ---------------------------------------------------------
    Private Async Function ResolveRuleExpressionAsync(src As DxccSource) As Task(Of Geometry)

        Dim parts = Tokenize(src.rule)

        If parts.Count = 0 Then
            Return Nothing
        End If

        Dim current As Geometry = Nothing

        For Each part In parts
            Dim geom = Await ResolveSingleSourceAsync(src.source, part.Token)
            If geom Is Nothing Then Continue For

            If current Is Nothing Then
                ' First geometry always initializes the accumulator
                current = geom
            Else
                If part.Op = "+"c Then
                    current = GeometryEngine.Union(current, geom)
                ElseIf part.Op = "-"c Then
                    current = GeometryEngine.Difference(current, geom)
                End If
            End If
        Next

        Return current
    End Function

    Private Function Tokenize(expr As String) As List(Of (Op As Char, Token As String))
        Dim result As New List(Of (Char, String))
        Dim current As String = ""
        Dim currentOp As Char = "+"c   ' default for first token

        For Each ch In expr
            If ch = "+"c OrElse ch = "-"c Then
                If current.Trim().Length > 0 Then
                    result.Add((currentOp, current.Trim()))
                End If
                current = ""
                currentOp = ch
            Else
                current &= ch
            End If
        Next

        If current.Trim().Length > 0 Then
            result.Add((currentOp, current.Trim()))
        End If

        Return result
    End Function

    ' ---------------------------------------------------------
    '  SINGLE SOURCE RESOLVER (ASYNC)
    ' ---------------------------------------------------------
    Private Async Function ResolveSingleSourceAsync(source As String,
                                                    token As String) As Task(Of Geometry)

        Select Case source.ToUpperInvariant()

            Case "NE"
                Return Await NaturalEarth.Lookup(token)

            Case "OSM"
                Return Await BuildFromOSM(token)

            Case "FAKE"
                Return BuildFromFake(token)

            Case Else
                Throw New Exception($"Unsupported resolver type '{source}' for token '{token}'.")
        End Select
    End Function
End Module
