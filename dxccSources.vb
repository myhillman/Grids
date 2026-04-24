Imports System.Text.RegularExpressions
Imports Microsoft.Data.Sqlite
Imports NetTopologySuite.Simplify
Imports NetTopologySuite.Geometries
Imports NetTopologySuite.Operation.Valid
Imports NetTopologySuite.Operation.Buffer
Imports NetTopologySuite.Precision
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
        Dim ids As New Dictionary(Of String, Integer)

        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            EnsureHashVRegistered(connect)

            ' Make list of ids to process
            Dim sql = "SELECT DXCCnum, Entity FROM DXCC WHERE (`hash` <> hashV(`rule`,`bbox`,`tolerance_m`,`geometry`) OR geometry is NULL) AND `Deleted`=0 AND `DXCCnum`<>999 ORDER BY `Entity`"
            Using cmd As New SqliteCommand(sql, connect)
                Using r = cmd.ExecuteReader()
                    While r.Read()
                        ids.Add(r("Entity"), r("DXCCnum"))
                    End While
                End Using
            End Using
        End Using

        ' Process the list
        With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = ids.Count
            End With

            For Each id In ids
                count += 1
                Debug.WriteLine($"Getting geometry for {id.Key}")
                Try
                Dim result = Await BuildAndSaveDxccGeometry(id.Value)
                If result Then
                    Debug.WriteLine($"Built geometry for Entity: {id.Key}")
                    success += 1
                    End If
                Catch ex As Exception
                    Debug.WriteLine($"Error building geometry for Entity {id.Key}: {ex.Message}")
                End Try
            Next
            Debug.WriteLine($"Done. Processed {count} DXCC entries. Successfully built {success} geometries.")
    End Sub

    Public Async Function BuildAndSaveDxccGeometry(dxccId As Integer) As Task(Of Boolean)
        Try
            Dim src = LoadDxccSource(dxccId)
            Dim geom = Await BuildDXCCGeometryAsync(src)
            If geom Is Nothing OrElse geom.IsEmpty Then
                Debug.WriteLine($"No geometry built for DXCC ID {dxccId}. Skipping save.")
                Return False
            End If
            SaveDxccGeometry(dxccId, geom)
        Catch ex As Exception
            Debug.WriteLine($"Error building geometry for DXCC ID {dxccId}: {ex.Message}")
            Return False
        End Try
        Return True
    End Function

    Private Function LoadDxccSource(dxccId As Integer) As DxccSource

        Dim sql = "SELECT DXCCnum, Entity, source, rule, bbox, tolerance_m, notes FROM DXCC WHERE DXCCnum=@id"
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()

            Using cmd As New SqliteCommand(sql, connect)
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

    Private Function BuildFromFake(fakeName As String,
                                       gf As GeometryFactory) As Geometry
        ' Some geometries just dont exist anywhere, especially small rocks, reefs and atolls
        ' This makes fake geometry for these

        Dim lon As Double
        Dim lat As Double
        Dim radius_m As Double

        Select Case fakeName

            'Case "Willis Is"
            '    lon = 149.9645 : lat = -16.28733
            '    radius_m = 200

            'Case "Mellish Reef"
            '    lon = 155.8644 : lat = -17.4126
            '    radius_m = 1000

            Case "St Peter & St Paul Rocks"
                lon = -29.335 : lat = 0.9158
                radius_m = 1000

            'Case "Malpelo Is"
            '    lon = -81.5833 : lat = 3.9833
            '    radius_m = 1000

            Case "SMOM"
                ' Sovereign Military Order of Malta (SMOM). The Magistral Palace in Rome
                lon = 12.4663 : lat = 41.8894
                radius_m = 500

            Case Else
                Throw New Exception($"FAKE parameter '{fakeName}' not recognised")
        End Select

        ' Convert metres → degrees (approx)
        Dim deg = radius_m / 110540.0

        ' Build point → buffer → polygon
        Dim pt As Point = gf.CreatePoint(New Coordinate(lon, lat))
        Return pt.Buffer(deg)
    End Function

    Private Sub SaveDxccGeometry(dxccId As Integer, geom As Geometry)
        If geom Is Nothing OrElse geom.IsEmpty Then
            Debug.WriteLine($"Geometry for DXCC {dxccId} is empty. Not saved")
            Return      ' don't save empty geometry
        End If

        ' Validate and repair geometry before saving
        If Not geom.IsValid Then
            ' Preferred: GeometryFixer (NTS 2.x)
            geom = geom.Buffer(0) ' This is a common quick fix for invalid geometries, but it can be slow and may not fix all issues. If GeometryFixer is available, it would be better to use that instead.
        End If

        ' Optional: normalize orientation
        If NetTopologySuite.Algorithm.Orientation.IsCCW(geom.Coordinates) Then
            geom = geom.Reverse()
        End If

        ' Save to DB
        Dim GeoJson As String = FromNTSToGeoJson(geom)   ' GeoJSON
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            EnsureHashVRegistered(connect) ' Register hashV ONCE per connection
            ' Update geometry and hash in one statement
            Dim sql = "UPDATE DXCC SET geometry = $g, hash = hashV(rule, bbox, tolerance_m, $g) WHERE DXCCnum = $id"

            Using cmd As New SqliteCommand(sql, connect)
                cmd.Parameters.AddWithValue("$g", GeoJson)
                cmd.Parameters.AddWithValue("$id", dxccId)
                cmd.ExecuteNonQuery()
            End Using
        End Using
    End Sub

    Public Async Function BuildDXCCGeometryAsync(src As DxccSource) As Task(Of Geometry)

        ' ---------------------------------------------------------
        ' STEP 1 — Resolve base geometry (already NTS)
        ' ---------------------------------------------------------
        Dim baseGeom As Geometry = Await ResolveRuleExpressionAsync(src)

        If baseGeom Is Nothing OrElse baseGeom.IsEmpty Then
            Debug.WriteLine($"Failed to resolve base geometry for DXCC '{src.name}' using rule: {src.rule}")
            Return Nothing
        End If

        ' ---------------------------------------------------------
        ' STEP 2 — Pre-simplify (optional but useful)
        ' ---------------------------------------------------------
        Dim preTol = 0.00002 ' ~2m
        baseGeom = TopologyPreservingSimplifier.Simplify(baseGeom, preTol)

        ' ---------------------------------------------------------
        ' STEP 3 — Apply bbox/poly clip
        ' ---------------------------------------------------------
        If Not String.IsNullOrWhiteSpace(src.bbox) Then

            Dim clipGeom As Geometry = ParseBBoxOrPoly(src.bbox)
            If clipGeom Is Nothing OrElse clipGeom.IsEmpty Then Return Nothing

            ' Topology fix on clipper
            clipGeom = clipGeom.Buffer(0)

            ' Clip using NTS intersection
            baseGeom = baseGeom.Intersection(clipGeom)

            ' Cleanup — buffer(0) fixes slivers, self-intersections, etc.
            baseGeom = baseGeom.Buffer(0)
        End If

        If baseGeom Is Nothing OrElse baseGeom.IsEmpty Then Return Nothing

        ' ---------------------------------------------------------
        ' STEP 4 — Generalize AFTER clipping
        ' ---------------------------------------------------------
        If src.tolerance_m > 0 Then
            Dim degTol = src.tolerance_m / 110540.0
            baseGeom = TopologyPreservingSimplifier.Simplify(baseGeom, degTol)
        End If

        ' ---------------------------------------------------------
        ' STEP 5 — Round coordinates
        ' ---------------------------------------------------------
        baseGeom = RoundGeometry(baseGeom, 6)

        Return baseGeom
    End Function

    Public Function RoundGeometry(geom As Geometry, decimals As Integer) As Geometry
        If geom Is Nothing Then Return Nothing

        Dim gf As GeometryFactory = geom.Factory

        Select Case True

        ' -------------------------
        ' POINT
        ' -------------------------
            Case TypeOf geom Is Point
                Dim p = DirectCast(geom, Point)
                Dim c = p.Coordinate
                Dim rc As New Coordinate(Math.Round(c.X, decimals),
                                     Math.Round(c.Y, decimals))
                Return gf.CreatePoint(rc)

        ' -------------------------
        ' LINESTRING
        ' -------------------------
            Case TypeOf geom Is LineString
                Dim ls = DirectCast(geom, LineString)
                Dim coords = ls.Coordinates.Select(
                Function(c) New Coordinate(Math.Round(c.X, decimals),
                                           Math.Round(c.Y, decimals))
            ).ToArray()
                Return gf.CreateLineString(coords)

        ' -------------------------
        ' POLYGON
        ' -------------------------
            Case TypeOf geom Is Polygon
                Dim poly = DirectCast(geom, Polygon)

                ' Shell
                Dim shellCoords = poly.Shell.Coordinates.Select(
                Function(c) New Coordinate(Math.Round(c.X, decimals),
                                           Math.Round(c.Y, decimals))
            ).ToArray()
                Dim shell = gf.CreateLinearRing(shellCoords)

                ' Holes
                Dim holes(poly.NumInteriorRings - 1) As LinearRing
                For i = 0 To poly.NumInteriorRings - 1
                    Dim holeCoords = poly.GetInteriorRingN(i).Coordinates.Select(
                    Function(c) New Coordinate(Math.Round(c.X, decimals),
                                               Math.Round(c.Y, decimals))
                ).ToArray()
                    holes(i) = gf.CreateLinearRing(holeCoords)
                Next

                Return gf.CreatePolygon(shell, holes)

        ' -------------------------
        ' MULTILINESTRING
        ' -------------------------
            Case TypeOf geom Is MultiLineString
                Dim mls = DirectCast(geom, MultiLineString)
                Dim lines As New List(Of LineString)
                For i = 0 To mls.NumGeometries - 1
                    lines.Add(DirectCast(RoundGeometry(mls.GetGeometryN(i), decimals), LineString))
                Next
                Return gf.CreateMultiLineString(lines.ToArray())

        ' -------------------------
        ' MULTIPOLYGON
        ' -------------------------
            Case TypeOf geom Is MultiPolygon
                Dim mp = DirectCast(geom, MultiPolygon)
                Dim polys As New List(Of Polygon)
                For i = 0 To mp.NumGeometries - 1
                    polys.Add(DirectCast(RoundGeometry(mp.GetGeometryN(i), decimals), Polygon))
                Next
                Return gf.CreateMultiPolygon(polys.ToArray())

        ' -------------------------
        ' GEOMETRYCOLLECTION
        ' -------------------------
            Case TypeOf geom Is GeometryCollection
                Dim gc = DirectCast(geom, GeometryCollection)
                Dim geoms As New List(Of Geometry)
                For i = 0 To gc.NumGeometries - 1
                    geoms.Add(RoundGeometry(gc.GetGeometryN(i), decimals))
                Next
                Return gf.CreateGeometryCollection(geoms.ToArray())

                ' -------------------------
                ' FALLBACK
                ' -------------------------
            Case Else
                Return geom
        End Select
    End Function



    ' ---------------------------------------------------------
    '  RULE PARSER (UNION SUPPORT) — NOW TAKES DxccSource
    ' ---------------------------------------------------------
    Private Async Function ResolveRuleExpressionAsync(src As DxccSource) As Task(Of Geometry)

        Dim parts As List(Of RuleToken) = Tokenize(src.rule)

        If parts.Count = 0 Then
            Return Nothing
        End If

        Dim current As Geometry = Nothing

        For Each part In parts

            Dim geom As Geometry = Await ResolveSingleSourceAsync(src.source, part.Token)
            If src.name = "Agalega & St Brandon" Then
                Debug.WriteLine($"Resolved geometry for {src.name} part: {geom.NumGeometries} parts")
            End If

            geom = RepairGeometry(geom)     ' repair the geometry

            If geom Is Nothing Then Continue For

            If current Is Nothing Then
                ' First geometry initializes accumulator
                current = geom

            Else
                Select Case part.Op

                    Case "+"c
                        current =
                            NetTopologySuite.Operation.Union.UnaryUnionOp.Union(
                                New Geometry() {current, geom}
                                )

                    Case "-"c
                        current =
                            current.Difference(geom)

                End Select
            End If
        Next

        Return current
    End Function

    Public Function RepairGeometry(g As Geometry) As Geometry
        If g Is Nothing OrElse g.IsEmpty Then Return g

        ' DO NOT use Buffer(0) on raw NE multipolygons
        ' First, ensure rings are closed and coordinates are valid
        g = g.Copy()

        ' Fix duplicate points
        g = NetTopologySuite.Precision.GeometryPrecisionReducer.Reduce(g,
                                                                       New PrecisionModel(1000000000.0))

        ' Now try buffer(0) ONLY on each polygon separately
        If TypeOf g Is MultiPolygon Then
            Dim polys As New List(Of Geometry)
            For i = 0 To g.NumGeometries - 1
                Dim p = g.GetGeometryN(i)
                If Not p.IsValid Then p = p.Buffer(0)
                polys.Add(p)
            Next
            Return g.Factory.CreateMultiPolygon(polys.Cast(Of Polygon).ToArray())
        End If

        ' Single polygon case
        If Not g.IsValid Then g = g.Buffer(0)

        Return g
    End Function


    Private Function Tokenize(rule As String) As List(Of RuleToken)

        Dim tokens As New List(Of RuleToken)

        If String.IsNullOrWhiteSpace(rule) Then
            Return tokens
        End If

        Dim parts = rule.Split({"+"c, "-"c}, StringSplitOptions.RemoveEmptyEntries)
        Dim ops = New List(Of Char)

        ' Extract operators in order
        For Each ch In rule
            If ch = "+"c OrElse ch = "-"c Then
                ops.Add(ch)
            End If
        Next

        ' First token defaults to '+'
        If ops.Count < parts.Length Then
            ops.Insert(0, "+"c)
        End If

        ' Build RuleToken objects
        For i = 0 To parts.Length - 1
            tokens.Add(New RuleToken With {
                          .Op = ops(i),
                          .Token = parts(i).Trim()
                          })
        Next

        Return tokens
    End Function

    Public Class RuleToken
        Public Property Op As Char
        Public Property Token As String
    End Class


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
                Return BuildFromFake(token, factory)

            Case Else
                Throw New Exception($"Unsupported resolver type '{source}' for token '{token}'.")
        End Select
    End Function
End Module
