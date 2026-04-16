Imports System.Text.RegularExpressions
Imports Microsoft.Data.Sqlite
Public Module QueryValidator

    '===========================
    '  REGEX DEFINITIONS
    '===========================

    Private ReadOnly RE_QUERY_START As New Regex("^\s*(rel|way|node|\()", RegexOptions.IgnoreCase)
    Private ReadOnly RE_FORBID_POLY_OR_BBOX_IN_QUERY As New Regex("poly:|[-+]?\d+(\.\d+)?\s*,\s*[-+]?\d+(\.\d+)?", RegexOptions.IgnoreCase)

    Private ReadOnly RE_TAG_FILTER As New Regex(
    "^\[\s*((""[^""]+""|[A-Za-z0-9:_-]+))\s*(=|~)\s*((""[^""]+""|[A-Za-z0-9:_\-\.\(\)\|\^\$]+))\s*\]$",
    RegexOptions.IgnoreCase
)


    Private ReadOnly RE_ID_SELECTOR As New Regex("^(rel|way|node)\(\d+\)$", RegexOptions.IgnoreCase)

    Private ReadOnly RE_BBOX_RECT As New Regex(
        "^\s*([-+]?\d+(?:\.\d+)?)\s*,\s*" &
        "([-+]?\d+(?:\.\d+)?)\s*,\s*" &
        "([-+]?\d+(?:\.\d+)?)\s*,\s*" &
        "([-+]?\d+(?:\.\d+)?)\s*$"
    )

    Private ReadOnly RE_BBOX_POLY As New Regex("^poly:""([^""]+)""$", RegexOptions.IgnoreCase)

    '===========================
    '  PUBLIC ENTRY POINT
    '===========================
    ''' <summary>
    ''' Scans all DXCC records that contain a non-empty query and validates each one
    ''' using <c>ValidateRow</c>. Any validation errors are appended to the UI log and
    ''' progress is reflected in the main form's progress bar.
    ''' </summary>
    ''' <remarks>
    ''' This routine only validates rows where the <c>query</c> field is non-empty.
    ''' Rows that rely solely on polygon-based bounding boxes (e.g., Antarctica) are
    ''' intentionally skipped. Validation results are written to <c>Form1.TextBox1</c>.
    ''' </remarks>
    Public Sub ValidateDXCC()
        Dim errors As New List(Of String), count As Integer, errorCount As Integer = 0
        Using connect As New SqliteConnection(DXCC_DATA)
            connect.Open()
            Dim sql As SqliteCommand
            sql = connect.CreateCommand
            sql.CommandText = "SELECT COUNT(*) FROM DXCC WHERE Deleted=0 AND (query IS NOT NULL AND query <> '' AND DXCCnum<>999)"
            count = sql.ExecuteScalar()
            With Form1.ProgressBar1
                .Minimum = 0
                .Value = 0
                .Maximum = count
            End With
            sql.CommandText = "SELECT Entity,query,bbox FROM DXCC WHERE Deleted=0 AND (query IS NOT NULL AND query <> '' AND DXCCnum<>999)"
            Dim sqldr As SqliteDataReader = sql.ExecuteReader()
            count = 0
            While sqldr.Read
                Dim entity As String = sqldr("Entity")
                Dim query As String = SafeStr(sqldr("query"))
                Dim bbox As String = SafeStr(sqldr("bbox"))
                Dim entryErrors = ValidateRow(query, bbox)
                errorCount += entryErrors.Count
                If entryErrors.Count > 0 Then
                    AppendText(Form1.TextBox1, $"DXCC entry '{entity}' has {entryErrors.Count} validation error(s):{vbCrLf}query <{query}> bbox <{bbox}>{vbCrLf}")
                    For Each e In entryErrors
                        AppendText(Form1.TextBox1, $"{e}{vbCrLf}")
                    Next
                End If
                count += 1
                Form1.ProgressBar1.Value = count
            End While
        End Using
        AppendText(Form1.TextBox1, $"Validation complete: {errorCount} error(s) found in {count} entries.{vbCrLf}")
    End Sub

    Public Function ValidateRow(query As String, bbox As String) As List(Of String)
        Dim errors As New List(Of String)

        ValidateQuery(query, bbox, errors)
        ValidateBbox(bbox, errors)
        ValidateCrossRules(query, bbox, errors)

        Return errors
    End Function


    '===========================
    '  QUERY VALIDATION
    '===========================

    Private Sub ValidateQuery(query As String, bbox As String, errors As List(Of String))
        Dim q As String = If(query, "").Trim()
        Dim b As String = If(bbox, "").Trim()

        ' Q001 — Query empty but bbox not polygon
        If q = "" Then
            If b <> "" AndAlso Not RE_BBOX_POLY.IsMatch(b) Then
                errors.Add("Q001: Query is empty but bbox is not a poly:""..."".")
            End If
            Exit Sub
        End If

        ' Q003 — Must start with rel/way/node/(
        If Not RE_QUERY_START.IsMatch(q) Then
            errors.Add("Q003: Query must start with rel, way, node, or '('.")
        End If

        ' Q002 — Query contains forbidden bbox/poly
        If RE_FORBID_POLY_OR_BBOX_IN_QUERY.IsMatch(q) Then
            errors.Add("Q002: Query contains forbidden bbox/poly content.")
        End If

        ' Q004 — Invalid characters
        Dim allowed As String = " ()[]"";:=~^$|_-.+*/0-9A-Za-z'" & vbTab & vbCr & vbLf
        For Each c In q
            If Not Char.IsLetterOrDigit(c) AndAlso Not allowed.Contains(c) Then
                errors.Add("Q004: Query contains invalid characters.")
                Exit For
            End If
        Next

        ' Q005 — Unbalanced parentheses
        If CountChar(q, "("c) <> CountChar(q, ")"c) Then
            errors.Add("Q005: Query has unbalanced parentheses.")
        End If

        ' Q006 — Unbalanced brackets
        If CountChar(q, "["c) <> CountChar(q, "]"c) Then
            errors.Add("Q006: Query has unbalanced tag filter brackets [].")
        End If

        ' Q014 — Empty parentheses
        If q.Contains("()") Then
            errors.Add("Q014: Query contains empty parentheses.")
        End If

        ' Q015 — Empty tag filters
        If q.Contains("[]") Then
            errors.Add("Q015: Query contains empty tag filters [].")
        End If

        ' Q016/Q017/Q018 — Strict tag filter grammar
        For Each part In ExtractBracketedParts(q)
            If Not RE_TAG_FILTER.IsMatch(part) Then
                errors.Add("Q017: Query contains malformed tag filter: expected [""key""=""value""].")
            End If
        Next

        ' Q019 — ID selectors must be inside parentheses blocks
        For Each token In Tokenize(q)
            If token.Contains("(") AndAlso token.Contains(")") Then
                If RE_ID_SELECTOR.IsMatch(token) Then
                    ' OK
                ElseIf token.Contains("(") AndAlso token.Contains(")") Then
                    ' Might be way(r) or rel(r) — allowed
                Else
                    errors.Add("Q019: Invalid ID selector syntax.")
                End If
            End If
        Next

        ' Q010 — Semicolon outside parentheses
        If q.Contains(";") AndAlso Not q.TrimStart().StartsWith("(") Then
            errors.Add("Q010: Semicolon found outside a set block.")
        End If

        ' Q012 — Dangling operator
        If q.EndsWith("=") OrElse q.EndsWith("~") OrElse q.EndsWith(":") Then
            errors.Add("Q012: Query ends with an incomplete expression.")
        End If
    End Sub


    '===========================
    '  BBOX VALIDATION
    '===========================

    Private Sub ValidateBbox(bbox As String, errors As List(Of String))
        Dim b As String = If(bbox, "").Trim()

        If b = "" Then Exit Sub

        ' POLYGON
        Dim mPoly = RE_BBOX_POLY.Match(b)
        If mPoly.Success Then
            Dim coords = mPoly.Groups(1).Value.Trim().Split({" "}, StringSplitOptions.RemoveEmptyEntries)

            ' B003 — odd number of coords
            If coords.Length Mod 2 <> 0 Then
                errors.Add("B003: Polygon bbox must contain an even number of coordinates.")
                Exit Sub
            End If

            ' numeric + range checks
            Dim nums As New List(Of Double)
            For Each c In coords
                Dim v As Double
                If Not Double.TryParse(c, v) Then
                    errors.Add("B004: Polygon bbox contains non-numeric values.")
                    Exit Sub
                End If
                nums.Add(v)
            Next

            ' B005/B006 — range checks
            For i = 0 To nums.Count - 1 Step 2
                Dim lat = nums(i)
                Dim lon = nums(i + 1)
                If lat < -90 OrElse lat > 90 Then errors.Add("B005: Polygon bbox latitude out of range.")
                If lon < -180 OrElse lon > 180 Then errors.Add("B006: Polygon bbox longitude out of range.")
            Next

            ' B011 — polygon must be closed
            If nums(0) <> nums(nums.Count - 2) OrElse nums(1) <> nums(nums.Count - 1) Then
                errors.Add("B011: Polygon bbox must be closed (first and last coordinate must match).")
            End If

            ' B012 — must have at least 3 points
            If nums.Count < 6 Then
                errors.Add("B012: Polygon bbox must contain at least three points.")
            End If

            Exit Sub
        End If

        ' RECTANGLE
        Dim mRect = RE_BBOX_RECT.Match(b)
        If Not mRect.Success Then
            errors.Add("B001: Bbox must be empty, a poly:""..."", or a rectangular bbox.")
            Exit Sub
        End If

        Dim lat1 = Double.Parse(mRect.Groups(1).Value)
        Dim lon1 = Double.Parse(mRect.Groups(2).Value)
        Dim lat2 = Double.Parse(mRect.Groups(3).Value)
        Dim lon2 = Double.Parse(mRect.Groups(4).Value)

        If lat1 < -90 OrElse lat1 > 90 OrElse lat2 < -90 OrElse lat2 > 90 Then
            errors.Add("B008: Rectangular bbox latitude out of range.")
        End If

        If lon1 < -180 OrElse lon1 > 180 OrElse lon2 < -180 OrElse lon2 > 180 Then
            errors.Add("B009: Rectangular bbox longitude out of range.")
        End If

        If lat1 >= lat2 OrElse lon1 >= lon2 Then
            errors.Add("B010: Rectangular bbox must be ordered south<north and west<east.")
        End If
    End Sub


    '===========================
    '  CROSS-FIELD RULES
    '===========================

    Private Sub ValidateCrossRules(query As String, bbox As String, errors As List(Of String))
        Dim q As String = If(query, "").Trim()
        Dim b As String = If(bbox, "").Trim()

        ' QB003 — If query contains rel(id), bbox must be empty
        If q.Contains("rel(") AndAlso b <> "" Then
            errors.Add("QB003: ID-based selectors require empty bbox.")
        End If

        ' QB005 — If query contains way(r) or rel(r), bbox must be empty
        If q.Contains("way(r") OrElse q.Contains("rel(r") Then
            If b <> "" Then
                errors.Add("QB005: Set-based selectors require empty bbox.")
            End If
        End If

    End Sub


    '===========================
    '  UTILITIES
    '===========================

    Private Function CountChar(s As String, c As Char) As Integer
        Return s.Count(Function(ch) ch = c)
    End Function

    Private Function ExtractBracketedParts(q As String) As IEnumerable(Of String)
        Dim parts As New List(Of String)
        Dim start As Integer = -1
        Dim depth As Integer = 0

        For i = 0 To q.Length - 1
            Dim ch = q(i)

            If ch = "["c Then
                If depth = 0 Then start = i
                depth += 1
            ElseIf ch = "]"c Then
                depth -= 1
                If depth = 0 AndAlso start >= 0 Then
                    parts.Add(q.Substring(start, i - start + 1))
                    start = -1
                End If
            End If
        Next

        Return parts
    End Function

    Private Function Tokenize(q As String) As IEnumerable(Of String)
        Return q.Split({" ", vbTab, vbCr, vbLf, ";"}, StringSplitOptions.RemoveEmptyEntries)
    End Function

End Module

