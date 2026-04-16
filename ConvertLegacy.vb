Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.Json
Imports Microsoft.Data.Sqlite

' ========================================================================
' MODULE: ConvertLegacy
'
' PURPOSE:
'   Converts legacy Overpass selector strings stored in Dxcc.Query into
'   modern relation/way expressions (e.g., "3547592+W27264466").
'
'   This module:
'     • Reads DXCC rows that still contain legacy queries
'     • Sends each query to Overpass (with throttling + retries)
'     • Parses the returned JSON
'     • Extracts relation IDs (no prefix) and way IDs ("W" prefix)
'     • Ignores nodes (no geometry)
'     • Writes the resolved expression + hash back into SQLite
'
'   This is a ONE-TIME conversion pipeline. Once a relation/way ID is
'   stored, the system never needs Overpass again for that DXCC entity.
'
' DESIGN FEATURES:
'   • 1 request per second throttle to avoid rate limits
'   • 90-second timeout for large multipolygons
'   • Retry logic with exponential backoff
'   • Hashing for change detection
'   • Normalization of legacy selector syntax
' ========================================================================

Module ConvertLegacy

    ' Timestamp of the last Overpass request.
    ' Used to enforce a minimum 1-second delay between requests.
    Private lastRequestTime As DateTime = DateTime.MinValue

    ' Lock object for thread-safe throttling.
    Private throttleLock As New Object()

    ' ====================================================================
    ' FUNCTION: ConvertLegacyQueriesAsync
    '
    ' PURPOSE:
    '   Iterates through all DXCC rows that still contain legacy queries
    '   and resolves them into relation/way expressions.
    '
    ' PROCESS:
    '   1. Select rows where Relation is NULL or empty
    '   2. For each row:
    '       • Display progress in Form1.TextBox1
    '       • Resolve the legacy query → relation expression
    '       • If successful, compute hash and update the database
    '       • If failed, log the failure and continue
    '
    ' NOTES:
    '   DXCCnum 999 is excluded (special case entry)
    ' ====================================================================
    Public Async Function ConvertLegacyQueriesAsync() As Task

        Using conn As New SqliteConnection(DXCC_DATA)
            conn.Open()

            Dim selectCmd As New SqliteCommand(
                "SELECT DXCCnum, Entity, Query 
                 FROM Dxcc 
                 WHERE Deleted=0 
                   AND (Relation IS NULL OR Relation ='') 
                   AND DXCCnum<>999", conn)

            Using reader = selectCmd.ExecuteReader()
                While reader.Read()

                    Dim dxccnum = reader("DXCCnum")
                    Dim entity = reader("Entity").ToString()
                    Dim legacyQuery = reader("Query").ToString()

                    AppendText(Form1.TextBox1, $"Getting {entity}{vbCrLf}")

                    ' Resolve the legacy Overpass selector into R/W IDs
                    Dim relationExpr As String =
                        Await ResolveQueryWithRetries(dxccnum, legacyQuery)

                    If relationExpr = "" Then
                        ' All retry attempts failed
                        AppendText(Form1.TextBox1,
                                   $"Failed to resolve DXCC {entity} after multiple attempts.{vbCrLf}")
                        Continue While
                    Else
                        ' Compute a hash of the resolved expression
                        Dim hash As String = HashText(relationExpr)

                        ' Write the resolved expression + hash back to SQLite
                        Dim updateCmd As New SqliteCommand(
                            "UPDATE Dxcc 
                             SET Relation=@expr, hash=@hash 
                             WHERE DXCCnum=@num", conn)

                        updateCmd.Parameters.AddWithValue("@expr", relationExpr)
                        updateCmd.Parameters.AddWithValue("@num", dxccnum)
                        updateCmd.Parameters.AddWithValue("@hash", hash)

                        updateCmd.ExecuteNonQuery()

                        AppendText(Form1.TextBox1,
                                   $"Updated {entity} with relation {relationExpr}{vbCrLf}")
                    End If

                End While
            End Using
        End Using

        AppendText(Form1.TextBox1, "Done")
    End Function

    ' ====================================================================
    ' FUNCTION: ResolveQueryWithRetries
    '
    ' PURPOSE:
    '   Wraps ResolveQueryOnce() with retry logic.
    '
    ' BEHAVIOR:
    '   • Up to 5 attempts
    '   • Exponential backoff (1s → 2s → 4s → 8s → 16s)
    '   • Returns "" on final failure
    ' ====================================================================
    Private Async Function ResolveQueryWithRetries(dxccnum As Integer,
                                                   legacyQuery As String) As Task(Of String)

        Dim delay As Integer = 1000

        For attempt = 1 To 5
            Debug.WriteLine($"DXCC {dxccnum}: Attempt {attempt} to resolve query.")

            Try
                Dim result = Await ResolveQueryOnce(legacyQuery)
                Return result

            Catch ex As Exception
                If attempt = 5 Then
                    Debug.WriteLine($"DXCC {dxccnum}: Attempt {attempt} failed: {ex.Message}")
                    Return ""
                End If
            End Try

            Await Task.Delay(delay)
            delay *= 2
        Next

        Return ""
    End Function

    ' ====================================================================
    ' FUNCTION: ResolveQueryOnce
    '
    ' PURPOSE:
    '   Sends a normalized legacy selector to Overpass and extracts:
    '     • relation IDs (no prefix)
    '     • way IDs ("W" prefix)
    '     • ignores nodes
    '
    ' NOTES:
    '   • Uses FR Overpass server (faster, less strict)
    '   • 90-second timeout for large multipolygons
    '   • Returns a '+'-joined expression (e.g., "3547592+W27264466")
    ' ====================================================================
    Private Async Function ResolveQueryOnce(legacyQuery As String) As Task(Of String)

        ' Enforce 1 request per second
        Await ThrottleAsync()

        Dim query = NormalizeLegacyQuery(legacyQuery).Replace(vbCr, "")
        Dim overpassQuery As String = $"[out:json][timeout:90];{query};out ids;" & vbLf

        Dim content = New StringContent(overpassQuery, Encoding.UTF8, "text/plain")
        Dim response = Await Http.PostAsync("https://overpass.kumi.systems/api/interpreter", content)

        If Not response.IsSuccessStatusCode Then
            Throw New Exception($"HTTP {response.StatusCode}")
        End If

        Dim json = Await response.Content.ReadAsStringAsync()

        If json.TrimStart().StartsWith("<") Then
            Throw New Exception("Received HTML instead of JSON")
        End If

        Dim doc As JsonDocument
        Try
            doc = JsonDocument.Parse(json)
        Catch
            Throw New Exception("Invalid JSON returned by Overpass")
        End Try

        Dim ids As New List(Of String)

        If doc.RootElement.TryGetProperty("elements", Nothing) Then
            For Each el In doc.RootElement.GetProperty("elements").EnumerateArray()

                Dim t = el.GetProperty("type").GetString()
                Dim id = el.GetProperty("id").GetInt64()

                Select Case t
                    Case "relation"
                        ids.Add(id.ToString())      ' Relations: no prefix

                    Case "way"
                        ids.Add("W" & id.ToString()) ' Ways: prefix W

                    Case "node"
                        ' Nodes have no area → ignored
                End Select

            Next
        End If

        Return String.Join("+", ids)
    End Function

    ' ====================================================================
    ' FUNCTION: NormalizeLegacyQuery
    '
    ' PURPOSE:
    '   Converts legacy selector strings into valid Overpass syntax.
    '
    ' HANDLES:
    '   • Missing semicolons
    '   • Multi-line selectors
    '   • Already-wrapped parentheses
    '   • Ensures final form is "(selector1; selector2; ...)"
    ' ====================================================================
    Private Function NormalizeLegacyQuery(legacy As String) As String

        Dim q = legacy.Trim()

        ' CASE 1: Already wrapped in parentheses
        If q.StartsWith("(") AndAlso q.EndsWith(")") Then

            Dim lines = q.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
            Dim fixed As New List(Of String)

            For Each line In lines
                Dim t = line.Trim()

                If t.StartsWith("rel") AndAlso Not t.EndsWith(";") Then
                    t &= ";"
                End If

                fixed.Add(t)
            Next

            Return String.Join(" ", fixed)
        End If

        ' CASE 2: Not wrapped — treat as list of selectors
        Dim selectors = legacy.Split({vbCr, vbLf}, StringSplitOptions.RemoveEmptyEntries)
        Dim normalized As New List(Of String)

        For Each sel In selectors
            Dim t = sel.Trim()
            If t <> "" AndAlso Not t.EndsWith(";") Then
                t &= ";"
            End If
            normalized.Add(t)
        Next

        Return "(" & String.Join(" ", normalized) & ")"
    End Function

    ' ====================================================================
    ' FUNCTION: ThrottleAsync
    '
    ' PURPOSE:
    '   Ensures a minimum of 1 second between Overpass requests.
    '
    ' WHY:
    '   • Prevents rate-limit errors
    '   • Prevents "server too busy" load shedding
    ' ====================================================================
    Private Async Function ThrottleAsync() As Task

        Dim waitTime As Integer = 0

        SyncLock throttleLock

            If lastRequestTime <> DateTime.MinValue Then
                Dim elapsed = (DateTime.UtcNow - lastRequestTime).TotalMilliseconds

                If elapsed < 1000 Then
                    waitTime = CInt(1000 - elapsed)
                End If
            End If

            lastRequestTime = DateTime.UtcNow.AddMilliseconds(waitTime)

        End SyncLock

        If waitTime > 0 Then
            Await Task.Delay(waitTime)
        End If

    End Function

End Module
