Imports System.Net.Http
Imports System.Security.Cryptography
Imports System.Text
Imports System.Text.RegularExpressions


Public Module helpers
    Public Const DXCC_DATA = "data Source=DXCC.sqlite"     ' the DXCC database
    Public Http As HttpClient
    Public HttpHandler As HttpClientHandler
    Private ReadOnly _hashVRegistered As New HashSet(Of SqliteConnection)
    Public Sub EnsureHashVRegistered(conn As SqliteConnection)
        SyncLock _hashVRegistered
            If _hashVRegistered.Contains(conn) Then Exit Sub
            conn.CreateFunction(name:="hashV", function:=Function(values As Object()) HashVariadic(values))
            _hashVRegistered.Add(conn)
        End SyncLock
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

    Public Sub InitHttp()
        HttpHandler = New HttpClientHandler()

        Http = New HttpClient(HttpHandler) With {
        .Timeout = TimeSpan.FromMinutes(10)}

        Http.DefaultRequestHeaders.UserAgent.Clear()
        Http.DefaultRequestHeaders.UserAgent.ParseAdd("DXCCMapper/1.0 (contact: myhillman@gmail.com)")
    End Sub

    Public Delegate Sub SetTextCallback(tb As System.Windows.Forms.TextBox, ByVal text As String)

    Public Sub SetText(tb As System.Windows.Forms.TextBox, ByVal text As String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf SetText), New Object() {tb, text})
        Else
            tb.Text = text
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub

    Public Sub AppendText(tb As System.Windows.Forms.TextBox, ByVal text As String)
        ' InvokeRequired required compares the thread ID of the
        ' calling thread to the thread ID of the creating thread.
        ' If these threads are different, it returns true.
        If tb.InvokeRequired Then
            tb.Invoke(New SetTextCallback(AddressOf AppendText), New Object() {tb, text})
        Else
            tb.AppendText(text)
        End If
        System.Windows.Forms.Application.DoEvents()
    End Sub
    Public Function IsDBNullorEmpty(ByVal value As Object) As Boolean
        ' test if DB field is null or empty
        Return IsDBNull(value) OrElse value = ""
    End Function
    Public Function SafeStr(value As Object) As String
        If value Is Nothing OrElse value Is DBNull.Value Then Return ""
        Return value.ToString()
    End Function
    Public Function SafeDbl(value As Object) As Double
        If value Is Nothing OrElse value Is DBNull.Value Then Return 0
        Return CDbl(value.ToString())
    End Function
    Public Function SafeInt(value As Object) As Integer
        If value Is Nothing OrElse value Is DBNull.Value Then Return 0
        Return CInt(value.ToString())
    End Function

    Public Function SQLescape(st As String) As String
        ' escape special characters for SQL
        Return st.Replace("'", "''")
    End Function
    Public Function KMLescape(st As String) As String
        ' escape special characters for KML
        Dim escapes As New Dictionary(Of String, String) From {
            {"<", "&lt;"},
            {">", "&gt;"},
            {"&", "&amp;"},
            {"""", "&quot;"},
            {"\'", "&apos;"}
            }

        Dim result As String
        result = st
        For Each s In escapes
            result = result.Replace(s.Key, s.Value)
        Next
        Return result
    End Function

    Public Function Hyperlink(links As String) As String
        ' convert a list of hyperlinks, separated by semi-colon, and return html hyperlinks separated by <br>
        Dim hyperlinkList As New List(Of String)
        Dim hyperlinks = links.Split(";").ToList
        For ndx = 0 To hyperlinks.Count - 1
            hyperlinks(ndx) = Regex.Replace(hyperlinks(ndx), "[^0-9a-zA-Z\-\:\./]", "")     ' remove noise characters
        Next
        For Each link In hyperlinks
            Dim matches = Regex.Match(link, "^https.*?-(.*)$")
            hyperlinkList.Add($"<a href=""{link}"">{matches.Groups(1)}</a>")
        Next
        Return String.Join(";<br>", hyperlinkList)
    End Function

    ' ====================================================================
    ' FUNCTION: HashText
    '
    ' PURPOSE:
    '   Produces a SHA-256 hash of a string.
    '
    ' WHY:
    '   • Allows change detection
    '   • Ensures stable identifiers even if expression order changes
    ' ====================================================================
    Public Function HashText(input As String) As String
        Using sha As SHA256 = SHA256.Create()
            Dim bytes = Encoding.UTF8.GetBytes(input)
            Dim hash = sha.ComputeHash(bytes)
            Return Convert.ToHexString(hash)
        End Using
    End Function


    Public Sub OpenHtml(path As String)
        Dim psi As New ProcessStartInfo With {
                .FileName = path,
                .UseShellExecute = True
                }
        System.Diagnostics.Process.Start(psi)
    End Sub
End Module
