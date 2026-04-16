Imports System.Data.SQLite
Imports System.Text

Module Module1

    Sub FixCoords()

        Using conn As New SqliteConnection(DXCC_DATA)
            conn.Open()

            ' Read all rows
            Using cmd As New SqliteCommand("SELECT * from DXCC WHERE bbox IS NOT NULL", conn)
                Using r = cmd.ExecuteReader()
                    While r.Read()

                        Dim id = r("DXCCnum")
                        Dim raw = r("bbox")

                        Dim newValue As String = ConvertField(raw)

                        Using upd As New SqliteCommand("UPDATE dxcc SET bbox=@b WHERE DXCCnum=@id", conn)
                            upd.Parameters.AddWithValue("@b", newValue)
                            upd.Parameters.AddWithValue("@id", id)
                            upd.ExecuteNonQuery()
                        End Using

                        ' Update row
                        Using upd As New SqliteCommand("UPDATE dxcc SET bbox=@b WHERE DXCCnum=@id", conn)
                            upd.Parameters.AddWithValue("@b", newValue)
                            upd.Parameters.AddWithValue("@id", id)
                            upd.ExecuteNonQuery()
                        End Using

                    End While
                End Using
            End Using

        End Using

    End Sub

    ' ---------------------------------------------------------
    ' Detects bbox vs poly, strips prefix, converts, outputs comma-separated
    ' ---------------------------------------------------------
    Function ConvertField(raw As String) As String

        Dim content As String = raw.Trim()

        ' Strip poly:" prefix if present
        If content.StartsWith("poly:""", StringComparison.OrdinalIgnoreCase) Then
            content = content.Substring(6) ' remove poly:"
            If content.EndsWith("""") Then
                content = content.Substring(0, content.Length - 1)
            End If
        End If

        ' Split on space or comma
        Dim tokens = content.Split({" ", ","}, StringSplitOptions.RemoveEmptyEntries)

        Dim newContent As String

        If tokens.Length = 4 Then
            ' -----------------------------------------
            ' BBOX: minLat,minLon,maxLat,maxLon
            ' → minLon,minLat,maxLon,maxLat
            ' -----------------------------------------
            Dim minLat = tokens(0)
            Dim minLon = tokens(1)
            Dim maxLat = tokens(2)
            Dim maxLon = tokens(3)

            newContent = $"{minLon},{minLat},{maxLon},{maxLat}"

        ElseIf tokens.Length Mod 2 = 0 Then
            ' -----------------------------------------
            ' POLY: lat lon lat lon ...
            ' → lon,lat,lon,lat,...
            ' -----------------------------------------
            Dim sb As New StringBuilder()

            For i = 0 To tokens.Length - 1 Step 2
                Dim lat = tokens(i)
                Dim lon = tokens(i + 1)
                sb.Append(lon).Append(",").Append(lat).Append(",")
            Next

            ' Remove trailing comma
            newContent = sb.ToString().TrimEnd(","c)

        Else
            ' Invalid geometry — return unchanged
            Return raw
        End If

        Return newContent

    End Function

End Module
