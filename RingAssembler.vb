Public Module WayRingAssembler

    Private Const SNAP_TOL As Double = 0.0001   ' ~11 meters

    ' ============================================================
    ' Convert OSM ways → stitched rings (List(Of List(Of PointD)))
    ' ============================================================
    Public Function BuildRingsFromWays(
        ways As List(Of OSMObject),
        nodeDict As Dictionary(Of Long, OSMNode)
    ) As List(Of List(Of PointD))

        ' 1. Convert ways into coordinate lists
        Dim segments As New List(Of List(Of PointD))()

        For Each w In ways
            If w.Nodes Is Nothing OrElse w.Nodes.Count < 2 Then Continue For

            Dim pts As New List(Of PointD)
            For Each nid In w.Nodes
                Dim n = nodeDict(nid)
                pts.Add(New PointD(n.Lat, n.Lon))   ' Lat, Lon
            Next

            segments.Add(pts)
        Next

        ' 2. Build endpoint index
        Dim endpointMap As New Dictionary(Of String, List(Of List(Of PointD)))()

        For Each seg In segments
            Dim a = SnapKey(seg.First())
            Dim b = SnapKey(seg.Last())

            If Not endpointMap.ContainsKey(a) Then endpointMap(a) = New List(Of List(Of PointD))()
            If Not endpointMap.ContainsKey(b) Then endpointMap(b) = New List(Of List(Of PointD))()

            endpointMap(a).Add(seg)
            endpointMap(b).Add(seg)
        Next

        ' 3. Walk graph to build rings
        Dim used As New HashSet(Of List(Of PointD))()
        Dim rings As New List(Of List(Of PointD))()

        For Each seg In segments
            If used.Contains(seg) Then Continue For

            Dim ringPts As New List(Of PointD)(seg)
            used.Add(seg)

            Dim extended As Boolean = True

            While extended
                extended = False

                Dim startKey = SnapKey(ringPts.First())
                Dim endKey = SnapKey(ringPts.Last())

                ' Try to extend at the end
                For Each cand In endpointMap(endKey)
                    If used.Contains(cand) Then Continue For

                    If PointsClose(ringPts.Last(), cand.First()) Then
                        ringPts.AddRange(cand.Skip(1))
                        used.Add(cand)
                        extended = True
                        Exit For
                    End If

                    If PointsClose(ringPts.Last(), cand.Last()) Then
                        Dim rev = cand.AsEnumerable().Reverse().ToList()
                        ringPts.AddRange(rev.Skip(1))
                        used.Add(cand)
                        extended = True
                        Exit For
                    End If
                Next

                If extended Then Continue While

                ' Try to extend at the start
                For Each cand In endpointMap(startKey)
                    If used.Contains(cand) Then Continue For

                    If PointsClose(ringPts.First(), cand.Last()) Then
                        Dim newPts = New List(Of PointD)(cand)
                        newPts.AddRange(ringPts.Skip(1))
                        ringPts = newPts
                        used.Add(cand)
                        extended = True
                        Exit For
                    End If

                    If PointsClose(ringPts.First(), cand.First()) Then
                        Dim rev = cand.AsEnumerable().Reverse().ToList()
                        Dim newPts = New List(Of PointD)(rev)
                        newPts.AddRange(ringPts.Skip(1))
                        ringPts = newPts
                        used.Add(cand)
                        extended = True
                        Exit For
                    End If
                Next
            End While

            ' 4. Close ring if endpoints match
            If PointsClose(ringPts.First(), ringPts.Last()) Then
                ringPts(ringPts.Count - 1) = ringPts.First()
            End If

            ' Only accept closed rings
            If PointsClose(ringPts.First(), ringPts.Last()) Then
                rings.Add(ringPts)
            End If
        Next

        Return rings
    End Function

    ' ============================================================
    ' Helpers
    ' ============================================================

    Private Function SnapKey(p As PointD) As String
        Dim lat = Math.Round(p.Lat / SNAP_TOL) * SNAP_TOL
        Dim lon = Math.Round(p.Lon / SNAP_TOL) * SNAP_TOL
        Return lat.ToString("F6") & "," & lon.ToString("F6")
    End Function

    Private Function PointsClose(a As PointD, b As PointD) As Boolean
        Dim dx = a.Lon - b.Lon
        Dim dy = a.Lat - b.Lat
        Return (dx * dx + dy * dy) < (SNAP_TOL * SNAP_TOL)
    End Function

End Module
