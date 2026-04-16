'Imports Windows.Devices.Geolocation
Imports System.Device.Location
Imports System.IO
Imports Esri.ArcGISRuntime.Geometry
Imports OsmSharp
Imports OsmSharp.Complete
Imports OsmSharp.Streams
Public Module PbfExtractor

    Public Class BoundaryExtractor

        ''' <summary>
        ''' Extracts a polygon boundary for a given relation ID from a .osm.pbf file.
        ''' </summary>
        Public Shared Function ExtractBoundary(pbfPath As String, relationId As Long) As List(Of MapPoint)
            Debug.WriteLine($"[Extract] Opening PBF: {pbfPath}")

            Using fs As FileStream = File.OpenRead(pbfPath)
                Dim source As New PBFOsmStreamSource(fs)

                ' Load everything into memory (fast for region PBFs)
                Dim all As List(Of OsmGeo) = source.ToList()

                Debug.WriteLine($"[Extract] Loaded {all.Count} OSM primitives")

                ' --- 1. Find relation ---
                Dim rel As Relation =
                all.OfType(Of Relation)().
                FirstOrDefault(Function(r) r.Id = relationId)

                If rel Is Nothing Then
                    Throw New Exception($"Relation {relationId} not found in PBF.")
                End If

                Debug.WriteLine($"[Extract] Found relation {rel.Id}")

                ' --- 2. Collect referenced ways ---
                Dim wayIds As HashSet(Of Long) =
                rel.Members.
                Where(Function(m) m.Type = OsmGeoType.Way).
                Select(Function(m) m.Id).
                ToHashSet()

                Dim ways As Dictionary(Of Long, Way) =
                all.OfType(Of Way)().
                Where(Function(w) wayIds.Contains(w.Id)).
                ToDictionary(Function(w) w.Id.Value)

                Debug.WriteLine($"[Extract] Loaded {ways.Count} ways")

                ' --- 3. Collect referenced nodes ---
                Dim nodeIds As HashSet(Of Long) =
                ways.Values.SelectMany(Function(w) w.Nodes).ToHashSet()

                Dim nodes As Dictionary(Of Long, Node) =
                all.OfType(Of Node)().
                Where(Function(n) nodeIds.Contains(n.Id.Value)).
                ToDictionary(Function(n) n.Id.Value)

                Debug.WriteLine($"[Extract] Loaded {nodes.Count} nodes")

                ' --- 4. Build CompleteWays ---
                Dim completeWays As New List(Of CompleteWay)

                For Each w In ways.Values
                    Dim cw As New CompleteWay()
                    cw.Id = w.Id
                    cw.Tags = w.Tags
                    cw.Nodes = w.Nodes.Select(Function(nid) nodes(nid)).ToArray()
                    completeWays.Add(cw)
                Next

                ' --- 5. Build CompleteRelation ---
                Dim compRel As New CompleteRelation()
                compRel.Id = CLng(rel.Id)
                compRel.Tags = rel.Tags
                compRel.Members = Nothing   ' start empty; will assign array later

                ' Build a temporary list because Members is an array
                Dim relMembers As New List(Of CompleteRelationMember)()

                For Each m In rel.Members
                    If m.Type = OsmGeoType.Way Then
                        Dim cw = completeWays.First(Function(x) x.Id = CLng(m.Id))

                        relMembers.Add(New CompleteRelationMember() With {
            .Member = cw,
            .Role = m.Role
        })
                    End If
                Next

                ' Assign the array to the CompleteRelation
                compRel.Members = relMembers.ToArray()

                Debug.WriteLine("[Extract] CompleteRelation assembled")

                ' --- 6. Convert to polygon coordinates ---
                Dim coords As New List(Of MapPoint)

                For Each member In compRel.Members
                    Dim cw As CompleteWay = CType(member.Member, CompleteWay)
                    For Each n In cw.Nodes
                        coords.Add(New MapPoint(n.Longitude, n.Latitude, SpatialReferences.Wgs84))
                    Next
                Next

                Debug.WriteLine($"[Extract] Polygon assembled with {coords.Count} vertices")

                Return coords
            End Using

        End Function

    End Class

    Public Function ExtractBoundaryFromPbf(pbfPath As String, relationId As Long) As List(Of MapPoint)
        Debug.WriteLine("[Extract] Opening PBF: " & pbfPath)

        Using fs As FileStream = File.OpenRead(pbfPath)
            Dim source As New PBFOsmStreamSource(fs)
            Dim all As IEnumerable(Of OsmGeo) = source.ToList()

            Debug.WriteLine("[Extract] Loaded " & all.Count().ToString() & " OSM primitives")

            ' --- 1. Find the relation ---
            Dim rel As Relation = all.
                OfType(Of Relation)().
                FirstOrDefault(Function(r) r.Id = relationId)

            If rel Is Nothing Then
                Throw New Exception("Relation " & relationId & " not found in PBF.")
            End If

            Debug.WriteLine("[Extract] Found relation: " & rel.Id)

            ' --- 2. Collect referenced ways ---
            Dim wayIds As HashSet(Of Long) = rel.Members.
                Where(Function(m) m.Type = OsmGeoType.Way).
                Select(Function(m) CLng(m.Id)).
                ToHashSet()

            Dim ways As Dictionary(Of Long, Way) =
                all.OfType(Of Way)().
                Where(Function(w) wayIds.Contains(CLng(w.Id))).
                ToDictionary(Function(w) CLng(w.Id))

            Debug.WriteLine("[Extract] Loaded " & ways.Count.ToString() & " ways")

            ' --- 3. Collect referenced nodes ---
            Dim nodeIds As HashSet(Of Long) =
                ways.Values.SelectMany(Function(w) w.Nodes).ToHashSet()

            Dim nodes As Dictionary(Of Long, Node) =
                all.OfType(Of Node)().
                Where(Function(n) nodeIds.Contains(n.Id.Value)).
                ToDictionary(Function(n) n.Id.Value)

            Debug.WriteLine("[Extract] Loaded " & nodes.Count.ToString() & " nodes")

            ' --- 4. Build CompleteWays ---
            Dim completeWays As New List(Of CompleteWay)

            For Each w In ways.Values
                Dim cw As New CompleteWay()
                cw.Id = w.Id
                cw.Tags = w.Tags
                cw.Nodes = w.Nodes.Select(Function(nid) nodes(nid)).ToArray()
                completeWays.Add(cw)
            Next

            ' --- 5. Build CompleteRelation ---
            Dim compRel As New CompleteRelation()
            compRel.Id = rel.Id
            compRel.Tags = rel.Tags
            compRel.Members =
                    rel.Members.
                        Select(Function(m)
                                   If m.Type = OsmGeoType.Way Then
                                       Return New CompleteRelationMember() With {
                                           .Member = completeWays.First(Function(cw) cw.Id = m.Id),
                                           .Role = m.Role
                                       }
                                   Else
                                       Return Nothing
                                   End If
                               End Function).
                        Where(Function(x) x IsNot Nothing).
                        ToArray()


            Debug.WriteLine("[Extract] CompleteRelation built")

            ' --- 6. Convert to polygon coordinates ---
            Dim coords As New List(Of MapPoint)

            For Each member In compRel.Members
                Dim cw As CompleteWay = CType(member.Member, CompleteWay)
                For Each n In cw.Nodes
                    coords.Add(New MapPoint(n.Latitude, n.Longitude))
                Next
            Next

            Debug.WriteLine("[Extract] Polygon assembled with " & coords.Count.ToString() & " vertices")

            Return coords
        End Using
    End Function

End Module

