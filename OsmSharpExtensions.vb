Imports OsmSharp.Streams
Module OsmSharpExtensions
    <Runtime.CompilerServices.Extension>
    Public Function MoveNextAndReturn(source As PBFOsmStreamSource) As OsmSharp.OsmGeo
        If source.MoveNext() Then
            Return source.Current()
        End If
        Return Nothing
    End Function
End Module
