Public Class MyFileInfo
    Public Property DestinationPath As String
    Public Property FilePath As String
    Public Property Bookmarks As New List(Of Long)
    Public Property Tags As New List(Of String)
    Public Sub Save()
        'Dim writer As New System.Xml.Serialization.XmlSerializer(GetType(FileInfo))
        'Dim file As New System.IO.StreamWriter(
        '    "c:\temp\SerializationOverview.xml")
        'writer.Serialize(file, overview)
        'file.Close()
    End Sub
    Public Sub Load()

    End Sub

End Class
