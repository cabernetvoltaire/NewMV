Imports System.Windows.Media.Imaging
Imports System.Windows.Media.MetadataLibrary

'Module MetadataTagger

'    Public Sub WriteTagsToMp4(ByVal filePath As String, ByVal title As String, ByVal artist As String, ByVal album As String, ByVal year As String, ByVal genre As String, ByVal trackNumber As Integer, ByVal coverImagePath As String)

'        ' Create a BitmapMetadata object and set the metadata values
'        Dim metadata As New BitmapMetadata("png")
'        metadata.Title = title
'        metadata.Author = artist
'        metadata.DateTaken = year
'        metadata.Subject = genre
'        metadata.AlbumTitle = album
'        metadata.TrackNumber = trackNumber

'        ' Add the cover image to the metadata
'        If coverImagePath <> "" AndAlso IO.File.Exists(coverImagePath) Then
'            Dim coverImage As New BitmapImage(New Uri(coverImagePath))
'            metadata.SetQuery("/xmp/dc:description[1]", coverImage)
'        End If

'        ' Open the MP4 file for writing
'        Using fs As New IO.FileStream(filePath, IO.FileMode.Open, IO.FileAccess.ReadWrite)

'            ' Create a MetadataEditor object and load the metadata from the MP4 file
'            Dim editor As New MetadataEditor(fs)
'            editor.Load()

'            ' Set the metadata values
'            editor.Title = metadata.Title
'            editor.Artist = metadata.Author
'            editor.Year = metadata.DateTaken
'            editor.Genre = metadata.Subject
'            editor.Album = metadata.AlbumTitle
'            editor.TrackNumber = metadata.TrackNumber

'            ' Save the metadata to the MP4 file
'            editor.Save()

'        End Using

'    End Sub

'End Module
