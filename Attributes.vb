Public Class Attributes
    Private Choices() As String = {"Rate", "Height", "Width", "Date", "Album", "Song"}
    Private Sub Attributes_Load(sender As Object, e As EventArgs) Handles MyBase.Load, MyBase.Shown
    End Sub
    Private mDestinationLabel As Label
    Public Property DestinationLabel() As Label
        Get
            Return mDestinationLabel
        End Get
        Set(ByVal value As Label)
            mDestinationLabel = value
        End Set
    End Property

    Public Sub UpdateLabel(filename As String)
        'Exit Sub
        Dim AttributeName As String = ""
        Dim AttList As List(Of KeyValuePair(Of String, String)) = FileMetaData.GetMetaText(filename, "C:\exiftool.exe")
        If AttList IsNot Nothing Then

            For Each f In AttList
                If ChosenString(f.Key) Then
                    AttributeName += f.Key & ": " & f.Value & vbCrLf
                End If
            Next
        End If
        mDestinationLabel.Text = AttributeName
    End Sub
    Public Sub AlbumArtist(filename As String)
        'Make folder 'Artist\Album'
        'Move file to it. 
        Dim f As New IO.FileInfo(filename)
        Dim artist As String = ""
        Dim album As String = ""
        Dim title As String = ""

        Dim AttList As List(Of KeyValuePair(Of String, String)) = FileMetaData.GetMetaText(filename, "C:\exiftool.exe")
        For Each m In AttList
            If m.Key = "Artist" Then artist = m.Value
            If m.Key = "Album" Then album = m.Value
            If m.Key = "Title" Then title = m.Value
        Next
        Dim dest As String
        Dim dir As New IO.DirectoryInfo(f.Directory.Root.FullName & artist & "\" & album)
        If Not dir.Exists Then
            dir.Create()
        End If
        If title <> "" Then
            dest = f.Directory.Root.FullName & artist & "\" & album & "\" & title & f.Extension
        Else
            dest = f.Directory.Root.FullName & artist & "\" & album & "\" & title & f.Name

        End If
        Dim r As New IO.FileInfo(dest)
        If r.Exists Then
            Deletefile(r.FullName)
        Else
            f.MoveTo(dest)
        End If
        ActiveForm.Refresh()
    End Sub
    Private Function ChosenString(s As String) As Boolean
        Dim Flag As Boolean = False
        For i = 0 To 5
            If Not Flag Then Flag = s.Contains(Choices(i))

        Next
        Return Flag
    End Function


End Class


Public NotInheritable Class FileMetaData
    Public Delegate Function GetMetaTextDelegate(ByVal filePath As String, ByVal exifToolsExePath As String) As List(Of KeyValuePair(Of String, String))

    Public Shared Function GetMetaText(ByVal filePath As String,
                    ByVal exifToolsExePath As String) As List(Of KeyValuePair(Of String, String))

        Dim retVal As List(Of KeyValuePair(Of String, String)) = Nothing

        Try
            If String.IsNullOrWhiteSpace(exifToolsExePath) Then
                Throw New ArgumentException("The file path of the EXIF Tools executable" & vbCrLf &
                                            "cannot be null or empty.")

            ElseIf Not My.Computer.FileSystem.FileExists(exifToolsExePath) Then
                Throw New IO.FileNotFoundException("The file path of the EXIF Tools executable" & vbCrLf &
                                                   "could not be located.")

            ElseIf String.IsNullOrWhiteSpace(filePath) Then
                Throw New ArgumentException("The file path of the file to be examined" & vbCrLf &
                                            "cannot be null or empty.")

            ElseIf Not My.Computer.FileSystem.FileExists(filePath) Then
                Throw New IO.FileNotFoundException("The file path of the file to be examined" & vbCrLf &
                                                   "could not be located.")

            Else
                Using tool As New Process
                    Dim sb As New System.Text.StringBuilder

                    With tool
                        With .StartInfo
                            .FileName = exifToolsExePath
                            .Arguments = String.Format(" -s  {0}{1}{0}", Chr(34), filePath)
                            .UseShellExecute = False
                            .RedirectStandardOutput = True
                            .CreateNoWindow = True
                        End With

                        .Start()

                        sb.Append(tool.StandardOutput.ReadToEnd)
                        sb.Remove(sb.Length - 2, 2)

                        Dim pairs() As String = sb.ToString.Split(CChar(vbCrLf))
                        Dim delimiters() As String = New String() {" : "}

                        If retVal Is Nothing Then
                            retVal = New List(Of KeyValuePair(Of String, String))
                        End If

                        For Each element As String In pairs
                            Dim kv() As String = element.Split(delimiters, StringSplitOptions.RemoveEmptyEntries)

                            If kv.Length = 2 Then
                                retVal.Add(New KeyValuePair(Of String, String)(kv(0).Trim, kv(1).Trim))
                            End If
                        Next
                    End With
                End Using
            End If

        Catch ex As Exception
            'Throw
        End Try

        Return retVal

    End Function

End Class