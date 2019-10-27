Imports System.Threading
Public Class VideoThumbnailer
    Private WithEvents p As Process
    Property Fileref As String
    Property ThumbnailHeight As Integer
    Property Thumbnail As String
    Event Thumbnailed(s As String)
    ''' <summary>
    ''' Creates a Thumbnail of Filename in Q:\ taken from Frame
    ''' Raises an event containing the filename of the Thumbnail when done.
    ''' </summary>
    ''' <param name="Filename"></param>
    ''' <param name="Frame"></param>
    ''' <returns></returns>
    Function GetThumbnail(Filename As String, Frame As Long) As String
        ' Exit Function
        ' Generate thumbnail
        Dim finfo As New IO.FileInfo(Filename)

        Fileref = Filename
        'Dim shortname() As String = Split(Filename, "\")
        Dim ss As String ' = shortname(shortname.Length - 1)
        ss = finfo.Name
        ss.Replace(".", "")

        ThumbnailProcess(finfo, Filename, Frame, ss)

        While (Not p.HasExited)
            Thread.Sleep(10)
        End While
        Return Thumbnail
    End Function

    Private Sub ThumbnailProcess(finfo As IO.FileInfo, Filename As String, Frame As Long, ss As String)
        Dim pInfo As New ProcessStartInfo()
        pInfo.FileName = "C:ffmpeg.exe"
        pInfo.WindowStyle = ProcessWindowStyle.Hidden

        Thumbnail = ThumbnailName(ss)

        '        Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -vf scale='192:-1' -y " & " -frames:v 1 -f image2 " & " " & Thumbnail

        Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -vf scale='125:-1' -y " & " -frames:v 1 -f image2 " & " """ & Thumbnail & """"
        pInfo.Arguments = s
        p = Process.Start(pInfo)
    End Sub

    Sub ThumbnailCompleted() Handles p.Exited
        RaiseEvent Thumbnailed(Thumbnail)
    End Sub
End Class
