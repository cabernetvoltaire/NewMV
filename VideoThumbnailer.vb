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
            System.Threading.Thread.Sleep(10)
        End While
        Return Thumbnail
    End Function

    Private Sub ThumbnailProcess(finfo As IO.FileInfo, Filename As String, Frame As Long, ss As String)
        Dim pInfo As New ProcessStartInfo()
        pInfo.FileName = "C:ffmpeg.exe"
        pInfo.WindowStyle = ProcessWindowStyle.Hidden

        Thumbnail = """Q:\Thumbs\" & ss & "thn.png"""
        ' Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -y -an -codec " & finfo.Extension & " -frames:v 1 -f image2 " & Thumbnail
        Dim size As String = Str(ThumbnailHeight) & ":" & Str(ThumbnailHeight / 2)
        size = Str(ThumbnailHeight)
        size.Replace(" ", "")
        ' Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -y " & " -frames:v 1 -f image2 -vf scale=" & size & ":force_original_aspect_ratio=increase,crop=" & size & " " & Thumbnail
        'Dim s As String = "-ss " & Str(30) & " -i """ & Filename & """ -vf scale='200:100' -y " & " -frames:v 1 -f image2 " & " " & Thumbnail
        Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -vf scale='192:-1' -y " & " -frames:v 1 -f image2 " & " " & Thumbnail
        'Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -y " & " -frames:v 1 -f image2 " & Thumbnail

        pInfo.Arguments = s
        Thumbnail = "Q:\Thumbs\" & ss & "thn.png"

        p = Process.Start(pInfo)
    End Sub

    Sub ThumbnailCompleted() Handles p.Exited
        RaiseEvent Thumbnailed(Thumbnail)
    End Sub
End Class
