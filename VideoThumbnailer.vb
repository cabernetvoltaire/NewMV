﻿Imports System.Threading
Public Class VideoThumbnailer
    Private WithEvents p As Process
    Property Fileref As String
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

        Thumbnail = """Q:\Thumbs\" & ss & "thn.jpeg"""
        ' Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -y -an -codec " & finfo.Extension & " -frames:v 1 -f image2 " & Thumbnail
        Dim s As String = "-ss " & Str(Frame) & " -i """ & Filename & """ -y -an " & " -frames:v 1 -f image2 " & Thumbnail
        pInfo.Arguments = s
        Thumbnail = "Q:\Thumbs\" & ss & "thn.jpeg"

        p = Process.Start(pInfo)
    End Sub

    Sub ThumbnailCompleted() Handles p.Exited
        RaiseEvent Thumbnailed(Thumbnail)
    End Sub
End Class
