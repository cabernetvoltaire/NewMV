﻿Public Class DirectoryLister
    Private _DirectoryPath As String
    Public Property Path() As String
        Get
            Return _DirectoryPath
        End Get
        Set(ByVal value As String)
            _DirectoryPath = value
        End Set
    End Property
    Private _Depth As String
    Public Property Depth() As String
        Get
            Return _Depth
        End Get
        Set(ByVal value As String)
            _Depth = value
            _DepthCounter = _Depth
        End Set
    End Property
    Private _DirList As New List(Of String)
    Public Property DirectoryList() As List(Of String)
        Get
            Return _DirList
        End Get
        Set(ByVal value As List(Of String))
            _DirList = value
        End Set
    End Property
    Private _DepthCounter As Integer
    Public Sub GenerateDirs()
        Dim list As New List(Of String)
        Dim nlist As New List(Of String)
        list = FindDirs(_DirectoryPath)
        While _DepthCounter > 0
            _DirList.AddRange(list)
            For i = 0 To list.Count - 1
                nlist.AddRange(FindDirs(list(i)))
            Next
            list = nlist
            _DepthCounter -= 1
        End While
    End Sub
    Private Function FindDirs(s As String) As List(Of String)
        Dim founddirs As New List(Of String)
        Dim dir As New IO.DirectoryInfo(s)
        For Each m In dir.EnumerateDirectories("*", IO.SearchOption.TopDirectoryOnly)

            If (((m.Attributes And System.IO.FileAttributes.Hidden) = 0) AndAlso
                                ((m.Attributes And System.IO.FileAttributes.System) = 0)) Then
                founddirs.Add(m.FullName)
            End If
        Next
        Return founddirs
    End Function
End Class