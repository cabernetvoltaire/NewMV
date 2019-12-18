Imports System.IO
Imports System.Threading
Public Class FileHandler
    Public FilterState As FilterHandler
    Public Random As RandomHandler
    Public NMS As StateHandler
    Public Listbox As ListBox
    Public Media As New MediaHandler("Media")
    Public MSFiles As New MediaSwapper
    Public Allfaveminder As New FavouritesMinder("Q:\Favourites")
    Public Faveminder As New FavouritesMinder("Q:\Favourites")

#Region "Event Handlers"

#End Region

#Region "Methods"
    Public Sub MoveFolder(SourcePath As String, DestPath As String)
        Dim TargetFolder As New DirectoryInfo(DestPath)
        Dim SourceFolder As New DirectoryInfo(SourcePath)

        MoveDirectoryContents(TargetFolder, SourceFolder, SourceFolder, True)
        For Each d In SourceFolder.EnumerateDirectories("*", SearchOption.AllDirectories)
            MoveDirectoryContents(TargetFolder, SourceFolder, d, True)
        Next
        DirectoriesList.Remove(SourcePath)
        DirectoriesList.Add(DestPath)
    End Sub
    Private Sub MoveDirectoryContents(TargetDir As DirectoryInfo, SourceDir As DirectoryInfo, d As DirectoryInfo, Optional Parent As Boolean = False)
        Dim flist As New List(Of String)
        Dim newpath As String = d.FullName
        If Parent Then
            newpath = newpath.Replace(SourceDir.Parent.FullName, TargetDir.FullName)
        Else
            newpath = newpath.Replace(SourceDir.FullName, TargetDir.FullName)

        End If
        Dim NewDir As New DirectoryInfo(newpath)
        If NewDir.Exists = False Then
            NewDir.Create()
        End If
        For Each f In d.EnumerateFiles("*", SearchOption.TopDirectoryOnly)
            flist.Add(f.FullName)
        Next
        blnSuppressCreate = True
        'MoveFiles(flist, newpath, Listbox, True)

    End Sub

    Public Sub MoveFiles(files As List(Of String), DestPath As String, Optional Folder As Boolean = False)
    End Sub

#End Region

#Region "Properties"

#End Region

End Class
