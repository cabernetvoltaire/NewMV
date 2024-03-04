Imports System.IO
Module FolderSearchModule
    ' Function to collect all folders and subfolders starting from the root
    Public Function GetAllFolders(rootFolder As String) As List(Of String)
        Dim allFolders As New List(Of String)
        Try
            CollectFoldersRecursively(rootFolder, allFolders)
        Catch ex As Exception
            MsgBox("An error occurred: " & ex.Message)
        End Try
        Return allFolders

    End Function

    ' Recursive helper function to collect folders
    Private Sub CollectFoldersRecursively(folder As String, ByRef allFolders As List(Of String))
        Try
            ' Add the current folder to the list
            allFolders.Add(folder)

            ' Recursively collect all subfolders
            For Each subFolder As String In Directory.GetDirectories(folder)
                CollectFoldersRecursively(subFolder, allFolders)
            Next
        Catch ex As UnauthorizedAccessException
            ' Log the error or ignore it
            'MsgBox($"Access denied to {folder}. Skipping...")
        Catch ex As Exception
            ' Handle other types of exceptions if necessary
            MsgBox($"An error occurred accessing {folder}: {ex.Message}")
        End Try
    End Sub


    ' Function to filter folders based on the substring in their end folder name
    Public Function FilterFoldersBySubstringOld(folders As List(Of String), searchString As String) As List(Of String)
        Return folders.Where(Function(f) Path.GetFileName(f).IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0).ToList()
    End Function
    Public Function FilterFoldersBySubstring(folders As List(Of String), searchString As String) As List(Of String)
        Return folders.Where(Function(f) f.IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0).ToList()
    End Function
End Module


