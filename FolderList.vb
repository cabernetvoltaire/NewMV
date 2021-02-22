Public Class FolderList
    Private _FolderPath As String

    Private _FolderList As New Stack(Of String)
    Public Property Depth
    Public Property FolderPath As String
        Get
            Return _FolderPath
        End Get
        Set
            _FolderPath = Value

        End Set
    End Property

    Public ReadOnly Property FolderList As Stack(Of String)
        Get
            SubFolders(_FolderPath)
            Return _FolderList
        End Get
    End Property


    Private Sub SubFolders(Path As String)
        Dim d As New IO.DirectoryInfo(Path)
        Dim parts() As String = Path.Split("\")
        Dim n As Integer = parts.Length
        If n <= Depth Then
            _FolderList.Push(Path)
            For Each m In d.EnumerateDirectories()
                SubFolders(m.FullName)
            Next
        Else
            Exit Sub
        End If

    End Sub
End Class
