Public Class DirectoryLister
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
    Public Sub GenerateDirsOld()
        Dim list As New List(Of String)
        Dim nlist As New List(Of String)
        list = FindDirs(_DirectoryPath, IO.SearchOption.TopDirectoryOnly)
        While _DepthCounter > 0
            '_DirList.AddRange(list)
            For i = 0 To list.Count - 1
                nlist.AddRange(FindDirs(list(i), IO.SearchOption.TopDirectoryOnly))
            Next
            list = nlist
            nlist.Clear()
            _DepthCounter -= 1
        End While
        _DirList = list
    End Sub
    Public Sub GenerateDirs()
        _DirList = FindDirs(_DirectoryPath, IO.SearchOption.AllDirectories)

    End Sub

    ''' <summary>
    ''' Returns a list of safe directories under string path s
    ''' </summary>
    ''' <param name="s"></param>
    ''' <returns></returns>
    Private Function FindDirs(s As String, choice As IO.SearchOption) As List(Of String)
        Dim founddirs As New List(Of String)
        Dim dir As New IO.DirectoryInfo(s)
        For Each m In dir.EnumerateDirectories("*", choice)

            If (((m.Attributes And System.IO.FileAttributes.Hidden) = 0) AndAlso
                                ((m.Attributes And System.IO.FileAttributes.System) = 0)) Then
                founddirs.Add(m.FullName)

            End If
        Next
        Return founddirs
    End Function
End Class
