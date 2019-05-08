Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As New ShortcutHandler
    Public Event FoundParent As EventHandler
    Public Event NewOrphanTesting As EventHandler
    Public TotalFoundCount As Integer
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            mFolderPath = value
        End Set
    End Property

    Private mOrphanList As List(Of String)
    ''' <summary>
    ''' List of potential orphans
    ''' </summary>
    ''' <returns></returns>
    Public Property OrphanList() As List(Of String)
        Get
            Return mOrphanList
        End Get
        Set(ByVal value As List(Of String))
            mOrphanList = value
        End Set
    End Property
    Private mPathList As List(Of String)
    Public WriteOnly Property PathList(root As String) As List(Of String)
        Set(ByVal value As List(Of String))
            If value IsNot mPathList Then
                GetDirectoriesList(root)
                mPathList = value
            End If
        End Set
    End Property
    ''' <summary>
    ''' Dictionary containing (new parent,link) pairs
    ''' </summary>
    Private mFoundParents As New Dictionary(Of String, String)



    Public Property FoundParents() As Dictionary(Of String, String)
        Get
            Return mFoundParents
        End Get
        Set(ByVal value As Dictionary(Of String, String))
            mFoundParents = value
        End Set
    End Property
    ''' <summary>
    ''' Searches in subfolders of original destination of each link
    ''' </summary>
    Public Property SearchPath As String
    Public Sub FindOrphans()
        'for each file in orphanlist
        'change target to all possible directories.
        For Each n In mOrphanList
            RaiseEvent NewOrphanTesting(Me, Nothing)
            Dim filename As String = FilenameFromLink(n)
            Dim i As Integer = 0
            Dim max = DirectoriesList.Count
            Dim newlink = DirectoriesList(0) & "\" & filename
            While My.Computer.FileSystem.FileExists(newlink) = False And i < max
                If Len(filename) > 8 Then
                    newlink = DirectoriesList(i) & "\" & filename
                End If
                i += 1
                'TODO: What about files which occur in multiple places, or different files with the same name?
            End While
            If My.Computer.FileSystem.FileExists(newlink) Then
                mFoundParents.Add(n, newlink)
            End If
        Next
        If mFoundParents.Count <> 0 Then
            RaiseEvent FoundParent(Me, Nothing)
        End If
        Reunite()
    End Sub


    'Find Orphans
    'With orphan list
    'Get all putative destinations
    '


    Private Function CheckListForParents(ByVal list As List(Of String), f As IO.FileInfo) As Boolean

        For Each x In list
            If x.Contains(f.Name.Replace(f.Extension, "")) Then
                Return True
                Exit Function
            Else
                Return False
            End If
        Next
        Return False
    End Function

    Public Sub Reunite()
        TotalFoundCount = TotalFoundCount + mFoundParents.Count
        For Each m In mFoundParents
            mSHandler.ShortcutName = m.Key
            mSHandler.MarkOffset = 0
            Dim s As String = Right(m.Key, 7)
            Dim mn As String = m.Key
            If s(0) = "#" Then
                mn = Replace(m.Key, s, "")
            End If
            Dim f As New IO.FileInfo(m.Value)
            If f.Exists = True Then mSHandler.ReAssign_ShortCutPath(m.Value, mn)
        Next
    End Sub
End Class
