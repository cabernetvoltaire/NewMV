Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As New ShortcutHandler
    Public Event FoundParent As EventHandler
    Public Event NewOrphanTesting As EventHandler
    Public Event ParentNotFound As EventHandler
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

    Private mOrphanList As New List(Of String)
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
    ''' <summary>
    ''' Gets a list of all possible folders to be searched
    ''' </summary>

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
            Dim foundparent As Boolean = False
            Dim filename As String = FilenameFromLink(n)
            Dim aimfile As String = LinkTarget(n)
            Dim newlink = DirectoriesList(0) & "\" & filename
            If Not mFoundParents.ContainsKey(n) Then
                ' RaiseEvent NewOrphanTesting(Me, Nothing)
                Dim i As Integer = 0
                Dim max = DirectoriesList.Count
                While Not foundparent
                    If Len(aimfile) > 8 Then

                        'search sub folders first
                        Dim finfo As New IO.FileInfo(aimfile)
                        Try
                            For Each m In finfo.Directory.EnumerateDirectories
                                If foundparent Then
                                    Exit For
                                Else
                                    newlink = m.FullName & "\" & finfo.Name
                                    foundparent = My.Computer.FileSystem.FileExists(newlink)

                                End If
                            Next

                        Catch ex As Exception

                        End Try
                    End If
                End While

                While Not foundparent And i < max
                    If Len(filename) > 8 Then
                        newlink = DirectoriesList(i) & "\" & filename
                        foundparent = My.Computer.FileSystem.FileExists(newlink)
                    End If
                    i += 1
                    'TODO: What about files which occur in multiple places, or different files with the same name?
                End While
                If foundparent Then
                    mFoundParents.Add(n, newlink)
                Else
                    Dim s = filename & " was not found anywhere inside the hierarchy " & Rootpath
                    MsgBox(s)
                    RaiseEvent ParentNotFound(Me, Nothing)
                End If
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
    ''' <summary>
    ''' Creates new links from the linklist, unless they already exist. 
    ''' </summary>
    ''' <param name="LinkList"></param>
    Public Sub ImportLinks(LinkList As List(Of String))
        mOrphanList.Clear()

        For Each m In LinkList
            Dim x As New IO.FileInfo(m)
            If x.Exists Then
            Else
                mOrphanList.Add(m)
            End If
        Next
        FindOrphans()
    End Sub
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
