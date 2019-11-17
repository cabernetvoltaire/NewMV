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
    Public Sub FindOrphans2()
        'for each file in the orphan list
        'first look at all subfolders of original file destination
        'as most likely that it's been moved to a lower folder
        Dim foundparent As Boolean = False
        For Each n In mOrphanList
            Dim filename = FilenameFromLink(n) 'name of the file 
            'Dim aimfile As String = LinkTarget(n) 'non-existent file which is the linktarget
            Dim finfo As New IO.FileInfo(filename)
            filename = finfo.Name 'in case it was wrong?
            If finfo.Directory.Exists Then 'Not going to use this method if the original folder has been destroyed
                Dim dirs = finfo.Directory.EnumerateDirectories
                For Each subdir In dirs
                    If foundparent Then
                        Exit For
                    Else
                        Dim newlink = subdir.FullName & "\" & finfo.Name
                        If My.Computer.FileSystem.FileExists(newlink) Then
                            foundparent = True
                            If Not mFoundParents.Keys.Contains(n) Then
                                mFoundParents.Add(n, newlink)

                            End If

                        End If
                    End If

                Next
                foundparent = False
            End If
        Next
        For Each x In mFoundParents.Keys
            mOrphanList.Remove(x)
        Next
        'otherwise, have to do a complete search
        'for each directory, 
        Dim i As Integer = 0
        Dim max = DirectoriesList.Count
        While mFoundParents.Count < mOrphanList.Count And i < max
            For Each j In mOrphanList
                Dim filename = LinkTarget(j) 'name of the file 
                Dim spl = filename.Split("\")
                filename = spl(spl.Length - 1)
                If Len(filename) > 8 Then
                    Dim newlink = DirectoriesList(i) & "\" & filename
                    '          Report("Trying " & newlink, 0)
                    If My.Computer.FileSystem.FileExists(newlink) Then
                        If Not mFoundParents.Keys.Contains(j) Then
                            mFoundParents.Add(j, newlink)
                        End If
                    End If
                End If
            Next
            i += 1
            'TODO: What about files which occur in multiple places, or different files with the same name?
        End While
        'check each file
        'This takes number of directories x number of orphans at worst

    End Sub
    Public Sub FindOrphans()
        FindOrphans2()
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

    Public Sub ReuniteWithFile(list As List(Of String), filename As String)
        'All the shortcuts in list are redirected to filename
        mFoundParents.Clear()

        For Each m In list
            If LinkTargetExists(m) Then
            Else
                If mFoundParents.Keys.Contains(m) Then
                Else
                    mFoundParents.Add(m, filename)
                End If
            End If
        Next
        Reunite()
    End Sub
End Class
