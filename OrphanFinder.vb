Imports System.Threading
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
            For Each f In mOrphanList
                If LinkTarget(f) <> "" Then mOrphanTargetPairs.Add(f, LinkTarget(f))
            Next
        End Set
    End Property
    Private mOrphanTargetPairs As New Dictionary(Of String, String)
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
    Public Sub FindOrphans2(Optional Deep As Boolean = False)
        'for each file in the orphan list
        'first look at all subfolders of original file destination
        'as most likely that it's been moved to a lower folder
        Dim i As Integer = 0
        If Deep Then
            'Check those.
            i = DeepSearch()
            Exit Sub
        End If
        Dim foundparent As Boolean = False
        foundparent = FindBracketed(foundparent)

        For Each x In mFoundParents.Keys
            mOrphanList.Remove(x)
        Next

        ' foundparent = FindInSubFolders(foundparent)

        'For Each x In mFoundParents.Keys
        'mOrphanList.Remove(x)
        'Next

        'Alternative
        'Get parent folder name
        i = FindInNamedFolders(i)

        For Each x In mFoundParents.Keys
            mOrphanList.Remove(x)
        Next

        'check each file
        'This takes number of directories x number of orphans at worst

    End Sub

    Private Function DeepSearch() As Integer

        Dim i As Integer = 0
        'otherwise, have to do a complete search
        'for each directory, 
        Dim max = DirectoriesList.Count
        While mFoundParents.Count < mOrphanTargetPairs.Count And i < max
            Dim filename As String

            For Each j In mOrphanTargetPairs

                filename = j.Value 'name of the file 
                filename = FilenameFromPath(filename, True)
                Dim newlink As String
                If SearchTreeForFile("Q:\Watch\", filename, newlink) = 1 Then

                    'If My.Computer.FileSystem.FileExists(newlink) Then
                    If Not mFoundParents.Keys.Contains(j.Key) Then
                        mFoundParents.Add(j.Key, newlink)
                        If mFoundParents.Count > 10 Then Reunite()
                    End If

                End If

            Next
            i += 1
            'TODO: What about files which occur in multiple places, or different files with the same name?
        End While

        Return i
    End Function
    Private Function DeepSearchOld() As Integer

        Dim i As Integer = 0
        'otherwise, have to do a complete search
        'for each directory, 
        Dim max = DirectoriesList.Count
        While mFoundParents.Count < mOrphanTargetPairs.Count And i < max
            Dim filename As String

            For Each j In mOrphanTargetPairs

                filename = j.Value 'name of the file 

                Dim spl = filename.Split("\")
                filename = spl(spl.Length - 1)
                If Len(filename) > 8 Then
                    Dim newlink = DirectoriesList(i) & "\" & filename
                    Report("Trying " & newlink, 0)
                    If SearchTreeForFile("Q:\Watch\", filename, newlink) = 1 Then

                        'If My.Computer.FileSystem.FileExists(newlink) Then
                        If Not mFoundParents.Keys.Contains(j.Key) Then
                            mFoundParents.Add(j.Key, newlink)
                            If mFoundParents.Count > 10 Then Reunite()
                        End If

                    End If

                End If
            Next
            i += 1
            'TODO: What about files which occur in multiple places, or different files with the same name?
        End While

        Return i
    End Function
    Private Function FindInNamedFolders(i As Integer) As Integer
        For Each n In mOrphanTargetPairs
            If n.Value = "" Then
                Continue For
            Else
                Dim tgt As New IO.DirectoryInfo(n.Value)
                Dim found As Boolean = False
                While Not found And tgt.Parent.Name <> tgt.Root.FullName
                    Dim parentname = tgt.Parent.Name
                    Dim searchdir As New List(Of String)
                    searchdir = DirectoriesList.FindAll(Function(x) x.Contains(parentname))
                    Dim filename = FilenameFromPath(n.Value, True)
                    While Not found And i < searchdir.Count
                        Dim newtarget = searchdir(i) & "\" & filename
                        found = FindTarget(n.Key, found, newtarget)
                        i += 1
                    End While
                    i = 0
                    If tgt.Parent.Name <> "Watch" Then

                        tgt = tgt.Parent
                    Else
                        Exit While
                    End If
                End While
            End If
            i = 0
        Next
        Return i
    End Function
    Private Function FindInNamedFoldersOld(i As Integer) As Integer
        For Each f In mOrphanList
            Dim fname As String = LinkTarget(f)
            If fname <> "" Then

                Dim tgt As New IO.DirectoryInfo(fname)
                Dim found As Boolean = False
                While Not found And tgt.Parent.Name <> tgt.Root.FullName
                    Dim parentname As String = tgt.Name

                    'Search Directories list for path containing folder name
                    Dim searchdir As New List(Of String)
                    searchdir = DirectoriesList.FindAll(Function(x) x.Contains(parentname))
                    Dim filename = LinkTarget(f) 'name of the file 
                    Dim spl = filename.Split("\")
                    filename = spl(spl.Length - 1)
                    While Not found And i < searchdir.Count
                        Dim newtarget = searchdir(i) & "\" & filename
                        For j = 0 To 3
                            found = FindTarget(f, found, newtarget)
                            If found Then
                                Exit For
                            Else
                                newtarget = newtarget.Replace(".", "(" & Right(Str(j), 1) & ").")

                            End If
                        Next
                        i += 1
                    End While
                    i = 0
                    'Failed so use parent folder as search name
                    If tgt.Parent.Name <> "Watch" Then
                        tgt = tgt.Parent
                    Else
                        Exit While
                    End If

                End While
            End If
            i = 0
        Next

        Return i
    End Function

    Private Function FindTarget(f As String, found As Boolean, newtarget As String) As Boolean
        If My.Computer.FileSystem.FileExists(newtarget) Then
            If Not mFoundParents.Keys.Contains(f) Then
                mFoundParents.Add(f, newtarget)
                found = True
            End If

        End If

        Return found
    End Function
    Private Function FindInSubFolders(foundparent As Boolean) As Boolean
        For Each n In mOrphanTargetPairs
            If n.Value <> "" Then
                Dim aimfile As New IO.FileInfo(n.Value)
                If aimfile.Directory.Exists Then
                    Dim dirs = GenerateSafeFolderList(aimfile.Directory.FullName)
                    For Each dirname In dirs
                        If foundparent Then
                            Exit For
                        Else
                            Dim newtarget = dirname & "\" & aimfile.Name
                            If My.Computer.FileSystem.FileExists(newtarget) Then
                                foundparent = True
                                If Not mFoundParents.Keys.Contains(n.Key) Then
                                    mFoundParents.Add(n.Key, newtarget)

                                End If
                            Else
                                foundparent = False
                            End If

                        End If
                    Next
                End If
            End If
        Next
        Return foundparent
    End Function

    Private Function FindInSubFoldersold(foundparent As Boolean) As Boolean
        For Each n In mOrphanList
            Dim filename = FilenameFromLink(n) 'name of the file 
            Dim aimfile As String = LinkTarget(n) 'non-existent file which is the linktarget
            If aimfile <> "" And aimfile <> "C:\exiftools.exe" Then

                Dim finfo As New IO.FileInfo(aimfile)
                filename = finfo.Name 'in case it was wrong?
                If finfo.Directory.Exists Then 'Not going to use this method if the original folder has been destroyed
                    Dim dirs = GenerateSafeFolderList(FolderPathFromPath(aimfile))
                    '                    Dim dirs = finfo.Directory.EnumerateDirectories("*", IO.SearchOption.AllDirectories)
                    For Each subdir In dirs
                        If foundparent Then
                            Exit For
                        Else
                            Dim subd As New IO.DirectoryInfo(subdir)
                            Dim newlink = subd.FullName & "\" & finfo.Name
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
            End If
        Next

        Return foundparent
    End Function
    Private Function FindBracketed(foundparent As Boolean) As Boolean
        For Each n In mOrphanTargetPairs
            Dim finfo As New IO.FileInfo(n.Value)
            If finfo.Extension = "" Then Exit For
            If n.Value <> "" Then
                For i = 0 To 3

                    Dim newtarget = finfo.FullName.Replace(finfo.Extension, "(" & Right(Str(i), 1) & ")" & finfo.Extension)
                    If My.Computer.FileSystem.FileExists(newtarget) Then
                        foundparent = True
                        If Not mFoundParents.Keys.Contains(n.Key) Then
                            mFoundParents.Add(n.Key, newtarget)
                        End If
                    Else
                        foundparent = False
                    End If
                Next
            End If
        Next
        Return foundparent
    End Function

    Public Function FindOrphans(Optional Deep As Boolean = False) As Boolean
        FindOrphans2(Deep)
        Reunite()
        MsgBox("Reunite finished - " & mFoundParents.Count & " files reunited")
        If mFoundParents.Count > 0 Then
            Return True
        Else
            Return False
        End If
        mFoundParents.Clear()
    End Function


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
        For Each m In mFoundParents.ToList
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
        ' mFoundParents.Clear()
    End Sub
    Public Function ListOfDeadFiles() As List(Of String)
        Dim list As New List(Of String)
        For Each m In mOrphanList
            Dim target As String = LinkTarget(m)
            If Not list.Contains(target) Then
                list.Add(target)
            End If

        Next
        Return list
    End Function
    Public Sub ListReuniter(ListofFiles As List(Of String))
        For Each f In ListofFiles
            ReuniteWithFile(AllFaveMinder.GetLinksOf(f), f)
        Next
    End Sub
    Public Sub ReuniteWithFile(ListofLinkFiles As List(Of String), filename As String) 'TODO: Needs to extract filename from list element
        'All the shortcuts in list are redirected to filename
        mFoundParents.Clear()
        Try

            For Each linkfile In ListofLinkFiles
                If LinkTargetExists(linkfile) Then
                Else

                    If mFoundParents.Keys.Contains(linkfile) Then
                    Else
                        mFoundParents.Add(linkfile, filename)
                    End If
                End If
            Next
        Catch ex As Exception

        End Try

        Reunite()
    End Sub
End Class
