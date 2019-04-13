Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As New ShortcutHandler
    Public Event FoundParent As EventHandler
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
    Public Sub FindOrphans(SearchPath As String)
        ' Dim FoundParent As Boolean
        '                Dim StartDest = New IO.DirectoryInfo(f.Directory.FullName)
        Dim StartDest = New IO.DirectoryInfo(SearchPath)

        While Not StartDest.Exists
            StartDest = StartDest.Parent
        End While
        'Check all the files in startdest, and put matches in Names
        ' If Split(StartDest.FullName, "\").Length > 2 Then
        Dim Names As New Dictionary(Of String, String)
            Dim FilesToCheck As IO.FileInfo() = StartDest.GetFiles("*", IO.SearchOption.AllDirectories)
            For Each xx In FilesToCheck
                If Not Names.ContainsKey(xx.Name) Then Names.Add(xx.Name, xx.FullName)
            Next
        'add names to foundparents
        For Each m In mOrphanList
            Dim s As String = LinkTarget(m)
            If s <> "" Then

                Dim f As New IO.FileInfo(s)
                If Names.ContainsKey(f.Name) Then
                    If mFoundParents.ContainsKey(Names(f.Name)) Then
                        mFoundParents.Add(Names(f.Name) & "#" & Format(Int(Rnd() * 1000000), "######"), m)
                    Else
                        mFoundParents.Add(Names(f.Name), m)
                    End If
                End If
            End If
        Next
        If mFoundParents.Count <> 0 Then
            RaiseEvent FoundParent(Me, Nothing)
        End If
        Reunite()
        If mFoundParents.Count = mOrphanList.Count Then

        End If

    End Sub


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
            mSHandler.ShortcutName = m.Value
            Dim s As String = Right(m.Key, 7)
            Dim mn As String = m.Key
            If s(0) = "#" Then
                mn = Replace(m.Key, s, "")
            End If
            Dim f As New IO.FileInfo(mn)
            If f.Exists = True Then mSHandler.ReAssign_ShortCutPath(mn, m.Value)
        Next
    End Sub
End Class
