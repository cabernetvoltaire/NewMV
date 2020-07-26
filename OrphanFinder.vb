Imports System.Threading
Imports System.IO
Public Class OrphanFinder
    Private mFolderPath As String
    Private mSHandler As New ShortcutHandler
    Public DB As Database
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
                If LinkTarget(f) <> "" Then mOrphanTargetPairs.Add(f, TargetFileFromLinkName(f))
            Next
        End Set
    End Property

    Private Function TargetFileFromLinkName(str As String) As String
        Dim parts() = Split(Path.GetFileName(str), "%")
        Return parts(0)
        Exit Function
    End Function

    Private mOrphanTargetPairs As New Dictionary(Of String, String)
    ''' <summary>
    ''' Dictionary containing (new parent,link) pairs
    ''' </summary>
    Private mFoundParents As New Dictionary(Of String, String)



    Public Function FindOrphans(Optional Deep As Boolean = False) As Boolean
        Orphanfinder()
        Reunite()
        MsgBox("Reunite finished - " & mFoundParents.Count & " files reunited")
        If mFoundParents.Count > 0 Then
            Return True
        Else
            Return False
        End If
        mFoundParents.Clear()
    End Function

    Public Sub Orphanfinder()

        Dim mNewTargets As New Dictionary(Of String, String)
        For Each p In mOrphanTargetPairs
            Dim file As String = Path.GetFileNameWithoutExtension(p.Value)
            Dim x As New DatabaseEntry
            x = DB.FindEntry(file)
            If x IsNot Nothing Then
                mNewTargets.Add(p.Key, x.FullPath)
            End If
        Next
        mFoundParents = mNewTargets

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
            If f.Exists = True Then mSHandler.ReAssign(m.Value, mn)
        Next

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
