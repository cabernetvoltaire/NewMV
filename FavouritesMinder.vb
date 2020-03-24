Public Class FavouritesMinder
    Private mFavesList As New List(Of String)
    Public Property FavesList() As List(Of String)
        Get
            Return mFavesList
        End Get
        Set(ByVal value As List(Of String))
            mFavesList = value
        End Set
    End Property
    ''' <summary>
    ''' Where the favourite will be moved to
    ''' </summary>
    Private mDestPath As String

    Public Property DestinationPath() As String
        Get
            Return mDestPath
        End Get
        Set(ByVal value As String)
            mDestPath = value
        End Set

    End Property
    ''' <summary>
    ''' Finds all favourites pointing at path and deletes them, removing them from the FavesList
    ''' </summary>
    ''' <param name="path"></param>
    Public Sub DeleteFavourite(path As String)
        Dim templist As New List(Of String)

        templist = mFavesList.FindAll(Function(x) LinkTarget(x) = path)
        For Each m In templist
            Dim f As New IO.FileInfo(m)
            Deletefile(f.FullName)
            mFavesList.Remove(m)
        Next
    End Sub
    Public Sub New(path As String)
        NewPath(path)
    End Sub
    Public Sub NewPath(path As String)
        mFavesList.Clear()
        Dim faves As New IO.DirectoryInfo(path)
        If Not faves.Exists Then faves = IO.Directory.CreateDirectory(Environment.SpecialFolder.MyPictures & "\Favourites")
        For Each f In faves.GetFiles("*.lnk", IO.SearchOption.AllDirectories)

            mFavesList.Add(f.FullName)
        Next
    End Sub

    Public Sub CheckFile(f As IO.FileInfo)
        If f.Extension = ".lnk" Then
        Else
            Check(f, mDestPath)

        End If
    End Sub
    ''' <summary>
    ''' Runs through all files in f, and redirects their targets
    ''' </summary>
    ''' <param name="f"></param>
    Public Sub CheckFiles(f As List(Of String))
        For Each m In f
            Dim k As New IO.FileInfo(m)
            Check(k, mDestPath)
        Next
    End Sub

    Public Sub RedirectShortCutList(f As IO.FileInfo, shortcuts As List(Of String))
        ReallocateShortCuts(f, shortcuts)
    End Sub

    Private Sub ReallocateShortCut(f As IO.FileInfo, m As String, bk As Long, minfo As IO.FileInfo)
        Dim sch As New ShortcutHandler(f.FullName, minfo.DirectoryName, minfo.Name)
        sch.MarkOffset = 0
        Dim fn As String = sch.Create_ShortCut(bk)
        'Remove the old shortcut
        FavesList.Remove(m)
        FavesList.Add(fn)
    End Sub

    ''' <summary>
    ''' Finds f in the FavesList and makes it point to destinationpath, preserving the bookmark
    ''' </summary>
    ''' <param name="f"></param>
    ''' <param name="destinationpath"></param>
    Public Sub Check(f As IO.FileInfo, destinationpath As String)
        OkToDelete = False
        Dim mf As New List(Of String)
        mf = FavesList.FindAll(Function(x) x.Contains("\" & f.Name))
        If mf.Count = 0 Then
            OkToDelete = True
        ElseIf destinationpath = "" Then
            If MsgBox(f.Name & " - There are links to this file. Delete?", MsgBoxStyle.YesNoCancel, "Delete file?") = MsgBoxResult.Yes Then
                OkToDelete = True
                ' FavesList.RemoveAll(Function(e) mf.Exists(Function(x) x = e))
            Else
                OkToDelete = False
            End If
        Else
            ReallocateShortCuts(f, mf)
        End If
    End Sub

    Private Sub ReallocateShortCuts(f As IO.FileInfo, mf As List(Of String))
        Dim bk As Long = 0
        For Each m In mf
            Dim minfo As New IO.FileInfo(m)
            'Get the bookmark
            If InStr(m, "%") <> 0 Then
                Dim s As String()
                s = m.Split("%")
                bk = Val(s(1))
            Else
                bk = 0
            End If
            'Create a new shortcut where the old one was.
            ReallocateShortCut(f, m, bk, minfo)

        Next
    End Sub

    Public Function GetLinksOf(Path As String) As List(Of String)

        ' Exit Function
        Dim finfo As New IO.FileInfo(Path)
        Dim list As New List(Of String)
        If finfo.Exists Then
            'FavesList.Sort()
            'For Each s In FavesList
            '    If InStr(s, "\" & Left(finfo.Name, 20)) <> 0 Then
            '        list.Add(s)
            '    End If
            'Next
            list = FavesList.FindAll(Function(x) InStr(x, "\" & finfo.Name) <> 0)
        End If
        Return List
    End Function
    Public Property OkToDelete As Boolean = True
End Class
