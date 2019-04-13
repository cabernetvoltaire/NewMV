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
        templist = Duplicatelist(mFavesList)

        For Each m In mFavesList
            Dim x = LinkTarget(m)
            '      Debug.Print(x)
            If x = path Then
                templist.Remove(m)
                Dim f As New IO.FileInfo(m)
                f.Delete()
            End If
        Next
        mFavesList = Duplicatelist(templist)
    End Sub
    Public Sub New(path As String)
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
    ''' <summary>
    ''' Finds f in the FavesList and makes it point to destinationpath, preserving the bookmark
    ''' </summary>
    ''' <param name="f"></param>
    ''' <param name="destinationpath"></param>
    Public Sub Check(f As IO.FileInfo, destinationpath As String)
        Dim m As String = "a"
        'm is a file in the favourites, and we have to update its target to the new destination.
        Dim bk As Long = 0
        While m IsNot Nothing

            m = FavesList.Find(Function(x) x.Contains(f.Name))
            If m Is Nothing Then Exit While
            'Debug.Print(m)
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
            Dim sch As New ShortcutHandler(destinationpath, minfo.Directory.FullName, f.Name)
            sch.MarkOffset = 0
            Dim fn As String = sch.Create_ShortCut(bk)
            'Remove the old shortcut
            FavesList.Remove(m)
        End While
    End Sub
End Class
