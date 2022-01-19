Public Class FavouriteFiles
    Property FavouritesLocation As String = GlobalFavesPath
    Public WithEvents LBH As New ListBoxHandler

    Public Sub SortFiles(s As SortHandler)
        LBH.SortOrder = s
        LBH.Filter.State = FilterHandler.FilterState.LinkOnly
        LBH.FillBox()
    End Sub
    Private Sub FavouriteFiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        FormMain.FavouritesFormOpen = True
        LBH.ListBox = ListBox1
        LBH.Filter.State = FilterHandler.FilterState.LinkOnly
        LBH.ItemList = FilesBelow(FavouritesLocation)
        LBH.SingleLinks = True

        LBH.FillBox()

    End Sub
    Private Sub ShowFile() Handles LBH.ListIndexChanged
        FormMain.HighlightCurrent(LBH.ListBox.SelectedItem)
    End Sub



    Private Function FilesBelow(fpath As String) As List(Of String)
        Dim list As New List(Of String)
        Dim dir As New IO.DirectoryInfo(fpath)
        For Each f In dir.GetFiles
            list.Add(f.FullName)

        Next
        For Each d In dir.GetDirectories
            list.AddRange(FilesBelow(d.FullName))
        Next
        Return list

    End Function

    Private Sub FavouriteFiles_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        FormMain.HandleKeys(sender, e)
    End Sub

    'Private Sub FavouriteFiles_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
    '    FileHandling.MSFiles.Listbox = ListBox1
    'End Sub

    'Private Sub FavouriteFiles_LostFocus(sender As Object, e As EventArgs) Handles Me.LostFocus
    '    FileHandling.MSFiles.Listbox = FormMain.FBH.ListBox
    'End Sub
End Class