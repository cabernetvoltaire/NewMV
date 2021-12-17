Public Class Form3

    Public LBH As New ListBoxHandler
    Public MLH As New MediaListHandler

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        LBH.ListBox = ListBox1 '
        MLH.List = FormMain.LBH.ItemList
        MLH.Index = FormMain.LBH.CurrentIndex
        LBH.FillBox(MLH.SetOfThree)
    End Sub

End Class