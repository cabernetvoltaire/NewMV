Imports Microsoft.Office.Interop.Excel

Public Class TreeTry
    ' In your form's code
    Dim myFolderTreeView As FolderTreeView

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        myFolderTreeView = New FolderTreeView(TreeView1)
        myFolderTreeView.PopulateTreeView("Q:\Favourites")
    End Sub

End Class