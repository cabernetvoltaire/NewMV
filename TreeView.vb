Public Class TVPopup
    Public StartFolder As String = "Q:\"
    Private Sub FileSystemTree1_Load(sender As Object, e As EventArgs) Handles FileSystemTree1.Load
        FileSystemTree1.SelectedFolder = StartFolder
    End Sub
End Class