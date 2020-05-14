Imports System.IO
Public Class FormDeadFiles
    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        If DatabaseForm.Visible Then
            DatabaseForm.ShortFilepath = Path.GetFileName(ListBox1.SelectedItem)
        End If
    End Sub
End Class