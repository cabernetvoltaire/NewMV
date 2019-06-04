Public Class ShowListForm
    Dim LBH As New ListBoxHandler(ListBox1)

    Public Sub ShowListForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        MainForm.frmMain_KeyDown(sender, e)
        e.SuppressKeyPress = True
        Me.Show()
    End Sub
End Class