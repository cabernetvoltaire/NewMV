Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim tx As New TextBox
        TableLayoutPanel1.Controls.Add(tx, 0, 0)
        '       Me.Controls.Add(tx)
    End Sub
End Class