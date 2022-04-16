Public Class FormTagging
    Private Sub FormTagging_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'MVDataSet.MarksNoCrypt' table. You can move, or remove it, as needed.
        Me.MarksNoCryptTableAdapter.Fill(Me.MVDataSet.MarksNoCrypt)

    End Sub
End Class