Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'MVDataSet.MarksNoCrypt' table. You can move, or remove it, as needed.
        For Each dbel In FormMain.DB.Entries
            Me.MVDataSet.MarksNoCrypt.Append(getdbrow(dbel))
        Next
        Me.MarksNoCryptTableAdapter.Fill(Me.MVDataSet.MarksNoCrypt)
    End Sub
    Private Function getdbrow(dbelement As DatabaseEntry) As MVDataSet.MarksNoCryptRow
        Dim x As New MVDataSet.MarksNoCryptRow(Nothing) With {
            .fCreateDate = dbelement.Dt,
            .fFileName = dbelement.Filename,
            .fPath = dbelement.Path,
            .fSize = dbelement.Size
        }
        Return x
    End Function
End Class