Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles Me.Load
        'TODO: This line of code loads data into the 'MVDataSet.MarksNoCrypt' table. You can move, or remove it, as needed.
        For Each dbel In FormMain.DB.Entries
            If dbel.Size < 2 ^ 30 Then
                MVDataSet.MarksNoCrypt.Rows.Add(getdbrow(dbel))
            End If

        Next
        MarksNoCryptTableAdapter.Update(MVDataSet.MarksNoCrypt)
        MarksNoCryptTableAdapter.Fill(MVDataSet.MarksNoCrypt)
    End Sub

    Private Sub dgv1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv1.RowEnter
        Dim j = e.RowIndex
        If j >= 0 Then
            Dim s = dgv1.Rows(j).Cells(2).Value & dgv1.Rows(j).Cells(1).Value
            If j > 0 Then FormMain.HighlightCurrent(s)
        End If
    End Sub

    Private Sub Form4_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up
            Case Else

                FormMain.Main_KeyDown(sender, e)
        End Select
    End Sub

    Private Function getdbrow(dbelement As DatabaseEntry) As MVDataSet.MarksNoCryptRow

        Dim m As DataRow
        m = Me.MVDataSet.MarksNoCrypt.NewRow()
        'm.Item("ID") = iCurrentAlpha + 1
        m.Item("fCreateDate") = dbelement.Dt
        m.Item("fFileName") = dbelement.Filename
        m.Item("fPath") = dbelement.Path
        m.Item("fSize") = dbelement.Size

        Return m
    End Function

End Class