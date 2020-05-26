Public Class DatabaseShower
    Public DB As New Database
    Public Sub LoadEntries()

        dgv.ColumnCount = 3
        dgv.RowCount = DB.Entries.Count

        Dim i = 0
        For Each entry In DB.Entries
            dgv.Rows(i).Cells(0).Value = entry.Filename
            dgv.Rows(i).Cells(1).Value = entry.Path
            dgv.Rows(i).Cells(2).Value = entry.Size
            i += 1
        Next
    End Sub

    Private Sub dgv_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.RowEnter
        Dim j = e.RowIndex
        If j >= 0 Then
            Dim s = dgv.Rows(j).Cells(1).Value & dgv.Rows(j).Cells(0).Value
            MainForm.HighlightCurrent(s)
        End If
    End Sub

    Private Sub DatabaseShower_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up, Keys.PageDown, Keys.PageUp
                'e.SuppressKeyPress = True
            Case Else
                MainForm.Main_KeyDown(sender, e)
        End Select
    End Sub
End Class