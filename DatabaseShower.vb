Public Class DatabaseShower
    Public MDB As New Database
    Private SubDB As New Database
    Public Sub LoadEntries(DB As Database)

        dgv.ColumnCount = 3
        dgv.RowCount = DB.Entries.Count

        Dim i = 0
        For Each entry In DB.Entries

            dgv.Rows(i).Cells(0).Value = entry.Filename
            dgv.Rows(i).Cells(1).Value = entry.Path
            dgv.Rows(i).Cells(2).Value = entry.Size
            i += 1
        Next
        Me.Text = String.Format("{0} files out of {1}", SubDB.Entries.Count, DB.Entries.Count)
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
            Case Keys.Down, Keys.Up
            Case Else
                MainForm.Main_KeyDown(sender, e)
        End Select
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        UpdateFilter(ListBox1.SelectedItem)

    End Sub

    Private Sub UpdateFilter(Search As String)
        SubDB.Entries.Clear()
        SubDB.Entries = MDB.FindPartEntry(Search)
        If SubDB.Entries.Count > 0 Then LoadEntries(SubDB)
    End Sub


    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp
        If e.KeyCode = Keys.Return Then
            UpdateFilter(TextBox1.Text)
        End If
    End Sub
End Class