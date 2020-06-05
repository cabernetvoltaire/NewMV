Public Class DatabaseShower
    Public MDB As New Database
    Private SubDB As New Database
    Private Filenames As New List(Of String)
    Public Sub LoadEntries(DB As Database)

        dgv.ColumnCount = 3
        dgv.RowCount = DB.Entries.Count

        Dim i = 0
        For Each entry In DB.Entries
            Filenames.Add(entry.Filename)
            dgv.Rows(i).Cells(0).Value = entry.Filename
            dgv.Rows(i).Cells(1).Value = entry.Path
            dgv.Rows(i).Cells(2).Value = entry.Size
            Application.DoEvents()
            i += 1
        Next
        Me.Text = String.Format("{0} files out of {1}", SubDB.Entries.Count, DB.Entries.Count)
    End Sub

    Private Sub dgv_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.RowEnter
        Dim j = e.RowIndex
        If j >= 0 Then
            Dim s = dgv.Rows(j).Cells(1).Value & dgv.Rows(j).Cells(0).Value
            If j > 0 Then MainForm.HighlightCurrent(s)
        End If
    End Sub

    Private Sub DatabaseShower_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up
            Case Else
                If TextBox1.Focused Then
                Else
                    MainForm.Main_KeyDown(sender, e)
                End If
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
            FindEntry(TextBox1.Text)
        End If
    End Sub

    Private Sub FindEntry(text As String)
        'MDB.Sort(New CompareDBByFilename)
        Dim i As Integer = Filenames.BinarySearch(text)
        dgv.Rows(Math.Abs(i)).Selected = True
        'Throw New NotImplementedException()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus

    End Sub

    Private Sub DatabaseShower_Load(sender As Object, e As EventArgs) Handles Me.Load
        dgv.Sort(dgv.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub DatabaseShower_GotFocus(sender As Object, e As EventArgs) Handles Me.Enter
        MainForm.DB = MDB

    End Sub

End Class