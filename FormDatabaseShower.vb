Imports System.Threading
Public Class FormDatabaseShower
    Public MDB As New Database
    Private SubDB As New Database
    Private Filenames As New List(Of String)

    Public Sub LoadEntries(DB As Database)

        LoadEntriesThread(DB)
        Me.Text = MDB.Name & " " & String.Format("{0} files out of {1}", SubDB.Entries.Count, DB.Entries.Count)
    End Sub

    Private Sub LoadEntriesThread(DB As Database)
        dgv.ColumnCount = 5
        dgv.RowCount = DB.Entries.Count

        Dim i = 0
        For Each entry In DB.Entries
            Filenames.Add(entry.Filename)
            dgv.Rows(i).Cells(0).Value = entry.Filename
            dgv.Rows(i).Cells(1).Value = entry.Path
            dgv.Rows(i).Cells(2).Value = entry.Size
            dgv.Rows(i).Cells(3).Value = entry.Dt
            dgv.Rows(i).Cells(4).Value = entry.Rand
            Application.DoEvents()
            i += 1
        Next
    End Sub

    Private Sub dgv_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgv.RowEnter
        Dim j = e.RowIndex
        If j >= 0 Then
            Dim s = dgv.Rows(j).Cells(1).Value & dgv.Rows(j).Cells(0).Value


            If j > 0 Then FormMain.HighlightCurrent(s)

        End If
    End Sub

    Private Sub DatabaseShower_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up
            Case Else
                If FinderTextBox.Focused Then
                Else
                    FormMain.Main_KeyDown(sender, e)
                End If
        End Select
    End Sub


    Private Sub UpdateFilter(Search As String)
        SubDB.Entries.Clear()
        SubDB.Entries = MDB.FindPartEntry(Search)
        If SubDB.Entries.Count > 0 Then LoadEntries(SubDB)
    End Sub


    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles FinderTextBox.KeyUp
        If e.KeyCode = Keys.Return Then
            UpdateFilter(FinderTextBox.Text)
        End If
    End Sub

    Private Sub FindEntry(text As String)
        'MDB.Sort(New CompareDBByFilename)
        For Each r In dgv.Rows
            If r.cells(0).value.contains(text) Then
                r.selected = True
            Else
                r.selected = False
            End If

        Next
        'Throw New NotImplementedException()
    End Sub


    Private Sub DatabaseShower_Load(sender As Object, e As EventArgs) Handles Me.Load
        dgv.Sort(dgv.Columns(0), System.ComponentModel.ListSortDirection.Ascending)

    End Sub

    Private Sub DatabaseShower_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        FormMain.DB = MDB
        ' MSFiles.NextF.List = ListFromDatabase(MDB)
    End Sub

    Private Sub btnAnalyzeDups_Click(sender As Object, e As EventArgs) Handles btnAnalyzeDups.Click
        Dim x As New FormDuplicates
        x.DB = MDB
        x.FindDuplicates()
        x.Show()
    End Sub

    Private Sub FinderTextBox_TextChanged(sender As Object, e As EventArgs) Handles FinderTextBox.TextChanged

    End Sub
End Class