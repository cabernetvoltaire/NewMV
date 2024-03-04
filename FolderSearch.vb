Imports System.IO
Public Class FolderSearch ' Replace with your actual form's class name

    Public Rootfolder As String = "Q:\"
    Private allFolders As List(Of String) = New List(Of String) ' Store all folders here
    Private lbh As New ListBoxHandler
    Public t As New Timer With {.Interval = 300}

    Private Sub FolderSearch_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Collect all folders and subfolders at the initial stage
        allFolders = FolderSearchModule.GetAllFolders(Rootfolder)
        For Each f In allFolders
            ListBox1.Items.Add(f)
        Next
        lbh.ListBox = ListBox1
        AddHandler t.Tick, AddressOf ProcessText

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        t.Enabled = True
    End Sub

    Private Sub ProcessText()
        ListBox1.Items.Clear()
        Dim SearchString As String = TextBox1.Text.Trim()

        ' Ensure there's a search string to avoid clearing the list unnecessarily
        If String.IsNullOrEmpty(SearchString) Then Return

        ' Filter the pre-collected list of folders based on the search string
        Dim matchingFolders As List(Of String) = FolderSearchModule.FilterFoldersBySubstring(allFolders, SearchString)

        For Each folderPath In matchingFolders
            ListBox1.Items.Add(folderPath)
        Next
        lbh.SortOrder = FormMain.PlayOrder
        t.Enabled = False
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        'FormMain.ChangeFolder(ListBox1.SelectedItem)

        FormMain.HighlightCurrent(GetFirstFilePath(ListBox1.SelectedItem))
    End Sub

    Private Sub FolderSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        FormMain.HandleKeys(sender, e)
        'e.SuppressKeyPress = True
    End Sub

    Function GetFirstFilePath(folderPath As String) As String
        Try
            Dim files As String() = Directory.GetFiles(folderPath)
            If files.Length > 0 Then
                Return files(0) ' Return the first file's path
            Else
                Return Nothing ' Return Nothing if there are no files
            End If
        Catch ex As Exception
            MsgBox("An error occurred: " & ex.Message)
            Return Nothing
        End Try
    End Function

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        Dim s As New List(Of String)
        For Each m In ListBox1.Items
            s.Add(m)
        Next
        FormMain.BH.AssignList(s)
    End Sub

    Private Sub FolderSearch_Leave(sender As Object, e As EventArgs) Handles Me.Leave

    End Sub

    Private Sub TextBox1_GotFocus(sender As Object, e As EventArgs) Handles TextBox1.GotFocus
        TextBox1.SelectAll()
    End Sub
    ' Other methods...
End Class
