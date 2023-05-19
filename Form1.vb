Imports System.IO
Imports System.Threading.Tasks

Public Class Form1
    Private allFolders As New List(Of String)

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Task.Run(Sub() ListAllFolders())
    End Sub

    Private Sub ListAllFolders()
        Dim drives As DriveInfo() = DriveInfo.GetDrives()

        For Each drive As DriveInfo In drives
            If drive.Name <> "Q:\" Then Continue For
            If drive.IsReady Then
                Try
                    ListFolders(drive.RootDirectory.FullName)
                Catch ex As Exception
                    Invoke(New Action(Sub() allFolders.Add($"Error accessing {drive.RootDirectory.FullName}: {ex.Message}")))
                End Try
            End If
        Next
        UpdateListBox()
    End Sub

    Private Sub ListFolders(path As String)
        Try
            For Each folder As String In Directory.GetDirectories(path)
                Invoke(New Action(Sub() allFolders.Add(folder)))
                ListFolders(folder)
            Next
        Catch ex As UnauthorizedAccessException
            Invoke(New Action(Sub() allFolders.Add($"Unauthorized access to {path}: {ex.Message}")))
        Catch ex As Exception
            Invoke(New Action(Sub() allFolders.Add($"Error accessing {path}: {ex.Message}")))
        End Try
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        UpdateListBox()
    End Sub

    Private Sub UpdateListBox()
        ListBox1.Invoke(New Action(Sub()
                                       ListBox1.Items.Clear()
                                       Dim filteredFolders = allFolders.Where(Function(f) f.IndexOf(TextBox1.Text, StringComparison.CurrentCultureIgnoreCase) >= 0)
                                       For Each folder In filteredFolders
                                           ListBox1.Items.Add(folder)
                                       Next
                                   End Sub))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub
End Class
