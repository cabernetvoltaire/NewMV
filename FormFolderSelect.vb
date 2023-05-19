Imports System.ComponentModel
Imports AxWMPLib
Imports MasaSam.Forms.Controls
Imports System.IO
Imports System.Threading.Tasks

Public Class FormFolderSelect
    Private PreMH As New MediaHandler("PreMH")
    Private ClickedFolder As Boolean = False
    Public BH As New ButtonHandler
    Public Property Alpha() As Int16
    Private newButtonNumber As Byte

    Public Property ButtonNumber() As Byte
        Get
            Return newButtonNumber
        End Get
        Set(ByVal value As Byte)
            newButtonNumber = value
            Folder = BH.buttons.CurrentRow.Buttons(value).Path
        End Set
    End Property

    Private newFolder As String
    Public Property Folder() As String
        Get
            newFolder = fst1.SelectedFolder
            Return newFolder
        End Get
        Set(ByVal value As String)
            newFolder = value
            Label1.Text = value
            If value <> "" Then
                PlayMedia(value)
            End If
        End Set
    End Property
    Public Property IsLoaded As Boolean = False
    Public Property Button As Button
    Private Sub PlayMedia(value As String)
        If value = "" Then Exit Sub
        ''  Exit Sub
        Dim x As New IO.DirectoryInfo(value)
        If x.Exists = False Then Exit Sub
        Dim count As Integer = x.GetFiles.Count
        If count = 0 Then

        Else

            Dim i As Integer = Int(Rnd() * (count - 1))

            Dim f As IO.FileInfo

            f = x.GetFiles.ToArray(i)
            PreMH.Player = PreviewWMP
            PreMH.Player.settings.mute = True
            PreMH.PicHandler.State = PictureHandler.Screenstate.Fitted
            PreMH.PicHandler.PicBox = PictureBox1
            PreMH.SPT.State = StartPointHandler.StartTypes.Random
            PreMH.MediaPath = f.FullName
            If PreMH.MediaType = Filetype.Movie Or PreMH.MediaType = Filetype.Pic Then

                If f.Exists Then
                    Select Case PreMH.MediaType
                        Case Filetype.Movie
                            SplitContainer1.Panel2Collapsed = True
                            SplitContainer1.Panel1Collapsed = False
                            PreviewWMP.settings.mute = True
                            PreviewWMP.uiMode = "none"
                        Case Filetype.Pic
                            SplitContainer1.Panel1Collapsed = True
                            SplitContainer1.Panel2Collapsed = False

                    End Select
                End If
            End If

        End If
    End Sub

    Private newChosenFolder As String
    Public ReadOnly Property ChosenFolder() As String
        Get
            Dim s As String
            If TextBox1.Text <> "" AndAlso Not ClickedFolder Then
                s = Folder & "\" & TextBox1.Text
            Else
                s = Folder

            End If

            newChosenFolder = s
            Return newChosenFolder
            ClickedFolder = False
        End Get
    End Property
    Public Sub ChangeMedia()
        PlayMedia(newFolder)
    End Sub
    Public Sub UpdateFolder()
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        FormMain.BH.AssignButton(ButtonNumber, ChosenFolder)
        If Not My.Computer.FileSystem.DirectoryExists(ChosenFolder) Then
            CreateNewDirectory(fst1, ChosenFolder, False)
        End If
    End Sub
    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub FolderSelect_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        changepic.Enabled = False

        IsLoaded = False
    End Sub
    Private Sub fst1_Paint(sender As Object, e As PaintEventArgs) Handles fst1.Paint
        fst1.SelectedFolder = newFolder
    End Sub
    Private Sub fst1_DirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles fst1.DirectorySelected
        Folder = e.Directory.FullName

    End Sub


    Private WithEvents changepic As New Timer With {
        .Interval = 3000,
        .Enabled = False
    }

    Public Sub UpdatePic() Handles changepic.Tick
        ChangeMedia()
    End Sub


    Private Sub FormFolderSelect_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        Timer1.Enabled = False
    End Sub

    Public Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        Me.Hide()

    End Sub
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
                                       For Each f In filteredFolders
                                           ListBox1.Items.Add(f)
                                       Next
                                   End Sub))
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        fst1.SelectedFolder = ListBox1.SelectedItem
        Folder = ListBox1.SelectedItem
        ClickedFolder = True
        btnAssign_Click(sender, e)
    End Sub

    Private Sub FormFolderSelect_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then Me.Hide()
    End Sub
End Class