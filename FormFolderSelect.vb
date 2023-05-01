Imports System.ComponentModel
Imports AxWMPLib
Imports MasaSam.Forms.Controls

Public Class FormFolderSelect
    Private PreMH As New MediaHandler("PreMH")
    Public BH As New ButtonHandler
    Public Property Alpha() As Int16
    Private newButtonNumber As Byte
    Public KeepDisplaying As Boolean = True
    Public Property ButtonNumber() As Byte
        Get
            Return newButtonNumber
        End Get
        Set(ByVal value As Byte)
            newButtonNumber = value
            Folder = BH.buttons.CurrentRow.Buttons(value).Path
        End Set
    End Property
    Public Property Showing As Boolean = True
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
            ' med.Bookmark = med.Duration * Rnd()
            PreMH.Player = PreviewWMP
            PreMH.Player.settings.mute = True
            PreMH.PicHandler.State = PictureHandler.Screenstate.Fitted
            PreMH.PicHandler.PicBox = PictureBox1
            PreMH.SPT.State = StartPointHandler.StartTypes.Random
            PreMH.MediaPath = f.FullName
            If PreMH.MediaType = Filetype.Movie Or PreMH.MediaType = Filetype.Pic Then

                If f.Exists Then
                    '                    PreMH.MediaPath
                    Select Case PreMH.MediaType
                        Case Filetype.Movie
                            SplitContainer1.Panel2Collapsed = True
                            SplitContainer1.Panel1Collapsed = False
                            'PreviewWMP.URL = f.FullName
                            PreviewWMP.settings.mute = True
                            PreviewWMP.uiMode = "none"
                            'PreviewWMP.stretchToFit = True
                            ''PreviewWMP.Visible = True
                            'PreviewWMP.Ctlcontrols.play()
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
            If TextBox1.Text <> "" Then
                s = Folder & "\" & TextBox1.Text
            Else
                s = Folder

            End If

            newChosenFolder = s
            Return newChosenFolder
        End Get
    End Property
    Public Sub ChangeMedia()
        PlayMedia(newFolder)
    End Sub
    Public Sub UpdateFolder()
        fst1.SelectedFolder = newFolder
        '    fst1_Paint(Me, Nothing)
    End Sub

    Private Sub btnAssign_Click(sender As Object, e As EventArgs) Handles btnAssign.Click
        FormMain.BH.AssignButton(ButtonNumber, ChosenFolder)
        If Not My.Computer.FileSystem.DirectoryExists(ChosenFolder) Then
            CreateNewDirectory(fst1, ChosenFolder, False)
        End If
        'Me.Close()
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


    'Private Sub FolderSelect_Shown(sender As Object, e As CancelEventArgs) Handles Me.Shown

    '    IsLoaded = True
    'End Sub
    Private WithEvents changepic As New Timer With {
        .Interval = 3000,
        .Enabled = False
    }

    Public Sub UpdatePic() Handles changepic.Tick
        ChangeMedia()
        '        changepic.Enabled = False
    End Sub


End Class