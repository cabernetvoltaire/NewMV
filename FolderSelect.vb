Imports System.ComponentModel
Imports AxWMPLib
Imports MasaSam.Forms.Controls

Public Class FolderSelect
    Private PreMH As New MediaHandler("PreMH")

    Public Property Alpha() As Int16
    Private newButtonNumber As Byte
    Public KeepDisplaying As Boolean = True
    Public Property ButtonNumber() As Byte
        Get
            Return newButtonNumber
        End Get
        Set(ByVal value As Byte)
            newButtonNumber = value
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
                '    fst1_Paint(Me, Nothing)
            End If
        End Set
    End Property
    Public Property IsLoaded As Boolean = False
    Public Property Button As Button
    Private Sub PlayMedia(value As String)
        If value = "" Then Exit Sub
        ''  Exit Sub
        Dim x As New IO.DirectoryInfo(value)
        Dim count As Integer = x.GetFiles.Count
        If count = 0 Then

        Else

            Dim i As Integer = Int(Rnd() * (count - 1))

            Dim f As IO.FileInfo

            f = x.GetFiles.ToArray(i)
            ' med.Bookmark = med.Duration * Rnd()
            PreMH.Player = PreviewWMP
            PreMH.Player.settings.mute = True
            PreMH.Picture = PictureBox1
            PreMH.StartPoint.State = StartPointHandler.StartTypes.Random
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

    Private newChosenFolder As Integer
    Public ReadOnly Property ChosenFolder() As Integer
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
        AssignButton(ButtonNumber, Alpha, 1, Folder, True)
        'Me.Close()
    End Sub
    Private Sub Label1_DoubleClick(sender As Object, e As EventArgs) Handles Label1.DoubleClick
        fst1.SelectedFolder = newFolder
    End Sub

    Private Sub FolderSelect_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        changepic.Enabled = False
        If Not My.Computer.FileSystem.DirectoryExists(Folder) Then
            CreateNewDirectory(fst1, Folder, False)
        End If
        'PreMH.Player.URL = ""
        'PreMH.Player.Dispose()
        'PreMH.Picture.Dispose()
        'GC.Collect()
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
        .Interval = 1500,
        .Enabled = True
    }

    Private Sub UpdatePic() Handles changepic.Tick
        ChangeMedia()
        '        changepic.Enabled = False
    End Sub

    Private Sub FolderSelect_MouseEnter(sender As Object, e As EventArgs) Handles Me.MouseEnter
        Me.BringToFront()
    End Sub
End Class