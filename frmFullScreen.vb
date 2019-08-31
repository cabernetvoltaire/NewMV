Imports AxWMPLib

Public Class FullScreen
    Public Shared Property Changing As Boolean
    Public Event FullScreenClosing()
    Public FirstMediaIndex As Integer
    Public FSFiles As New MediaSwapper()
    Private Sub FullScreen_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        InitialisePlayer(FSWMP)
        InitialisePlayer(FSWMP3)
        InitialisePlayer(FSWMP2)
        FSFiles = FileHandling.MSFiles
        FSFiles.Media1.Player = FSWMP
        FSFiles.Media2.Player = FSWMP2
        FSFiles.Media3.Player = FSWMP3
        FSFiles.Media1.Picture = FSPB1
        FSFiles.Media2.Picture = FSPB2
        FSFiles.Media3.Picture = FSPB3


        FSFiles.ListIndex = MSFiles.ListIndex

        'FSWMP.URL = MainForm.MainWMP1.URL
        'FSWMP2.URL = MainForm.MainWMP2.URL
        'FSWMP3.URL = MainForm.MainWMP3.URL
        'MSFiles.AssignPlayers(FSWMP, FSWMP2, FSWMP3)

        'fullScreenPicBox = MainForm.PictureBox1
        'PictureBox1 = MainForm.PictureBox2
        'PictureBox2 = MainForm.PictureBox3
        'MSFiles.AssignPictures(fullScreenPicBox, PictureBox1, PictureBox2)
        'MSFiles.ListIndex = FirstMediaIndex

    End Sub

    Private Sub InitialisePlayer(WMP As AxWindowsMediaPlayer)
        WMP.uiMode = "None"
        If separate Then
        Else
            WMP.Dock = DockStyle.Fill
        End If
        WMP.settings.mute = True

    End Sub

    Private Sub FullScreen_Keydown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control

        If e.KeyCode = Keys.Escape Then
            MainForm.GoFullScreen(False)

            ' RaiseEvent FullScreenClosing()
        End If

        MainForm.HandleKeys(Me, e)
        e.SuppressKeyPress = True
    End Sub


    Private Sub FSWMP_KeyDownEvent(sender As Object, e As _WMPOCXEvents_KeyDownEvent) Handles FSWMP.KeyDownEvent
        If e.nKeyCode = Keys.Escape Then
            MainForm.GoFullScreen(False)
        End If
    End Sub



    Private Sub fullScreenPicBox_MouseWheel(sender As Object, e As MouseEventArgs) Handles FSPB1.MouseWheel, Me.MouseWheel, FSPB3.MouseWheel, FSPB2.MouseWheel
        PictureFunctions.Mousewheel(sender, e)
    End Sub

    Private Sub fullScreenPicBox_MouseMove(sender As Object, e As MouseEventArgs) Handles FSPB1.MouseMove, Me.MouseMove, FSPB2.MouseMove, FSPB3.MouseMove
        picBlanker = FSBlanker
        PictureFunctions.MouseMove(sender, e)
    End Sub

    Private Sub fullScreenPicBox_KeyDown(sender As Object, e As KeyEventArgs) Handles FSPB1.KeyDown, Me.KeyUp
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub


    Private Sub FullScreen_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub fullScreenPicBox_MouseDown(sender As Object, e As MouseEventArgs) Handles FSPB1.MouseDown
        Select Case e.Button
            Case MouseButtons.XButton1, MouseButtons.XButton2
                MainForm.AdvanceFile(e.Button = MouseButtons.XButton2)
                e = Nothing

            Case Else
                PicClick(FSPB1)
        End Select
    End Sub

End Class