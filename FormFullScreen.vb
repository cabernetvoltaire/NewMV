Imports AxWMPLib

Public Class FullScreen
    Public Shared Property Changing As Boolean
    Public Event FullScreenClosing()
    Public FirstMediaIndex As Integer
    Private mFSFiles As New MediaSwapper
    Public Property FSFiles() As MediaSwapper
        Get
            Return mFSFiles
        End Get
        Set(ByVal value As MediaSwapper)
            mFSFiles = value
            mFSFiles.AssignPictures(FSPB1, FSPB2, FSPB3)
            mFSFiles.AssignPlayers(FSWMP, FSWMP2, FSWMP3)
            mFSFiles.Media1.Player.uiMode = "none"
            mFSFiles.Media2.Player.uiMode = "none"
            mFSFiles.Media3.Player.uiMode = "none"
            mFSFiles.ListIndex = FileHandling.MSFiles.ListIndex
        End Set
    End Property
    Private Sub FullScreen_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        mFSFiles.DockMedias(General.separate)
        AddHandler FSPB1.MouseDown, AddressOf fullScreenPicBox_MouseDown
        AddHandler FSPB2.MouseDown, AddressOf fullScreenPicBox_MouseDown
        AddHandler FSPB3.MouseDown, AddressOf fullScreenPicBox_MouseDown
    End Sub



    Private Sub FullScreen_Keydown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control

        If e.KeyCode = Keys.Escape Then
            FormMain.GoFullScreen(False)

            ' RaiseEvent FullScreenClosing()
        End If

        FormMain.HandleKeys(Me, e)
        e.SuppressKeyPress = True
    End Sub


    Private Sub FSWMP_KeyDownEvent(sender As Object, e As _WMPOCXEvents_KeyDownEvent) Handles FSWMP.KeyDownEvent, FSWMP2.KeyDownEvent, FSWMP3.KeyDownEvent
        If e.nKeyCode = Keys.Escape Then
            FormMain.GoFullScreen(False)
        End If
    End Sub





    Private Sub FullScreen_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Maximized

    End Sub

    Private Sub fullScreenPicBox_MouseDown(sender As Object, e As MouseEventArgs)
        Select Case e.Button
            Case MouseButtons.XButton1, MouseButtons.XButton2
                FormMain.AdvanceFile(e.Button = MouseButtons.XButton1)
                e = Nothing

            Case Else
                'PicClick(FSPB1)
        End Select
    End Sub

End Class