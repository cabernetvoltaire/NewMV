Public Class SoundController
    Public SoundPlayer As AxWMPLib.AxWindowsMediaPlayer
    Private WithEvents mCurrentPlayer As AxWMPLib.AxWindowsMediaPlayer
    Public WithEvents SPH As New SpeedHandler
    Public WithEvents ResetPos As New Timer With {.Interval = 1000}
    Public AllowedLag As Long = 15

    Private mMuted As Boolean
    Public Property CurrentPlayer() As AxWMPLib.AxWindowsMediaPlayer
        Get
            Return mCurrentPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            If value IsNot mCurrentPlayer Then
                mCurrentPlayer = value
                ' mCurrentPlayer.URL = value.URL
                SoundPlayer.URL = value.URL
            End If
        End Set
    End Property
    Public Property Muted() As Boolean
        Get
            Return mMuted
        End Get
        Set(ByVal value As Boolean)
            If value <> mMuted Then
                mMuted = value
                mCurrentPlayer.settings.mute = mMuted
                SoundPlayer.settings.mute = True
            End If
        End Set
    End Property
    Private mSlow As Boolean

    Public Sub New(Sound As AxWindowsMediaPlayer, Player As AxWindowsMediaPlayer)
        SoundPlayer = Sound
        CurrentPlayer = Player
    End Sub
    Public Sub New()
        mSlow = False
    End Sub

    Public Property Slow() As Boolean
        Get
            Return mSlow
        End Get
        Set(ByVal value As Boolean)

            mSlow = value
            SlowSound()
        End Set
    End Property

    Private Sub SlowSound()

        If mSlow Then
            SoundPlayer.URL = mCurrentPlayer.URL
            SoundPlayer.settings.mute = False
            mCurrentPlayer.settings.mute = True
            SoundPlayer.Ctlcontrols.currentPosition = CurrentPlayer.Ctlcontrols.currentPosition
            SoundPlayer.settings.rate = SPH.FrameRate / 30
        Else
            SoundPlayer.URL = ""
            SoundPlayer.settings.mute = True
            mCurrentPlayer.settings.mute = False
        End If
    End Sub
    'Private Sub ChangePos() Handles mCurrentPlayer.PositionChange
    '    SoundPlayer.Ctlcontrols.currentPosition = CurrentPlayer.Ctlcontrols.currentPosition
    'End Sub

    Private Sub OnSpeedChange() Handles SPH.SpeedChanged
        SoundPlayer.settings.rate = FormMain.SP.FrameRate
        SlowSound()
        'SoundPlayer.settings.rate = SPH.FrameRate / 30
    End Sub
    Private Sub OnURLChange() Handles mCurrentPlayer.CurrentItemChange
        ToggleReset()
        SoundPlayer.URL = mCurrentPlayer.URL
        ' SoundPlayer.settings.mute = mSlow
        SoundPlayer.settings.rate = mCurrentPlayer.settings.rate
    End Sub
    Private Sub OnPosChange() Handles mCurrentPlayer.PositionChange
        SoundPlayer.Ctlcontrols.currentPosition = mCurrentPlayer.Ctlcontrols.currentPosition
        ToggleReset()
    End Sub
    Private Sub ResetPos_Tick() Handles ResetPos.Tick
        Dim x = SoundPlayer.Ctlcontrols.currentPosition
        Dim y = mCurrentPlayer.Ctlcontrols.currentPosition
        If x - y > AllowedLag Then
            SoundPlayer.Ctlcontrols.currentPosition = mCurrentPlayer.Ctlcontrols.currentPosition
        End If
    End Sub
    Private Sub ToggleReset()
        ResetPos.Enabled = False
        ResetPos.Enabled = True
    End Sub
End Class
