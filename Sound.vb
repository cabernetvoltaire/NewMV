Public Class SoundController
    Public SoundPlayer As AxWMPLib.AxWindowsMediaPlayer
    Private WithEvents mCurrentPlayer As AxWMPLib.AxWindowsMediaPlayer
    Public WithEvents SPH As New SpeedHandler
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
                Mute(mMuted)
            End If
        End Set
    End Property
    Private mSlow As Boolean

    Public Sub New(Sound As AxWindowsMediaPlayer, Player As AxWindowsMediaPlayer)
        SoundPlayer = Sound
        CurrentPlayer = Player
    End Sub
    Public Sub New()

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
            mCurrentPlayer.settings.mute = True
            SoundPlayer.URL = mCurrentPlayer.URL
            SoundPlayer.settings.mute = False
            SoundPlayer.Ctlcontrols.currentPosition = CurrentPlayer.Ctlcontrols.currentPosition
            SoundPlayer.settings.rate = SPH.FrameRate / 30
        Else
            SoundPlayer.URL = ""
            mCurrentPlayer.settings.mute = False
        End If
    End Sub
    'Private Sub ChangePos() Handles mCurrentPlayer.PositionChange
    '    SoundPlayer.Ctlcontrols.currentPosition = CurrentPlayer.Ctlcontrols.currentPosition
    'End Sub
    Private Sub Mute(mute As Boolean)
         mCurrentPlayer.settings.mute = mute
        If mSlow Then
            SoundPlayer.settings.mute = mute
        Else
            SoundPlayer.settings.mute = Not mute
        End If

    End Sub
    Private Sub OnSpeedChange() Handles SPH.SpeedChanged
        SoundPlayer.settings.rate = FormMain.SP.FrameRate
        SlowSound()
        'SoundPlayer.settings.rate = SPH.FrameRate / 30
    End Sub
    Private Sub OnURLChange() Handles mCurrentPlayer.CurrentItemChange
        SoundPlayer.URL = mCurrentPlayer.URL
        SoundPlayer.settings.rate = mCurrentPlayer.settings.rate
    End Sub
    Private Sub OnPosChange() Handles mCurrentPlayer.PositionChange
        SoundPlayer.Ctlcontrols.currentPosition = mCurrentPlayer.Ctlcontrols.currentPosition
    End Sub

End Class
