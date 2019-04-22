Public Class SoundController
    Public SoundPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Private WithEvents mCurrentPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Public WithEvents SPH As New SpeedHandler
    Private mMuted As Boolean
    Public Property CurrentPlayer() As AxWMPLib.AxWindowsMediaPlayer
        Get
            Return mCurrentPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            If value IsNot mCurrentPlayer Then
                mCurrentPlayer = value
                mCurrentPlayer.URL = value.URL
                OnURLChange()
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
    Public Property Slow() As Boolean
        Get
            Return mSlow
        End Get
        Set(ByVal value As Boolean)
            If value <> mSlow Then
                mSlow = value
                SlowSound()
            End If
        End Set
    End Property
    Private Sub SlowSound()
        If mSlow Then
            CurrentPlayer.settings.mute = True
            SoundPlayer.settings.mute = False
            SoundPlayer.URL = CurrentPlayer.URL
            SoundPlayer.Ctlcontrols.currentPosition = CurrentPlayer.Ctlcontrols.currentPosition
            SoundPlayer.settings.rate = SPH.FrameRate / 30
        Else
            SoundPlayer.URL = ""
            CurrentPlayer.settings.mute = False
        End If
    End Sub
    Private Sub Mute(mute As Boolean)
         mCurrentPlayer.settings.mute = mute
        If mSlow Then
            SoundPlayer.settings.mute = mute
        Else
            SoundPlayer.settings.mute = Not mute
        End If

    End Sub
    Private Sub OnSpeedChange() Handles SPH.SpeedChanged
        SoundPlayer.settings.rate = MainForm.SP.FrameRate

        'SoundPlayer.settings.rate = SPH.FrameRate / 30
    End Sub
    Private Sub OnURLChange()
        SoundPlayer.URL = CurrentPlayer.URL
        SoundPlayer.settings.rate = CurrentPlayer.settings.rate
    End Sub
    Private Sub OnPosChange() Handles mCurrentPlayer.PositionChange
        SoundPlayer.Ctlcontrols.currentPosition = mCurrentPlayer.Ctlcontrols.currentPosition
    End Sub
    'Public Sub New(Sound As AxWMPLib.AxWindowsMediaPlayer, Player As AxWMPLib.AxWindowsMediaPlayer, Speedhandler As SpeedHandler)
    '    SoundPlayer = Sound
    '    CurrentPlayer = Player
    '    SP = Speedhandler
    'End Sub
End Class
