Public Class AutoTrail
    Private WithEvents Timer As New Timer()

    Private ReadOnly MediaPlayer As AxWMPLib.AxWindowsMediaPlayer
    Private ReadOnly RandomGenerator As New Random()
    Private ReadOnly InitialMomentsDuration As Integer = 500 ' Duration for initial moments in milliseconds
    Private ReadOnly MainRandomJumpInterval As Integer = 5000 ' Interval for main random jumping in milliseconds
    Private InitialMomentsPlayed As Integer = 0
    Private ReadOnly InitialMomentsCount As Integer = 3 ' Number of initial brief moments to play
    Private ReadOnly Markers As New List(Of Double) ' Assuming this list is populated elsewhere

    Public Sub New(mediaPlayer As AxWMPLib.AxWindowsMediaPlayer, markers As List(Of Double))
        Me.MediaPlayer = mediaPlayer
        Me.Markers = markers
        ' Initialize the timer but do not start it yet
        Timer.Interval = InitialMomentsDuration
    End Sub
    Public Sub StopPlay()
        Timer.Stop()
    End Sub

    Public Sub Play()
        ' Reset counter for initial moments
        InitialMomentsPlayed = 0
        ' Start or restart the timer
        Timer.Start()
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles Timer.Tick
        If InitialMomentsPlayed < InitialMomentsCount Then
            ' Play initial moments from the beginning for a brief time

            InitialMomentsPlayed += 1
        Else
            ' Transition to main random jumping
            Timer.Interval = MainRandomJumpInterval
            JumpRandomly()
        End If
    End Sub

    Private Sub JumpRandomly()
        If Markers.Count > 0 Then
            ' Example logic to jump to a random marker
            Dim randomIndex = RandomGenerator.Next(Markers.Count)
            MediaPlayer.Ctlcontrols.currentPosition = Markers(randomIndex)
        Else
            ' If no markers, jump to a random position in the video
            Dim randomPosition = RandomGenerator.NextDouble() * MediaPlayer.currentMedia.duration
            MediaPlayer.Ctlcontrols.currentPosition = randomPosition
        End If
    End Sub
End Class
