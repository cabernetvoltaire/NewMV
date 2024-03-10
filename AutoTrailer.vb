Public Class AutoTrailerOld
    Private ReadOnly rand As New Random()
    ' Change RandomTimes to Single to support fractional seconds.
    Public Property RandomTimes As Single() = {0.7R, 1.2R, 2.0R, 2.5R}
    Public Property SelectionWeights As Decimal() = {0.1D, 0.3D, 0.6D, 1D}
    Public Property Framerates As Integer() = {5, 12, 20}
    Public Property AdvanceChance As Integer = 8

    ' Counter for tracking conditions.
    Public Property Counter As Int16 = 0

    ' Simplified way to get a weighted random index for speed.
    Public ReadOnly Property SpeedIndex As Byte
        Get
            Return GetWeightedIndex(SelectionWeights)
        End Get
    End Property


    ' Adjusted Duration property to return a Single, allowing for fractional seconds.
    Public ReadOnly Property Duration As Single
        Get
            Return RandomTimes(rand.Next(RandomTimes.Length))
        End Get
    End Property

    ' Utility function to get a weighted random index.
    Private Function GetWeightedIndex(weights As Decimal()) As Byte
        Dim value As Decimal = rand.NextDouble()
        Dim cumulative As Decimal = 0
        For i As Integer = 0 To weights.Length - 1
            cumulative += weights(i)
            If value < cumulative Then
                Return i
            End If
        Next
        Return weights.Length - 1 ' Fallback to the last index
    End Function
End Class
Public Class AutoTrailNew
    Private WithEvents Timer As New Timer()
    Private ReadOnly MediaPlayer As AxWMPLib.AxWindowsMediaPlayer
    Private ReadOnly RandomGenerator As New Random()
    Private ReadOnly InitialMomentsDuration As Integer = 500 ' Duration for initial moments in milliseconds
    Private ReadOnly MainRandomJumpInterval As Integer = 5000 ' Interval for main random jumping in milliseconds
    Private InitialMomentsPlayed As Integer = 0
    Private ReadOnly InitialMomentsCount As Integer = 3 ' Number of initial brief moments to play
    Private ReadOnly Markers As List(Of Double) ' Assuming this list is populated elsewhere

    Public Sub New(mediaPlayer As AxWMPLib.AxWindowsMediaPlayer, markers As List(Of Double))
        Me.MediaPlayer = mediaPlayer
        Me.Markers = markers
        ' Initialize the timer but do not start it yet
        Timer.Interval = InitialMomentsDuration
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
            MediaPlayer.Ctlcontrols.currentPosition = 0
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

Public Class AutoTrailer
    Private ReadOnly rand As New Random()
    Public Property Framerates As Integer() = {5, 12, 20}
    Public Property RandomDurations As Single() = {0.7R, 1.2R, 2.0R, 2.5R}
    Public Property PlaybackRates As Single() = {0.5R, 1.0R, 1.5R, 2.0R}
    Public Property AdvanceChance As Integer = 20

    ' New property to track the initial moments
    Public Property InitialMomentsCount As Integer = 3 ' Number of initial moments to showcase

    ' Selects a random duration for the timer interval
    Public Function GetRandomDuration() As Single
        Return RandomDurations(rand.Next(RandomDurations.Length))
    End Function

    ' Determines whether to advance based on a random chance
    Public Function ShouldAdvance() As Boolean
        Return rand.Next(100) < AdvanceChance
    End Function
    Public Function DetermineSpeedSetting() As Integer
        ' Implement your logic here to determine which framerate to use
        ' This is just an example logic; adjust it based on your requirements
        Dim randomValue As Double = New Random().NextDouble()
        If randomValue < 0.3 Then
            Return 0 ' Slow speed setting
        ElseIf randomValue < 0.6 Then
            Return 1 ' Medium speed setting
        Else
            Return 2 ' Fast speed setting
        End If
    End Function
    ' Returns a playback rate based on a speed index
    Public Function GetPlaybackRate(index As Byte) As Single
        If index >= 0 AndAlso index < PlaybackRates.Length Then
            Return PlaybackRates(index)
        End If
        Return 1.0R ' Default rate
    End Function
End Class

