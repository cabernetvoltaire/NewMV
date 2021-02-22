Public Class SpeedHandler
    Public Intervals() As Integer = {1000, 300, 10}
    Public FrameRates() As Byte = {5, 15, 20}
    Public Event SpeedChanged(sender As Object, e As EventArgs)
    Private mSlideshow As Boolean = False
    Public Property Slideshow() As Boolean
        Get
            Return mSlideshow
        End Get
        Set(ByVal value As Boolean)
            mSlideshow = value
            RaiseEvent SpeedChanged(Me, Nothing)
        End Set
    End Property
    Private mSpeed As Byte
    Public Property Fullspeed As Boolean

    Private Sub PauseVideo(Pause As Boolean)
        If Pause Then
            Media.Player.Ctlcontrols.pause()
            PausedPosition = Media.Player.Ctlcontrols.currentPosition
            Fullspeed = False
        Else
            Media.Player.Ctlcontrols.currentPosition = PausedPosition
            Media.Player.Ctlcontrols.play()
            PausedPosition = 0
        End If
        mPaused = Pause

    End Sub

    Private mPaused As Boolean '= False
    Public PausedPosition As Long
    Public Property Paused() As Boolean
        Get
            Return mPaused
        End Get
        Set(ByVal value As Boolean)
            mPaused = value
            PauseVideo(mPaused)

            'If mPaused Then
            '    PausedPosition = Media.Player.Ctlcontrols.currentPosition
            '    Fullspeed = False
            'Else
            '    PausedPosition = 0
            'End If
        End Set
    End Property
    Public Property AbsoluteJump As Integer ' = 35
        Get
            Return _AbsoluteJump
        End Get
        Set
            _AbsoluteJump = Value
        End Set
    End Property

    Public Property FractionalJump As Integer ' = 8
        Get
            Return _FractionalJump
        End Get
        Set
            _FractionalJump = Value
        End Set
    End Property

    Public Property Speed() As Byte
        Get
            Return mSpeed
        End Get
        Set(ByVal value As Byte)
            mSpeed = value
            FrameRate = FrameRates(mSpeed)
            Fullspeed = False
            RaiseEvent SpeedChanged(Me, Nothing)
        End Set
    End Property
    Private Property mSSSpeed As Byte
    Public Property SSSpeed() As Byte
        Get
            Return mSSSpeed
        End Get
        Set(ByVal value As Byte)
            mSSSpeed = value
            Interval = Intervals(mSSSpeed)
            RaiseEvent SpeedChanged(Me, Nothing)
        End Set
    End Property


    Private mInterval As Byte
    Public Property Interval() As Integer
        Get
            Return Intervals(mSSSpeed)
        End Get
        Set(ByVal value As Integer)
            Intervals(mSSSpeed) = value
        End Set
    End Property

    Private mFrameRate As Integer
    Private _Unpause As Boolean = False
    Private _AbsoluteJump As Integer = 35
    Private _FractionalJump As Integer = 8

    Public Property FrameRate() As Integer
        Get
            Return FrameRates(mSpeed)
        End Get
        Set(ByVal value As Integer)
            FrameRates(mSpeed) = value
            RaiseEvent SpeedChanged(Me, Nothing)
        End Set
    End Property
    Public Sub IncreaseSpeed()
        If mSlideshow Then
            Intervals(mSpeed) = Intervals(mSpeed) * 0.9
        Else
            FrameRates(mSpeed) = FrameRates(mSpeed) * 1.1
        End If
        RaiseEvent SpeedChanged(Me, Nothing)
    End Sub
    Public Sub DecreaseSpeed()
        If mSlideshow Then
            Intervals(mSpeed) = Intervals(mSpeed) * 1.1
        Else
            FrameRates(mSpeed) = FrameRates(mSpeed) * 0.9
        End If
        RaiseEvent SpeedChanged(Me, Nothing)

    End Sub

    Public Sub ChangeJump(Proportional As Boolean, Up As Boolean)
        If Proportional Then
            If Up Then
                FractionalJump = FractionalJump + 1
            Else
                If FractionalJump > 2 Then
                    FractionalJump = FractionalJump - 1
                End If
            End If

        Else
            If Up Then
                AbsoluteJump = AbsoluteJump + 1
            Else
                If AbsoluteJump > 2 Then
                    AbsoluteJump = AbsoluteJump - 1

                End If
            End If
        End If
    End Sub
End Class
