Public Class SpeedHandler
    Dim mIntervals() = {1000, 300, 200}
    Dim mFrameRates() = {5, 15, 20}
    Public Event SpeedChanged(mSlideshow As Boolean)
    Private mSlideshow As Boolean = False
    Public Property Slideshow() As Boolean
        Get
            Return mSlideshow
        End Get
        Set(ByVal value As Boolean)
            mSlideshow = value
            RaiseEvent SpeedChanged(mSlideshow)
        End Set
    End Property
    Private mSpeed As Byte
    Public Property Fullspeed As Boolean = True
    Public Property Paused As Boolean = False
    Public Property AbsoluteJump As Integer = 20
    Public Property FractionalJump As Integer = 12
    Public Property Speed() As Byte
        Get
            Return mSpeed
        End Get
        Set(ByVal value As Byte)
            mSpeed = value
            Fullspeed = False
            RaiseEvent SpeedChanged(mSlideshow)
        End Set
    End Property

    Private mInterval As Byte
    Public Property Interval() As Integer
        Get
            Return mIntervals(mSpeed)
        End Get
        Set(ByVal value As Integer)
            mIntervals(mSpeed) = value
        End Set
    End Property

    Private mFrameRate As Integer
    Public Property FrameRate() As Integer
        Get
            Return mFrameRates(mSpeed)
        End Get
        Set(ByVal value As Integer)
            mFrameRates(mSpeed) = value
        End Set
    End Property
    Public Sub IncreaseSpeed()
        If mSlideshow Then
            mIntervals(mSpeed) = mIntervals(mSpeed) * 0.9
        Else
            mFrameRates(mSpeed) = mFrameRates(mSpeed) * 1.1
        End If
        RaiseEvent SpeedChanged(mSlideshow)
    End Sub
    Public Sub DecreaseSpeed()
        If mSlideshow Then
            mIntervals(mSpeed) = mIntervals(mSpeed) * 1.1
        Else
            mFrameRates(mSpeed) = mFrameRates(mSpeed) * 0.9
        End If
        RaiseEvent SpeedChanged(mSlideshow)

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
                AbsoluteJump = AbsoluteJump * 1.1
            Else
                AbsoluteJump = AbsoluteJump * 0.9
            End If
        End If
    End Sub
End Class
