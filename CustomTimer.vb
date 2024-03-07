Public Class CustomTimer
    Private WithEvents timer As New System.Windows.Forms.Timer

    ' Properties to hold the parameters
    Public Property SmallJump As Boolean
    Public Property ForwardJump As Boolean

    ' Custom event with custom EventArgs
    Public Event Tick(sender As Object, e As CustomTickEventArgs)

    Public Sub New(interval As Integer)
        timer.Interval = interval
    End Sub

    Public Sub StartTimer()
        timer.Start()
    End Sub

    Public Sub StopTimer()
        timer.Stop()
    End Sub

    Private Sub Timer_Tick(sender As Object, e As EventArgs) Handles timer.Tick
        ' Raise the custom event with the parameters
        RaiseEvent Tick(Me, New CustomTickEventArgs(SmallJump, ForwardJump))
    End Sub
End Class

Public Class CustomTickEventArgs
    Inherits EventArgs

    Public ReadOnly Property SmallJump As Boolean
    Public ReadOnly Property ForwardJump As Boolean

    Public Sub New(smallJump As Boolean, forwardJump As Boolean)
        Me.SmallJump = smallJump
        Me.ForwardJump = forwardJump
    End Sub
End Class
