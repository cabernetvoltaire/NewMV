Public Class KeyHandler
    Public Event Waiting(Obj As Object, e As EventArgs)
    Public Event CancelDisplay(Obj As Object, e As EventArgs)
    Private mKeyEvent As KeyEventArgs
    Public Property e() As KeyEventArgs
        Get
            Return mKeyEvent
        End Get
        Set(ByVal value As KeyEventArgs)
            mKeyEvent = value
        End Set
    End Property

End Class
