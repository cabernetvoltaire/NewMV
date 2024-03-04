Public Class Progress
    Inherits ProgressBar
    Property Bar As ProgressBar
    Public Sub Advance()
        Value += 1
    End Sub

    Private Sub InitializeComponent()
        Me.SuspendLayout()
        Me.ResumeLayout(False)

    End Sub
End Class
