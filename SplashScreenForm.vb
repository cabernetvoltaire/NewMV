Public Class SplashScreenForm
    Public Sub New()
        InitializeComponent()
    End Sub
    Public Sub UpdateStatus(statusText As String)
        lblLoading.Text = statusText
        Application.DoEvents() ' Ensure the label is updated immediately
    End Sub


End Class