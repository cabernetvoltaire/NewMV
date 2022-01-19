Public Class Form6
    Private ax As New Axes(Me.PictureBox1)
    Private gr As Graphics

    Private Sub Form6_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub Form6_Click(sender As Object, e As EventArgs) Handles Me.Click
        ax.Draw()
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles Me.Load
    End Sub
End Class
Public Class Axes
    Private pb As PictureBox
    Private gr As Graphics
    Private x As Integer
    Private y As Integer
    Private _Origin As Point

    Public Sub New(picture As PictureBox)
        pb = picture
        gr = pb.CreateGraphics
        x = pb.Width / 15
        y = pb.Height / 15
        _Origin.X = x
        _Origin.Y = y
    End Sub
    Public Sub Draw()
        Xaxis()
        Yaxis()
    End Sub
    Private Sub Xaxis()
        Dim start As Point
        start = _Origin
        Dim endpt As Point
        endpt.X = pb.Width / 15 * 13
        Dim pen As New Pen(Color.Black, 1)
        gr.DrawLine(pen, start, endpt)
    End Sub
    Private Sub Yaxis()
        Dim start As Point
        start = _Origin
        Dim endpt As Point
        endpt.Y = pb.Height / 15 * 13
        Dim pen As New Pen(Color.Black, 1)
        gr.DrawLine(pen, start, endpt)

    End Sub
End Class