
Public Class MarkPlacement
    Private Property mBar As New PictureBox()
    Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Graphics = mBar.CreateGraphics
        End Set
        Get
            Return mBar
        End Get
    End Property

    Property Duration As Long
    Property Graphics As Graphics
    Property Bitmap As Bitmap
    Property Markers As List(Of Long)
    Public Sub Create()

        Clear()
        For Each m In Markers
            Dim start As Point
            start.Y = 0
            start.X = mBar.Width * m / Duration
            Dim endpt As Point
            endpt.Y = mBar.Height
            endpt.X = start.X
            Dim pen As New Pen(Color.Black, 1)
            Graphics.DrawLine(pen, start, endpt)
        Next


    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
    End Sub
End Class