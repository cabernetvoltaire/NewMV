
Public Class MarkPlacement
    Property Bar As PictureBox
    Property Duration As Long
    Property Bitmap As Bitmap
    Property Markers As List(Of Long)
    Public Sub Create()
        Dim graphics As Graphics = Bar.CreateGraphics
        For Each m In Markers
            Dim start As Point
            start.Y = 0
            start.X = Bar.Width * m / Duration
            Dim endpt As Point
            endpt.Y = Bar.Height
            endpt.X = start.X
            Dim pen As New Pen(Color.Black, 2)
            graphics.DrawLine(pen, start, endpt)
        Next

        Bitmap = Bar.Image
    End Sub

    Public Sub Clear()
        Dim g As Graphics = Bar.CreateGraphics
        g.Clear(Bar.BackColor)
    End Sub
End Class