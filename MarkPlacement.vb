
Public Class MarkPlacement
    Property Bar As PictureBox
    Property Duration As Long
    Property Bitmap As Bitmap
    Property Markers As List(Of Long)
    Public Sub Create()
        Dim graphics As Graphics = Bar.CreateGraphics
        graphics.Clear(Color.Aqua)
        For Each m In Markers
            Dim start As Point
            start.Y = 0
            start.X = Bar.Width * m / Duration
            Dim endpt As Point
            endpt.Y = Bar.Height
            endpt.X = start.X
            Dim pen As New Pen(Color.Black, 1)
            graphics.DrawLine(pen, start, endpt)
        Next
        Bar.Visible = True

        '   Bitmap = Bar.Image
    End Sub

    Public Sub Clear(c As Color)
        Dim g As Graphics = Bar.CreateGraphics
        g.Clear(c)
        '  Bar.Visible = False
    End Sub
End Class