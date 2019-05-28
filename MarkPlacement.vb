
Public Class MarkPlacement
    Property Frame As PictureBox
    Property Duration As Long
    Property Bitmap As Bitmap
    Property Markers As List(Of Long)
    Public Sub Create()
        Dim graphics As Graphics = Frame.CreateGraphics
        For Each m In Markers
            Dim start As Point
            start.Y = 0
            start.X = Frame.Width * m / Duration
            Dim endpt As Point
            endpt.Y = Frame.Height
            endpt.X = start.X
            Dim pen As New Pen(Color.Black, 1)
            graphics.DrawLine(pen, start, endpt)
            Bitmap = Frame.Image
        Next
    End Sub
    Public Sub Clear()
        Dim g As Graphics = Frame.CreateGraphics
        g.Clear(Frame.BackColor)
    End Sub
End Class
