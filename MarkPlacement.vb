
Public Class MarkPlacement
    Private Property mBar As New PictureBox()
    Public WriteOnly Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Graphics = mBar.CreateGraphics
            'Barcolor = mBar.BackColor
        End Set
    End Property
    Public Property Barcolor
    Property Duration As Long
    Property Fractions As Integer
    Property SmallJumps As Long
    Property Graphics As Graphics
    Property Bitmap As Bitmap
    Private Property mMarkers As List(Of Double)
    Property Markers As List(Of Double)
        Get
            Markers = mMarkers
        End Get
        Set(value As List(Of Double))
            mMarkers = value

        End Set

    End Property

    Public Sub Create()

        Clear() 'Erase marks
        'Draw Big Jumps
        Dim start As Point
        start.Y = 0
        Dim endpt As Point
        endpt.Y = mBar.Height

        Dim count As Integer = Duration / Fractions
        Dim x As Integer
        While x < Duration * mBar.Width / Duration

            Dim pen As New Pen(Color.Black, 2)
            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            x += count * mBar.Width / Duration
        End While
        x = 0
        While x < Duration * mBar.Width / Duration

            Dim pen As New Pen(Color.Gray, 1)
            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            x += SmallJumps * mBar.Width / Duration
        End While


        If mMarkers Is Nothing Then Exit Sub
        For Each m In mMarkers
            start.X = mBar.Width * m / Duration
            endpt.X = start.X
            Dim pen As New Pen(Color.Yellow, 3)
            Graphics.DrawLine(pen, start, endpt)
        Next

    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
    End Sub
End Class