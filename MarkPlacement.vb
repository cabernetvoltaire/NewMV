
Public Class MarkPlacement
    Private WithEvents tmr As New Timer With {.Enabled = False, .Interval = 100}
    Private Property mBar As New PictureBox()
    Private TooltipText As String = "Remove markers using CTRL"
    Public WriteOnly Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Graphics = mBar.CreateGraphics
            'Barcolor = mBar.BackColor

            AddToolTip(mBar, New ToolTip, TooltipText)
        End Set
    End Property
    Public Property Barcolor
    Property Duration As Long
    Property Fractions As Integer
    Property SmallJumps As Long
    Property Graphics As Graphics
    ' Property Bitmap As Bitmap
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
        tmr.Enabled = True
    End Sub

    Private Sub DrawMarks() Handles tmr.Tick

        Clear() 'Erase marks
        Dim start As Point
        start.Y = 0
        Dim endpt As Point
        endpt.Y = mBar.Height
        Duration = Math.Max(Duration, 1)
        'Draw Big Jumps
        Dim JumpSize As Integer = Duration / Fractions
        If JumpSize = 0 Then JumpSize = Math.Max(1, Duration / 2)
        Dim x As Integer
        While x < mBar.Width

            Dim pen As New Pen(Color.Black, 2)
            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            x += JumpSize * mBar.Width / Duration
        End While
        x = 0
        'Draw small jumps
        While x < mBar.Width

            Dim pen As New Pen(Color.Green, 1)
            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            x += SmallJumps * mBar.Width / Duration
        End While

        'Draw Markers
        If mMarkers Is Nothing Then Exit Sub
        For Each m In mMarkers
            start.X = mBar.Width * m / Duration
            endpt.X = start.X
            Dim pen As New Pen(Color.Yellow, 3)
            Graphics.DrawLine(pen, start, endpt)
        Next
        tmr.Enabled = False
    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
    End Sub
End Class