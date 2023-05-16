
Public Class MarkPlacement
    Private WithEvents tmr As New Timer With {.Enabled = False, .Interval = 300}
    Private Property mBar As New PictureBox()
    Private TooltipText As String = "Remove markers using CTRL"
    Public WriteOnly Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Bitmap = New Bitmap(mBar.Width, mBar.Height)
            Graphics = Graphics.FromImage(Bitmap)
            mBar.Width = mBar.Parent.Width

            AddToolTip(mBar, New ToolTip, TooltipText)
        End Set
    End Property
    Public Property Barcolor
    Property Duration As Long
        Get
            Return _Duration

        End Get
        Set(value As Long)
            _Duration = value
            DrawMarks()
        End Set
    End Property
    Private _Duration As Long

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
    Private mouseDownTimePosition As Double?
    Private mouseUpTimePosition As Double?

    Public Sub New()
        ' Attach the event handlers to the PictureBox control
        AddHandler mBar.MouseDown, AddressOf PictureBox_MouseDown
        AddHandler mBar.MouseUp, AddressOf PictureBox_MouseUp
        AddHandler mBar.SizeChanged, AddressOf PictureBox_SizeChanged
    End Sub
    Private Sub PictureBox_SizeChanged(sender As Object, e As EventArgs)
        ' Recreate the bitmap when the PictureBox is resized
        Bitmap = New Bitmap(mBar.Width, mBar.Height)
        Graphics = Graphics.FromImage(Bitmap)

        ' Redraw the marks
        DrawMarks()
        ' Update the PictureBox
        mBar.Image = Bitmap
    End Sub
    Private Sub PictureBox_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            mouseDownTimePosition = e.X * Duration / mBar.Width
        End If
    End Sub

    Private Sub PictureBox_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso mouseDownTimePosition.HasValue Then
            mouseUpTimePosition = e.X * Duration / mBar.Width

            ' Draw connected circles at the mouse down and up time positions
            DrawConnectedCircles(mouseDownTimePosition.Value, mouseUpTimePosition.Value)

            ' Reset the time positions
            mouseDownTimePosition = Nothing
            mouseUpTimePosition = Nothing
        End If
    End Sub
    Public Sub Create()
        DrawMarks()
    End Sub

    Private Sub DrawMarks() Handles tmr.Tick
        If Duration = 0 Or Duration > 10 ^ 6 Then Exit Sub
        Graphics.Clear(mBar.BackColor)
        Dim start As Point
        start.Y = 0
        Dim endpt As Point
        endpt.Y = mBar.Height
        '  Duration = Math.Min(Duration, 50000)
        'Draw Big Jumps
        Dim JumpSize As Integer = Duration / Fractions
        Dim px As Double = 0
        While px < mBar.Width
            Dim pen As New Pen(Color.Black, 2)
            Dim x As Integer = CInt(px)

            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            px += mBar.Width / Fractions
        End While
        px = 0
        'Draw small jumps
        While px < mBar.Width
            Dim pen As New Pen(Color.Green, 1)
            Dim x As Integer = CInt(px)

            start.X = x
            endpt.X = x
            Graphics.DrawLine(pen, start, endpt)
            px += SmallJumps * mBar.Width / Duration
        End While

        'Draw Markers
        If mMarkers Is Nothing Then Exit Sub
        For Each m In mMarkers
            start.X = mBar.Width * m / Duration
            endpt.X = start.X
            Dim pen As New Pen(Color.Yellow, 3)
            Graphics.DrawLine(pen, start, endpt)
        Next
        mBar.Image = Bitmap
        'Draw StartEnd
        If VideoTrim.Active Then DrawConnectedCircles(VideoTrim.StartTime, VideoTrim.Finish)
        'tmr.Enabled = False
    End Sub
    Public Sub DrawConnectedCircles(timePosition1 As Double, timePosition2 As Double)
        Dim circleRadius As Integer = 3
        Dim verticalOffset As Integer = mBar.Height \ 2

        Dim xPosition1 As Integer = CInt(mBar.Width * timePosition1 / Duration)
        Dim xPosition2 As Integer = CInt(mBar.Width * timePosition2 / Duration)

        Dim circle1Center As New Point(xPosition1, verticalOffset)
        Dim circle2Center As New Point(xPosition2, verticalOffset)

        Dim circlePen As New Pen(Color.Red, 1)
        Dim circleBrush As New SolidBrush(Color.Red)

        Graphics.FillEllipse(circleBrush, circle1Center.X - circleRadius, circle1Center.Y - circleRadius, circleRadius * 2, circleRadius * 2)
        Graphics.FillEllipse(circleBrush, circle2Center.X - circleRadius, circle2Center.Y - circleRadius, circleRadius * 2, circleRadius * 2)

        Graphics.DrawLine(circlePen, circle1Center, circle2Center)
    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
        mBar.Image = Bitmap
    End Sub
End Class
Public Class MarkPlacementOld
    Private WithEvents tmr As New Timer With {.Enabled = False, .Interval = 300}
    Private Property mBar As New PictureBox()
    Private TooltipText As String = "Remove markers using CTRL"
    Public WriteOnly Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Graphics = mBar.CreateGraphics
            mBar.Width = mBar.Parent.Width
            'Barcolor = mBar.BackColor

            AddToolTip(mBar, New ToolTip, TooltipText)
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
    Private mouseDownTimePosition As Double?
    Private mouseUpTimePosition As Double?

    Public Sub New()
        ' Attach the event handlers to the PictureBox control
        AddHandler mBar.MouseDown, AddressOf PictureBox_MouseDown
        AddHandler mBar.MouseUp, AddressOf PictureBox_MouseUp
    End Sub

    Private Sub PictureBox_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            mouseDownTimePosition = e.X * Duration / mBar.Width
        End If
    End Sub

    Private Sub PictureBox_MouseUp(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left AndAlso mouseDownTimePosition.HasValue Then
            mouseUpTimePosition = e.X * Duration / mBar.Width

            ' Draw connected circles at the mouse down and up time positions
            DrawConnectedCircles(mouseDownTimePosition.Value, mouseUpTimePosition.Value)

            ' Reset the time positions
            mouseDownTimePosition = Nothing
            mouseUpTimePosition = Nothing
        End If
    End Sub
    Public Sub Create()
        tmr.Enabled = True
    End Sub

    Private Sub DrawMarks() Handles tmr.Tick
        If Duration = 0 Or Duration > 10 ^ 6 Then Exit Sub
        Clear() 'Erase marks
        Dim start As Point
        start.Y = 0
        Dim endpt As Point
        endpt.Y = mBar.Height
        '  Duration = Math.Min(Duration, 50000)
        'Draw Big Jumps
        Dim JumpSize As Integer = Math.Max(10, Duration / Fractions)
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
        'Draw StartEnd
        If VideoTrim.Active Then DrawConnectedCircles(VideoTrim.StartTime, VideoTrim.Finish)
        tmr.Enabled = False
    End Sub
    Public Sub DrawConnectedCircles(timePosition1 As Double, timePosition2 As Double)
        Dim circleRadius As Integer = 3
        Dim verticalOffset As Integer = mBar.Height \ 2

        Dim xPosition1 As Integer = CInt(mBar.Width * timePosition1 / Duration)
        Dim xPosition2 As Integer = CInt(mBar.Width * timePosition2 / Duration)

        Dim circle1Center As New Point(xPosition1, verticalOffset)
        Dim circle2Center As New Point(xPosition2, verticalOffset)

        Dim circlePen As New Pen(Color.Red, 1)
        Dim circleBrush As New SolidBrush(Color.Red)

        Graphics.FillEllipse(circleBrush, circle1Center.X - circleRadius, circle1Center.Y - circleRadius, circleRadius * 2, circleRadius * 2)
        Graphics.FillEllipse(circleBrush, circle2Center.X - circleRadius, circle2Center.Y - circleRadius, circleRadius * 2, circleRadius * 2)

        Graphics.DrawLine(circlePen, circle1Center, circle2Center)
    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
    End Sub
End Class