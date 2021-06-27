
Public Class MarkPlacement
    Private Property mBar As New PictureBox()
    Public WriteOnly Property Bar As PictureBox
        Set(value As PictureBox)
            mBar = value
            Graphics = mBar.CreateGraphics
            'Barcolor = mBar.BackColor
        End Set
    End Property
    Private Property Barcolor
    Property Duration As Long
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

        Clear()
        If mMarkers Is Nothing Then Exit Sub
        For Each m In mMarkers
            Dim start As Point
            start.Y = 0
            start.X = mBar.Width * m / Duration
            Dim endpt As Point
            endpt.Y = mBar.Height
            endpt.X = start.X
            Dim pen As New Pen(Color.Black, 1)
            Graphics.DrawLine(pen, start, endpt)
        Next
        If mMarkers.Count <> 0 Then
            'Graphics.Clear(mBar.Parent.BackColor)

        End If

    End Sub

    Public Sub Clear()
        'Dim g As Graphics = Bar.CreateGraphics
        Graphics.Clear(mBar.BackColor)
    End Sub
End Class