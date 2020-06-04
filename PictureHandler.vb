Public Class PictureHandler
#Region "Properties"
    Public Event StateChanged(sender As Object, e As EventArgs)
    Public Event ZoomChange(sender As Object, e As EventArgs)
    Public Event AdvanceFile(sender As Object, e As EventArgs)
    Public RandomNext As Boolean
    Private WithEvents mPicBox As New PictureBox
    Private tb As New ToolTip
    Public MouseZone As New PictureBox
    Public Property PicBox() As PictureBox
        Get
            Return mPicBox
        End Get
        Set(ByVal value As PictureBox)
            mPicBox = value
            PicBoxContainer = mPicBox.Parent
        End Set
    End Property
    Private mPicImage As Image
    Public Property WheelScroll As Boolean = True
    Public Property TextOutput As String
    Public Property PicBoxContainer As New Control
    Private Property mZoomfactor As Integer = 100
    Public Property ZoomFactor As Integer
        Set(value As Integer)

            mZoomfactor = value

        End Set
        Get
            mZoomfactor = mPicBox.Width / mPicBox.Image.Width * 100
            Return mZoomfactor
        End Get
    End Property
    Private Property mState As Byte
    Public Property State As Byte
        Set(value As Byte)
            mState = SetState(value)
        End Set
        Get
            Return mState
        End Get
    End Property


    Dim ScreenStateDescriptions = {"Fitted", "True Size", "Zoomed"}

    Public ePicMousePoint As Point

    Public Property ImageDimensionState As Byte
    Public Property ShiftDown As Boolean
    Public Property CtrlDown As Boolean
    Public Property AltDown As Boolean
    Public Property KeyDownFlag As Boolean


#End Region

#Region "Enums and structures"
    Public Enum Screenstate As Byte
        Fitted
        TrueSize
        Zoomed
    End Enum
    Private Enum PicState As Byte
        Overscan
        Underscan
        TooWide
        TooTall

    End Enum

#End Region
#Region "Functions"
    Public Sub GetImage(strPath As String)
        If strPath = "" Then
        Else
            'If Not strPath.EndsWith(".gif") Then
            '    mPicImage = LoadImage(strPath)
            'Else
            Try
                Dim stream As New System.IO.FileStream(strPath, IO.FileMode.Open)
                Dim img As Image = Image.FromStream(stream)
                stream.Close()
                mPicImage = img
                OrientPic(mPicImage)
                PicBox.Image = mPicImage
                PicBox.Tag = strPath
                'tb.SetToolTip(PicBox, strPath)
            Catch ex As Exception
                mPicImage = Nothing
            End Try
            'End If
            'tb.AutoPopDelay = 1
            'PicBoxContainer.Controls.Add(MouseZone)
            'MouseZone.Width = PicBoxContainer.Width / 2
            'MouseZone.Height = PicBoxContainer.Height / 2
            'MouseZone.Left = MouseZone.Width / 4
            'MouseZone.Top = MouseZone.Top / 4
        End If
    End Sub
    Private Function ClassifyImage(ViewPortWidth As Long, ViewPortHeight As Long, ImageWidth As Long, ImageHeight As Long) As Byte
        If ViewPortWidth > ImageWidth Then 'Viewport wider than pic
            If ViewPortHeight > ImageHeight Then
                ClassifyImage = PicState.Underscan
            Else
                ClassifyImage = PicState.TooTall
            End If
        Else 'Viewport narrower than pic
            If ViewPortHeight > ImageHeight Then
                ClassifyImage = PicState.TooWide
            Else 'Viewport shorter than pic
                ClassifyImage = PicState.Overscan
            End If
        End If
        Return ClassifyImage
    End Function

#End Region
#Region "Events"
    Public Sub New(p As PictureBox)
        mPicBox = p


    End Sub
    Public Sub New()

    End Sub
    Public Sub PicClick(sender As Object, e As MouseEventArgs) Handles mPicBox.Click
        If e.Button = MouseButtons.Left Then
            WheelScroll = Not WheelScroll
            If WheelScroll Then
                mPicBox.Cursor = Cursors.Arrow
            Else
                mPicBox.Cursor = Cursors.Cross
            End If
        End If

    End Sub
    Private Sub PicFullScreen() Handles mPicBox.DoubleClick
        If ShiftDown Then
            blnSecondScreen = True
        Else
            blnSecondScreen = False
        End If
        MainForm.GoFullScreen(blnFullScreen)

    End Sub
    Public Sub Mousewheel(sender As Object, e As MouseEventArgs) Handles mPicBox.MouseWheel

        ePicMousePoint.X = e.X
        ePicMousePoint.Y = e.Y
        ePicMousePoint.X = ePicMousePoint.X + mPicBox.Left
        ePicMousePoint.Y = ePicMousePoint.Y + mPicBox.Top

        If WheelScroll Then 'State = Screenstate.Fitted Then
            'RaiseEvent AdvanceFile(sender, e)
            MainForm.AdvanceFile(e.Delta < 0, RandomNext)
            'MainForm.tmrSlideShow.Enabled = False 'Break slideshow if scrolled
            'Dim img As Image = Me.GetImage(Media.MediaPath)
        Else
            ZoomPicture(e.Delta > 0, 15)
        End If

    End Sub

    'Public Sub PlacePic()
    '    'Dim outside As Control = PicBoxContainer
    '    mPicBox.Left = PicBoxContainer.Width / 2 - mPicBox.Width / 2
    '    mPicBox.Top = PicBoxContainer.Height / 2 - mPicBox.Height / 2

    'End Sub

#End Region
#Region "Methods"
    Public Sub MouseMove(sender As Object, e As MouseEventArgs) Handles mPicBox.MouseMove
        If mState = Screenstate.Fitted Then Exit Sub
        'When the mouse moves
        Dim mouse As Point
        mouse.X = e.X
        mouse.Y = e.Y
        'Adjust the point
        mouse.X = mouse.X + mPicBox.Left
        mouse.Y = mouse.Y + mPicBox.Top

        'Move the picture using the new point

        MovePic(mouse, mPicBox, mPicBox.Parent)


    End Sub
    Private Sub MovePic(mouse As Point, inside As Control, outside As Control)
        Dim x As Long
        Dim y As Long
        x = mouse.X
        y = mouse.Y
        Dim xdist As Long
        Dim ydist As Long
        Select Case ImageDimensionState
            Case PicState.TooWide
                xdist = x * (inside.Width - outside.Width) / outside.Width
                inside.Left = -xdist
            Case PicState.TooTall
                ydist = y * (inside.Height - outside.Height) / outside.Height
                inside.Top = -ydist
            Case PicState.Overscan, PicState.Underscan
                ydist = y * (inside.Height - outside.Height) / (outside.Height)
                inside.Top = -ydist
                xdist = x * (inside.Width - outside.Width) / (outside.Width)
                inside.Left = -xdist

        End Select

    End Sub

    'Public Sub PreparePic() 'Not actually implemented at the moment
    '    mPicBox.Width = mPicBox.Image.Width
    '    mPicBox.Height = mPicBox.Image.Height
    '    'How does it exceeed the container, if at all?
    '    If PicBoxContainer IsNot Nothing Then
    '        ImageDimensionState = ClassifyImage(PicBoxContainer.Width, PicBoxContainer.Height, mPicBox.Width, mPicBox.Height)
    '    End If

    '    SetState(mState)

    '    PlacePic()

    'End Sub

    Public Function SetState(Sstate As Byte) As Byte
        ' Exit Sub
        'Sets the screenstate, docking style. Changes the sizemode of pbx
        'If iScreenstate = Sstate Then Exit Sub

        Select Case Sstate
            Case Screenstate.Fitted
                mPicBox.Dock = DockStyle.Fill
                mPicBox.SizeMode = PictureBoxSizeMode.Zoom

               ' mWheelScroll = True
            Case Screenstate.TrueSize

                mPicBox.Dock = DockStyle.None
                mPicBox.SizeMode = PictureBoxSizeMode.Zoom
                'Expand or shrink the picture box accordingly
                mPicBox.Width = mPicBox.Image.Width
                mPicBox.Height = mPicBox.Image.Height
            Case Screenstate.Zoomed
                'If zooming from fitted
                'Resize the picture box
                'then
                'reset the size mode
                'And undock

                If State <> Screenstate.Zoomed Then
                    mPicBox.Dock = DockStyle.None
                    mPicBox.Width = PicBoxContainer.Width
                    mPicBox.Height = PicBoxContainer.Height
                    mPicBox.Left = PicBoxContainer.Left
                    mPicBox.Top = PicBoxContainer.Top
                End If
                mPicBox.SizeMode = PictureBoxSizeMode.Zoom
                '  mWheelScroll = False
                'PlacePic(mPicBox)
        End Select
        ' PicTest.Label1.Text = "Picture State:" & ScreenStateDescriptions(State) & " Zoom factor: " & ZoomFactor
        Return Sstate
    End Function

    Public Sub ZoomPicture(blnEnlarge As Boolean, Percentage As Decimal)
        mState = SetState(Screenstate.Zoomed)
        Dim Enlargement As Decimal = 1 + Percentage / 100
        Dim Reduction As Decimal = 1 - Percentage / 100

        If blnEnlarge Then
            If mPicBox.Width < 12 * PicBoxContainer.Width Then
                mZoomfactor = mZoomfactor * Enlargement
                mPicBox.Width = mPicBox.Width * Enlargement
                mPicBox.Left = mPicBox.Left - (ePicMousePoint.X - mPicBox.Left) * Percentage / 100

                mPicBox.Top = mPicBox.Top - (ePicMousePoint.Y - mPicBox.Top) * Percentage / 100
                mPicBox.Height = mPicBox.Height * Enlargement
                '        If mPicBox.Top < -mPicBox.Height Then MsgBox("Stop")

                RaiseEvent ZoomChange(Me, Nothing)
            End If
        Else
            mZoomfactor = mZoomfactor * Reduction
            If mPicBox.Width * Reduction > PicBoxContainer.Width Or mPicBox.Height * Reduction > PicBoxContainer.Height Then
                mPicBox.Width = mPicBox.Width * Reduction
                mPicBox.Left = Math.Min(0, mPicBox.Left + (ePicMousePoint.X - mPicBox.Left) * Percentage / 100)
                mPicBox.Top = Math.Min(0, mPicBox.Top + (ePicMousePoint.Y - mPicBox.Top) * Percentage / 100)
                mPicBox.Height = mPicBox.Height * Reduction
                RaiseEvent ZoomChange(Me, Nothing)
            Else
                mState = SetState(PictureHandler.Screenstate.Fitted)
                WheelScroll = True
                RaiseEvent StateChanged(Me, Nothing)
            End If
        End If
        ' PicTest.Label1.Text = "Picture State:" & ScreenStateDescriptions(State) & " Zoom factor: " & ZoomFactor
    End Sub


#End Region



End Class

