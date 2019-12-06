Module PictureFunctions

#Region "Properties"
    Public Property iScreenstate As Byte = Screenstate.Fitted
    Public Property lSpeedfactor As Long = 10 'Larger is slower
    Public Property iZoomFactor As Integer = 100
    Private ZoomJustStarted As Boolean = True
    Public mediaLoopSize As Integer = 50
    Public picBlanker As PictureBox
    Dim strScreenState = {"Fitted", "True Size", "Zoomed"}

    Public ePicMousePoint As Point
    Public Property bImageDimensionState As Byte
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
    Public Function GetImage(strPath As String) As Image
        If strPath = "" Then Return Nothing
        If InStr(strPath, ".gif") = 0 Then
            Return LoadImage(strPath)

            ' Exit Function 'This Causes problems if extension is .gif
        Else
            Try
                Dim img As Image = Image.FromFile(strPath)
                Return img
            Catch ex As Exception
                'Reportfault("",ex.message)
                Return Nothing

            End Try
        End If
    End Function
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
    Public Sub PicClick(pic As PictureBox)
        iScreenstate = (iScreenstate + 1) Mod 2
        DisposePic(pic)
        Dim img As Image = GetImage(Media.MediaPath)
        PreparePic(pic, img)
    End Sub
    Public Sub Mousewheel(sender As Object, e As MouseEventArgs)
        Dim pbx1 As PictureBox = CType(sender, PictureBox)
        ePicMousePoint.X = e.X
        ePicMousePoint.Y = e.Y
        ePicMousePoint.X = ePicMousePoint.X + pbx1.Left
        ePicMousePoint.Y = ePicMousePoint.Y + pbx1.Top

        If iScreenstate = Screenstate.Fitted Then
            MainForm.AdvanceFile(e.Delta < 0, False)
            MainForm.tmrSlideShow.Enabled = False 'Break slideshow if scrolled
            Dim img As Image = GetImage(Media.MediaPath)
        Else


            If CtrlDown Then
                ZoomPicture(pbx1, e.Delta > 0, 2) 'TODO Options
            Else
                ZoomPicture(pbx1, e.Delta > 0, 15)
            End If
        End If

    End Sub

    Public Sub PlacePic(ByRef pbx As PictureBox)
        Dim outside As Control = pbx.Parent

        pbx.Left = outside.Width / 2 - pbx.Width / 2
        pbx.Top = outside.Height / 2 - pbx.Height / 2

    End Sub

#End Region
#Region "Methods"
    Public Sub MouseMove(sender As Object, e As MouseEventArgs)
        If TypeOf (sender) Is PictureBox Then
            Dim pbx1 As PictureBox = CType(sender, PictureBox)
            Dim mouse As Point
            mouse.X = e.X
            mouse.Y = e.Y
            mouse.X = mouse.X + pbx1.Left
            mouse.Y = mouse.Y + pbx1.Top
            If Not ShiftDown Then MovePic(mouse, pbx1, pbx1.Parent)
        End If
    End Sub
    Public Sub MovePic(mouse As Point, inside As Control, outside As Control)
        Dim x As Long
        Dim y As Long
        x = mouse.X
        y = mouse.Y
        Dim xdist As Long
        Dim ydist As Long
        Select Case bImageDimensionState
            Case PicState.TooWide
                xdist = x * (inside.Width - outside.Width) / outside.Width
                inside.Left = -xdist
            Case PicState.TooTall
                ydist = y * (inside.Height - outside.Height) / outside.Height
                inside.Top = -ydist
            Case PicState.Overscan
                ydist = y * (inside.Height - outside.Height) / outside.Height
                inside.Top = -ydist
                xdist = x * (inside.Width - outside.Width) / outside.Width
                inside.Left = -xdist

        End Select


        'If bImageDimensionState = PicState.TooWide Or bImageDimensionState = PicState.Overscan Then
        '    xdist = x * (inside.Width - outside.Width) / outside.Width
        '    inside.Left = -xdist
        'End If
        'If bImageDimensionState = PicState.TooTall Or bImageDimensionState = PicState.Overscan Then
        '    ydist = y * (inside.Height - outside.Height) / outside.Height
        '    inside.Top = -ydist
        'End If
    End Sub
    Public Sub MovePic(mouse As Point, blanker As PictureBox, inside As Control, outside As Control)
        Dim x As Long
        Dim y As Long
        ' mouse = PointToScreen(mouse)
        x = mouse.X
        y = mouse.Y
        'Move left
        Dim xdist As Long
        Dim ydist As Long

        If bImageDimensionState = PicState.TooWide Or bImageDimensionState = PicState.Overscan Then
            If x > blanker.Right Then
                'Move right
                xdist = x - blanker.Right
                'Large pic width
                If inside.Width > outside.Width Then
                    If inside.Left > outside.Width - inside.Width Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If
                Else 'Small pic width
                    If inside.Left > 0 Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If

                End If

            ElseIf x < blanker.Left Then
                xdist = x - blanker.Left

                'Large pic width
                If inside.Width > outside.Width Then
                    If inside.Left < 0 Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If
                Else 'small pic width
                    If inside.Left < outside.Width - inside.Width Then
                        inside.Left = inside.Left - xdist / lSpeedfactor
                    End If

                End If
            End If
        End If

        If bImageDimensionState = PicState.TooTall Or bImageDimensionState = PicState.Overscan Then

            If y < blanker.Top Then 'Move down
                ydist = y - blanker.Top
                'Large pic height
                If inside.Height > outside.Height Then
                    If inside.Top < 0 Then
                        inside.Top = inside.Top - ydist / lSpeedfactor
                    End If
                Else
                    If inside.Top < outside.Height - inside.Height Then
                        inside.Top = inside.Top - ydist / lSpeedfactor
                    End If
                End If
                '

            ElseIf y > blanker.Bottom Then 'Move Up
                ydist = y - blanker.Bottom
                'Large pic height
                If inside.Height > outside.Height Then
                    If inside.Top > outside.Height - inside.Height Then
                        inside.Top = inside.Top - 3 * ydist / lSpeedfactor
                    End If
                Else
                    If inside.Top > 0 Then
                        inside.Top = inside.Top - 3 * ydist / lSpeedfactor
                    End If
                End If

            End If


        End If
        Try

            inside.Refresh()
        Catch ex As Exception

        End Try

    End Sub
    'Public Sub ToggleTrueSize(pic As PictureBox)
    '    If iScreenstate = Screenstate.Fitted Then
    '        iScreenstate = Screenstate.TrueSize
    '    Else
    '        iScreenstate = Screenstate.Fitted
    '    End If
    '    DisposePic(pic)
    '    Dim img As Image = GetImage(Media.MediaPath)
    '    PreparePic(pic, picBlanker, img)
    'End Sub
    Public Sub DisposePic(box As PictureBox)
        If box.Image IsNot Nothing Then
            box.Image.Dispose()
            GC.SuppressFinalize(box)
            box.Image = Nothing
        End If
    End Sub
    Public Sub PreparePic(pbx As PictureBox, img As Image)
        Dim pbxBlanker As New PictureBox
        pbxBlanker.Parent = pbx.Parent
        If img Is Nothing Then Exit Sub
        If Not pbx.Image Is Nothing Then
            pbx.Image.Dispose()
        End If
        pbx.Image = img
        CentralCtrl(pbxBlanker, ZoneSize) 'insert central zone

        iZoomFactor = 100 * pbx.Width / pbx.Image.Width 'How much is image zoomed currently?

        'Make pic box same size as image (still within container)
        pbx.Width = pbx.Image.Width
        pbx.Height = pbx.Image.Height
        'How does it exceeed the container, if at all?
        If pbx.Parent IsNot Nothing Then
            bImageDimensionState = ClassifyImage(pbx.Parent.Width, pbx.Parent.Height, pbx.Width, pbx.Height)
        End If
        If Not iScreenstate = Screenstate.Fitted Then
        End If

        SetState(pbx, iScreenstate)

        PlacePic(pbx)

    End Sub
    Private Sub CentralCtrl(pbx As Control, proportion As Decimal)
        Dim ctr As Control = pbx.Parent
        With pbx
            .Width = ctr.Width * proportion
            .Height = ctr.Height * proportion
            .Left = (ctr.Width - .Width) / 2
            .Top = (ctr.Height - .Height) / 2

        End With
    End Sub



    Public Sub SetState(pbx As PictureBox, Sstate As Byte)
        'Sets the screenstate, docking style. Changes the sizemode of pbx
        'If iScreenstate = Sstate Then Exit Sub
        Select Case Sstate
            Case Screenstate.Fitted
                pbx.Dock = DockStyle.Fill
                pbx.SizeMode = PictureBoxSizeMode.Zoom
            Case Screenstate.TrueSize
                ' Exit Select
                pbx.Dock = DockStyle.None
                pbx.SizeMode = PictureBoxSizeMode.Normal
                'Expand or shrink the picture box accordingly
                pbx.Width = pbx.Image.Width
                pbx.Height = pbx.Image.Height
            Case Screenstate.Zoomed
                'If zooming from fitted
                'Resize the picture box
                'then
                'reset the size mode
                'And undock

                'If iScreenstate = Screenstate.Fitted Then
                If iScreenstate <> Screenstate.Zoomed Then
                    ' If ZoomJustStarted Then
                    Dim x As Control = pbx.Parent
                    pbx.Dock = DockStyle.None
                    pbx.Width = x.Width
                    pbx.Height = x.Height
                    pbx.Left = x.Left
                    pbx.Top = x.Top
                End If
                pbx.SizeMode = PictureBoxSizeMode.Zoom
                'PlacePic(pbx)
        End Select
        iScreenstate = Sstate
        MainForm.tsslPicState.Text = "Picture State:" & strScreenState(iScreenstate)
    End Sub
    Public Sub FadeInLabel(lbl As Label)
        lbl.Visible = True
    End Sub

    Public Sub ZoomPicture(pbx As PictureBox, blnEnlarge As Boolean, Percentage As Decimal)
        SetState(pbx, Screenstate.Zoomed)
        Dim Enlarge As Decimal = 1 + Percentage / 100
        Dim Reduce As Decimal = 1 - Percentage / 100
        Dim x As Control = pbx.Parent
        If blnEnlarge Then
            iZoomFactor = iZoomFactor * Enlarge
            pbx.Width = pbx.Width * Enlarge
            pbx.Left = pbx.Left - (ePicMousePoint.X - pbx.Left) * Percentage / 100

            pbx.Top = pbx.Top - (ePicMousePoint.Y - pbx.Top) * Percentage / 100
            pbx.Height = pbx.Height * Enlarge
            '        If pbx.Top < -pbx.Height Then MsgBox("Stop")
        Else
            iZoomFactor = iZoomFactor * Reduce
            If pbx.Width * Reduce >= x.Width Or pbx.Height * Reduce >= x.Height Then
                pbx.Width = pbx.Width * Reduce
                pbx.Left = Math.Min(0, pbx.Left + (ePicMousePoint.X - pbx.Left) * Percentage / 100)
                pbx.Top = Math.Min(0, pbx.Top + (ePicMousePoint.Y - pbx.Top) * Percentage / 100)
                pbx.Height = pbx.Height * Reduce
            End If
        End If

    End Sub

#End Region

End Module

