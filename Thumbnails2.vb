Public Class Thumbnails

    Public Frame As Long
    Public Event ThumbnailCreated(i As Short)
    Private Duration As TimeSpan
    Private mList As List(Of String)
    Private mRefresh As Boolean = False
    Private mMsg As String = " Creating thumbnails, please wait..."
    Public Property List() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))
            mList = value
        End Set
    End Property
    Public ThumbWidth As Int16
    Public Pics() As PictureBox
    Private Function LoadThumbnails() As FlowLayoutPanel

        Dim flp As New FlowLayoutPanel With {
            .BackColor = Color.WhiteSmoke,
            .Dock = DockStyle.Fill
        }
        '        flp.SendToBack()
        flp.Visible = True
        flp.AutoScroll = True
        AddHandler flp.MouseHover, AddressOf flp_Mouseover

        Dim pics(mList.Count) As PictureBox
        Dim i As Int16 = 0

        i = CreateThumb(flp, pics, i)
        Return flp
    End Function

    Private Function CreateThumb(ByRef flp As FlowLayoutPanel, pics() As PictureBox, i As Short) As Short
        For Each f In mList
            pics(i) = New PictureBox
            Try
                Dim typ As Filetype = FindType(f)
                If typ = Filetype.Pic Or typ = Filetype.Movie Or typ = Filetype.Link Then
                    Dim finfo = New IO.FileInfo(f)
                    flp.Controls.Add(pics(i))
                    '    flp.Update()

                    With pics(i)
                        .Height = ThumbWidth
                        If finfo.Extension = ".gif" Then
                            .Image = Image.FromFile(f)
                        Else

                            .Image = GetThumb(PaintArgs, .Height, f, typ)
                        End If
                        If .Image IsNot Nothing Then
                            .Width = .Image.Width / .Image.Height * .Height
                            .SizeMode = PictureBoxSizeMode.StretchImage
                        End If
                        If typ = Filetype.Link Then
                            f = LinkTarget(f)
                        End If
                        .Tag = f

                        AddHandler .MouseEnter, AddressOf pb_Mouseover
                        AddHandler .MouseDoubleClick, AddressOf pb_Click
                        '.Refresh()
                    End With
                End If
                RaiseEvent ThumbnailCreated(i)
                i += 1
            Catch ex As System.InvalidOperationException
                Continue For
            End Try
        Next
        Return i
    End Function

    Private Sub flp_Mouseover(sender As Object, e As EventArgs)
        ToolTip1.SetToolTip(sender, sender.name)
    End Sub
    Private Sub pb_Mouseover(sender As Object, e As EventArgs)
        Dim pb = DirectCast(sender, PictureBox)

        ToolTip1.SetToolTip(pb, pb.Tag)
        FormMain.ControlSetFocus(FormMain.lbxFiles)
        FormMain.Random.OnDirChange = False
        FormMain.HighlightCurrent(pb.Tag)
        FormMain.lbxFiles.SelectionMode = SelectionMode.One
        Me.Activate()
        'MainForm.tmrPicLoad.Enabled = True

    End Sub
    Private Sub pb_Click(sender As Object, e As MouseEventArgs)
        If MsgBox("Delete file?", MsgBoxStyle.YesNo, "Metavisua") = MsgBoxResult.Yes Then
            Dim pb = DirectCast(sender, PictureBox)
            pb.Visible = False
            pb.Enabled = False
            FormMain.HandleKeys(FormMain, New KeyEventArgs(KeyDelete))
        End If
    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function GetThumb(ByVal e As PaintEventArgs, h As Long, f As String, type As Filetype) As Image
        Dim myThumbnail As Image '= myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)

        If type = Filetype.Movie Or type = Filetype.Link Then
            Select Case type
                Case Filetype.Movie
                Case Filetype.Link
                    Dim tmpFr As Long = Frame
                    If InStr(f, "%") <> 0 Then
                        Dim s As String()
                        s = f.Split("%")
                        Frame = Val(s(1))
                    Else
                        Frame = tmpFr
                    End If
                    f = LinkTarget(f)
            End Select

            Try
                Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                Dim finfo As New IO.FileInfo(ThumbnailName(f))
                Dim THRef As String
                If Not finfo.Exists Or mRefresh Then
                    Dim VT As New VideoThumbnailer
                    VT.Fileref = f
                    VT.ThumbnailHeight = h
                    THRef = VT.GetThumbnail(f, Frame)
                Else
                    THRef = ThumbnailName(f)
                End If
                myThumbnail = Image.FromFile(THRef)
            Catch ex As Exception
                'MsgBox(ex.Message)
                Return Nothing
            End Try

        Else
            Try
                Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
                Dim myBitmap As New Bitmap(f)
                Dim ratio As Single = myBitmap.Height / myBitmap.Width
                myThumbnail = myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
                myBitmap.Dispose()
                Return myThumbnail
                e.Graphics.DrawImage(myThumbnail, ThumbWidth, CType(ThumbWidth / 2, Single))
            Catch ex As Exception
                Return Nothing
            End Try
        End If
        Return myThumbnail
        e.Graphics.DrawImage(myThumbnail, ThumbWidth, CType(ThumbWidth / 2, Single))
    End Function

    Private PaintArgs As PaintEventArgs
    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        Me.Text = Me.Text & mMsg
        Duration = TimeOperation(True)
        PaintArgs = e

        OnThumbnailed()

        '  Loadthumbs()


    End Sub

    Private Sub Loadthumbs()

        OnThumbnailed()


    End Sub

    Private Async Sub OnThumbnailed()
        If TableLayoutPanel1.GetControlFromPosition(0, 0) Is Nothing Then
            Dim flp3 As New FlowLayoutPanel
            flp3 = Await Task.Run(Function() LoadThumbnails())

            TableLayoutPanel1.Controls.Add(flp3, 0, 0)

            Duration = TimeOperation(False)
            Dim dDurStr As String = Duration.Seconds.ToString & "." & Duration.Milliseconds.ToString
            Me.Text = Me.Text.Replace(mMsg, " (" & dDurStr & "s to thumbnail " & List.Count & " files.)")
        End If

    End Sub

    Public Sub OnThumbnailCreated(i As Short) Handles Me.ThumbnailCreated

        'flp.Refresh()
    End Sub


    Public Property ThumbnailHeight() As Int16
        Get
            Return ThumbWidth
        End Get
        Set(ByVal value As Int16)
            ThumbWidth = value
        End Set
    End Property

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'If t.IsAlive Then
        '    '   Me.Refresh()
        '    Exit Sub
        'ElseIf Not t.IsAlive Then
        '    Duration = TimeOperation(False)
        '    Timer1.Enabled = False
        '    Me.Refresh()
        'End If

    End Sub



    Private Sub Thumbnails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        FormMain.HandleKeys(sender, e)
    End Sub



    Private Sub Slider_Scroll(sender As Object, e As EventArgs) Handles Slider.ValueChanged
        If List IsNot Nothing Then
            If List.Count <= 50 Then ResizePics(Slider.Value)
        End If
    End Sub
    Private Sub Slider_MouseUp(sender As Object, e As MouseEventArgs) Handles Slider.MouseUp
        ResizePics(Slider.Value)
    End Sub

    Private Sub ResizePics(value As Integer)
        For Each c In TableLayoutPanel1.Controls
            If TypeOf (c) Is FlowLayoutPanel Then
                For Each cc In c.controls
                    If TypeOf (cc) Is PictureBox Then
                        Dim ratio As Decimal = cc.Height / cc.Width
                        cc.Height = value * ThumbnailHeight / ((Slider.Minimum + Slider.Maximum) / 2)
                        cc.Width = value * ThumbnailHeight / (ratio * ((Slider.Minimum + Slider.Maximum) / 2))

                    End If
                Next
            End If
        Next
    End Sub

    Private Sub btnFindDupImages_Click(sender As Object, e As EventArgs) Handles btnFindDupImages.Click
        Dim TestList As New List(Of PictureBox)
        For Each c In TableLayoutPanel1.GetControlFromPosition(0, 0).Controls
            If TypeOf (c) Is PictureBox Then
                TestList.Add(c)
            End If
        Next
        Dim MatchingPics As New List(Of List(Of PictureBox))
        MatchingPics = PicsWithSameImages(TestList)
        For Each m In MatchingPics
            For Each n In m
                DealWithMatches(n)
            Next
        Next
    End Sub

    Friend Function PicsWithSameImages(ByRef PicList As List(Of PictureBox)) As List(Of List(Of PictureBox))
        Dim ReturnList As New List(Of List(Of PictureBox))
        Dim i, j As Integer
        For i = 0 To PicList.Count - 1
            Dim CurrentPairs As New List(Of PictureBox)
            If i < PicList.Count - 1 Then

                For j = i + 1 To PicList.Count - 1
                    Dim x As Image = PicList(i).Image
                    Dim y As Image = PicList(j).Image
                    If AreSameImage(x, y, False) Then
                        CurrentPairs.Add(PicList(i))
                        CurrentPairs.Add(PicList(j))

                    End If
                Next
                If Not ReturnList.Contains(CurrentPairs) Then
                    ReturnList.Add(CurrentPairs)
                End If
            End If
        Next
        Return ReturnList
    End Function
    Friend Sub DealWithMatches(x As PictureBox)
        x.Width = x.Width * 1.5
        x.Height = x.Height * 1.5
        Dim flp As FlowLayoutPanel = x.Parent
        '   flp.Controls.SetChildIndex(y, 0)
        flp.Controls.SetChildIndex(x, 0)
        ' flp.Controls(0).Tag = x.Tag


    End Sub

    Private Sub Refresh_Click(sender As Object, e As EventArgs) Handles RefreshThumbnails.Click
        mRefresh = True
        LoadThumbnails()
        mRefresh = False
    End Sub
End Class