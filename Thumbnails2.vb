Imports System.ComponentModel
Imports System.Threading

Public Class Thumbnails


    Public Frame As Long
    Public Event ThumbnailCreated(i As Short)
    Private Duration As TimeSpan
    Private mList As List(Of String)
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
    Private Function LoadThumbnails() As FlowLayoutPanel

        Dim flp As New FlowLayoutPanel With {
            .BackColor = Color.WhiteSmoke,
            .Dock = DockStyle.Fill
        }
        '        flp.SendToBack()
        flp.Visible = True
        flp.AutoScroll = True
        AddHandler flp.MouseMove, AddressOf flp_Mouseover

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

                If typ = Filetype.Pic Or typ = Filetype.Movie Then
                    Dim finfo = New IO.FileInfo(f)
                    flp.Controls.Add(pics(i))
                    flp.Update()

                    With pics(i)

                        .Height = ThumbWidth
                        Select Case typ
                            Case Filetype.Pic
                                If finfo.Extension = ".gif" Then
                                    .Image = Image.FromFile(f)

                                Else
                                    .Image = GetThumb(PaintArgs, .Height, f)
                                End If
                            Case Filetype.Movie
                                .Image = GetThumb(PaintArgs, .Height, f, True)
                        End Select
                        If .Image IsNot Nothing Then

                            .Width = .Image.Width / .Image.Height * .Height
                            .SizeMode = PictureBoxSizeMode.StretchImage
                        Else
                            .Image = pics(i).InitialImage
                        End If
                        .Tag = f

                        AddHandler .MouseEnter, AddressOf pb_Mouseover

                        AddHandler .MouseClick, AddressOf pb_Click
                        .Refresh()
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
        MainForm.HighlightCurrent(pb.Tag)
        MainForm.lbxFiles.SelectionMode = SelectionMode.One
        'MainForm.tmrPicLoad.Enabled = True

    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        If MsgBox("Delete file?", MsgBoxStyle.YesNo, "Metavisua") = MsgBoxResult.Yes Then
            Dim pb = DirectCast(sender, PictureBox)
            pb.Visible = False
            pb.Enabled = False
            MainForm.HandleKeys(MainForm, New KeyEventArgs(KeyDelete))
        End If
    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function GetThumb(ByVal e As PaintEventArgs, h As Long, f As String, Optional Movie As Boolean = False) As Image
        Dim myThumbnail As Image '= myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
        If Movie Then
            Try
                Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)

                Dim VT As New VideoThumbnailer

                VT.Fileref = f
                VT.ThumbnailHeight = h

                myThumbnail = Image.FromFile(VT.GetThumbnail(f, Frame))
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
        If t.IsAlive Then
            '   Me.Refresh()
            Exit Sub
        ElseIf Not t.IsAlive Then
            Duration = TimeOperation(False)
            Timer1.Enabled = False
            Me.Refresh()
        End If

    End Sub



    Private Sub Thumbnails_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        MainForm.HandleKeys(sender, e)
    End Sub

    Private Sub Thumbnails_Click(sender As Object, e As EventArgs) Handles Me.Click
        ' Loadthumbs()
    End Sub

    Private Sub Thumbnails_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        'Dim d As New IO.DirectoryInfo("Q:\Thumbs")
        'For Each f In d.GetFiles
        '    f.Delete()
        'Next
    End Sub

    Private Sub Slider_Scroll(sender As Object, e As EventArgs) Handles Slider.ValueChanged
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
End Class