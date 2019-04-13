Imports System.Threading

Public Class Thumbnails


    Public flp2 As New FlowLayoutPanel With {
        .BackColor = Color.Blue,
        .Dock = DockStyle.Bottom
    }
    Public t As Thread
    Private mList As List(Of String)
    Public Property List() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))
            mList = value
        End Set
    End Property
    Private mSize As Int16 = 200
    Private Sub LoadThumbnails()
        Dim flp As New FlowLayoutPanel With {
            .BackColor = Color.WhiteSmoke,
            .Dock = DockStyle.Fill
        }
        flp.SendToBack()
        flp.Visible = True
        flp.AutoScroll = True
        AddHandler flp.MouseMove, AddressOf flp_Mouseover

        Dim pics(mList.Count) As PictureBox
        Dim i As Int16 = 0


        For Each f In mList
            Dim mMedia As New MediaHandler("mMedia")
            mMedia.MediaPath = f
            Try
                If mMedia.MediaType = Filetype.Pic Then
                    Dim finfo = New IO.FileInfo(f)
                    pics(i) = New PictureBox
                    flp.Controls.Add(pics(i))
                    flp.Update()

                    With pics(i)

                        .Height = mSize

                        If finfo.Extension = ".gif" Then
                            .Image = Image.FromFile(f)

                        Else
                            .Image = GetThumb(p, .Height, f)
                        End If
                        If .Image IsNot Nothing Then

                            .Width = .Image.Width / .Image.Height * .Height
                            .SizeMode = PictureBoxSizeMode.StretchImage
                        End If
                        .Tag = f
                        AddHandler .MouseEnter, AddressOf pb_Mouseover

                        AddHandler .MouseClick, AddressOf pb_Click

                    End With
                End If
                '      Me.Refresh()
                i += 1
            Catch ex As System.InvalidOperationException
                Continue For
            End Try
        Next
        flp2 = flp
        '        flp2.Dock = DockStyle.Fill

        '   flp2.SendToBack()
        '  flp.Visible = False

    End Sub
    Private Sub flp_Mouseover(sender As Object, e As EventArgs)
        ToolTip1.SetToolTip(sender, sender.name)
    End Sub
    Private Sub pb_Mouseover(sender As Object, e As EventArgs)
        Dim pb = DirectCast(sender, PictureBox)

        ToolTip1.SetToolTip(pb, pb.Tag)
        Media.MediaPath = pb.Tag
        MainForm.lbxFiles.SelectionMode = SelectionMode.One
        MainForm.tmrPicLoad.Enabled = True

    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        Dim pb = DirectCast(sender, PictureBox)
        pb.Visible = False
        pb.Enabled = False
        MainForm.HandleKeys(MainForm, New KeyEventArgs(KeyDelete))

    End Sub
    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function GetThumb(ByVal e As PaintEventArgs, h As Long, f As String) As Image
        Try

            Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
            Dim myBitmap As New Bitmap(f)
            Dim ratio As Single = myBitmap.Height / myBitmap.Width
            Dim myThumbnail As Image = myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
            myBitmap.Dispose()
            Return myThumbnail
            e.Graphics.DrawImage(myThumbnail, 150, 75)
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private p As PaintEventArgs
    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        'p = e


        Loadthumbs()


    End Sub

    Private Sub Loadthumbs()
        Me.Controls.Add(flp2)
        t = New Thread(New ThreadStart(Sub() LoadThumbnails()))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)

        t.Start()
        Timer1.Enabled = True

    End Sub




    Public Property ThumbnailHeight() As Int16
        Get
            Return mSize
        End Get
        Set(ByVal value As Int16)
            mSize = value
        End Set
    End Property

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        If t.IsAlive Then
            '    Me.Refresh()
            Exit Sub
        ElseIf Not t.IsAlive Then

            Timer1.Enabled = False
            Me.Refresh()
        End If

    End Sub

    Private Sub Thumbnails_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class