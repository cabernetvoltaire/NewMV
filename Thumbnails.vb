Imports System.ComponentModel
Imports System.IO

Public Class Thumbnails

    Public Sub LoadThumbnails(Flist As List(Of String), e As PaintEventArgs)

        Dim pics(Flist.Count) As PictureBox
        Dim i As Int16 = 0
        For Each f In Flist
            If FindType(f) = 0 Then
                Dim finfo = New IO.FileInfo(f)
                pics(i) = New PictureBox
                FlowLayoutPanel1.Controls.Add(pics(i))
                With pics(i)

                    .Height = mSize

                    If finfo.Extension = ".gif" Then
                        .Image = Image.FromFile(f)
                    Else
                        .Image = Example_GetThumb(e, .Height, f)
                    End If
                    .Width = .Image.Width / .Image.Height * .Height
                    .SizeMode = PictureBoxSizeMode.StretchImage
                    .Tag = f
                    AddHandler .MouseEnter, AddressOf pb_Click
                End With
                Me.Update()
            End If
            i += 1
        Next

    End Sub
    Private Sub pb_Click(sender As Object, e As EventArgs)
        Dim pb = DirectCast(sender, PictureBox)
        strCurrentFilePath = pb.Tag

        frmMain.lbxFiles.SelectionMode = SelectionMode.One
        frmMain.tmrPicLoad.Enabled = True

    End Sub

    Public Function ThumbnailCallback() As Boolean
        Return False
    End Function

    Public Function Example_GetThumb(ByVal e As PaintEventArgs, h As Long, f As String) As Image
        Dim myCallback As New Image.GetThumbnailImageAbort(AddressOf ThumbnailCallback)
        Dim myBitmap As New Bitmap(f)
        Dim ratio As Single = myBitmap.Height / myBitmap.Width
        Dim myThumbnail As Image = myBitmap.GetThumbnailImage(h, h * ratio, myCallback, IntPtr.Zero)
        myBitmap.Dispose()
        Return myThumbnail
        'e.Graphics.DrawImage(myThumbnail, 150, 75)
    End Function


    Private Sub Thumbnails_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        LoadThumbnails(flist, e)
    End Sub

    Private flist As List(Of String)
    Public Property FileList() As List(Of String)
        Get
            Return flist
        End Get
        Set(ByVal value As List(Of String))
            flist = value

        End Set
    End Property
    Private mSize As Int16 = 100
    Public Property ThumbnailHeight() As Int16
        Get
            Return mSize
        End Get
        Set(ByVal value As Int16)
            mSize = value
        End Set
    End Property
End Class