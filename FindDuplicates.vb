Imports System.IO
Public Class FindDuplicates
    Private mDuplicates As List(Of List(Of String))
    Public Property Maxfiles As Integer = 8000
    Private mList As List(Of String)
    Public WriteOnly Property List() As List(Of String)
        Set(ByVal value As List(Of String))
            mList = value
            Dim nList = New List(Of String)
            For Each m In mList
                If Len(m) > 240 Then Exit For
                Dim inf = New IO.FileInfo(m)
                If inf.Exists Then
                    nList.Add(m)
                Else
                End If
            Next
            mList = nList
            t.InputList = mList
            mDuplicates = t.Duplicates()
            Me.Text = mDuplicates.Count & " files with duplicates found. Click to remove from delete list."
            If mDuplicates.Count > 0 Then

            Else
                Exit Property
            End If
        End Set
    End Property

    Public ReadOnly Property DuplicatesCount() As Int16
        Get
            Return mDuplicates.Count
        End Get
    End Property

    Private t As New Duplicates

    Private mThumbnailHeight As Integer = 200
    Public Property ThumbnailHeight() As Integer
        Get
            Return mThumbnailHeight
        End Get
        Set(ByVal value As Integer)
            mThumbnailHeight = value
        End Set
    End Property

    Private Sub LoadThumbnails()
        Dim i = 0
        For Each m In mDuplicates
            If i > Maxfiles Then Exit Sub
            Dim f As New MediaHandler("f")
            f.MediaPath = m(0)
            Panel3.Controls.Add(FPCreator(m, f.MediaType))
            i += 1
        Next

    End Sub

    Private Function FPCreator(list As List(Of String), mType As Filetype) As FlowLayoutPanel

        Dim x As New FlowLayoutPanel
        x.Dock = DockStyle.Top
        Dim i = 0

        For Each m In list
            Select Case mType
                Case Filetype.Pic, Filetype.Gif
                    Dim r As New PictureBox
                    If mType = Filetype.Gif Then
                        r.Image = Image.FromFile(m)
                    Else
                        With r
                            .Height = mThumbnailHeight
                            .Image = Thumbnails.GetThumb(p, mThumbnailHeight, m)
                            If Not .Image Is Nothing Then
                                If x.Controls.Count > 0 Then
                                    Dim x2 As PictureBox = x.Controls(0)
                                    If AreSameImage(.Image, x2.Image) Then
                                        r.BorderStyle = BorderStyle.None
                                    Else
                                        RemoveThumbnail(r)
                                    End If
                                End If
                                .Width = .Image.Width / .Image.Height * .Height
                                .SizeMode = PictureBoxSizeMode.StretchImage
                            End If
                        End With

                    End If

                    r.Tag = m
                    r.Tag = r.Tag & "(" & New FileInfo(m).Length & " bytes)"
                    x.Controls.Add(r)
                    Dim r2 As New PictureBox
                    r2.Width = 400 - r.Width
                    If i = 0 Then x.Controls.Add(r2)
                    Me.Refresh()
                    AddHandler r.MouseMove, AddressOf previewover
                    AddHandler r.Click, AddressOf picdoubleclick
                Case Filetype.Movie
                    Dim r As New AxWMPLib.AxWindowsMediaPlayer
                    r.InitializeLifetimeService()
                    r.Width = mThumbnailHeight
                    r.Height = mThumbnailHeight
                    x.Controls.Add(r)
                    If i = 0 Then x.Controls.Add(New PictureBox)

                    r.Tag = m
                    '    r.Tag = r.Tag & "(" & New FileInfo(m).Length & " bytes)"
                    '   r.URL = m
                    r.Visible = True
                    AddHandler r.MouseDownEvent, AddressOf previewover
                    AddHandler r.DoubleClickEvent, AddressOf picdoubleclick
            End Select
            i += 1
        Next

        Return x


    End Function
    Private Function AreSameImage(ByVal I1 As Image, ByVal I2 As Image) As Boolean
        Dim BM1 As Bitmap = I1
        Dim BM2 As Bitmap = I2
        If BM1 Is Nothing Or BM2 Is Nothing Then
            Return True
        Else
            For X = 1 To BM1.Width - 1
                For y = 1 To BM2.Height - 1
                    If BM1.GetPixel(X, y) <> BM2.GetPixel(X, y) Then
                        Return False
                    End If
                Next
            Next
            Return True
        End If
    End Function
    Private Function CreateDeleteList() As List(Of String)
        Dim mDeleteList As New List(Of String)
        Dim i As Integer
        For Each m In mDuplicates
            i = 0
            For Each s In m
                If i > 0 And i <= Maxfiles Then mDeleteList.Add(s)
                i += 1
            Next
        Next
        Return mDeleteList
    End Function

    Private p As PaintEventArgs


#Region "Events"
    Private Sub previewover(sender As Object, e As AxWMPLib._WMPOCXEvents_MouseDownEvent)
        ToolTipDups.SetToolTip(sender, sender.tag)
        Dim m As AxWMPLib.AxWindowsMediaPlayer
        m = sender
        If m.URL = "" Then
            m.URL = m.Tag
            m.Ctlcontrols.currentPosition = m.currentMedia.duration / 3
            m.Ctlcontrols.play()
            sender = m
        Else
            sender.url = ""
        End If
    End Sub

    Private Sub picdoubleclick(sender As Object, e As AxWMPLib._WMPOCXEvents_DoubleClickEvent)
        If e.nShiftState Then

            sender.visible = False
        End If
    End Sub
    Private Sub previewover(sender As Object, e As MouseEventArgs)
        ToolTipDups.SetToolTip(sender, sender.tag)
        Media.MediaPath = sender.tag
        MainForm.lbxFiles.SelectionMode = SelectionMode.One
        MainForm.tmrPicLoad.Enabled = True
    End Sub

    Private Sub picdoubleclick(sender As Object, e As MouseEventArgs)
        RemoveThumbnail(sender)
    End Sub

    Private Sub RemoveThumbnail(sender As Object)
        sender.visible = False
        Dim x As String = sender.tag
        For Each m In mDuplicates
            m.Remove(x)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        For Each m In CreateDeleteList()
            ' Dim k = New FileInfo(m)
            Me.Text = "Deleting file - " & m
            Deletefile(m)
        Next
        Me.Close()
    End Sub

    Private Sub FindDuplicates_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        If mDuplicates.Count > 0 Then

            LoadThumbnails()
        End If
    End Sub

    Private Sub Panel3_Paint(sender As Object, e As PaintEventArgs) Handles Panel3.Paint

    End Sub

#End Region
End Class