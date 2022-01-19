Public Class FormMovieScreen

    Private _List As New List(Of String)
    Private _ListIndex As Integer = 0
    Private _NumberOfMedia As Integer = 24
    Private _Mediahandlers As New List(Of MediaHandler)
    Public Property List() As List(Of String)
        Get
            Return _List
        End Get
        Set(ByVal value As List(Of String))
            _List = value
        End Set
    End Property
    Private Sub CreateMedia()
        Dim i As Integer
        For Each m In _List
            If i > _NumberOfMedia - 1 Then Exit For
            Dim wmp As New AxWindowsMediaPlayer
            Dim pic As New PictureBox
            wmp.Height = 300
            wmp.Width = 300
            pic.Size = wmp.Size
            FlowLayoutPanel1.Controls.Add(wmp)
            wmp.uiMode = "None"
            wmp.Tag = m
            Dim x As New MediaHandler("mp" & i) With {.Picture = pic, .Player = wmp, .MediaPath = m, .Visible = True, .SPT = Media.SPT}
            x.Mute()
            _Mediahandlers.Add(x)

            AddHandler wmp.MouseMoveEvent, AddressOf Mousover
            i += 1
        Next
    End Sub
    Private Sub Mousover(sender As Object, e As _WMPOCXEvents_MouseMoveEvent)
        Dim pb = DirectCast(sender, AxWindowsMediaPlayer)


        FormMain.ControlSetFocus(FormMain.lbxFiles)
        FormMain.Random.OnDirChange = False
        FormMain.HighlightCurrent(pb.Tag)
        FormMain.lbxFiles.SelectionMode = SelectionMode.One
        'Me.Activate()
    End Sub


    Private Sub Form6_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Dim scr As Screen
        If blnSecondScreen Then
            scr = Screen.AllScreens(1)
        Else
            scr = Screen.AllScreens(0)
        End If
        Me.Location = scr.Bounds.Location + New Point(100, 100)
        Me.Left = scr.Bounds.Left
        Me.Top = scr.Bounds.Top
        CreateMedia()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        List = FormMain.FBH.ItemList
        CycleMedia(_NumberOfMedia)

    End Sub
    Private Sub CycleMedia(n As Integer)
        For i = 0 To n - 1
            Dim mh As MediaHandler = _Mediahandlers(i)
            If _ListIndex + i >= List.Count Then
                _ListIndex = _ListIndex - List.Count
            End If
            mh.MediaPath = List(_ListIndex + i)
            mh.Player.Tag = List(_ListIndex + i)
            mh.SPT = Media.SPT
            mh.Mute()
        Next
        _ListIndex = AdvanceModular(List.Count, _ListIndex, True, _NumberOfMedia)

    End Sub

    Private Sub FormMovieScreen_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class