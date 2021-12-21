Public Class Form6
    Private _List As New List(Of String)
    Public Property List() As List(Of String)
        Get
            Return _List
        End Get
        Set(ByVal value As List(Of String))
            _List = value
        End Set
    End Property
    Private Sub CreateMedia()
        Dim i As Integer = 0
        For Each m In _List
            If i > 25 Then Exit For
            Dim wmp As New AxWindowsMediaPlayer
            wmp.Height = 300
            wmp.Width = 300
            FlowLayoutPanel1.Controls.Add(wmp)
            Dim x As New MediaHandler("mp" & i) With {.Player = wmp, .MediaPath = m, .Visible = True}
            wmp.uiMode = "None"
            wmp.settings.mute = True
            wmp.Tag = m
            AddHandler wmp.MouseMoveEvent, AddressOf Mousover
            i += 1
        Next
    End Sub
    Private Sub Mousover(sender As Object, e As _WMPOCXEvents_MouseMoveEvent)
        Dim pb = DirectCast(sender, AxWindowsMediaPlayer)

        'ToolTip1.SetToolTip(pb, pb.Tag)
        FormMain.ControlSetFocus(FormMain.lbxFiles)
        FormMain.Random.OnDirChange = False
        FormMain.HighlightCurrent(pb.Tag)
        FormMain.lbxFiles.SelectionMode = SelectionMode.One
        'Me.Activate()
    End Sub
    Private Sub Form6_Activated(sender As Object, e As EventArgs) Handles Me.Activated
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
        FlowLayoutPanel1.Controls.Clear()
        CreateMedia()
    End Sub
End Class