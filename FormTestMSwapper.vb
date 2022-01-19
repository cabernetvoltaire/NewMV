Public Class FormTestMSwapper
    Private MS2 As New MediaSwapper2
    Private LBH As New ListBoxHandler

    Private Sub FormTestMSwapper_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CreateMedia()
        If FormMain.LBH.ItemList.Count = 0 Then
            LBH = FormMain.FBH
        Else
            LBH = FormMain.LBH

        End If
        MS2.ListofFiles = LBH.ItemList
        MS2.CurrentIndex = LBH.CurrentIndex
        MS2.ShowFirst(True)
    End Sub
    Private Sub CreateMedia()
        For i = 0 To 2
            Dim wmp As New AxWindowsMediaPlayer
            Dim pic As New PictureBox
            Me.SplitContainer1.Panel1.Controls.Add(wmp)
            wmp.Height = 300
            wmp.Width = 300
            wmp.Left = 300 * i
            pic.Size = wmp.Size
            'wmp.Dock = DockStyle.Fill
            'wmp.stretchToFit = True
            Dim x As New MediaHandler("mp" & i) With {.Picture = pic, .Player = wmp, .Visible = True}
            wmp.uiMode = "None"
            wmp.settings.mute = True
            x.SPT = Media.SPT
            x.PlaceResetter(True)
            MS2.MediaHandlers.Add(x)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MS2.AdvanceList(True)
        LBH.SetNamed(MS2.Current)
    End Sub

    Private Sub FormTestMSwapper_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        FormMain.HandleKeys(sender, e)
    End Sub
End Class