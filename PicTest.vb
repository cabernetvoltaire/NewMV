Public Class PicTest
    Dim PH As New PictureHandler()

    Private Sub PicTest_Load(sender As Object, e As EventArgs) Handles Me.Load
        PH.PicBox = PictureBox1
        PH.GetImage("E:\Pictures\Pictures\1522253386990.jpg")
        '  PH.PreparePic()
    End Sub

    Private Sub PicTest_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown, Me.KeyUp
        PH.AltDown = e.Alt
        PH.ShiftDown = e.Shift
        PH.CtrlDown = e.Control

    End Sub
End Class