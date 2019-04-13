Imports AxWMPLib
Public Class MediaControl
    Private Sub MediaControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        wmp.uiMode = "None"
        wmp.Location = Me.Location
        PicBox.Location = Me.Location
    End Sub
    Private newFileName As String
    Public Property Filename() As String
        Get
            Return newFileName
        End Get
        Set(ByVal value As String)
            If Not IO.File.Exists(value) Then
                MsgBox("File not found")
            Else
                ProcessFile(value)

                newFileName = value
            End If
        End Set
    End Property

    Private Sub ProcessFile(value As String)
        fType = FindType(value)
        Select Case fType
            Case Filetype.Movie
                HandleMovie(blnRandomStartPoint)

            Case Filetype.Pic
                Dim img As Image
                If Not PicBox.Image Is Nothing Then
                    DisposePic(PicBox)
                End If
                If Not My.Computer.FileSystem.FileExists(strCurrentFilePath) Then Exit Select

                img = GetImage(newFileName)
                OrientPic(img)
                MovietoPic(img)
            Case Filetype.Unknown
        End Select
    End Sub
    Private Sub HandleMovie(blnRandom As Boolean)
        'If it is to jump to a random point, do not show first.
        wmp.Visible = True
        wmp.URL = strCurrentFilePath
        wmp.BringToFront()
        '   If PlaybackSpeed <> 1 Then wmp.settings.rate = PlaybackSpeed
        currentPicBox.Visible = False

    End Sub
    Private Sub MovietoPic(img As Image)
        wmp.Visible = False
        wmp.URL = ""
        PreparePic(PicBox, Blanker, img)
        PicBox.Visible = True
        PicBox.BringToFront()
        wmp.Visible = False
    End Sub
    Private Sub OrientPic(img As Image)
        Select Case ImageOrientation(img)
            Case ExifOrientations.BottomRight
                img.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case ExifOrientations.RightTop
                img.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case ExifOrientations.LeftBottom
                img.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select
    End Sub
    Public Sub MediaControl_Keydown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control

        'frmMain.HandleKeys(Me, e)
        e.suppresskeypress = True
    End Sub

    Public Sub wmp_PlayStateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles wmp.PlayStateChange
        PlaystateChange(sender, e)
    End Sub


    Private Sub picbox_Click(sender As Object, e As EventArgs) Handles PicBox.MouseClick

        picBlanker = Blanker
        PicClick(PicBox)
    End Sub

    Private Sub picbox_MouseWheel(sender As Object, e As MouseEventArgs) Handles PicBox.MouseWheel, Me.MouseWheel
        PictureFunctions.Mousewheel(PicBox, sender, e)
    End Sub

    Private Sub picbox_MouseMove(sender As Object, e As MouseEventArgs) Handles PicBox.MouseMove, Me.MouseMove
        picBlanker = Blanker
        PictureFunctions.MouseMove(PicBox, sender, e)
    End Sub

    Private Sub picbox_KeyDown(sender As Object, e As KeyEventArgs) Handles PicBox.KeyDown
        ShiftDown = e.Shift
        CtrlDown = e.Control
    End Sub

    Private Sub wmp_Enter(sender As Object, e As EventArgs) Handles wmp.Enter

    End Sub
End Class
