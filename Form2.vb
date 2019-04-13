Public Class Form2
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
    Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long


    Sub ScrnCap(Lt, Top, Rt, Bot)
        Dim rwidth, rheight, sourcedc, destdc, bhandle As Long

        rwidth = Rt - Lt
        rHeight = Bot - Top
        SourceDC = CreateDC("DISPLAY", 0, 0, 0)
        DestDC = CreateCompatibleDC(SourceDC)
        BHandle = CreateCompatibleBitmap(SourceDC, rWidth, rHeight)
        SelectObject(destdc, bhandle)
        BitBlt(destdc, 0, 0, rwidth, rheight, sourcedc, Lt, Top, &HCC0020)
        Dim Wnd As Long
        OpenClipboard(Wnd)
        EmptyClipboard
        SetClipboardData(2, bhandle)
        CloseClipboard
        DeleteDC(destdc)
        ReleaseDC(bhandle, sourcedc)
    End Sub

    '以下的示例把屏幕图象捕捉后，放到Picture1 中。
    Sub Command1_Click()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Me.Visible = False
        'ScrnCap(0, 0, 640, 480)
        'Me.Visible = True
        'PictureBox1.Image = Clipboard.GetData()
    End Sub
End Class