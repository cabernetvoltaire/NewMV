Public Class KeyHandler
    Public Event Waiting(Sender As Object, e As EventArgs)
    Public Event CancelDisplay(sender As Object, e As EventArgs)
    Private mKeyEvent As KeyEventArgs
    Public Property e() As KeyEventArgs
        Get
            Return mKeyEvent
        End Get
        Set(ByVal value As KeyEventArgs)
            mKeyEvent = value
        End Set
    End Property
#Region "Default Key Assignments"
    Public KeyJumpToPoint = Keys.Insert
    Public KeyCycleStartPoint = Keys.OemQuestion
    Public KeyCycleFilter = Keys.Scroll
    Public KeyCycleNavMoveState = Keys.Oem7
    Public KeyCycleSortOrder = Keys.Pause
    Public KeyFullscreen = Keys.Scroll + Keys.Control
    Public KeyNextFile = Keys.PageDown
    Public KeyPreviousFile = Keys.PageUp

    Public KeyJumpRandom = Keys.Divide
    Public KeyTraverseTreeBack = Keys.Subtract
    Public KeyTraverseTree = Keys.Add
    Public KeyMarkFavourite = Keys.NumPad7
    Public KeyBigJumpBack = Keys.NumPad8
    Public KeyBigJumpOn = Keys.NumPad9
    Public KeyJumpToMark = Keys.NumPad4
    Public KeySmallJumpDown = Keys.NumPad5
    Public KeySmallJumpUp = Keys.NumPad6
    Public KeySpeed1 = Keys.NumPad1
    Public KeySpeed2 = Keys.NumPad2
    Public KeySpeed3 = Keys.NumPad3
    Public KeyToggleSpeed = Keys.NumPad0
    Public KeyMuteToggle = Keys.Decimal

    Public LKeyFullscreen = Keys.F + Keys.Control
    Public LKeyRandomize = Keys.R + Keys.Control
    Public LKeyTraverseTreeBack = Keys.Control + Keys.OemMinus
    Public LKeyTraverseTree = Keys.Control + Keys.Oemplus
    Public LKeyJumpAutoT = Keys.Control + Keys.Divide
    Public LKeyMuteToggle = Keys.Control + Keys.Decimal
    Public LKeyNextFile = Keys.Control + Keys.N
    Public LKeyPreviousFile = Keys.Control + Keys.P
    Public LKeyZoomIn = Keys.Control + Keys.Z
    Public LKeyFolderJump = Keys.Control + Keys.Multiply
    Public LKeyToggleSpeed = Keys.Control + Keys.D0
    Public LKeySpeed1 = Keys.Control + Keys.D1
    Public LKeySpeed2 = Keys.Control + Keys.D2
    Public LKeySpeed3 = Keys.Control + Keys.D3
    Public LKeySmallJumpDown = Keys.Control + Keys.D5
    Public LKeySmallJumpUp = Keys.Control + Keys.D6
    Public LKeyJumpToPoint = Keys.Control + Keys.D7
    Public LKeyBigJumpBack = Keys.Control + Keys.D8
    Public LKeyBigJumpOn = Keys.Control + Keys.D9
    Public LKeyMarkPoint = Keys.Control + Keys.D4


    'Common Keys
    Public KeyBackUndo = Keys.Back
    Public KeyReStartSS = Keys.Space
    Public KeyLoopToggle = Keys.OemCloseBrackets + Keys.Alt
    Public KeyTrueSize = Keys.Oemplus
    Public KeyRotate = Keys.OemCloseBrackets
    Public KeyRotateBack = Keys.OemOpenBrackets
    Public KeyAddFile = Keys.Oemtilde

    Public KeyZoomOut = Keys.Oemcomma
    Public KeyToggleThumbs = Keys.Oem3

    Public KeyToggleButtons = Keys.OemSemicolon
    Public KeyEscape = Keys.Escape
    Public KeySelect = Keys.F3
    Public KeyDelete = Keys.Delete

#End Region
End Class
