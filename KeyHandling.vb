Public Module KeyHandling
    'Numpad options (not available on Laptop)
    Public KeyAssignments As New Dictionary(Of String, String)
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



    Public Sub InitialiseKeys()
        AssignKey("KeyJumpToPoint ", Keys.Insert)
        AssignKey("KeyCycleStartPoint ", Keys.OemQuestion)
        AssignKey("KeyCycleFilter ", Keys.Scroll)
        AssignKey("KeyCycleNavMoveState ", Keys.Oem7)
        AssignKey("KeyCycleSortOrder ", Keys.Pause)
        AssignKey("KeyFullscreen ", Keys.Scroll + Keys.Control)
        AssignKey("KeyNextFile ", Keys.PageDown)
        AssignKey("KeyPreviousFile ", Keys.PageUp)
        AssignKey("KeyJumpRandom ", Keys.Divide)
        AssignKey("KeyTraverseTreeBack ", Keys.Subtract)
        AssignKey("KeyTraverseTree ", Keys.Add)
        AssignKey("KeyMarkFavourite ", Keys.NumPad7)
        AssignKey("KeyBigJumpBack ", Keys.NumPad8)
        AssignKey("KeyBigJumpOn ", Keys.NumPad9)
        AssignKey("KeyJumpToMark ", Keys.NumPad4)
        AssignKey("KeySmallJumpDown ", Keys.NumPad5)
        AssignKey("KeySmallJumpUp ", Keys.NumPad6)
        AssignKey("KeySpeed1 ", Keys.NumPad1)
        AssignKey("KeySpeed2 ", Keys.NumPad2)
        AssignKey("KeySpeed3 ", Keys.NumPad3)
        AssignKey("KeyToggleSpeed ", Keys.NumPad0)
        AssignKey("KeyMuteToggle ", Keys.Decimal)
        AssignKey("LKeyFullscreen ", Keys.F + Keys.Control)
        AssignKey("LKeyRandomize ", Keys.R + Keys.Control)
        AssignKey("LKeyTraverseTreeBack ", Keys.Control + Keys.OemMinus)
        AssignKey("LKeyTraverseTree ", Keys.Control + Keys.Oemplus)
        AssignKey("LKeyJumpAutoT ", Keys.Control + Keys.Divide)
        AssignKey("LKeyMuteToggle ", Keys.Control + Keys.Decimal)
        AssignKey("LKeyNextFile ", Keys.Control + Keys.N)
        AssignKey("LKeyPreviousFile ", Keys.Control + Keys.P)
        AssignKey("LKeyZoomIn ", Keys.Control + Keys.Z)
        AssignKey("LKeyFolderJump ", Keys.Control + Keys.Multiply)
        AssignKey("LKeyToggleSpeed ", Keys.Control + Keys.D0)
        AssignKey("LKeySpeed1 ", Keys.Control + Keys.D1)
        AssignKey("LKeySpeed2 ", Keys.Control + Keys.D2)
        AssignKey("LKeySpeed3 ", Keys.Control + Keys.D3)
        AssignKey("LKeySmallJumpDown ", Keys.Control + Keys.D5)
        AssignKey("LKeySmallJumpUp ", Keys.Control + Keys.D6)
        AssignKey("LKeyJumpToPoint ", Keys.Control + Keys.D7)
        AssignKey("LKeyBigJumpBack ", Keys.Control + Keys.D8)
        AssignKey("LKeyBigJumpOn ", Keys.Control + Keys.D9)
        AssignKey("LKeyMarkPoint ", Keys.Control + Keys.D4)


        AssignKey("KeyBackUndo ", Keys.Back)
        AssignKey("KeyReStartSS ", Keys.Space)
        AssignKey("KeyLoopToggle ", Keys.OemCloseBrackets + Keys.Alt)
        AssignKey("KeyTrueSize ", Keys.Oemplus)
        AssignKey("KeyRotate ", Keys.OemCloseBrackets)
        AssignKey("KeyRotateBack ", Keys.OemOpenBrackets)
        AssignKey("KeyAddFile ", Keys.Oemtilde)
        AssignKey("KeyZoomOut ", Keys.Oemcomma)
        AssignKey("KeyToggleThumbs ", Keys.Oem3)
        AssignKey("KeyToggleButtons ", Keys.OemSemicolon)
        AssignKey("KeyEscape ", Keys.Escape)
        AssignKey("KeySelect ", Keys.F3)
        AssignKey("KeyDelete ", Keys.Delete)

    End Sub

    Public Sub AssignKey(name As String, n As Keys)

        KeyAssignments.Add(name, n.ToString)
    End Sub
    Public Structure MVKey
        Public KeyFunction As String
        Public Key As Integer

    End Structure
End Module

