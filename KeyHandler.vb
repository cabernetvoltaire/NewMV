'Public Class KeyHandler
'    Public Event Waiting(Sender As Object, e As EventArgs)
'    Public Event RenameFolder(sender As Object, e As EventArgs)
'    Public Event ExitApplication(Sender As Object, e As EventArgs)
'    Private mKeyEvent As KeyEventArgs
'    Public Property e() As KeyEventArgs
'        Get
'            Return mKeyEvent
'        End Get
'        Set(ByVal value As KeyEventArgs)
'            mKeyEvent = value
'        End Set
'    End Property
'    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)
'        'Principal key handling

'        Select Case e.KeyCode
'#Region "Function Keys"
'            Case Keys.F2
'                RaiseEvent RenameFolder(Me, e)
'            Case Keys.F4
'                If e.Alt Then
'                    RaiseEvent ExitApplication(Me, e)
'                End If
'            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
'                RaiseEvent FunctionKeyPressed(Me, e)
'                e.SuppressKeyPress = True

'            Case KeyDelete
'                RaiseEvent DeleteFiles(Me, e)
'#End Region

'#Region "Alpha and Numeric"

'            Case Keys.Enter
'                If e.Control Then
'                    'Find selected file in filetree/list (finds source of a link, rather than the link)
'                    RaiseEvent LocateFile(Me, e)

'                Else
'                    'General press return refreshes tree.
'                    RaiseEvent RefreshTree(Me, e)
'                End If
'            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
'                RaiseEvent AlphaNumeric(Me, e)

'#End Region

'#Region "Control Keys"
'            Case KeyTraverseTree, KeyTraverseTreeBack
'                RaiseEvent TraverseTree(Me, e)
'            Case Keys.Left, Keys.Right
'                RaiseEvent LeftRight(Me, e)
'            Case Keys.Up, Keys.Down
'                RaiseEvent UpDown(Me, e)
'            Case Keys.Escape
'                RaiseEvent EscapeMedia(Me, e)
'            Case KeyToggleButtons
'                RaiseEvent ToggleButtons(Me, e)
'            Case KeyNextFile, KeyPreviousFile, LKeyNextFile, LKeyPreviousFile
'                RaiseEvent NextFile(Me, e)
'                e.SuppressKeyPress = True
'                'e = SelectNextFile(e)


'#End Region


'#Region "Video Navigation"
'            Case KeySmallJumpDown, KeySmallJumpUp, LKeySmallJumpDown, LKeySmallJumpUp
'                If tmrJumpRandom.Enabled Then
'                    If e.KeyCode = KeySmallJumpUp Then
'                        tbScanRate.Value = Math.Min(tbScanRate.Maximum, tbScanRate.Maximum * 0.1 + tbScanRate.Value)
'                    Else
'                        tbScanRate.Value = Math.Max(tbScanRate.Minimum, -tbScanRate.Maximum * 0.1 + tbScanRate.Value)

'                    End If
'                Else

'                    If e.Alt Then
'                        If e.Control Then
'                            SP.ChangeJump(False, e.KeyCode = KeySmallJumpUp)
'                            UpdateFileInfo()
'                        Else

'                            SpeedIncrease(e)
'                        End If

'                    Else
'                        MediaSmallJump(e)
'                    End If
'                    ' e.SuppressKeyPress = True
'                End If

'            Case KeyBigJumpOn, KeyBigJumpBack
'                If e.Alt Then
'                    If e.Control Then
'                        SP.ChangeJump(True, e.KeyCode = KeyBigJumpOn)
'                        UpdateFileInfo()
'                    Else

'                    End If
'                Else
'                    MediaLargeJump(e, e.Modifiers = Keys.Control, e.KeyCode = KeyBigJumpOn)
'                End If

'                'e.SuppressKeyPress = True

'            Case KeyMarkFavourite
'                If e.Control And e.Alt And e.Shift Then
'                    RemoveAllFavourites(Media.MediaPath)
'                ElseIf e.Control And e.Alt Then
'                    If Media.IsLink Then 'TODO: Does this work?
'                        Dim finfo As New IO.FileInfo(Media.MediaPath)
'                        Dim s As String = Media.UpdateBookmark(Media.MediaPath, Media.Position)
'                        finfo.MoveTo(s)
'                        UpdatePlayOrder(FBH)
'                    End If
'                ElseIf e.Alt Then
'                    NavigateToFavourites()
'                ElseIf e.Control Then
'                    Dim m As Integer = Media.FindNearestCounter(True)
'                    If m < 0 Then Exit Select
'                    Dim l As Long = Media.Markers(m)
'                    If Media.IsLink Then
'                        RemoveMarker(Media.LinkPath, l)
'                    Else
'                        RemoveMarker(Media.MediaPath, l)

'                    End If
'                    DrawScrubberMarks()
'                Else
'                    AddMarker()
'                    '    DrawScrubberMarks()
'                End If

'            Case KeyJumpToPoint

'                Media.MediaJumpToMarker(ToEnd:=True)
'                e.SuppressKeyPress = True
'            Case KeyJumpToMark, LKeyMarkPoint
'                'Addmarker(Media.MediaPath)
'                e = JumpToMark(e)

'            Case KeyMuteToggle
'                Muted = Not Muted
'                If Muted Then
'                    Media.Player.settings.mute = False
'                    SwitchSound(False)
'                Else

'                    MSFiles.MuteAll()
'                End If
'                '                Media.Player.settings.mute = Not Media.Player.settings.mute
'                e.SuppressKeyPress = True

'            Case KeyLoopToggle

'                e.SuppressKeyPress = True

'            Case KeyJumpRandom
'                If e.Control AndAlso e.Shift Then
'                    JumpRandom(True)
'                ElseIf e.Alt Then
'                    ToggleMovieScan()
'                Else
'                    JumpRandom(False)
'                End If
'                'JumpRandom(e.Control And e.Shift) 'Autotrail if both held)

'            Case KeyToggleSpeed
'                If Media.Speed.Fullspeed Then
'                    Media.Speed.Paused = Not Media.Speed.Paused

'                Else
'                    SpeedChange(e)
'                End If
'            Case KeySpeed1, KeySpeed2, KeySpeed3, KeySpeed3 Or Keys.Control
'                SpeedChange(e)
'#End Region

'#Region "Picture Functions"
'            Case KeyRotateBack
'                RotatePic(Media.Picture, False)
'            Case KeyRotate
'                If e.Alt Then
'                    If Media.MediaType = Filetype.Movie Then Media.LoopMovie = Not Media.LoopMovie
'                Else
'                    RotatePic(Media.Picture, True)
'                    Media.HandlePic(Media.MediaPath)
'                End If
'#End Region
'#Region "States"
'            Case KeySelect
'                SelectSubList(FocusControl, False)

'            Case KeyFullscreen
'                e.SuppressKeyPress = True

'            Case KeyCycleSortOrder
'                If e.Alt Then
'                    PlayOrder.ReverseOrder = Not PlayOrder.ReverseOrder
'                ElseIf e.Shift Then
'                    PlayOrder.IncrementState()
'                Else
'                    PlayOrder.Toggle()
'                End If

'            Case KeyCycleFilter 'Cycle through listbox filters
'                If e.Alt Then
'                    If ShiftDown Then
'                        blnSecondScreen = True
'                    Else
'                        blnSecondScreen = False
'                    End If
'                    GoFullScreen(Not blnFullScreen)

'                Else
'                    CurrentfilterState.IncrementState(Not e.Shift)
'                End If
'                e.SuppressKeyPress = True

'            Case KeyCycleNavMoveState
'                If e.Shift Then
'                    NavigateMoveState.IncrementState()
'                Else
'                    ToggleMove()
'                End If
'            Case KeyCycleStartPoint
'                If e.Control Then Exit Sub
'                If e.Shift Then
'                    Media.SPT.IncrementState(7)
'                Else
'                    Media.SPT.IncrementState(2)

'                End If
'            Case KeyTrueSize
'                '   MSFiles.ClickAllPics()
'               ' PicClick(currentPicBox)

'#End Region
'            Case KeyBackUndo
'                If e.Alt Then
'                    If CBXButtonFiles.Items.Count > 1 Then
'                        If CBXButtonFiles.SelectedIndex = 0 Then
'                            CBXButtonFiles.SelectedIndex = CBXButtonFiles.Items.Count - 1
'                        Else
'                            CBXButtonFiles.SelectedIndex = (CBXButtonFiles.SelectedIndex - 1)
'                        End If

'                    End If
'                Else
'                    If LastFolder.Count > 0 Then
'                        tvmain2.SelectedFolder = LastFolder.Pop '= CurrentFolder
'                    End If
'                End If

'            Case KeyReStartSS
'                If e.Shift Then
'                    ToggleMovieSlideShow()

'                Else
'                    tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
'                    SP.Slideshow = tmrSlideShow.Enabled
'                    chbSlideShow.Checked = tmrSlideShow.Enabled

'                End If

'        End Select
'        '   End If
'        Me.Cursor = Cursors.Default
'        ' e.suppresskeypress = True
'        '    Response.Enabled = False
'    End Sub


'#Region "Default Key Assignments"
'    Public KeyJumpToPoint = Keys.Insert
'    Public KeyCycleStartPoint = Keys.OemQuestion
'    Public KeyCycleFilter = Keys.Scroll
'    Public KeyCycleNavMoveState = Keys.Oem7
'    Public KeyCycleSortOrder = Keys.Pause
'    Public KeyFullscreen = Keys.Scroll + Keys.Control
'    Public KeyNextFile = Keys.PageDown
'    Public KeyPreviousFile = Keys.PageUp

'    Public KeyJumpRandom = Keys.Divide
'    Public KeyTraverseTreeBack = Keys.Subtract
'    Public KeyTraverseTree = Keys.Add
'    Public KeyMarkFavourite = Keys.NumPad7
'    Public KeyBigJumpBack = Keys.NumPad8
'    Public KeyBigJumpOn = Keys.NumPad9
'    Public KeyJumpToMark = Keys.NumPad4
'    Public KeySmallJumpDown = Keys.NumPad5
'    Public KeySmallJumpUp = Keys.NumPad6
'    Public KeySpeed1 = Keys.NumPad1
'    Public KeySpeed2 = Keys.NumPad2
'    Public KeySpeed3 = Keys.NumPad3
'    Public KeyToggleSpeed = Keys.NumPad0
'    Public KeyMuteToggle = Keys.Decimal

'    Public LKeyFullscreen = Keys.F + Keys.Control
'    Public LKeyRandomize = Keys.R + Keys.Control
'    Public LKeyTraverseTreeBack = Keys.Control + Keys.OemMinus
'    Public LKeyTraverseTree = Keys.Control + Keys.Oemplus
'    Public LKeyJumpAutoT = Keys.Control + Keys.Divide
'    Public LKeyMuteToggle = Keys.Control + Keys.Decimal
'    Public LKeyNextFile = Keys.Control + Keys.N
'    Public LKeyPreviousFile = Keys.Control + Keys.P
'    Public LKeyZoomIn = Keys.Control + Keys.Z
'    Public LKeyFolderJump = Keys.Control + Keys.Multiply
'    Public LKeyToggleSpeed = Keys.Control + Keys.D0
'    Public LKeySpeed1 = Keys.Control + Keys.D1
'    Public LKeySpeed2 = Keys.Control + Keys.D2
'    Public LKeySpeed3 = Keys.Control + Keys.D3
'    Public LKeySmallJumpDown = Keys.Control + Keys.D5
'    Public LKeySmallJumpUp = Keys.Control + Keys.D6
'    Public LKeyJumpToPoint = Keys.Control + Keys.D7
'    Public LKeyBigJumpBack = Keys.Control + Keys.D8
'    Public LKeyBigJumpOn = Keys.Control + Keys.D9
'    Public LKeyMarkPoint = Keys.Control + Keys.D4


'    'Common Keys
'    Public KeyBackUndo = Keys.Back
'    Public KeyReStartSS = Keys.Space
'    Public KeyLoopToggle = Keys.OemCloseBrackets + Keys.Alt
'    Public KeyTrueSize = Keys.Oemplus
'    Public KeyRotate = Keys.OemCloseBrackets
'    Public KeyRotateBack = Keys.OemOpenBrackets
'    Public KeyAddFile = Keys.Oemtilde

'    Public KeyZoomOut = Keys.Oemcomma
'    Public KeyToggleThumbs = Keys.Oem3

'    Public KeyToggleButtons = Keys.OemSemicolon
'    Public KeyEscape = Keys.Escape
'    Public KeySelect = Keys.F3
'    Public KeyDelete = Keys.Delete

'#End Region
'End Class
