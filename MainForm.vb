
Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports AxWMPLib
Imports System.Threading
Imports MasaSam.Forms.Controls

Public Class MainForm

    Public Initialising As Boolean = True
    Public AutoLoadButtons As Boolean = False
    Public ButtonsHidden As Boolean = False
    Public defaultcolour As Color = Color.Aqua
    Public movecolour As Color = Color.Orange
    Public sound As New AxWindowsMediaPlayer
    Public AutoTraverse As Boolean = False
    Public WithEvents FNG As New FileNamesGrouper
    Public WithEvents Random As New RandomHandler
    Public WithEvents NavigateMoveState As New StateHandler()
    Public WithEvents CurrentFilterState As New FilterHandler
    Public WithEvents PlayOrder As New SortHandler
    Private WithEvents FM As New FilterMove()
    Private WithEvents DM As New DateMove
    Private WithEvents Att As New Attributes
    Private WithEvents Op As New OrphanFinder
    Public WithEvents SP As New SpeedHandler
    Public WithEvents AT As New AutoTrailer
    Public WithEvents X As New OrphanFinder
    Public WithEvents VT As New VideoThumbnailer
    Public WithEvents FKH As FunctionKeyHandler
    Public Splash As New SplashScreen1

    ' Public WithEvents Response As New Timer
    Public FocusControl As New Control
    Public DraggedFolder As String
    Public CurrentFileList As New List(Of String)
    Public T As Thread
    Public FirstButtons As New ButtonForm
    Public bmp As Bitmap
    Public ScrubberProportion As Decimal = 0.97

    Public Marks As New MarkPlacement
    Public speedkeys = {KeySpeed1, KeySpeed2, KeySpeed3}
    Public WithEvents FBH As New FileboxHandler(lbxFiles)
    Public WithEvents LBH As New ListBoxHandler(lbxShowList)


#Region "Event Responders"
    Sub OnThumbnailed(file As String) Handles VT.Thumbnailed
        emblem.ImageLocation = file
        emblem.Load()
    End Sub
    Public Sub OnItemChanged(sender As Object, e As EventArgs) Handles FBH.ListItemChanged
        Dim m = New IO.FileInfo(FBH.OldItem)
        Try
            m.MoveTo(FBH.NewItem)

        Catch ex As IOException
            FBH.NewItem = AppendExistingFilename(FBH.NewItem, m)
            m.MoveTo(FBH.NewItem)
        End Try
    End Sub
    Public Sub OnEndReached(sender As Object, e As EventArgs) Handles FBH.EndReached, LBH.EndReached
        If AutoTraverse Then

            tvMain2.Traverse(False)
        End If
    End Sub
    Public Sub OnFolderRename(sender As Object, e As NodeLabelEditEventArgs) Handles tvMain2.LabelEdited
        If Not e.CancelEdit Then
            CurrentFolder = CurrentFolder.Replace(e.Node.Text, e.Label)

        End If
    End Sub
    Public Sub OnParentNotFound(sender As Object, e As EventArgs) Handles X.ParentNotFound, Op.ParentNotFound
        '  lbxReport.Items.Add("Not found")
    End Sub
    Public Sub OnRandomChanged() Handles Random.RandomChanged
        FBH.Random = Random
        If Random.All Then
            PlayOrder.State = SortHandler.Order.Random
            Media.SPT.State = StartPointHandler.StartTypes.Random
        Else
            'PlayOrder.State = SortHandler.Order.DateTime
            '       StartPoint.State = StartPointHandler.StartTypes.NearBeginning
        End If

        If Random.StartPointFlag Then
            Media.SPT.State = StartPointHandler.StartTypes.Random
        End If
        MSFiles.RandomNext = Random.NextSelect

        ToggleRandomAdvanceToolStripMenuItem.Checked = Random.NextSelect
        ToggleRandomSelectToolStripMenuItem.Checked = Random.OnDirChange
        ToggleRandomStartToolStripMenuItem.Checked = Random.StartPointFlag
        chbNextFile.Checked = Random.NextSelect
        chbOnDir.Checked = Random.OnDirChange


    End Sub
    Public Sub OnStartChanged(Sender As Object, e As EventArgs)
        'Media.Bookmark = -1
        tbxAbsolute.Text = New TimeSpan(0, 0, Media.SPT.StartPoint).ToString("hh\:mm\:ss") 'TODO: DUbious

        tbxPercentage.Text = Int(100 * Media.SPT.StartPoint / Media.SPT.Duration) & "%"
        tbPercentage.Value = Media.SPT.Percentage
        tbAbsolute.Maximum = Media.SPT.Duration
        tbAbsolute.Value = Media.SPT.StartPoint
        Select Case Media.SPT.State
            Case StartPointHandler.StartTypes.ParticularAbsolute
                tbxPercentage.Enabled = False
                tbxAbsolute.Enabled = True
            Case StartPointHandler.StartTypes.ParticularPercentage
                tbxAbsolute.Enabled = False
                tbxPercentage.Enabled = True
            Case Else
                tbxAbsolute.Enabled = False
                tbxPercentage.Enabled = False
        End Select
        ' MSFiles.ListIndex = lbxFiles.SelectedIndex
        'Set all MH to have the same StartPoint handler
        MSFiles.SetStartStates(Media.SPT)

        FullScreen.Changing = False
        cbxStartPoint.SelectedIndex = Media.SPT.State
        tbStartpoint.Text = "START:" & Media.SPT.Description
        If Not Media.DontLoad Then Media.MediaJumpToMarker()

    End Sub
    Public Sub OnStateChanged(sender As Object, e As EventArgs) Handles NavigateMoveState.StateChanged, CurrentFilterState.StateChanged, PlayOrder.StateChanged
        'If StartingUpFlag Then Exit Sub
        Dim s() As String = Split(sender.ToString, ".")


        Select Case s(s.Length - 1)
            Case "StateHandler"
                lblNavigateState.Text = NavigateMoveState.Instructions
                tbState.Text = UCase(NavigateMoveState.Description)
            Case "FilterHandler"
                FBH.Filter = CurrentFilterState
                cbxFilter.BackColor = CurrentFilterState.Colour
                cbxFilter.SelectedIndex = CurrentFilterState.State
                tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
            Case "SortHandler"
                FBH.SortOrder = PlayOrder
                cbxOrder.SelectedIndex = PlayOrder.State
                cbxOrder.BackColor = PlayOrder.Colour
                tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        End Select
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

        If sender IsNot NavigateMoveState Then
            If Not Media.DontLoad Then
                If FocusControl Is lbxShowList Then
                    UpdatePlayOrder(LBH)
                Else
                    UpdatePlayOrder(FBH)
                End If
            End If
        Else
            If NavigateMoveState.State = StateHandler.StateOptions.ExchangeLink Then
                CurrentFilterState.State = FilterHandler.FilterState.LinkOnly
            End If
        End If
        If Not Media.DontLoad Then PreferencesSave()
    End Sub
    Private Sub SetControlColours(MainColor As Color, FilterColor As Color)
        tvMain2.BackColor = FilterColor
        tvMain2.HighlightSelectedNodes()
        lbxFiles.BackColor = FilterColor
        lbxShowList.BackColor = FilterColor

        FocusControl.BackColor = MainColor

    End Sub
    Public Sub OnSpeedChange(sender As Object, e As EventArgs) Handles SP.SpeedChanged
        Dim SH As SpeedHandler = CType(sender, SpeedHandler)
        If SH.Slideshow Then
            tbSpeed.Text = "Slide Interval=" & SH.Interval
        Else
            tbSpeed.Text = "Speed:" & SH.FrameRate & "fps"

        End If
    End Sub
    Public Sub OnRenameFolderStart(sender As Object, e As KeyEventArgs) Handles tvMain2.KeyDown
        If e.KeyCode = Keys.F2 Then
            CancelDisplay(True)
        End If
        If e.KeyCode = KeyTraverseTree Or e.KeyCode = KeyTraverseTreeBack Or e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
            tvMain2.tvFiles_KeyDown(sender, e)
        End If
    End Sub
    Friend Sub OnFolderMoved(ByVal path As String)
        Dim d As New IO.DirectoryInfo(tvMain2.SelectedFolder)
        tvMain2.RemoveNode(path)
        ' tvMain2.RefreshTree(d.FullName)
        'Dim dir As New IO.DirectoryInfo(path)
        'Dir.Delete()
    End Sub
#End Region
    Private Sub AddMarker()
        If Media.IsLink Then
            CreateFavourite(Media.LinkPath)
        Else
            CreateFavourite(Media.MediaPath)
        End If
        PopulateLinkList(Media)
        '       e.SuppressKeyPress = True
        DrawScrubberMarks()
    End Sub
    Private Sub RefreshLinks(List As List(Of String))
        For Each filename In List
            Dim x As List(Of String) = AllFaveMinder.GetLinksOf(filename)
            If x.Count <> 0 Then Op.ReuniteWithFile(x, filename)
        Next
    End Sub
    Public Sub RemoveMarker(filepath As String, timecode As Long)
        Dim x As List(Of String) = AllFaveMinder.GetLinksOf(filepath)
        For Each f In x
            If BookmarkFromLinkName(f) = timecode Then
                Dim file As New IO.FileInfo(f)
                file.Delete()
            End If
        Next
        AllFaveMinder.NewPath(GlobalFavesPath)
        PopulateLinkList(Media)

    End Sub
    Public Sub PopulateLinkList(Media As MediaHandler)
        Dim filepath As String
        If Media.IsLink Then
            filepath = Media.LinkPath
        Else
            filepath = Media.MediaPath
        End If
        If filepath = "" Then Exit Sub
        Media.Markers.Clear()
        Dim x As List(Of String) = AllFaveMinder.GetLinksOf(filepath)
        If Media.IsLink Then
        Else
            If x.Count = 0 Then
            Else

                T = New Thread(New ThreadStart(Sub() Op.ReuniteWithFile(x, filepath))) With {
            .IsBackground = True
        }
                T.SetApartmentState(ApartmentState.STA)

                T.Start() 'Causing memory leak?
                ' Op.ReuniteWithFile(x, filepath)
            End If
        End If

        Dim i = 0
        If x.Count = 0 Then
            chbPreviewLinks.Font = New Font(chbPreviewLinks.Font, FontStyle.Regular)
            chbPreviewLinks.Text = "Preview links (None)"
            Scrubber.BackColor = Me.BackColor
            If chbPreviewLinks.Checked Then
                LBH.ItemList.Clear()
                ControlSetFocus(lbxFiles)
            End If
        Else
            chbPreviewLinks.Font = New Font(chbPreviewLinks.Font, FontStyle.Bold)
            chbPreviewLinks.Text = "Preview links (" & x.Count & ")"
            'Scrubber.Visible = True
            Scrubber.BackColor = Color.HotPink
            If chbPreviewLinks.Checked Then
                x.Sort(New CompareByEndNumber)
                FillShowbox(lbxShowList, FilterHandler.FilterState.LinkOnly, x)
            End If
            Media.Markers = Media.GetMarkersFromLinkList
            Media.Markers.Sort()
            ' Marks.Markers = Media.Markers
        End If
        DrawScrubberMarks()



    End Sub

    Public Sub DrawScrubberMarks()

        If Media.Duration = 0 Then Exit Sub

        Marks.Duration = Media.Duration
        'Marks.Bar = Scrubber
        '  Marks.Clear()
        Marks.Markers = Media.GetMarkersFromLinkList
        Scrubber.Width = ctrPicAndButtons.Width * ScrubberProportion
        Scrubber.Left = Scrubber.Width * ((1 - ScrubberProportion) / 2)
        'Scrubber.Visible = False
        Marks.Create()
        'Marks.Bar.BackColor = Me.BackColor

        'Scrubber.Image = Marks.Bitmap NEVER add this back.



    End Sub


    Public Sub SwitchSound(slow As Boolean)
        'SndH.Slow = slow
    End Sub


    Private Sub SaveShowlist()
        Dim path As String
        With SaveFileDialog1
            .DefaultExt = "msl"
            .Filter = "Metavisua list files|*.msl|All files|*.*"
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            End If
            path = .FileName
        End With
        StoreList(Showlist, path)
    End Sub
    Public Function LoadButtonList() As String
        Dim path As String = ""
        With OpenFileDialog1
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"
            .Title = "Choose file containing button assignments..."
            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
                ButtonFilePath = path
            End If
        End With
        AddToButtonFilesList(path)
        Return path
    End Function
    Private Sub LoadShowList(Optional DesiredFile As String = "")
        Dim path As String = ""
        If DesiredFile <> "" Then
            path = DesiredFile
        Else
            With OpenFileDialog1
                .DefaultExt = "msl"
                .Filter = "Metavisua list files|*.msl|All files|*.*"
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        LastShowList = path
        If path = "" Then Exit Sub

        Dim finfo As New FileInfo(path)
        Dim size As Long = finfo.Length
        Dim time As TimeSpan
        Dim loadrate As Single
        lngListSizeBytes = size
        If path <> "" Then
            CollapseShowlist(False)
            tbLastFile.Text = TimeOperation(True).TotalMilliseconds
            ProgressBarOn(lngListSizeBytes)
            'LBH.ItemList = Showlist
            Getlist(Showlist, path, lbxShowList)
            time = TimeOperation(False)
            loadrate = size / time.TotalMilliseconds
        End If
        ProgressBarOff()
        'tbShowfile.Text = "SHOWFILE LOADED:" & path
    End Sub


    Public Sub CancelDisplay(SlideshowsOff As Boolean)
        MSFiles.CancelURL(Media.MediaPath)
        'MSFiles.ResettersOff()


        If SlideshowsOff Then
            tmrMovieSlideShow.Enabled = False
            tmrSlideShow.Enabled = False
            tmrAutoTrail.Enabled = False
            SP.Slideshow = False
        End If
    End Sub

    Private Sub DeleteFolder(FolderName As String, tvw As FileSystemTree, blnConfirm As Boolean)

        If My.Computer.FileSystem.DirectoryExists(FolderName) Then
            Dim m As MsgBoxResult = MsgBoxResult.No
            If blnConfirm Then
                m = MsgBox("Delete folder " & FolderName & "?", MsgBoxStyle.YesNoCancel)
            End If
            If Not blnConfirm OrElse m = MsgBoxResult.Yes Then
                MoveFolder(FolderName, "")
                tvw.Traverse(False)
                tvw.RemoveNode(FolderName)
            End If
        End If
        Exit Sub
        With My.Computer.FileSystem
            Dim m As MsgBoxResult = MsgBoxResult.No
            If .DirectoryExists(FolderName) Then
                If blnConfirm Then
                    m = MsgBox("Delete folder " & FolderName & "?", MsgBoxStyle.YesNoCancel)
                End If
                If Not blnConfirm OrElse m = MsgBoxResult.Yes Then
                    Dim f As New DirectoryInfo(FolderName)
                    Try
                        For Each ff In f.EnumerateFiles
                            If MSFiles.CurrentURLS.Contains(ff.FullName) Then
                                MSFiles.CancelURL(ff.FullName)
                            End If
                            .DeleteFile(ff.FullName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                        Next
                        For Each fol In f.EnumerateDirectories
                            DeleteFolder(fol.FullName, tvw, False)
                        Next

                        f.Delete(FileIO.RecycleOption.SendToRecycleBin)
                        DirectoriesList.Remove(FolderName)
                        '.DeleteDirectory(CurrentFolder, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    Catch x As System.IO.IOException
                        MsgBox(Err.Description)


                        'MsgBox(Err.Description)

                    End Try

                    tvw.Traverse(False)
                    tvw.RemoveNode(f.FullName)
                    ' tvw.RefreshTree(f.Parent.FullName)
                End If

            ElseIf m = MsgBoxResult.No Or m = MsgBoxResult.Cancel Then
                ControlSetFocus(lbxFiles)

            End If
        End With
    End Sub


    Public Sub UpdatePlayOrder(LBHandler As ListBoxHandler)
        LBHandler.SortOrder = PlayOrder
        If LBHandler Is LBH Then
            Dim s = lbxShowList.SelectedItem
            LBH.FillBox()
            LBH.SetNamed(s)
        Else
            Dim e = New DirectoryInfo(CurrentFolder)
            '  FillFileBox(lbxFiles, e, False)
            FBH.DirectoryPath = e.FullName
            FBH.SetNamed(Media.MediaPath)
        End If


    End Sub

    Private Function SpeedChange(e As KeyEventArgs) As KeyEventArgs


        If Not Media.Playing Then
            tmrSlideShow.Enabled = True
            Media.Speed.Slideshow = tmrSlideShow.Enabled
            tmrSlideShow.Interval = Media.Speed.Interval
            If e.KeyCode < KeySpeed1 Then
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
                Return e
                Exit Function
            End If
            Media.Speed.SSSpeed = e.KeyCode - KeySpeed1 'Set slideshow speed if pic showing, and start slideshow
            'PlaybackSpeed = 30

        ElseIf e.KeyCode >= KeySpeed1 Then
            Media.Speed.Speed = e.KeyCode - KeySpeed1
            PlaybackSpeed = 1000 / Media.Speed.FrameRate 'Otherwise, set playback speed 'TODO Options
            Media.Speed.Fullspeed = False
        End If
        If e.KeyCode = KeyToggleSpeed Then
            If Media.Playing Then
                'If Media.Player.playState = WMPLib.WMPPlayState.wmppsPaused And Media.Speed.Fullspeed = False Then
                TogglePause()
            Else
                tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
            End If

        Else
            If Media.Playing Then

                tmrSlowMo.Interval = 1000 / Media.Speed.FrameRate
                tmrSlowMo.Enabled = True
                Media.Player.Ctlcontrols.pause()
            Else
            End If
        End If

        'End If
        'Report
        'blnSpeedRestart = True
        '  tbSpeed.Text = "SPEED (" & Media.Speed.FrameRate & " fps)"
        e.SuppressKeyPress = True
        Return e
    End Function

    Private Sub TogglePause()
        If Media.Speed.Paused Then
            Media.Position = Media.Player.Ctlcontrols.currentPosition

            Media.Speed.Paused = False
            tmrSlowMo.Enabled = False
            Media.Speed.Fullspeed = True
        Else
            Media.Player.Ctlcontrols.pause()
            'Media.Speed.Fullspeed = False
            Media.Speed.Paused = True
        End If
    End Sub

    Private Sub ToggleRandomStartPoint()
        Random.StartPointFlag = Not Random.StartPointFlag
    End Sub
    Public Sub GoFullScreen(blnGo As Boolean)
        FullScreen.Changing = True
        MSFiles.ForceLoad = True
        'TogglePause()
        'Media.Speed.Paused = True
        If blnGo Then
            Dim s As String = Media.MediaPath
            Dim screen As Screen
            If blnSecondScreen Then
                screen = Screen.AllScreens(1)
            Else
                screen = Screen.PrimaryScreen
                'SplitterPlace(0.75)
            End If
            FullScreen.StartPosition = FormStartPosition.CenterScreen
            FullScreen.Location = screen.Bounds.Location + New Point(100, 100)
            CancelDisplay(False)
            'Media.Player.Size = screen.Bounds.Size
            TogglePause()
            FullScreen.FSFiles = MSFiles
            FullScreen.Show()
            ' TogglePause()
        Else
            '            SplitterPlace(0.25)
            MSFiles = FullScreen.FSFiles
            MSFiles.AssignPlayers(MainWMP1, MainWMP2, MainWMP3)
            MSFiles.AssignPictures(PictureBox1, PictureBox2, PictureBox3)
            MSFiles.ListIndex = MSFiles.Listbox.SelectedIndex

            FullScreen.Close()
            'TogglePause()
        End If
        FullScreen.Changing = False
        blnFullScreen = Not blnFullScreen
        MSFiles.ForceLoad = False
    End Sub
    Private Sub SplitterPlace(i As Decimal)
        ctrMainFrame.SplitterDistance = Me.Width * i
    End Sub
    Private Sub ToggleButtons()
        Buttons_Load()
        blnButtonsLoaded = Not blnButtonsLoaded
        ctrPicAndButtons.Panel2Collapsed = Not blnButtonsLoaded
        UpdateButtonAppearance()
    End Sub
    Public Sub AdvanceFile(blnForward As Boolean, Optional Random As Boolean = False)
        Dim LBHH As New ListBoxHandler(LBH.ListBox)
        If FocusControl Is lbxShowList Or CtrlDown Then
            LBHH = LBH
        Else
            LBHH = FBH
        End If
        LBHH.ListBox.SelectionMode = SelectionMode.One
        Dim count = LBHH.ItemList.Count
        ReDim Preserve FBCShown(count)
        MSFiles.RandomNext = Random
        If Random Then
            LBHH.ListBox.SelectedItem = MSFiles.NextItem
        Else
            LBHH.IncrementIndex(blnForward)
        End If

        NofShown += 1
        If NofShown >= count Then 'Re-sets when all shown. Quite nice. 
            ReDim FBCShown(count)
            NofShown = 0
        End If
        GC.Collect()

    End Sub

    Public Sub CollapseShowlist(Collapse As Boolean)
        ButtonsHidden = Collapse
        MasterContainer.Panel2Collapsed = Collapse
        lbxFiles.SelectionMode = SelectionMode.One
        ' ControlSetFocus(lbxFiles)
        MasterContainer.SplitterDistance = MasterContainer.Height / 3
        UpdateFileInfo()
    End Sub

    Public Sub MediaSmallJump(e As KeyEventArgs)
        Dim iJumpFactor As Integer
        Dim blnBack As Boolean = e.KeyCode < (KeySmallJumpUp + KeySmallJumpDown) / 2
        If e.Control Then
            iJumpFactor = 10
        Else
            iJumpFactor = 1
        End If
        If blnBack Then
            Media.Position = Media.Position - SP.AbsoluteJump / iJumpFactor
        Else
            Media.Position = Media.Position + SP.AbsoluteJump / iJumpFactor
        End If


    End Sub

    Public Sub MediaLargeJump(e As KeyEventArgs, Small As Boolean, Forward As Boolean)
        Dim iJumpFactor As Integer

        If Small Then  'Holding down ctrl reduces the jump distance by a factor
            iJumpFactor = 10
        Else
            iJumpFactor = 1
        End If
        If Forward Then
            Media.Position = Math.Min(Media.Duration, Media.Position + Media.Duration / (iJumpFactor * SP.FractionalJump))
        Else
            Media.Position = Math.Min(Media.Duration, Media.Position - Media.Duration / (iJumpFactor * SP.FractionalJump))
        End If

    End Sub
    Public Sub JumpRandom(blnAutoTrail As Boolean)

        If Media.MediaType = Filetype.Movie Or Media.IsLink Then

            If Not blnAutoTrail Then
                'Random.StartPoint = True
                Media.Position = (Rnd(1) * (Media.Duration))
                ' JumpVideo(Media.Player, SoundWMP)

                tbStartpoint.Text = "START:RANDOM"


            Else
                'MsgBox("Autotrail")
                ToggleAutoTrail()
            End If
        ElseIf Media.MediaType = Filetype.Pic Then
            AdvanceFile(True, True)
        Else
        End If

    End Sub
    Public Sub ToggleAutoTrail()

        tmrAutoTrail.Enabled = Not tmrAutoTrail.Enabled
        chbAutoTrail.Checked = tmrAutoTrail.Enabled
        TrailerModeToolStripMenuItem.Checked = tmrAutoTrail.Enabled
        If tmrAutoTrail.Enabled Then
            Media.SPT.SavedState = Media.SPT.State
            Media.SPT.State = StartPointHandler.StartTypes.Random
            SwitchSound(True)
        Else
            Media.SPT.State = Media.SPT.SavedState
            SwitchSound(False)
            Debug.Print("Normal")
            tmrSlowMo.Enabled = False
            Media.Player.settings.rate = 1
            Media.Player.Ctlcontrols.play()
        End If
    End Sub
    Public Function NumKeyEquivalent(e As KeyEventArgs) As Keys
        If e.Modifiers = Keys.Shift Then
            Dim newkey As New Keys
            Select Case e.KeyCode
                Case Keys.Home
                    newkey = Keys.NumPad7
                Case Keys.Up
                    newkey = Keys.NumPad8
                Case Keys.PageUp
                    newkey = Keys.NumPad9
                Case Keys.Left
                    newkey = Keys.NumPad4
                Case Keys.Right
                    newkey = Keys.NumPad6
                Case Keys.End
                    newkey = Keys.NumPad1
                Case Keys.Down
                    newkey = Keys.NumPad2
                Case Keys.PageDown
                    newkey = Keys.NumPad3
                Case Else
                    newkey = e.KeyCode
            End Select
            Return newkey
        Else
            Return e.KeyCode
        End If
    End Function

    Public Sub HandleKeys(sender As Object, e As KeyEventArgs)

        Me.Cursor = Cursors.WaitCursor
        'MsgBox(e.KeyCode.ToString)
        Dim nke As New Keys
        nke = NumKeyEquivalent(e)
        '  If nke <> Keys.None Then

        Select Case e.KeyCode
#Region "Function Keys"
            Case Keys.F2
                CancelDisplay(False)
            Case Keys.F4
                If e.Alt Then
                    PreferencesSave()
                    e.SuppressKeyPress = True
                    Application.Exit()
                End If
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                HandleFunctionKeyDown(sender, e)
                e.SuppressKeyPress = True

            Case KeyDelete
                DeleteFiles(e)
#End Region

#Region "Alpha and Numeric"

            Case Keys.Enter And e.Control
                If Media.IsLink Then
                    tmrUpdateFileList.Enabled = False
                    CurrentFilterState.State = FilterHandler.FilterState.All
                    HighlightCurrent(Media.LinkPath)
                ElseIf FocusControl Is lbxShowList Then
                    CurrentFilterState.State = FilterHandler.FilterState.All
                    HighlightCurrent(Media.MediaPath)
                Else
                    'Get link file
                    'Highlight 
                    If Media.Markers.Count > 0 Then

                        Dim s As String = AllFaveMinder.GetLinksOf(Media.MediaPath)(Media.LinkCounter)
                        CurrentFilterState.State = FilterHandler.FilterState.All
                        HighlightCurrent(s)
                    End If
                    '                    MsgBox("Highlight Links")

                End If
            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If Not e.Control Then
                    ChangeButtonLetter(e)
                Else
                    Select Case e.KeyCode
                        Case Keys.I
                            AddFolders.Show()
                            AddFolders.Folder = CurrentFolder
                            tvMain2.RefreshTree(CurrentFolder)
                    End Select
                End If
#End Region

#Region "Control Keys"
            Case KeyTraverseTree, KeyTraverseTreeBack
                '   If FocusControl IsNot tvMain2 Then
                tvMain2.Traverse(e.KeyCode = KeyTraverseTreeBack)
                e.SuppressKeyPress = True
              '  End If

            Case Keys.Left, Keys.Right, Keys.Up, Keys.Down
                'If FocusControl IsNot lbxShowList Then
                ControlSetFocus(tvMain2)
                    tvMain2.tvFiles_KeyDown(sender, e)
                'End If

            Case Keys.Escape
                CancelDisplay(True)

            Case KeyToggleButtons
                ToggleButtons()
            Case KeyNextFile, KeyPreviousFile, LKeyNextFile, LKeyPreviousFile
                e = SelectNextFile(e)


#End Region


#Region "Video Navigation"
            Case KeySmallJumpDown, KeySmallJumpUp, LKeySmallJumpDown, LKeySmallJumpUp
                If tmrJumpRandom.Enabled Then
                    If e.KeyCode = KeySmallJumpUp Then
                        tbScanRate.Value = Math.Min(tbScanRate.Maximum, tbScanRate.Maximum * 0.1 + tbScanRate.Value)
                    Else
                        tbScanRate.Value = Math.Max(tbScanRate.Minimum, -tbScanRate.Maximum * 0.1 + tbScanRate.Value)

                    End If
                Else

                    If e.Alt Then
                        If e.Control Then
                            SP.ChangeJump(False, e.KeyCode = KeySmallJumpUp)
                            UpdateFileInfo()
                        Else

                            SpeedIncrease(e)
                        End If

                    Else
                        MediaSmallJump(e)
                    End If
                    ' e.SuppressKeyPress = True
                End If

            Case KeyBigJumpOn, KeyBigJumpBack
                If e.Alt Then
                    If e.Control Then
                        SP.ChangeJump(True, e.KeyCode = KeyBigJumpOn)
                        UpdateFileInfo()
                    Else

                    End If
                Else
                    MediaLargeJump(e, e.Modifiers = Keys.Control, e.KeyCode = KeyBigJumpOn)
                End If

                'e.SuppressKeyPress = True

            Case KeyMarkFavourite
                If e.Control And e.Alt And e.Shift Then
                    RemoveAllFavourites(Media.MediaPath)
                ElseIf e.Control And e.Alt Then
                    If Media.IsLink Then
                        Dim finfo As New IO.FileInfo(Media.MediaPath)
                        Dim s As String = Media.UpdateBookmark(Media.MediaPath, Media.Position)
                        finfo.MoveTo(s)
                        UpdatePlayOrder(FBH)
                    End If
                ElseIf e.Alt Then
                    NavigateToFavourites()
                ElseIf e.Control Then
                    Dim m As Integer = Media.FindNearestCounter(True)
                    If m < 0 Then Exit Select
                    Dim l As Long = Media.Markers(m)
                    If Media.IsLink Then
                        RemoveMarker(Media.LinkPath, l)
                    Else
                        RemoveMarker(Media.MediaPath, l)

                    End If
                    DrawScrubberMarks()
                Else
                    AddMarker()
                    '    DrawScrubberMarks()
                End If

            Case KeyJumpToPoint

                Media.MediaJumpToMarker(ToEnd:=True)
                e.SuppressKeyPress = True
            Case KeyJumpToMark, LKeyMarkPoint
                'Addmarker(Media.MediaPath)
                e = JumpToMark(e)

            Case KeyMuteToggle
                Muted = Not Muted
                If Muted Then
                    Media.Player.settings.mute = False
                    SwitchSound(False)
                Else

                    MSFiles.MuteAll()
                End If
                '                Media.Player.settings.mute = Not Media.Player.settings.mute
                e.SuppressKeyPress = True

            Case KeyLoopToggle

                e.SuppressKeyPress = True

            Case KeyJumpRandom
                If e.Control AndAlso e.Shift Then
                    JumpRandom(True)
                ElseIf e.Alt Then
                    ToggleRandomJump()
                Else
                    JumpRandom(False)
                End If
                'JumpRandom(e.Control And e.Shift) 'Autotrail if both held)

            Case KeyToggleSpeed
                If Media.Speed.Fullspeed Then
                    Media.Speed.Paused = Not Media.Speed.Paused

                Else
                    SpeedChange(e)
                End If
            Case KeySpeed1, KeySpeed2, KeySpeed3, KeySpeed3 Or Keys.Control
                SpeedChange(e)
#End Region

#Region "Picture Functions"
            Case KeyRotateBack
                RotatePic(Media.Picture, False)
            Case KeyRotate
                If e.Alt Then
                    If Media.MediaType = Filetype.Movie Then Media.LoopMovie = Not Media.LoopMovie
                Else
                    RotatePic(Media.Picture, True)
                    Media.HandlePic(Media.MediaPath)
                End If
#End Region
#Region "States"
            Case KeySelect
                SelectSubList(FocusControl, False)

            Case KeyFullscreen
                e.SuppressKeyPress = True

            Case KeyCycleSortOrder
                If e.Alt Then
                    PlayOrder.ReverseOrder = Not PlayOrder.ReverseOrder
                ElseIf e.Shift Then
                    PlayOrder.IncrementState()
                Else
                    PlayOrder.Toggle()
                End If

            Case KeyCycleFilter 'Cycle through listbox filters
                If e.Alt Then
                    If ShiftDown Then
                        blnSecondScreen = True
                    Else
                        blnSecondScreen = False
                    End If
                    GoFullScreen(Not blnFullScreen)

                Else
                    CurrentFilterState.IncrementState(Not e.Shift)
                End If
                e.SuppressKeyPress = True

            Case KeyCycleNavMoveState
                If e.Shift Then
                    NavigateMoveState.IncrementState()
                Else
                    ToggleMove()
                End If
            Case KeyCycleStartPoint
                If e.Control Then Exit Sub
                If e.Shift Then
                    Media.SPT.IncrementState(7)
                Else
                    Media.SPT.IncrementState(2)

                End If
            Case KeyTrueSize
                '   MSFiles.ClickAllPics()
               ' PicClick(currentPicBox)

#End Region
            Case KeyBackUndo
                If e.Alt Then
                    If CBXButtonFiles.Items.Count > 1 Then
                        If CBXButtonFiles.SelectedIndex = 0 Then
                            CBXButtonFiles.SelectedIndex = CBXButtonFiles.Items.Count - 1
                        Else
                            CBXButtonFiles.SelectedIndex = (CBXButtonFiles.SelectedIndex - 1)
                        End If

                    End If
                Else
                    If LastFolder.Count > 0 Then
                        tvMain2.SelectedFolder = LastFolder.Pop '= CurrentFolder
                    End If
                End If

            Case KeyReStartSS
                If e.Shift Then
                    tmrMovieSlideShow.Interval = 950
                    tmrMovieSlideShow.Enabled = Not tmrMovieSlideShow.Enabled
                    SP.Slideshow = Not tmrMovieSlideShow.Enabled
                    chbSlideShow.Checked = tmrMovieSlideShow.Enabled

                Else
                    tmrSlideShow.Enabled = Not tmrSlideShow.Enabled
                    SP.Slideshow = tmrSlideShow.Enabled
                    chbSlideShow.Checked = tmrSlideShow.Enabled

                End If

        End Select
        '   End If
        Me.Cursor = Cursors.Default
        ' e.suppresskeypress = True
        '    Response.Enabled = False
    End Sub

    Private Function SelectNextFile(e As KeyEventArgs) As KeyEventArgs
        'Dim LLBH As New ListBoxHandler(FocusControl)

        If FocusControl Is lbxFiles Then
            If e.Control Then
                ControlSetFocus(lbxShowList)
                LBH.ListBox = lbxShowList
                LBH.IncrementIndex(e.KeyCode = KeyNextFile)
            Else
                FBH.IncrementIndex(e.KeyCode = KeyNextFile)
            End If
        Else
            If e.Control Then
                ControlSetFocus(lbxFiles)
                FBH.ListBox = lbxFiles
                FBH.IncrementIndex(e.KeyCode = KeyNextFile)
            Else
                LBH.IncrementIndex(e.KeyCode = KeyNextFile)

            End If
        End If
        '        LLBH.IncrementIndex(e.KeyCode = KeyNextFile)

        e.SuppressKeyPress = True
        tmrSlideShow.Enabled = False
        tmrMovieSlideShow.Enabled = False
        Return e
    End Function

    Private Function JumpToMark(e As KeyEventArgs) As KeyEventArgs
        If Media.Markers.Count > 0 Then
            If e.Alt Then
                Media.LinkCounter = Media.RandomCounter
            Else
                'Media.StartPoint.IncrementMarker()
                Media.IncrementLinkCounter(e.Modifiers <> Keys.Control)
                '     Media.LinkCounter = Media.FindNearestCounter(e.Modifiers = Keys.Control)
            End If
            '  Media.Bookmark = -2
            Media.MediaJumpToMarker(ToMarker:=True)

        Else
            MediaLargeJump(e, e.Modifiers = Keys.Control, True)

            'JumpRandom(e.Control And e.Shift)
        End If
        e.Handled = True
        e.SuppressKeyPress = True
        Return e
    End Function

    Private Sub RemoveAllFavourites(mediaPath As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub DeleteFiles(e As KeyEventArgs)
        'Use Movefiles with current selected list, and option to delete. 
        Dim lbx As New ListBox
        CancelDisplay(False)
        If e.Shift Then
            DeleteFolder(CurrentFolder, tvMain2, NavigateMoveState.State = StateHandler.StateOptions.Navigate)
        Else
            If FocusControl Is tvMain2 Then
            Else
                lbx = FocusControl
            End If

            Dim m As List(Of String) = ListfromSelectedInListbox(lbx)
            If lbx Is lbxShowList And NavigateMoveState.State = StateHandler.StateOptions.Navigate Then
                LBH.RemoveItems(m)
            Else
                FBH.RemoveItems(m)
                blnSuppressCreate = True
                MoveFiles(m, "")
                'FBH.Refresh()
            End If


        End If
    End Sub

    Private Sub SpeedIncrease(e As KeyEventArgs)
        If e.KeyCode = KeySmallJumpUp Then
            SP.IncreaseSpeed()
        Else
            SP.DecreaseSpeed()
        End If
        tmrSlideShow.Interval = SP.Interval
        tmrSlowMo.Interval = 1000 / SP.FrameRate

    End Sub

    Private Sub NavigateToFavourites()
        CurrentFilterState.OldState = CurrentFilterState.State
        CurrentFilterState.State = FilterHandler.FilterState.LinkOnly
        ChangeFolder(CurrentFavesPath)

        tvMain2.SelectedFolder = CurrentFolder
    End Sub
    Public Sub ChangeFolder(strPath As String)
        'If strPath <> FavesFolderPath Then
        '    CurrentfilterState.State = CurrentfilterState.OldState
        'End If
        If strPath = CurrentFolder Then
        Else
            If Not LastFolder.Contains(CurrentFolder) Then
                LastFolder.Push(CurrentFolder)

            End If
            FNG.Clear()
            ' MainForm.tvMain2.SelectedFolder = strPath
            ChangeWatcherPath(strPath)
            CurrentFolder = strPath
            ' FBH.DirectoryPath = strPath
            ReDim FBCShown(0)
            NofShown = 0

            If AutoButtons Then
                AssignLinear(CurrentFolder, LetterNumberFromAscii(Asc("0")), True)
                ChangeButtonLetter(New KeyEventArgs(Keys.D0))
            End If
            '   My.Computer.Registry.CurrentUser.SetValue("File", Media.MediaPath)
        End If

    End Sub

    Private Sub LoopToggle()
        blnLoopPlay = Not blnLoopPlay
    End Sub

    Public Sub OnTVMainReady(sender As Object, e As EventArgs) Handles tvMain2.TreeBuilt
        HighlightCurrent(Media.MediaPath)

    End Sub


    Friend Sub HighlightCurrent(strPath As String)

        If strPath = "" Then Exit Sub 'Empty
        Dim finfo As New FileInfo(strPath)
        'Change the tree if it needs changing

        Dim s As String = Path.GetDirectoryName(strPath)

        If tvMain2.SelectedFolder <> s Then
            tvMain2.SelectedFolder = s
            tmrUpdateFileList.Enabled = False
            FBH.DirectoryPath = s
        End If
        'Select file in filelist
        FBH.SetNamed(strPath)
        Att.DestinationLabel = lblAttributes
        If Not tmrSlideShow.Enabled And chbShowAttr.Checked Then
            Att.UpdateLabel(strPath)
        Else
            Att.Text = ""
        End If

    End Sub

    Private Sub AddFiles(blnRecurse As Boolean)
        ProgressBarOn(1000)
        Dim x As New List(Of String)
        x = Showlist

        Showlist = AddFilesToCollection("", blnRecurse)
        Showlist = x.Concat(Showlist)
        LBH.FillBox(Showlist)
        '        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
        ProgressBarOff()
    End Sub






    Private Sub RemoveFilesFromCollection(ByVal list As List(Of String), extensions As String)
        Dim d As New System.IO.DirectoryInfo(CurrentFolder)
        FindAllFilesBelow(d, list, extensions, True, "", True, False)
    End Sub
#Region "Initialisation"

    Private Sub PopulateLists()
        Dim m As New StartPointHandler
        cbxFilter.Items.Clear()
        cbxOrder.Items.Clear()
        For Each s In CurrentFilterState.Descriptions
            cbxFilter.Items.Add(s)
        Next
        For Each s In PlayOrder.Descriptions
            cbxOrder.Items.Add(s)
        Next
        For Each s In m.Descriptions
            cbxStartPoint.Items.Add(s)
        Next
    End Sub

    Sub SetupPlayers()
        MainWMP1.stretchToFit = True
        MainWMP2.stretchToFit = True
        MainWMP3.stretchToFit = True
        'Media.Player.uiMode = "FULL"

        MainWMP1.settings.volume = 100
        MainWMP2.settings.volume = 100
        MainWMP3.settings.volume = 100
        MSFiles.DockMedias(separate)

    End Sub
    Private Sub GlobalInitialise()
        Randomize()
        PopulateLists()
        CollapseShowlist(True)
        SetupPlayers()
        PositionUpdater.Enabled = False
        Media.DontLoad = True
        PreferencesGet()


        'Media.Picture = PictureBox1
        tbPercentage.Enabled = True
        Marks.Bar = Scrubber

        AddHandler FileHandling.FolderMoved, AddressOf OnFolderMoved
        AddHandler FileHandling.FileMoved, AddressOf OnFilesMoved

        InitialiseButtonsOld()
        NavigateMoveState.State = StateHandler.StateOptions.Navigate
        OnRandomChanged()
        ControlSetFocus(lbxFiles)
        Media.DontLoad = False
        ' LoadShowList("C:\Users\paulc\AppData\Roaming\Metavisua\Lists\ITC.msl")
        FBH.ListBox = lbxFiles
        LBH.ListBox = lbxShowList
    End Sub

    Private Sub InitialiseButtonsOld()
        If ButtonFilePath = "" Then Exit Sub
        Try
            KeyAssignmentsRestore(ButtonFilePath)
            If Not blnButtonsLoaded Then
                ToggleButtons()
            End If

        Catch ex As FileNotFoundException
            ctrPicAndButtons.Panel2Collapsed = True
            Exit Try
        Catch ex As DirectoryNotFoundException
            ctrPicAndButtons.Panel2Collapsed = True

            Exit Try
        End Try
    End Sub

    'Private Sub InitialiseButtons()
    '    'x.MdiParent = Me
    '    FirstButtons.sbtns = {btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8}
    '    FirstButtons.lbls = {lbl1, lbl2, lbl3, lbl4, lbl5, lbl6, lbl7, lbl8}
    '    FirstButtons.handler.LoadButtonSet(ButtonFilePath)
    '    If Not blnButtonsLoaded Then
    '        ToggleButtons()
    '    End If
    '    FirstButtons.TranscribeButtons(FirstButtons.buttons.CurrentRow)
    '    FirstButtons.Show()
    '    FirstButtons.Top = ctrPicAndButtons.Panel2.Top
    '    FirstButtons.Left = ctrPicAndButtons.Panel2.Left
    'End Sub

    Public Sub WatchStart(path As String)
        Exit Sub
        ' watchfolder = New System.IO.FileSystemWatcher()
        FileSystemWatcher1.Path = path
        FileSystemWatcher1.NotifyFilter = IO.NotifyFilters.LastWrite

        AddHandler FileSystemWatcher1.Changed, AddressOf Logchange
        AddHandler FileSystemWatcher1.Created, AddressOf Logchange
        AddHandler FileSystemWatcher1.Deleted, AddressOf Logchange
        ' AddHandler FileSystemWatcher1.Renamed, AddressOf logrename
        FileSystemWatcher1.EnableRaisingEvents = True
    End Sub
    Private Sub Logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Changed
        If e.ChangeType = System.IO.WatcherChangeTypes.Changed Then
            Exit Sub 'TODO This gets repeatedly called. 
            MsgBox("File " & e.FullPath & " has been modified")
        End If

    End Sub
#End Region

#Region "Control Handlers"

    Private Sub Form1_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Splash.Hide()
        GlobalInitialise()

        Me.Visible = True


    End Sub
    Public Sub ControlSetFocus(control As Control)
        ' Set focus to the control, if it can receive focus.
        'Exit Sub
        If control Is tvMain2 Then
        Else
            FocusControl = control
        End If
        If control.CanFocus Then
            control.Focus()
        End If
        If TypeOf (control) Is ListBox Then
            MSFiles.Listbox = control
        End If
    End Sub



    Public Sub Main_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        'ry
        KeyDownFlag = True
        ShiftDown = e.Shift
        CtrlDown = e.Control
        AltDown = e.Alt
        UpdateButtonAppearance()
        HandleKeys(sender, e)
        If e.KeyCode = KeyBackUndo Then
            e.SuppressKeyPress = True
        End If
        'atch ex As Exception
        'sgBox(ex.Message)
        'End Try
    End Sub
    Private Sub Main_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        KeyDownFlag = False
        ShiftDown = e.Shift
        CtrlDown = e.Control
        AltDown = e.Alt
        If FocusControl IsNot tvMain2 Then
            'IndexHandler(FocusControl, New EventArgs)

        End If

        If e.KeyData <> (Keys.F4 Or AltDown) Then
            Try
                UpdateButtonAppearance()
            Catch ex As Exception

            End Try
        End If
        e.Handled = True
        e.SuppressKeyPress = True

    End Sub
    Private Sub Main_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Media.PlaceResetter(False)
        ' DeleteThumbs()
        PreferencesSave()
        Application.Exit()
    End Sub

    Private Sub ClearShowList()
        Showlist.Clear()
        LBH.ItemList.Clear()
        CollapseShowlist(True)
    End Sub


    Private Sub RemoveFiles(sender As Object, e As EventArgs)
        ' RemoveFilesFromCollection(Showlist, strVideoExtensions)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)

    End Sub



    Sub EscapeMultiSelect(sender As Object, e As KeyEventArgs) Handles lbxFiles.KeyDown, lbxShowList.KeyDown
        If e.KeyCode = Keys.Escape Then

            Dim m As New ListBox
            m = sender

            Dim i As Integer
            If m.SelectedIndices.Count > 0 Then
                i = m.SelectedIndices(0)
            Else
                i = 0
            End If
            m.SelectionMode = SelectionMode.One
            m.SelectedIndex = i
            sender = m
        End If

    End Sub
    Private Sub Listbox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LBH.ListIndexChanged, FBH.ListIndexChanged
        '  IndexHandler(FocusControl, e)
        NewIndex.Enabled = False

        NewIndex.Interval = 50

        NewIndex.Enabled = True

    End Sub
    Public Sub IndexHandler(sender As Object, e As EventArgs) ' Handles lbxShowList.SelectedIndexChanged, lbxFiles.SelectedIndexChanged
        Dim lbx As ListBox = sender
        If lbx.SelectionMode = SelectionMode.One Then
            Dim i As Long = lbx.SelectedIndex
            If i = -1 Then
            Else
                ' Debug.Print(vbCrLf & vbCrLf & "NEXT SELECTION ---------------------------------------")
                MSFiles.Listbox = lbx
                If blnFullScreen Then
                    FullScreen.FSFiles.ListIndex = i
                Else
                    MSFiles.ListIndex = i

                End If

                'Media.MediaPath = lbx.Items(i)
            End If
        End If
    End Sub

    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        'SP.Slideshow = True
        AdvanceFile(True, Random.NextSelect)
    End Sub

    Private Sub RotatePic(currentPicBox As PictureBox, blnLeft As Boolean)
        If currentPicBox.Image Is Nothing Then Exit Sub
        If Media.MediaType <> Filetype.Pic Then Exit Sub
        Media.PicHandler.GetImage(Media.MediaPath)
        With currentPicBox.Image
            If blnLeft Then
                .RotateFlip(RotateFlipType.Rotate90FlipNone)
            Else
                .RotateFlip(RotateFlipType.Rotate270FlipNone)
            End If
            currentPicBox.Refresh()

            Dim finfo As New FileInfo(Media.MediaPath)
            Dim dt As New Date
            'avoid the updating of the write time
            dt = finfo.LastWriteTime
            Dim bt As Bitmap = currentPicBox.Image
            Dim b As Image = bt.Clone


            ' Dim b As Bitmap = InitializeStandaloneImageCopy(Media.MediaPath)
            currentPicBox.Image.Dispose()
            Try


                finfo.Delete()
                b.Save(Media.MediaPath)
                finfo.LastWriteTime = dt
                '  currentPicBox.ImageLocation = Media.MediaPath
            Catch ex As System.UnauthorizedAccessException
                MsgBox("Enable write permissions on this folder")

            End Try
        End With
    End Sub


    Private Sub ControlEnter(sender As Object, e As EventArgs) Handles lbxShowList.Enter, lbxFiles.Enter
        'PFocus = CtrlFocus.Tree

        FocusControl = sender
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

    End Sub
    Private Sub TreeEnter(sender As Object, e As EventArgs) Handles tvMain2.Enter
        'PFocus = CtrlFocus.Tree
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
    End Sub






    Private Sub tmrLoadLastFolder_Tick(sender As Object, e As EventArgs) Handles tmrLoadLastFolder.Tick
        'Exit Sub
        tmrLoadLastFolder.Enabled = False
        'MsgBox("LLF")
        If Media.MediaPath = "" Then Exit Sub
        ' Media.MediaPath="E:\"
        HighlightCurrent(Media.MediaPath)
        'LoadDefaultShowList()
    End Sub


    Public Sub SetMotion(KeyCode As Integer)

        Dim intSpeed As Integer
        Media.Player.Ctlcontrols.pause()
        tmrMediaSpeed.Enabled = True
        Select Case KeyCode
            Case KeySpeed1
                intSpeed = 1000 / iPlaybackSpeed(0)
            Case KeySpeed2
                intSpeed = 1000 / iPlaybackSpeed(1)
            Case KeySpeed3
                intSpeed = 1000 / iPlaybackSpeed(2)
        End Select

        tmrMediaSpeed.Interval = intSpeed ' 020406

    End Sub

    Private Sub tmrMediaSpeed_Tick(sender As Object, e As EventArgs) Handles tmrMediaSpeed.Tick
        Media.Player.Ctlcontrols.step(1)
    End Sub

    Private Sub tmrInitialise_Tick(sender As Object, e As EventArgs) Handles tmrInitialise.Tick
        MsgBox("Initialise")
        ' ToolStripButton3_Click(Me, e)
        tmrInitialise.Enabled = False
    End Sub



    Private Sub ButtonListToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        SaveButtonlist()
    End Sub

    Private Sub tmrUpdateForm_Tick(sender As Object, e As EventArgs) Handles tmrUpdateForm.Tick
        MsgBox("UF")
        Me.Update()
    End Sub

    Private Sub lbxFiles_DoubleClick(sender As Object, e As EventArgs) Handles lbxFiles.DoubleClick, lbxShowList.DoubleClick
        Dim lbx As New ListBox
        lbx = CType(sender, ListBox)
        Process.Start("explorer.exe", lbx.SelectedItem)
    End Sub

    Public Sub UpdateFileInfo()
        If Media.DontLoad Then Exit Sub
        If Media.MediaPath = "" Then Exit Sub
        ' If Not FileLengthCheck(Media.MediaPath) Then Exit Sub
        Dim filepath As String
        If Media.IsLink Then
            filepath = Media.LinkPath
        Else
            filepath = Media.MediaPath
        End If
        Dim f As New FileInfo(filepath)
        If Not f.Exists Then Exit Sub
        Dim listcount = lbxFiles.Items.Count
        Dim showcount = lbxShowList.Items.Count
        Dim dt As Date
        dt = f.LastAccessTime
        Text = "Metavisua - " & filepath
        If f.LastWriteTime < dt Then dt = f.LastWriteTime
        If f.CreationTime < dt Then dt = f.CreationTime
        Select Case Math.Log10(f.Length)
            Case < 5
                tbDate.ForeColor = Color.Red
                tbDate.BackColor = Me.BackColor
            Case < 6
                tbDate.ForeColor = Color.Orange
                tbDate.BackColor = Color.Black

            Case < 7
                tbDate.ForeColor = Color.Yellow

                tbDate.BackColor = Color.Black
            Case < 8
                tbDate.ForeColor = Color.DarkGreen
                tbDate.BackColor = Me.BackColor
            Case < 9
                tbDate.ForeColor = Color.Blue
                tbDate.BackColor = Me.BackColor

            Case < 10
                tbDate.ForeColor = Color.Indigo
                tbDate.BackColor = Me.BackColor

            Case < 11
                tbDate.ForeColor = Color.Violet
                tbDate.BackColor = Me.BackColor

        End Select
        If Media.IsLink Then f = New IO.FileInfo(Media.LinkPath)
        tbDate.Text = dt.ToShortDateString & " " & dt.ToShortTimeString + " (" + Format(f.Length / 1024, "#,0.") + " Kb) " + Str(Int(f.Length / (128 * Media.Duration))) + " Kps"
        Dim c As Integer = lbxFiles.SelectedItems.Count
        Dim sl As Integer = lbxShowList.SelectedItems.Count
        If c > 1 And sl > 1 Then
            tbFiles.Text = "FOLDER:" & listcount & "(" & c & " selected)" & " SHOW:" & showcount & "(" & sl & " selected)"
        ElseIf c > 1 Then
            tbFiles.Text = "FOLDER:" & listcount & "(" & c & " selected)" & " SHOW:" & showcount
        ElseIf sl > 1 Then
            tbFiles.Text = "FOLDER:" & listcount & " SHOW:" & showcount & "(" & sl & " selected)"
        Else
            tbFiles.Text = "FOLDER:" & listcount & " SHOW:" & showcount
        End If
        tbFiles.Text = tbFiles.Text.Replace("SHOW:", "(" & GetDirSizeString(CurrentFolder) & ") SHOW:")
        tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
        tbLastFile.Text = Media.MediaPath
        tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        tbShowfile.Text = "SHOWFILE: " & LastShowList
        '   tbSpeed.Text = tbSpeed.Text = "SPEED (" & PlaybackSpeed & "fps)"
        tbButton.Text = "BUTTONFILE: " & ButtonFilePath
        TBFractionAbsolute.Text = "Fraction: " & Str(SP.FractionalJump) & "ths Absolute:" & Str(SP.AbsoluteJump) & "s"

        '  tbZoom.Text = iZoomFactor
        If Random.StartPointFlag Then
            tbStartpoint.Text = "START:RANDOM"
        Else
            tbStartpoint.Text = "START:NORMAL"
        End If

        Text = Text & " - " & Media.DisplayerName 'TODO remove displayer name for release. 
    End Sub
    ''' <summary>
    ''' Selects all files in the listbox matching a search string or regex
    ''' Returns chosen directory name
    ''' </summary>
    ''' <param name="blnRegex"></param>
    ''' <returns></returns>
    Private Function SelectSubList(lbx As ListBox, blnRegex As Boolean) As String
        Dim s As String
        If blnRegex Then
            s = InputBox("Enter regex search string")
        Else
            s = InputBox("Enter selection string")
        End If
        If s = "" Then
            Return s
            Exit Function
        End If

        SelectFromListbox(lbx, s, blnRegex)
        UpdateFileInfo()
        lastselection = s

        CancelDisplay(False)
        Return s
    End Function



    Private Sub DeleteShowListFiles()
        If MsgBox("This deletes all files in showlist. Sure?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
            For Each file In lbxShowList.Items
                My.Computer.FileSystem.DeleteFile(file, FileIO.UIOption.AllDialogs, FileIO.RecycleOption.SendToRecycleBin)

            Next
            lbxShowList.Items.Clear()
            CollapseShowlist(True)
        End If
    End Sub

    Private Sub LinearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinearToolStripMenuItem.Click
        AssignLinear(CurrentFolder, iCurrentAlpha, True)
    End Sub
    Public Sub HandleFunctionKeyDown(sender As Object, e As KeyEventArgs)

        Dim i As Byte = e.KeyCode - Keys.F5
        Dim s As StateHandler.StateOptions = NavigateMoveState.State
        '  CancelDisplay() 'Need to cancel display to prevent 'already in use' problems when moving files or deleting them. 
        If (e.Shift And e.Control And e.Alt) Or strVisibleButtons(i) = "" Then
            'Assign button
            AssignButton(i, iCurrentAlpha, 1, CurrentFolder, True) 'Just assign in all modes when all three control buttons held
            'Always update the button file. 
            If My.Computer.FileSystem.FileExists(ButtonFilePath) Then
                KeyAssignmentsStore(ButtonFilePath)
            Else
                SaveButtonlist()
            End If
            Exit Sub
        End If
        If s <> StateHandler.StateOptions.Navigate Then
            'Non navigate behaviour

            If e.Control And e.Shift Then
                'Jump to folder
                If strVisibleButtons(i) <> CurrentFolder Then
                    ChangeFolder(strVisibleButtons(i))
                    tvMain2.SelectedFolder = CurrentFolder
                ElseIf Random.OnDirChange Then
                    AdvanceFile(True, True) 'TODO: Whaat?
                End If
            ElseIf e.Shift Then
                'Move Folder
                MovingFolder(tvMain2.SelectedFolder, strVisibleButtons(i))
            Else
                'Move files
                Dim flag = blnSuppressCreate
                blnSuppressCreate = True
                If lbxShowList.Visible Then
                    Dim m = LBH.SelectedItemsList
                    LBH.RemoveItems(m)
                    MoveFiles(m, strVisibleButtons(i))

                Else
                    Dim m = FBH.SelectedItemsList
                    Media.Forceload = True
                    FBH.RemoveItems(m)
                    MoveFiles(m, strVisibleButtons(i))
                End If
                blnSuppressCreate = flag
            End If
        Else
            'Navigate behaviour
            If e.Shift And e.Control And strVisibleButtons(i) <> "" Then
                'Move folder
                MovingFolder(tvMain2.SelectedFolder, strVisibleButtons(i))

            ElseIf e.Shift Then
                'Move files
                MoveFiles(ListfromSelectedInListbox(lbxFiles), strVisibleButtons(i))
            Else
                'SWITCH folder
                If strVisibleButtons(i) <> CurrentFolder Then
                    ChangeFolder(strVisibleButtons(i))
                    'CancelDisplay()
                    tvMain2.SelectedFolder = strVisibleButtons(i)
                    ' FBH.DirectoryPath = strVisibleButtons(i)
                ElseIf Random.OnDirChange Then
                    'Change file if same folder
                    AdvanceFile(True, True)

                End If
            End If
        End If

        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

    End Sub

    Private Sub MovingFolder(Source As String, Dest As String)
        T = New Thread(New ThreadStart(Sub() MoveFolder(Source, Dest))) With {
            .IsBackground = True
        }
        T.SetApartmentState(ApartmentState.STA)

        T.Start()

        OnFolderMoved(Source)

    End Sub



    Private Async Sub DeleteEmptyFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteEmptyFoldersToolStripMenuItem.Click, IncludingAllSubfoldersToolStripMenuItem.Click

        'If Not MsgBox("This deletes all empty directories", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
        '    Exit Sub
        'Else
        Await DeleteEmptyFolders(New DirectoryInfo(CurrentFolder), sender Is IncludingAllSubfoldersToolStripMenuItem)
        tvMain2.RefreshTree(CurrentFolder)
        'End If

    End Sub

    Private Async Sub HarvestFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HarvestFoldersToolStripMenuItem.Click
        Await HarvestBelow(New DirectoryInfo(CurrentFolder))
        'tvMain2.RefreshTree(CurrentFolder)
        'Dim x As New BundleHandler(tvMain2, lbxFiles, CurrentFolder)

    End Sub



    Public Sub AddCurrentType(Recurse As Boolean)
        Showlist = AddFilesToCollection(strFilterExtensions(CurrentFilterState.State), Recurse)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)

    End Sub
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))

        If lst.Count = 0 Then Exit Sub

        LBH.Filter.State = Filter
        LBH.FillBox(lst)

        If lbx.Name = "lbxShowList" Then
            CollapseShowlist(False)
        End If
    End Sub


    Private Sub AddCurrentAndSubfoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentAndSubfoldersToolStripMenuItem.Click
        ProgressBarOn(1000)
        AddCurrentType(True)
        ProgressBarOff()
    End Sub

    Private Sub SaveListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveButtonFileToolStripMenuItem.Click
        SaveButtonlist()
    End Sub

    Private Sub loadbuttonfiletoolstripmenuitem_Click(sender As Object, e As EventArgs) Handles LoadButtonFileToolStripMenuItem.Click

        ButtonFilePath = LoadButtonList()
        KeyAssignmentsRestore(ButtonFilePath)
        AddToButtonFilesList(ButtonFilePath)
    End Sub

    Private Sub SaveListasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveListToolStripMenuItem.Click
        SaveShowlist()

    End Sub

    Private Sub NewListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        NewButtonList()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Visible = False

        ConstructMenuShortcuts()
        ConstructMenutooltips()
    End Sub

    Private Sub ToggleMove()
        NavigateMoveState.ToggleState()

        UpdateButtonAppearance()
    End Sub

    Private Async Sub BurstFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BurstFolderToolStripMenuItem.Click
        Await BurstFolder(New DirectoryInfo(CurrentFolder))
        'tvMain2.RemoveNode(CurrentFolder)
        '        tvMain2.RefreshTree(New IO.DirectoryInfo(CurrentFolder).Parent.FullName)
    End Sub

    Private Sub OnDirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles tvMain2.DirectorySelected
        tmrUpdateFileList.Enabled = False
        'PreferencesSave()
        lbxGroups.Items.Clear()
        CurrentFolder = e.Directory.FullName
        ' FBH.DirectoryPath = CurrentFolder
        tmrUpdateFileList.Enabled = True
        tmrUpdateFileList.Interval = 200



    End Sub

    Private Sub OnFilenamesParsed() Handles FNG.WordsParsed
        lbxGroups.Items.Clear()
        Dim i As Integer = 0
        For Each g In FNG.Groups
            lbxGroups.Items.Add(FNG.GroupNames(i) & " (" & g.Count & ")")
            Console.WriteLine(FNG.GroupNames(i))

            i += 1
            For Each m In g
                Console.WriteLine(m)
            Next
            Console.WriteLine()
        Next

    End Sub
    Private Sub tmrSlowMo_Tick(sender As Object, e As EventArgs) Handles tmrSlowMo.Tick
        Media.Player.Ctlcontrols.step(1)
        'Throw New Exception

    End Sub

    Private Sub ClearCurrentListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearCurrentListToolStripMenuItem.Click
        ClearShowList()
    End Sub
    Private Sub PicFullScreen() Handles PictureBox1.DoubleClick
        If ShiftDown Then
            blnSecondScreen = True
        Else
            blnSecondScreen = False
        End If
        GoFullScreen(True)

    End Sub


    Private Sub ThumbnailsStart()
        Dim t As New Thumbnails With {
            .ThumbnailHeight = 250,
            .Frame = 120
                    }
        Dim scr As Screen
        If blnSecondScreen Then
            scr = Screen.AllScreens(1)
        Else
            scr = Screen.AllScreens(0)
        End If
        t.Location = scr.Bounds.Location + New Point(100, 100)
        t.Left = scr.Bounds.Left
        t.Top = scr.Bounds.Top

        If FocusControl Is lbxFiles Or FocusControl Is lbxShowList Then
            t.List = Duplicatelist(AllfromListbox(FocusControl))
        Else
            t.List = Duplicatelist(AllfromListbox(lbxFiles))

        End If

        t.Text = CurrentFolder
        t.Show()

    End Sub



    Private Sub RandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomStartToolStripMenuItem.Click
        ToggleRandomStartPoint()
    End Sub


    'Private Sub tsbOnlyOne_Click(sender As Object, e As EventArgs)
    '    blnChooseOne = Not blnChooseOne
    'End Sub


    Private Sub AlphabeticToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlphaToolStripMenuItem.Click
        AssignAlphabetic(True)
        SaveButtonlist()
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        CurrentFilterState.State = cbxFilter.SelectedIndex

    End Sub

    Private Sub SingleFilePerFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SingleFilePerFolderToolStripMenuItem.Click
        blnChooseOne = True
        AddFilesToCollectionSingle(Showlist, PICEXTENSIONS & VIDEOEXTENSIONS, True)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)

    End Sub

    'Private Sub BundleToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    BundleFiles(lbxFiles, CurrentFolder)
    'End Sub

    Public Sub BundleFiles(lbx1 As ListBox, strFolder As String)
        If lbx1.SelectedItems.Count > 1 Then
        Else
            SelectSubList(FocusControl, False)

        End If
        Dim x As New BundleHandler(tvMain2, lbx1, strFolder)
        x.List = ListfromSelectedInListbox(lbx1)
        x.ListBox = lbx1
        Dim blnold = blnSuppressCreate
        blnSuppressCreate = False
        x.Bundle("")
        blnSuppressCreate = blnold
        'blnSuppressCreate = False
        'MoveFiles(ListfromSelectedInListbox(lbx1), strFolder, lbx1)
        'tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True

    End Sub
    Private Sub Groupfiles(ByVal m As FileNamesGrouper)
        MSFiles.URLSZero()
        blnSuppressCreate = True
        For i As Integer = 0 To m.Groups.Count - 1
            If lbxGroups.SelectedIndices.Contains(i) Then
                MoveFiles(m.Groups.Item(i), CurrentFolder & "\" & m.GroupNames.Item(i))
            End If
        Next
        blnSuppressCreate = False

        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub





    Private Sub DeadLinksSelect(lbx As ListBox)
        'Dim s As New List(Of String)

        'Dim lbx As ListBox = CType(FocusControl, ListBox)
        'T = New Thread(New ThreadStart(Sub() SelectDeadLinks(lbx))) With {
        '    .IsBackground = True
        '}
        'T.SetApartmentState(ApartmentState.STA)

        'T.Start()
        SelectDeadLinks(lbx)
        UpdateFileInfo()
    End Sub



    Private Sub tmrAutoTrail_Tick(sender As Object, e As EventArgs) Handles tmrAutoTrail.Tick
        If Not CtrlDown And Not Media.DontLoad Then AutoTrail(sender)

    End Sub

    Private Sub AutoTrail(sender As Object) 'Called on each autotrail tick, using tbAutotrail value for Time unit
        'Exit Sub
        Dim trail As New TrailMode

        AT.Framerates = SP.FrameRates
        Dim i As Byte = AT.SpeedIndex
        Dim TimeUnit As Byte = tbAutoTrail.Value
        AT.RandomTimes = {1 * TimeUnit, 3 * TimeUnit, 5 * TimeUnit, 10 * TimeUnit}
        AT.AdvanceChance = 25

        Select Case i
            Case 0, 1, 2
                SP.FrameRate = AT.Framerates(i)
                SpeedChange(New KeyEventArgs(speedkeys(i)))
                '       AT.Probability = DurBinParam
            Case 3
                tmrSlowMo.Enabled = False
                Media.Player.settings.rate = 1
                Media.Player.Ctlcontrols.play()
                '        tmrAutoTrail.Interval = Int(Rnd() * AT.RandomTimes(AT.Duration) * 500) + 500
        End Select
        tmrAutoTrail.Interval = Math.Max(Int(Rnd() * AT.RandomTimes(AT.Duration) * 500), 1000)

        'If we are holding Alt, change the file
        If AltDown Or Int(Rnd() * AT.AdvanceChance) + 1 = 1 Then
            If CtrlDown Then
                'Ctrl means dont change
            Else
                '    To change the file.
                Debug.Print("Change file")

                HandleKeys(sender, New KeyEventArgs(KeyNextFile))
                'count = Media.Markers.Count
            End If

        End If

        'This actually does the main job
        Dim ex1 As New KeyEventArgs(KeyJumpRandom Or Keys.Alt)
        Dim ex2 As New KeyEventArgs(KeyJumpRandom)
        'Go to markers first
        If AT.Counter > 0 Or (AT.Counter = 0 And Rnd() > 0.6) Then
            JumpToMark(ex1)
            If AT.Counter > 0 Then AT.Counter -= 1
        Else
            JumpRandom(False)
        End If
    End Sub



    'Private Sub tsbToggleRandomAdvance_Click(sender As Object, e As EventArgs)
    '    blnRandomAdvance(PFocus) = Not blnRandomAdvance(PFocus)
    'End Sub


    Private Sub ConstructMenuShortcuts()
        'Dim prefixkeys As Keys
        ''CTRL+
        'prefixkeys = Keys.Control
        ''AddCurrentFileListToolStripMenuItem.ShortcutKeys = KeyAddFile
        'LoadListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.L
        'SaveListToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        'BundleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        'AddCurrentFileToShowlistToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        ''Alt+
        'prefixkeys = Keys.Alt
        'ToggleRandomAdvanceToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        'ToggleMoveModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.M
        'ToggleJumpToMarkToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.J
        'ToggleRandomSelectToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.R
        'SlowToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'NormalToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        'FastToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.F
        'TrailerModeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        'RandomiseNormalToggleToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.Z

        ''CTRL+SHIFT
        'prefixkeys = Keys.Control + Keys.Shift
        'DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        'ClearCurrentListToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.C
        'NewButtonFileStripMenuItem.ShortcutKeys = prefixkeys + Keys.N
        'LoadButtonFileToolstripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        'SaveButtonfileasToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'DuplicatesToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.D
        'ThumbnailsToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T
        'DeleteEmptyFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.E
        'HarvestFoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.H
        'BurstFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.B
        'AddCurrentAndSubfoldersToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        ''CTRL + ALT
        'prefixkeys = Keys.Control + Keys.Alt

        'SingleFilePerFolderToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.S
        'SlowToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.S
        'NormalToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.N
        'FastToolStripMenuItem1.ShortcutKeys = prefixkeys + Keys.F
        'LinearToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.L
        'AlphabeticToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.A
        'TreeToolStripMenuItem.ShortcutKeys = prefixkeys + Keys.T


        'ToggleRandomStartToolStripMenuItem.ShortcutKeys = Keys.Shift + Keys.D
    End Sub
    Private Sub ConstructMenutooltips()
        AddCurrentFileListToolStripMenuItem.ToolTipText = "Add the current file list to the show list"
        DeleteEmptyFoldersToolStripMenuItem.ToolTipText = "Deletes all empty folders below the currently selected one"
        LoadListToolStripMenuItem.ToolTipText = "Load a previously saved show list"
        SaveListToolStripMenuItem.ToolTipText = "Save the current show list"
        BundleToolStripMenuItem.ToolTipText = "Moves all the selected files to a subfolder of their current location"



        ToggleRandomAdvanceToolStripMenuItem.ToolTipText = "Toggles advancing the file randomly or sequentially"
        'ToggleMoveModeToolStripMenuItem.ToolTipText = "Toggles move mode, which changes what the f keys do."
        ToggleJumpToMarkToolStripMenuItem.ToolTipText = "Makes movies start at the same fixed point"
        ToggleRandomSelectToolStripMenuItem.ToolTipText = "Toggles whether the first file, or a random file, is selected when the folder is changed."
        ToggleRandomStartToolStripMenuItem.ToolTipText = "Toggles whether movies begin at a random point"
        'SlowToolStripMenuItem.ToolTipText = "Sets slowest slideshow speed"
        'NormalToolStripMenuItem.ToolTipText = "Sets middle slideshow speed"
        'FastToolStripMenuItem.ToolTipText = "Sets fastest slideshow speed"
        TrailerModeToolStripMenuItem.ToolTipText = "Toggles auto-trail mode for movies"
        RandomiseNormalToggleToolStripMenuItem.ToolTipText = "Sets either all, or none of the random functions in one go"

        ClearCurrentListToolStripMenuItem.ToolTipText = "Clears the current show list"
        ' NewButtonFileStripMenuItem.ToolTipText = "Creates a new button file"
        LoadListToolStripMenuItem.ToolTipText = "Loads a previously-saved button file"
        SaveListToolStripMenuItem.ToolTipText = "Saves the current button file"
        DuplicatesToolStripMenuItem1.ToolTipText = "Opens the duplicates analyser"
        ThumbnailsToolStripMenuItem.ToolTipText = "Creates an interactive page of thumbnails"
        HarvestFoldersToolStripMenuItem.ToolTipText = "Takes all files from subfolders having fewer than a given number of files, and places them in the selected folder"
        BurstFolderToolStripMenuItem.ToolTipText = "Takes all files from the current folder, places them in the parent, and deletes the folder"
        AddCurrentAndSubfoldersToolStripMenuItem.ToolTipText = "Adds all files in the current folder, and all subfolder, and adds to the show list"
        'CTRL + ALT
        SelectDeadLinksToolStripMenuItem.ToolTipText = "Selects any .lnk files which are orphans"
        SingleFilePerFolderToolStripMenuItem.ToolTipText = "Adds a single file from all subfolders to the show list"

        LinearToolStripMenuItem.ToolTipText = "Assigns next 8 folders to the f buttons"
        AlphaToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, alphabetically"
        TreeToolStripMenuItem.ToolTipText = "Assigns subfolders to the f buttons, hierarchically (preference given to larger folders)"


    End Sub



    Private Sub ToggleRandomSelectToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomSelectToolStripMenuItem.Click
        ToggleRandomSelect()

    End Sub
    Private Sub ToggleRandomSelect()
        Random.OnDirChange = Not Random.OnDirChange
    End Sub


    Private Sub ToggleRandomAdvance() Handles ToggleRandomAdvanceToolStripMenuItem.Click
        Random.NextSelect = Not Random.NextSelect
        '  blnRandomAdvance(PFocus) = Random.NextSelect
    End Sub

    Private Sub ToggleRandomStartToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToggleRandomStartToolStripMenuItem.Click
        Random.StartPointFlag = Not Random.StartPointFlag
    End Sub

    Private Sub BundleToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles BundleToolStripMenuItem.Click
        If FocusControl IsNot tvMain2 Then
            BundleFiles(FocusControl, CurrentFolder)
        End If
    End Sub

    Private Sub ThumbnailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThumbnailsToolStripMenuItem.Click
        ThumbnailsStart()
    End Sub

    Private Sub TrailerModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrailerModeToolStripMenuItem.Click

        tmrAutoTrail.Enabled = Not tmrAutoTrail.Enabled
    End Sub



    Private Sub RandomiseNormalToggleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RandomiseNormalToggleToolStripMenuItem.Click
        ' Static Randomised As Boolean
        'Randomised = Not Randomised
        RandomFunctionsToggle()


    End Sub

    Public Sub RandomFunctionsToggle()
        Random.All = Not Random.All


    End Sub




    Private Sub SelectDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectDeadLinksToolStripMenuItem.Click
        DeadLinksSelect(FocusControl)

    End Sub



    Private Sub cbxFilter_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxFilter.SelectedIndexChanged
        If CurrentFilterState.State <> cbxFilter.SelectedIndex Then
            CurrentFilterState.State = cbxFilter.SelectedIndex
        End If
    End Sub

    Private Sub cbxOrder_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxOrder.SelectedIndexChanged
        If PlayOrder.State <> cbxOrder.SelectedIndex Then
            PlayOrder.State = cbxOrder.SelectedIndex
        End If
    End Sub



    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        SelectSubList(FocusControl, False)
    End Sub

    Private Sub RegexSearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RegexSearchToolStripMenuItem.Click
        SelectSubList(FocusControl, True)
    End Sub



    Private Sub AddCurrentFileToShowlistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentFileListToolStripMenuItem.Click
        AddAlltoShowlist() 'This one

    End Sub
    Private Sub AddAlltoShowlist()
        CollapseShowlist(False)
        LBH.AddList(FBH.ItemList)
    End Sub
    Private Sub AddCurrentFilesToShowList()
        Dim list As New List(Of String)

        For Each f In lbxFiles.SelectedItems
            list.Add(f)
        Next

        CollapseShowlist(False)
        LBH.AddList(list)
    End Sub
    Private Sub AddCurrentFilesToShowListold()

        For Each f In lbxFiles.SelectedItems
            AddSingleFileToList(Showlist, f)
        Next
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
        CollapseShowlist(False)
    End Sub






    Private Sub chbNextFile_CheckedChanged(sender As Object, e As EventArgs) Handles chbNextFile.CheckedChanged
        Random.NextSelect = chbNextFile.Checked
    End Sub



    Private Sub FilterMoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilterMoveToolStripMenuItem.Click
        '   FM.Recursive = False
        FM.FilterMoveFiles(CurrentFolder, False)
        tvMain2.RefreshTree(CurrentFolder)

        tmrUpdateFileList.Enabled = True

    End Sub

    Private Sub FilterMoveRecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FM.FilterMoveFiles(CurrentFolder, True)
        ReportAction(Format("Filtering {0}", CurrentFolder))
        tvMain2.RefreshTree(CurrentFolder)

        tmrUpdateFileList.Enabled = True

    End Sub


    Private Sub ReUniteFavesLinks()
        Dim s As List(Of String)
        s = ListfromSelectedInListbox(CType(FocusControl, ListBox))
        X.OrphanList = s
        ProgressBarOn(s.Count)
        T = New Thread(New ThreadStart(Sub() X.FindOrphans())) With {
            .IsBackground = True
        }
        T.SetApartmentState(ApartmentState.STA)

        T.Start()
        lbxFiles.SelectionMode = SelectionMode.One
        'While T.IsAlive
        '    ProgressIncrement(1)
        'End While
        '        X.FindOrphans()
        ProgressBarOff()
        UpdatePlayOrder(FBH)
    End Sub
    Private Sub OnOrphanTest(sender As Object, e As EventArgs) Handles X.NewOrphanTesting
        ProgressIncrement(1)

    End Sub
    Public Sub ReportAction(Msg As String)
        Console.WriteLine(Msg)
    End Sub

    Private Sub Filter(index As DateMove.DMY)
        CancelDisplay(False)
        DM.FilterByDate(CurrentFolder, False, index)
        tvMain2.RefreshTree(CurrentFolder)
    End Sub
    Private Sub FilterAlphabetic(Folders As Boolean, Files As Boolean)
        CancelDisplay(False)
        DM.FilterByAlpha(CurrentFolder, Folders:=Folders, Files:=Files)
        tvMain2.RefreshTree(CurrentFolder)

    End Sub

    Private Sub FilterAlphabetic(Folders As Boolean, Files As Boolean, Letters As Boolean)
        CancelDisplay(False)
        DM.FilterByAlpha(CurrentFolder, Folders, Files, Letters)
        tvMain2.RefreshTree(CurrentFolder)

    End Sub

    Private Sub tbAbsolute_MouseUp(sender As Object, e As MouseEventArgs) Handles tbAbsolute.MouseUp
        Media.SPT.State = StartPointHandler.StartTypes.ParticularAbsolute
        Media.SPT.Absolute = tbAbsolute.Value
        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(Media.SPT.Percentage) & "%"
        tbPercentage.Value = Media.SPT.Percentage
        Media.MediaJumpToMarker()
        MSFiles.SetStartpoints(Media.SPT)
    End Sub

    Private Sub tbPercentage_MouseUp(sender As Object, e As MouseEventArgs) Handles tbPercentage.MouseUp
        Media.SPT.State = StartPointHandler.StartTypes.ParticularPercentage
        Media.SPT.Percentage = tbPercentage.Value
        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(Media.SPT.Percentage) & "%"
        tbPercentage.Value = Media.SPT.Percentage

        Media.MediaJumpToMarker()
        MSFiles.SetStartpoints(Media.SPT)

    End Sub


    Private Sub ReclaimDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedToolStripMenuItem.Click
        ReUniteFavesLinks()

    End Sub


    Private Sub tvMain2_DragEnter(sender As Object, e As DragEventArgs) Handles tvMain2.DragEnter
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            If (e.KeyState And 8) = 8 Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.Move
            End If
        End If
    End Sub

    Private Async Sub RecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecursiveToolStripMenuItem.Click
        Dim x As New BundleHandler(tvMain2, lbxFiles, CurrentFolder)
        x.Maxfiles = Val(InputBox("Min no. of files to preserve folder? (Blank means all folders burst)",, ""))
        Dim d As New IO.DirectoryInfo(CurrentFolder)
        If d.Exists Then
            Await x.Burst(New IO.DirectoryInfo(CurrentFolder), True)
            tvMain2.RefreshTree(CurrentFolder)
            tmrUpdateFileList.Enabled = True
        End If

    End Sub

    Private Sub BySizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BySizeToolStripMenuItem.Click
        DM.FilterBySize(CurrentFolder, False)
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub


    Private Sub ShowlistToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles ShowlistToolStripMenuItem.Click
        If FocusControl Is lbxShowList Then

            Dim x As New ShowListForm
            x.Show()
            x.ListofFiles = FBH.ItemList
        End If
    End Sub


    Private Sub ButtonFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ButtonFormToolStripMenuItem.Click
        Dim x As New ButtonForm
        x.Show()
        x.handler.buttons.CurrentLetter = iCurrentAlpha

    End Sub


    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ByMonthToolStripMenuItem.Click
        Filter(DateMove.DMY.Month)
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ByDateToolStripMenuItem.Click
        Filter(DateMove.DMY.Day)
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ByYearToolStripMenuItem.Click
        Filter(DateMove.DMY.Year)
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub


    Private Sub tmrUpdateFileList_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFileList.Tick
        If AutoLoadButtons Then
            AddToButtonFilesList(CurrentFolder)
            UpdateFileInfo()
        End If
        FBH.DirectoryPath = CurrentFolder
        FBH.Random = Random
        FBH.SetFirst()
        tmrUpdateFileList.Enabled = False
    End Sub

    Public Sub AddToButtonFilesList(path As String)
        Dim folder As String = Autoload(FilenameFromPath(path, False))
        If CBXButtonFiles.Items.Contains(folder) Then
        ElseIf folder <> "" Then
            CBXButtonFiles.Items.Add(folder)
        End If
        CBXButtonFiles.SelectedItem = folder
    End Sub
    'Private Sub ChangedTree() Handles tvMain2.DirectorySelected
    '    ChangeFolder(tvMain2.SelectedFolder)
    'End Sub


    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles CalendarToolStripMenuItem.Click
        DM.FilterByCalendar(New DirectoryInfo(CurrentFolder))
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ByExtToolStripMenuItem.Click
        DM.FilterByType(CurrentFolder)
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub









    Private Sub Groupfiles(sender As Object, e As EventArgs) Handles ByNameToolStripMenuItem.Click
        Groupfiles(FNG)
    End Sub


    Private Sub LoadListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadListToolStripMenuItem.Click
        LoadShowList()
    End Sub

    Private Sub LevelFolders_Click(sender As Object, e As EventArgs) Handles PromoteFolderToolStripMenuItem.Click
        LevelAllFolders()
        '        PromoteFolder(New DirectoryInfo(CurrentFolder))
        tvMain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub


    Private Sub DashboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DashboardToolStripMenuItem.Click
        Dashboard.Show()
    End Sub

    Private Sub DuplicatesToolStripMenuItem1_Click_1(sender As Object, e As EventArgs) Handles DuplicatesToolStripMenuItem1.Click
        With FindDuplicates

            If FocusControl Is lbxShowList Then
                .List = LBH.ItemList
            Else

                .List = FBH.ItemList


            End If
            If .DuplicatesCount > 0 Then
                .ThumbnailHeight = 150
                .SetBounds(-1920, 0, 750, 900)
                .Show()
            Else
                MsgBox("No duplicates found.")
            End If
        End With
    End Sub



    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        PreferencesReset(True)
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        PreferencesSave()

    End Sub

    Private Sub RestoreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RestoreToolStripMenuItem.Click
        PreferencesGet()

    End Sub

    Private Sub FavouritesFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FavouritesFolderToolStripMenuItem.Click
        Try
            If CurrentFavesPath = "" Then
                CurrentFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) & "\MVFavourites"
            End If
            If MsgBox("Make " & CurrentFolder & " the new Favourites folder?", MsgBoxStyle.OkCancel, "Assign Favourites folder") = MsgBoxResult.Ok Then
                CurrentFavesPath = CurrentFolder
                FaveMinder.NewPath(CurrentFavesPath)
                PreferencesSave()

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub tmrMovieSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrMovieSlideShow.Tick
        tmrMovieSlideShow.Interval = Rnd() * 3000 + 1500
        Dim x = Int(Rnd() * 100)
        Dim FSPercent = 70
        If x < FSPercent Then
            Media.Speed.Fullspeed = True
            Media.Position = Media.Player.Ctlcontrols.currentPosition
            Media.Player.settings.rate = 1
            Media.Player.Ctlcontrols.play()
            tmrSlowMo.Enabled = False
            Media.Speed.Fullspeed = True
        Else
            Media.Speed.Speed = Int((x - FSPercent) / ((100 - FSPercent) / 2.8))
            SpeedChange(New KeyEventArgs(speedkeys(Media.Speed.Speed)))
        End If

        '  tmrMovieSlideShow.Interval = tmrMovieSlideShow.Interval * 29 / Media.Speed.FrameRates(Media.Speed.Speed)
        AdvanceFile(True, False)
    End Sub

    Private Sub cbxStartPoint_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxStartPoint.SelectedIndexChanged
        Media.SPT.State = cbxStartPoint.SelectedIndex
        'If cbxStartPoint.SelectedIndex <> Media.StartPoint.State Then cbxStartPoint.SelectedIndex = Media.StartPoint.State

    End Sub

    Private Sub btn8_DragEnter(sender As Object, e As DragEventArgs) Handles btn8.DragEnter, btn1.DragEnter
        e.Effect = DragDropEffects.Copy
        MsgBox(e.Data.GetData(DataFormats.Text))
    End Sub

    Private Sub btn8_DragDrop(sender As Object, e As DragEventArgs) Handles btn8.DragDrop, btn1.DragDrop
        Dim i As Integer = Val(sender.name(3))
        MsgBox(i)
        '  AssignButton(i - 1, e.Data.GetData(DataFormats.Text).ToString)
    End Sub

    Private Sub RefreshSelectedLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshSelectedLinksToolStripMenuItem.Click, SelectedDeepToolStripMenuItem.Click
        RefreshSelectedLinks(sender)
    End Sub

    Private Sub RefreshSelectedLinks(sender As Object)
        ProgressBarOn(100)
        Dim op2 As New OrphanFinder
        Dim lbx As New ListBox
        Dim files As New List(Of String)
        lbx = FocusControl

        For Each m In lbx.SelectedItems
            files.Add(m)
        Next
        op2.OrphanList = files
        If op2.FindOrphans(sender Is SelectedDeepToolStripMenuItem) Then FBH.FillBox()
        ProgressBarOff()
    End Sub

    Private Sub lbxGroups_DoubleClick(sender As Object, e As EventArgs) Handles lbxGroups.DoubleClick
        FNG.AdvanceOption()
        FNG.Filenames = FBH.ItemList
    End Sub


    Private Sub SelectNonFavouritsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectNonFavouritsToolStripMenuItem.Click
        'Op.ImportLinks(Showlist)
        'Select files with links
        Dim hilist As New List(Of String)
        Dim FLBH As New ListBoxHandler(lbxFiles)
        If FocusControl Is lbxFiles Then
            FLBH = FBH
        Else
            FLBH = LBH
        End If
        For Each f In FLBH.ItemList
            If FaveMinder.GetLinksOf(f).Count = 0 Then
                hilist.Add(f)
            End If
        Next
        FLBH.SelectItems = hilist
        ' FBH.ListBox.SelectionMode = SelectionMode.One
    End Sub

    Private Sub AddCurrentToShowlistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddCurrentToShowlistToolStripMenuItem.Click
        AddCurrentFilesToShowList()
    End Sub

    Private Sub ExperimentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExperimentToolStripMenuItem.Click
        'MovieSwapTest.Show()
    End Sub

    Private Sub ByLinkFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ByLinkFolderToolStripMenuItem.Click
        DM.FilterByLinkFolder(CurrentFolder)
    End Sub

    Private Sub CreateListFromLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateListFromLinksToolStripMenuItem.Click
        ListFromLinks(ListfromSelectedInListbox(lbxFiles))
    End Sub

    Private Sub NewIndex_Tick(sender As Object, e As EventArgs) Handles NewIndex.Tick

        IndexHandler(FocusControl, e)
        NewIndex.Enabled = False
    End Sub



    Private Sub ToolStripTextBox1_Click_1(sender As Object, e As EventArgs) Handles AutoButton.Click
        AutoButtonsToggle()

    End Sub

    Private Sub Response_Tick(sender As Object, e As EventArgs) Handles Response.Tick
        Response.Enabled = False
    End Sub

    Private Sub DisplayedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DisplayedToolStripMenuItem.Click
        FaveMinder.RedirectShortCutList(New FileInfo(lbxFiles.SelectedItem), AllfromListbox(lbxShowList))
    End Sub

    Private Sub chbInDir_CheckedChanged(sender As Object, e As EventArgs) Handles chbOnDir.CheckedChanged
        Random.OnDirChange = chbOnDir.Checked
    End Sub

    Private Sub chbEncrypt_CheckedChanged(sender As Object, e As EventArgs) Handles chbEncrypt.CheckedChanged
        Encrypted = chbEncrypt.Checked
    End Sub

    Private Sub ForceFavouritesReloadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceFavouritesReloadToolStripMenuItem.Click
        AllFaveMinder.NewPath(GlobalFavesPath)
    End Sub

    Private Sub ForceDirectoriesReloadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ForceDirectoriesReloadToolStripMenuItem.Click
        If MsgBox("Reload directories list? Could take a while!", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
            DirectoriesList = GetDirectoriesList(Rootpath, True)
        End If
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveToolStripMenuItem.Click
        If FocusControl Is LBH.ListBox Then
            RemoveBrackets(LBH.ItemList)
        Else
            RemoveBrackets(FBH.ItemList)
            FBH.Refresh()
        End If
    End Sub


    Private Sub lbxFiles_KeyDown(sender As Object, e As KeyEventArgs) Handles lbxFiles.KeyDown, lbxShowList.KeyDown
        If e.KeyCode = KeyTraverseTree Or e.KeyCode = KeyTraverseTreeBack Then
            tvMain2.tvFiles_KeyDown(sender, e)
        End If
    End Sub



    Private Sub HideDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HideDeadLinksToolStripMenuItem.Click
        SelectDeadLinks(lbxFiles)
        FBH.RemoveItems(ListfromSelectedInListbox(lbxFiles))
    End Sub

    Private Sub InvertSelectionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InvertSelectionToolStripMenuItem.Click
        If FocusControl Is FBH.ListBox Then
            FBH.InvertSelected()
        Else
            LBH.InvertSelected()
        End If
    End Sub


    Private Sub MainForm_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint

    End Sub

    Private Sub AlphabeticGroupsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlphabeticGroupsToolStripMenuItem.Click
    End Sub

    Private Sub chbPreviewLinks_CheckedChanged(sender As Object, e As EventArgs) Handles chbPreviewLinks.CheckedChanged

    End Sub

    Private Sub Scrubber_Paint(sender As Object, e As PaintEventArgs) Handles Scrubber.Paint
        DrawScrubberMarks()

        'MsgBox("Uh-oh")
    End Sub
    Private Sub ToggleRandomJump()
        tmrJumpRandom.Interval = tbScanRate.Value
        tmrJumpRandom.Enabled = Not tmrJumpRandom.Enabled
        chbScan.Checked = tmrJumpRandom.Enabled
    End Sub
    Private Sub chbSeparate_CheckedChanged(sender As Object, e As EventArgs) Handles chbSeparate.CheckedChanged
        separate = chbSeparate.Checked

        MSFiles.DockMedias(separate)

    End Sub

    Private Sub tvMain2_Load(sender As Object, e As EventArgs) Handles tvMain2.Load

    End Sub

    Private Sub PromoteFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PromoteFileToolStripMenuItem.Click
        Dim folder As New IO.DirectoryInfo(FBH.DirectoryPath)
        MoveFiles(FBH.SelectedItemsList, folder.Parent.FullName)
    End Sub

    Private Sub FilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilesToolStripMenuItem.Click
        FilterAlphabetic(False, True)

    End Sub

    Private Sub FoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FoldersToolStripMenuItem.Click
        FilterAlphabetic(True, False)
    End Sub

    Private Sub NewButtonFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewButtonFileToolStripMenuItem.Click
        NewButtonList()
    End Sub

    Private Sub SelectedDeepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedDeepToolStripMenuItem.Click

    End Sub

    Private Sub ListDeadFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListDeadFilesToolStripMenuItem.Click
        FormDeadFiles.Show()
        SelectDeadLinks(FocusControl)
        Dim op2 As New OrphanFinder
        Dim lbx As New ListBox
        Dim files As New List(Of String)
        lbx = FocusControl

        For Each m In lbx.SelectedItems
            files.Add(m)
        Next
        op2.OrphanList = files
        Dim deadfiles As New List(Of String)
        deadfiles = op2.ListOfDeadFiles
        deadfiles.Sort()

        For Each f In deadfiles
            FormDeadFiles.TextBox1.Text = FormDeadFiles.TextBox1.Text & f & vbCrLf

        Next
    End Sub

    Private Sub chbLoadButtonFiles_CheckedChanged(sender As Object, e As EventArgs) Handles chbLoadButtonFiles.CheckedChanged
        AutoLoadButtons = chbLoadButtonFiles.Checked
    End Sub

    Private Sub LettersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LettersToolStripMenuItem.Click
        FilterAlphabetic(False, True, True)
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As Object, e As CancelEventArgs) Handles ContextMenuStrip1.Opening

    End Sub

    Private Sub SelectThenRefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectThenRefreshToolStripMenuItem.Click
        ProgressBarOn(1000)
        SelectDeadLinks(FocusControl)
        RefreshSelectedLinks(sender)
        ProgressBarOff()
    End Sub

    Private Sub tmrProgressBar_Tick(sender As Object, e As EventArgs) Handles tmrProgressBar.Tick
        tmrProgressBar.Interval = 1000
        ProgressIncrement(10)
    End Sub

    Private Sub SelectAndBundleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAndBundleToolStripMenuItem.Click
        SelectDeadLinks(FocusControl)
        Dim m As New BundleHandler(tvMain2, FocusControl, CurrentFolder)
        m.Bundle("Dead")

    End Sub

    Private Sub RefreshAllFilesWithLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshAllFilesWithLinksToolStripMenuItem.Click
        SelectDeadLinks(FocusControl)
        InvertSelectionToolStripMenuItem_Click(sender, e)
        RefreshLinks(ListfromSelectedInListbox(FocusControl))
    End Sub

    Private Sub CBXButtonFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBXButtonFiles.SelectedIndexChanged
        Dim folder As String = CBXButtonFiles.SelectedItem
        Autoload(FilenameFromPath(folder, False))
    End Sub

    Private Sub CHBAutoAdvance_CheckedChanged(sender As Object, e As EventArgs) Handles CHBAutoAdvance.CheckedChanged
        AutoTraverse = CHBAutoAdvance.Checked
        chbOnDir.Checked = Not AutoTraverse
    End Sub

    Private Sub chbShowAttr_CheckedChanged(sender As Object, e As EventArgs) Handles chbShowAttr.CheckedChanged
        'MsgBox("Attribute Changed")
    End Sub

    Private Sub chbAutoTrail_CheckedChanged(sender As Object, e As EventArgs) Handles chbAutoTrail.CheckedChanged
        tmrAutoTrail.Enabled = chbAutoTrail.Checked
    End Sub

    Private Sub MovieScanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovieScanToolStripMenuItem.Click
        ToggleRandomJump()
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles tmrJumpRandom.Tick
        tmrJumpRandom.Interval = tbScanRate.Value
        JumpRandom(False)
    End Sub

    Private Sub ResetButtonsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetButtonsToolStripMenuItem.Click
        CBXButtonFiles.SelectedItem = "Watch.msb" 'TODO Make general

    End Sub

    Private Sub PreferencesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PreferencesToolStripMenuItem1.Click
        Mysettings.LoadForm()
    End Sub

    Private Sub LBH_ListboxChanged(sender As Object, e As EventArgs) Handles LBH.ListIndexChanged
        HighlightCurrent(lbxShowList.SelectedItem)
    End Sub

    Private Sub lbxFiles_Leave(sender As Object, e As EventArgs) Handles lbxFiles.Leave
        If lbxShowList.Visible Then
            lbxShowList.Focus()
        Else
            tvMain2.Focus()
        End If
    End Sub

    Private Sub tbScanRate_ValueChanged(sender As Object, e As EventArgs) Handles tbScanRate.ValueChanged
        tmrJumpRandom.Interval = tbScanRate.Value
    End Sub
    Private Sub MainForm_Mousewheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        If Media.PicHandler.WheelScroll And Media.MediaType = Filetype.Movie Then
            AdvanceFile(e.Delta < 0, Random.NextSelect)
        End If

    End Sub


    Private Sub MainForm_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick, PictureBox1.MouseClick, PictureBox2.MouseClick, PictureBox3.MouseClick

        If e.Button = MouseButtons.XButton1 Then AdvanceFile(True, Random.NextSelect)
        If e.Button = MouseButtons.XButton2 Then AdvanceFile(False, Random.NextSelect)
        If e.Button = MouseButtons.Right Then AddMarker()
        '  If e.Button = MouseButtons.Left Then MSFiles.ClickAllPics(sender, e)
    End Sub

    Private Sub tmrUpdateFolderSelection_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFolderSelection.Tick

    End Sub

    Private Sub ReplaceOldLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReplaceOldLinksToolStripMenuItem.Click
        Dim x As New List(Of String)
        x = FBH.SelectedItemsList
        For Each m In x
            ConvertLink(m)
        Next

    End Sub

    Private Sub FBH_ListIndexChanged(sender As Object, e As EventArgs) Handles FBH.ListIndexChanged

    End Sub

    Private Sub CreateToolTips()
        For Each ctl In Me.Controls
            Select Case ctl.name
                Case "cbxFilter"
                Case "cbxOrder"
                Case "cbxStartPoint"
                Case "cbxStartPoint"
                Case "chbNextFile"
                Case "lbxFiles"
                Case "lbxShowList"
                Case "lbxGroups"

            End Select
        Next
    End Sub

    Private Sub chbSlideShow_CheckedChanged(sender As Object, e As EventArgs) Handles chbSlideShow.CheckedChanged
        tmrMovieSlideShow.Enabled = chbSlideShow.Checked
    End Sub

    Private Sub tbMovieSlideShowSpeed_Scroll(sender As Object, e As EventArgs) Handles tbMovieSlideShowSpeed.Scroll
        tmrMovieSlideShow.Interval = tbMovieSlideShowSpeed.Value
    End Sub

    Private Sub tbAutoTrail_Scroll(sender As Object, e As EventArgs) Handles tbAutoTrail.Scroll

    End Sub
#End Region

End Class
