
Option Explicit On
Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports MasaSam.Forms.Controls
'Imports Microsoft.Office.Interop.Excel
Imports Microsoft.VisualBasic.FileIO


Public Class FormMain
    Inherits Form

#Region "Members"
    Public DB As New Database
    Public Initialising As Boolean = True
    Public AutoLoadButtons As Boolean = False
    Public defaultcolour As Color = Color.Aqua
    Public movecolour As Color = Color.Orange
    Public sound As New AxWindowsMediaPlayer
    Public AutoTraverse As Boolean = False
    Public WithEvents FNG As New FilenamesGrouper
    Public WithEvents FG As New FileGrouper
    Public WithEvents Random As New RandomHandler
    Public WithEvents NavigateMoveState As New StateHandler()
    Public WithEvents CurrentFilterState As New FilterHandler
    Public WithEvents PlayOrder As New SortHandler
    Private WithEvents FM As New FilterMove()
    Private WithEvents DM As New DateMove
    Public WithEvents Att As New Attributes
    Private WithEvents Op As New OrphanFinder
    Public WithEvents SP As New SpeedHandler
    Public WithEvents AT As New AutoTrailer
    Public WithEvents X As New OrphanFinder
    Public WithEvents VT As New VideoThumbnailer
    Public folderwatcher As FileSystemWatcher
    Public WithEvents FOP As FileOperations
    Public WithEvents BH As New ButtonHandler With {.RowProgressBar = ProgressBar1}

    'Public openvr As New openvrsystem()
    Public SlowMoSoundOpt As SlowMoSoundOptions = SlowMoSoundOptions.Silent

    ' Public WithEvents Response As New Timer
    Public FocusControl As New Control
    Public DraggedFolder As String

    Public T As Thread
    ' Public FirstButtons As New ButtonForm
    Public ScrubberProportion As Decimal = 0.98
    Public FavouritesFormOpen As Boolean = False
    Public Marks As New MarkPlacement With {.Fractions = SP.FractionalJump, .SmallJumps = SP.AbsoluteJump}
    Public speedkeys = {KeySpeed1, KeySpeed2, KeySpeed3}
    Public WithEvents FBH As New FileboxHandler()
    Public WithEvents LBH As New ListBoxHandler()
    Private splashScreen As SplashScreenForm
    Private ScanInstructions = {"Movie scan in progress" & vbCrLf & "Selects next movie automatically.",
        "Movie Scan in progress" & vbCrLf & "Change speed with slider.",
        "Auto Trail in progress" & vbCrLf & "Change core speed change with slider"}

#End Region
    Private WithEvents keyPressTimer As New Windows.Forms.Timer()
    Public Parameters As New AppSettings
    Private Sub CurrentDomain_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs)

        'LastTimeSuccessful = False
        ' PreferencesSave()
        SavePreferences(Parameters)
    End Sub


    Public Sub New()

        splashScreen = New SplashScreenForm
        splashScreen.Show()
        Application.DoEvents()
        ' This call is required by the designer.

        InitializeComponent()

        ' Enable high-DPI support
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi

        InitializeFolderWatcher()
        Parameters = SettingsManager.LoadSettings
    End Sub

    Private Sub InitializeFolderWatcher()
        folderwatcher = New FileSystemWatcher()
        ' Set the folder to watch
        folderwatcher.NotifyFilter = NotifyFilters.FileName Or NotifyFilters.DirectoryName ' Watch for changes in file and folder names
        folderwatcher.Path = "Q:\"
        ' Add event handlers
        AddHandler folderwatcher.Changed, AddressOf Folderwatcher_Changed
        AddHandler folderwatcher.Created, AddressOf Folderwatcher_Changed
        AddHandler folderwatcher.Deleted, AddressOf Folderwatcher_Changed
        AddHandler folderwatcher.Renamed, AddressOf Folderwatcher_Changed

        folderwatcher.EnableRaisingEvents = True ' Enable the FileSystemWatcher
    End Sub
    Public Sub Folderwatcher_Changed()
        If FBH.ListBox.InvokeRequired Then
            FBH.ListBox.Invoke(New MethodInvoker(AddressOf FBH.Refresh))
        Else
            FBH.Refresh()
        End If
    End Sub
#Region "Control Handlers"


    Public Sub ControlSetFocus(control As Control)
        ' Set focus to the control, if it can receive focus.
        'Exit Sub
        If control Is tvmain2 Then
            tvmain2.Focus()
        Else
            FocusControl = control
        End If
        If control.CanFocus Then
            control.Focus()
        End If
        If TypeOf (control) Is ListBox Then
            Dim listboxcontrol As ListBox = CType(control, ListBox)
            MSFiles.FileList = listboxcontrol.Items.Cast(Of String).ToList()
        End If
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
    End Sub

    Private accumulatedKeys As String = String.Empty

    Public Sub Main_KeyDown(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        KeyDownFlag = True
        ShiftDown = e.Shift
        CtrlDown = e.Control
        AltDown = e.Alt
        If e.KeyCode >= Keys.A AndAlso e.KeyCode <= Keys.Z Then
            accumulatedKeys &= Chr(e.KeyCode)
        End If
        BH.UpdateButtonAppearance()
        Try
            HandleKeys(sender, e)
            Report(sender.ToString)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        If e.KeyCode = KeyBackUndo Then
            e.SuppressKeyPress = True
        End If
        keyPressTimer.Stop()
        keyPressTimer.Start()
    End Sub
    Private Sub Main_KeyUp(sender As Object, e As KeyEventArgs) Handles MyBase.KeyUp
        KeyDownFlag = False
        ShiftDown = e.Shift
        CtrlDown = e.Control
        AltDown = e.Alt

        If e.KeyData <> (Keys.F4 Or AltDown) Then
            Try
                BH.UpdateButtonAppearance()
            Catch ex As Exception

            End Try
        End If

        e.Handled = True
        '  e.SuppressKeyPress = True

    End Sub
    Private Sub Main_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        tmrAutoTrail.Enabled = False
        tmrMovieSlideShow.Enabled = False
        MSFiles.Soundhandler.ResetPos.Enabled = False
        LastTimeSuccessful = True
        SavePreferences(Parameters)
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



    Sub EscapeMultiSelect(sender As Object, e As KeyEventArgs) Handles lbxFiles.KeyDown
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
        NewIndex.Enabled = True

    End Sub
    Public Sub IndexHandler(sender As Object, e As EventArgs)
        Dim lbx As ListBox = sender
        If lbx.SelectionMode = SelectionMode.One Then
            Dim i As Long = lbx.SelectedIndex
            If i = -1 Then
            Else
                MSFiles.FileList = lbx.Items.Cast(Of String)().ToList()
                If blnFullScreen Then
                    FullScreen.FSFiles.ListIndex = i
                Else
                    MSFiles.ListIndex = i
                End If
            End If
        End If
    End Sub

    Private Sub tmrSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrSlideShow.Tick
        'SP.Slideshow = True
        If AltDown Then Exit Sub
        If Media.Playing Then
            If CtrlDown Then
                AdvanceFile(True, Random.NextSelect)
            Else

                If chbScan.Checked And Media.Markers.Count <> 0 Then
                    'Scan movie for marks
                    If Media.IncrementLinkCounter(True) = 0 Then
                        'Static pos As Long = Media.Position
                        If Media.Position > Media.Markers(Media.Markers.Count - 1) Then
                            MediaLargeJump(False, True)
                        Else
                            AdvanceFile(True, Random.NextSelect)
                        End If
                    Else
                        Media.MediaJumpToMarker(ToMarker:=True)
                    End If
                Else
                    MediaLargeJump(False, True)
                End If
            End If
        Else
            AdvanceFile(True, Random.NextSelect)
        End If
    End Sub

    Private Sub RotatePic(PicBox As PictureBox, blnLeft As Boolean)
        If PicBox.Image Is Nothing Then Exit Sub
        If Media.MediaType <> Filetype.Pic Then Exit Sub
        '  Media.PicHandler.GetImage(Media.MediaPath)
        With PicBox.Image
            If blnLeft Then
                .RotateFlip(RotateFlipType.Rotate90FlipNone)
            Else
                .RotateFlip(RotateFlipType.Rotate270FlipNone)
            End If
            PicBox.Refresh()

            Dim finfo As New FileInfo(Media.MediaPath)
            Dim dt As New Date
            'avoid the updating of the write time
            dt = finfo.LastWriteTime
            Dim bt As Bitmap = PicBox.Image
            Dim b As Image = bt.Clone
            DisposePic(PicBox)
            'MSFiles.CancelURL(Media.MediaPath)
            ' Dim b As Bitmap = InitializeStandaloneImageCopy(Media.MediaPath)

            Try


                My.Computer.FileSystem.DeleteFile(finfo.FullName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                b.Save(Media.MediaPath)
                PicBox.Image = LoadImage(Media.MediaPath)
                finfo.LastWriteTime = dt
                '  currentPicBox.ImageLocation = Media.MediaPath
            Catch ex As System.UnauthorizedAccessException
                MsgBox("Enable write permissions on this folder")

            End Try
        End With
    End Sub


    Private Sub ControlEnter(sender As Object, e As EventArgs) Handles lbxFiles.Enter
        'PFocus = CtrlFocus.Tree
        If sender Is lbxFiles Then
            LBHH = FBH

        Else
            LBHH = LBH
        End If
        FocusControl = sender
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

    End Sub
    Private Sub TreeEnter(sender As Object, e As EventArgs)
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
    Private Sub tmrInitialise_Tick(sender As Object, e As EventArgs) Handles tmrInitialise.Tick
        MsgBox("Initialise")
        ' ToolStripButton3_Click(Me, e)
        tmrInitialise.Enabled = False
    End Sub
    Private Sub ButtonListToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        BH.SaveButtonSet()
    End Sub
    Private Sub tmrUpdateForm_Tick(sender As Object, e As EventArgs) Handles tmrUpdateForm.Tick
        MsgBox("UF")
        Me.Update()
    End Sub

    Private Sub lbxFiles_DoubleClick(sender As Object, e As EventArgs) Handles lbxFiles.DoubleClick
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
        Text = "Metavisua - " & filepath
        Dim f As New FileInfo(filepath)
        If Not f.Exists Then Exit Sub
        Dim listcount = lbxFiles.Items.Count
        Dim showcount = lbxShowList.Items.Count
        Dim dt As Date = GetDate(f)
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
        'Dim tx As Task(Of String) = GetDirSizeString(CurrentFolder)
        Dim x As String = "XX"
        'x = tx.Result
        tbFiles.Text = tbFiles.Text.Replace("SHOW:", "(" & x & ") SHOW:")
        tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
        tbLastFile.Text = Media.MediaPath
        tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        tbShowfile.Text = "SHOWFILE: " & LBH.ListBox.Tag
        '   tbSpeed.Text = tbSpeed.Text = "SPEED (" & PlaybackSpeed & "fps)"
        TBFractionAbsolute.Text = "Fraction: " & Str(SP.FractionalJump) & "ths Absolute:" & Str(SP.AbsoluteJump) & "s"

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
        BH.AssignLinear(New DirectoryInfo(CurrentFolder))
    End Sub
    Private Sub AlphaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AlphaToolStripMenuItem.Click
        ProgressBarOn(100, "Assigning buttons alphabetically")
        BH.AssignAlphabetical(New DirectoryInfo(CurrentFolder))
        ProgressBarOff()
    End Sub
    Private Sub TreeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TreeToolStripMenuItem.Click
        BH.AssignAlphaHierarchical(CurrentFolder, 15)

    End Sub





    Private Sub RespondToKey(sender As Object, e As KeyEventArgs)
        Dim mBH As New ButtonHandler
        If TypeOf (sender) Is FormButton Then
            mBH = sender.bh
        Else
            mBH = BH
        End If
        Select Case e.KeyCode

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If e.Control AndAlso e.Alt Then
                    If mBH IsNot BH Then
                        mBH.buttons.CurrentLetter = mBH.buttons.CurrentLetter

                        Select Case e.KeyCode
                            Case Keys.L
                                mBH.AssignLinear(New DirectoryInfo(CurrentFolder))
                            Case Keys.A
                                mBH.AssignAlphabetical(New DirectoryInfo(CurrentFolder))
                            Case Keys.T
                                mBH.AssignTreeNew(CurrentFolder, Val(InputBox("Size Magnitude")))
                        End Select
                    End If

                ElseIf e.Control Then
                    If e.KeyCode = Keys.I Then
                        AddFolders.Show()
                        AddFolders.Folder = CurrentFolder
                        tvmain2.RefreshTree(CurrentFolder)
                    End If

                ElseIf e.Alt And e.Shift Then
                    Select Case e.KeyCode
                        Case Keys.I
                            VideoTrim.StartTime = Media.Position
                            VideoTrim.Active = True
                            MsgBox("Extraction at current point begun. Choose end point (Alt shift O)")
                        Case Keys.O
                            VideoTrim.Finish = Media.Position
                            Marks.Create()
                            MsgBox("Now Alt Shift P to process, or carry on to cancel.")
                        Case Keys.P
                            AddHandler VideoTrim.Finished, AddressOf VideoTrimFinished
                            VideoTrim.InputFile1 = MSFiles.CurrentItem
                            VideoTrim.Main()

                            VideoTrim.Active = False
                    End Select

                Else

                    'Advance row, or change letter
                    If mBH.buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode) Then

                        mBH.buttons.NextRow(LetterNumberFromAscii(e.KeyCode), e.Alt)
                    Else
                        mBH.buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode)
                    End If
                    mBH.SwitchRow(mBH.buttons.CurrentRow)


                End If

                ' e.SuppressKeyPress = True
            Case Else

        End Select
    End Sub
    Private Sub VideoTrimFinished(sender As Object, e As EventArgs)
        Dim output As String = sender
        MsgBox("Finished extracting")
    End Sub
    Private Sub MovingFolder(Source As String, Dest As String)
        T = New Thread(New ThreadStart(Sub() MoveFolder(Source, Dest))) With {
            .IsBackground = True
        }
        T.SetApartmentState(ApartmentState.STA)

        T.Start()

        OnFolderMoved(Source, Dest)

    End Sub



    Private Sub DeleteEmptyFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteEmptyFoldersToolStripMenuItem.Click, IncludingAllSubfoldersToolStripMenuItem.Click

        'If Not MsgBox("This deletes all empty directories", MsgBoxStyle.OkCancel) = MsgBoxResult.Ok Then
        '    Exit Sub
        'Else

        DeleteEmptyFolders(CurrentFolder)

        tvmain2.RefreshTree(CurrentFolder)
        'End If

    End Sub

    Private Async Sub HarvestFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HarvestFoldersToolStripMenuItem.Click
        '        Await HarvestBelow(New DirectoryInfo(CurrentFolder))
        Dim FC As New FolderCleaner
        FC.FolderPath = CurrentFolder
        FC.MaxFileCount = 0
        FC.CleanupFolder()
        AddHandler FC.CleanupFinished, AddressOf CleanupFinishedHandler

        'Dim x As New BundleHandler(tvMain2, lbxFiles, CurrentFolder)

    End Sub

    Private Sub CleanupFinishedHandler(ByVal folderPath As String)
        tvmain2.RefreshTree(folderPath)
    End Sub

    Public Sub AddCurrentType(Recurse As Boolean)
        Showlist = AddFilesToCollection(strFilterExtensions(CurrentFilterState.State), Recurse)
        If Showlist.Count = 0 Then
            MsgBox("No files found matching your search")
        Else
            FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)
        End If

    End Sub
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))
        Dim f As New FormShowList

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
        BH.SaveButtonSet()

    End Sub

    Private Sub Loadbuttonfiletoolstripmenuitem_Click(sender As Object, e As EventArgs) Handles LoadButtonFileToolStripMenuItem.Click
        ProgressBarOn(100, "Loading button file " & ButtonFilePath)
        ButtonFilePath = LoadButtonList()
        BH.LoadButtonSetAsync(ButtonFilePath)
        AddToButtonFilesList(ButtonFilePath)
        ProgressBarOff()
    End Sub

    Private Sub SaveListasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveListToolStripMenuItem.Click
        SaveShowlist()

    End Sub

    Private Sub NewListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewToolStripMenuItem.Click
        BH.buttons.Clear()
        BH.SaveButtonSet()
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf CurrentDomain_UnhandledException
        splashScreen.Show()
        Try
            CurrentFolder = "C:\"
            GlobalInitialise()
            keyPressTimer.Interval = 500 ' Interval set to 500 milliseconds (half a second)
            AddHandler keyPressTimer.Tick, AddressOf KeyPressTimer_Tick
        Catch ex As Exception

        End Try
        ConstructMenuShortcuts()
        ConstructMenutooltips()
        Marks.Bar = Scrubber
        Scrubber.Size = New Size(Scrubber.Width + 1, Scrubber.Height + 1)
        Scrubber.Size = New Size(Scrubber.Width - 1, Scrubber.Height - 1)
        Marks.Create()

    End Sub
    Private Sub KeyPressTimer_Tick(sender As Object, e As EventArgs)
        ' Timer has ticked - no more keys have been pressed within the interval
        ' Use the accumulatedKeys string as needed here
        keyPressTimer.Stop()
        'MessageBox.Show("Accumulated keys: " & accumulatedKeys)
        If Len(accumulatedKeys) > 1 Then BH.FindRowContaining(accumulatedKeys)
        ' Reset the accumulated keys and stop the timer
        accumulatedKeys = String.Empty

    End Sub


    Private Sub ToggleMove()
        NavigateMoveState.ToggleState()

        BH.UpdateButtonAppearance()
    End Sub

    Private Sub BurstFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BurstFolderToolStripMenuItem.Click
        ' MSFiles.CancelURLS()
        Dim d As New IO.DirectoryInfo(CurrentFolder)
        BurstFolder(d)
        FBH.FillBox()
        'tvMain2.RemoveNode(CurrentFolder)
        '        tvMain2.RefreshTree(New IO.DirectoryInfo(CurrentFolder).Parent.FullName)
    End Sub

    Private Sub OnDirectorySelected(sender As Object, e As DirectoryInfoEventArgs) Handles tvmain2.DirectorySelected
        tmrUpdateFileList.Enabled = False
        'PreferencesSave()
        lbxGroups.Items.Clear()
        MSFiles.CancelURL()
        CurrentFolder = e.Directory.FullName
        ' FBH.DirectoryPath = CurrentFolder
        tmrUpdateFileList.Enabled = True
        tmrUpdateFileList.Interval = 200



    End Sub

    Private Sub OnFilenamesParsed() Handles FNG.WordsParsed
        'lbxGroups.Items.Clear()
        ''    Dim i As Integer = 0
        'For Each g In FG.
        '    lbxGroups.Items.Add(FNG.GroupNames(i) & " (" & g.Count & ")")
        '    '        Console.WriteLine(FNG.GroupNames(i))

        '    '        i += 1
        '    '        For Each m In g
        '    '            Console.WriteLine(m)
        '    '        Next
        '    '        Console.WriteLine()
        'Next

    End Sub
    ''' <summary>
    ''' Still used for Slowmo
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub tmrSlowMo_Tick(sender As Object, e As EventArgs) Handles tmrSlowMo.Tick
        Media.Player.Ctlcontrols.step(1)
        Media.Speed.PausedPosition = Media.Player.Ctlcontrols.currentPosition
        'Throw New Exception

    End Sub

    Private Sub ClearCurrentListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearCurrentListToolStripMenuItem.Click
        ClearShowList()
    End Sub
    'Private Sub PicFullScreen() Handles PictureBox1.DoubleClick
    '    If ShiftDown Then
    '        blnSecondScreen = True
    '    Else
    '        blnSecondScreen = False
    '    End If
    '    GoFullScreen(True)

    'End Sub
    Private Function LoadImage(fname As String) As Image
        Try
            Dim FileStream1 As New System.IO.FileStream(fname, IO.FileMode.Open, IO.FileAccess.Read)


            Try
                Dim MyImage As Image = Image.FromStream(FileStream1)
                FileStream1.Close()
                '  FileStream1.Dispose()
                Return MyImage
            Catch ex As System.ArgumentException
                FileStream1.Close()
                FileStream1.Dispose()
                Return Nothing
            End Try
        Catch ex As Exception

            Return Nothing
        End Try


    End Function

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


    Private Sub AlphabeticToolStripMenuItem_Click(sender As Object, e As EventArgs)
        BH.AssignAlphabetical(New DirectoryInfo(CurrentFolder))
        BH.SaveButtonSet()
    End Sub


    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs)
        CurrentFilterState.State = cbxFilter.SelectedIndex

    End Sub

    Private Sub SingleFilePerFolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SingleFilePerFolderToolStripMenuItem.Click
        blnChooseOne = True
        AddFilesToCollectionSingle(Showlist, True)
        FillShowbox(lbxShowList, CurrentFilterState.State, Showlist)

    End Sub

    'Private Sub BundleToolStripMenuItem_Click(sender As Object, e As EventArgs)
    '    BundleFiles(lbxFiles, CurrentFolder)
    'End Sub

    Public Sub BundleFiles(lbx1 As ListBox, Folder As String)
        Dim s As String
        If lbx1.SelectedItems.Count > 1 Then
        Else
            s = SelectSubList(FocusControl, False)
        End If
        Dim x As New BundleHandler(tvmain2, lbx1, Folder)
        Dim blnold = blnSuppressCreate
        blnSuppressCreate = False
        x.Bundle("")
        blnSuppressCreate = blnold
        tmrUpdateFileList.Enabled = True

    End Sub
    Private Sub Groupfiles(ByVal m As FilenamesGrouper)
        MSFiles.URLSZero()
        blnSuppressCreate = True
        For i As Integer = 0 To m.Groups.Count - 1
            If lbxGroups.SelectedIndices.Contains(i) Then
                MoveFiles(m.Groups.Item(i), CurrentFolder & "\" & m.GroupNames.Item(i))
            End If
        Next
        blnSuppressCreate = False
        RefreshFiles()
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

        If Not CtrlDown And Not Media.DontLoad Then AutoTrail()
    End Sub

    Private Sub AutoTrail()
        'Dim i As Byte = AT.SpeedIndex ' Assuming SpeedIndex calculation logic is implemented as needed
        UpdateTimerInterval()
        ChangePlaybackSpeed()

        ' Check if we are still in the initial moments phase
        If AT.InitialMomentsCount > 0 Then
            ' Choose moments from the very beginning
            SelectInitialMoments()
            AT.InitialMomentsCount -= 1 ' Decrement the counter
        Else
            ' Standard random jumping behavior
            JumpToVideoSegment()
        End If
    End Sub
    Private Sub SelectInitialMoments()
        ' Example: Jump to 5%, 10%, and 15% of the video duration in the first 3 ticks
        Dim startPoint As Double = (Media.Duration * 0.05 * AT.InitialMomentsCount)
        Media.MediaJumpToMarker(ToPoint:=startPoint)
    End Sub

    Private Sub JumpToVideoSegment()
        ' Determine whether to jump to a mark based on AdvanceChance
        If AT.ShouldAdvance() Then
            JumpToMark(True, True)
        Else
            JumpRandom(False)
        End If
    End Sub

    Private Sub ChangePlaybackSpeed()
        Dim speedSetting As Integer = AT.DetermineSpeedSetting()
        Select Case speedSetting
            Case 0, 1, 2
                SP.FrameRate = AT.Framerates(speedSetting)
                SpeedChange(New KeyEventArgs(speedkeys(speedSetting)))
            Case 3
                tmrSlowMo.Enabled = False
                Media.Player.settings.rate = 1
                Media.Player.Ctlcontrols.play()
                ' Handle other cases as needed...
        End Select
    End Sub
    Private Sub UpdateTimerInterval()

        ' Convert the duration from seconds to milliseconds for the Timer.Interval property
        Dim interval As Integer = CInt(AT.GetRandomDuration * 1000)

        ' Ensure there's a minimum interval to avoid too rapid execution
        interval = Math.Max(interval, 500) ' Setting a minimum interval of 500 milliseconds

        ' Update the timer's interval
        tmrAutoTrail.Interval = interval
    End Sub


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

    Private Sub BundleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BundleToolStripMenuItem.Click
        If FocusControl IsNot tvmain2 Then
            BundleFiles(FocusControl, CurrentFolder)
        End If
    End Sub

    Private Sub ThumbnailsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ThumbnailsToolStripMenuItem.Click
        ThumbnailsStart()
    End Sub

    Private Sub TrailerModeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TrailerModeToolStripMenuItem.Click
        'Media.AT.Play()
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
        For Each f In Application.OpenForms
            If f.Name = "FavouriteFiles" Then
                f.sortfiles(PlayOrder)
            End If
        Next
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

    Private Sub chbNextFile_CheckedChanged(sender As Object, e As EventArgs) Handles chbNextFile.CheckedChanged
        Random.NextSelect = chbNextFile.Checked
    End Sub



    Private Sub FilterMoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilterMoveToolStripMenuItem.Click
        '   FM.Recursive = False
        FM.FilterMoveFiles(CurrentFolder, False)
        RefreshFiles()

    End Sub

    Private Sub FilterMoveRecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs)
        FM.FilterMoveFiles(CurrentFolder, True)
        ReportAction(Format("Filtering {0}", CurrentFolder))
        RefreshFiles()

    End Sub


    Private Sub ReUniteFavesLinks()
        Dim s As List(Of String)
        s = ListfromSelectedInListbox(CType(FocusControl, ListBox))
        Op.OrphanList = s
        Op.DB = DB
        ProgressBarOn(s.Count)
        T = New Thread(New ThreadStart(Sub() Op.FindOrphans())) With {
            .IsBackground = True
        }
        T.SetApartmentState(ApartmentState.STA)
        T.Start()
        lbxFiles.SelectionMode = SelectionMode.One
        ProgressBarOff()
        UpdatePlayOrder(FBH)
    End Sub

    Public Sub ReportAction(Msg As String)
        Console.WriteLine(Msg)
    End Sub

    Private Sub Filter(index As DateMove.DMY)
        CancelDisplay(False)
        DM.FilterByDate(CurrentFolder, False, index)
        tvmain2.RefreshTree(CurrentFolder)
    End Sub
    Private Sub FilterAlphabetic(Folders As Boolean, Files As Boolean)
        CancelDisplay(False)
        DM.FilterByAlpha(CurrentFolder, Folders:=Folders, Files:=Files)
        tvmain2.RefreshTree(CurrentFolder)

    End Sub

    Private Sub FilterAlphabetic(Folders As Boolean, Files As Boolean, Letters As Boolean)
        CancelDisplay(False)
        DM.FilterByAlpha(CurrentFolder, Folders, Files, Letters)
        tvmain2.RefreshTree(CurrentFolder)

    End Sub

    Private Sub tbAbsolute_MouseUp(sender As Object, e As MouseEventArgs) Handles tbAbsolute.MouseUp
        Media.SPT.State = StartPointHandler.StartTypes.ParticularAbsolute
        Media.SPT.Absolute = tbAbsolute.Value
        Media.MediaJumpToMarker() 'At mouse up (Absolute)
        MSFiles.SetStartpoints(Media.SPT)
    End Sub
    Private Sub tbAbsolute_ValueChanged(sender As Object, e As EventArgs) Handles tbAbsolute.ValueChanged

        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(Media.SPT.Percentage) & "%"
        tbPercentage.Value = Media.SPT.Percentage
    End Sub
    Private Sub tbPercentage_ValueChanged(sender As Object, e As EventArgs) Handles tbPercentage.ValueChanged

    End Sub
    Private Sub tbPercentage_MouseUp(sender As Object, e As MouseEventArgs) Handles tbPercentage.MouseUp
        Media.SPT.State = StartPointHandler.StartTypes.ParticularPercentage
        Media.SPT.Percentage = tbPercentage.Value
        tbxAbsolute.Text = New TimeSpan(0, 0, tbAbsolute.Value).ToString("hh\:mm\:ss")
        tbxPercentage.Text = Str(Media.SPT.Percentage) & "%"
        tbPercentage.Value = Media.SPT.Percentage
        Media.MediaJumpToMarker() 'At Mouse Up (Percentage)
        MSFiles.SetStartpoints(Media.SPT)
    End Sub


    Private Sub ReclaimDeadLinksToolStripMenuItem_Click(sender As Object, e As EventArgs)
        ReUniteFavesLinks()

    End Sub


    Private Sub tvMain2_DragEnter(sender As Object, e As DragEventArgs) Handles tvmain2.DragEnter
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            If (e.KeyState And 8) = 8 Then
                e.Effect = DragDropEffects.Copy
            Else
                e.Effect = DragDropEffects.Move
            End If
        End If
    End Sub

    Private Async Sub RecursiveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RecursiveToolStripMenuItem.Click
        Dim x As New BundleHandler(tvmain2, lbxFiles, CurrentFolder)
        x.Maxfiles = Val(InputBox("Min no. of files to preserve folder? (Blank means all folders burst)",, ""))
        Dim d As New IO.DirectoryInfo(CurrentFolder)
        If d.Exists Then
            'MSFiles.CancelURLS()
            x.HarvestFolder(New IO.DirectoryInfo(CurrentFolder)) 'TODO: sometimes crashes because CurrentFolder doesn't exist
            ' tvMain2.RefreshTree(CurrentFolder)
            tmrUpdateFileList.Enabled = True
        End If

    End Sub

    Private Sub BySizeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BySizeToolStripMenuItem.Click
        DM.FilterBySize(CurrentFolder, False)
        RefreshFiles()
    End Sub


    Private Sub ShowlistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowlistToolStripMenuItem.Click
        ' If FocusControl Is lbxShowList Then

        Dim x As New FormShowList
        x.Show()
        x.Order = PlayOrder.State
        x.ListofFiles = LBH.ItemList
        ' End If
        CollapseShowlist(True)
    End Sub


    Private Sub RefreshFiles()
        tvmain2.RefreshTree(CurrentFolder)
        tmrUpdateFileList.Enabled = True
    End Sub


    Private Sub ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ByMonthToolStripMenuItem.Click
        Filter(DateMove.DMY.Month)
        RefreshFiles()

    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As Object, e As EventArgs) Handles ByDateToolStripMenuItem.Click
        Filter(DateMove.DMY.Day)
        RefreshFiles()
    End Sub

    Private Sub ToolStripMenuItem4_Click(sender As Object, e As EventArgs) Handles ByYearToolStripMenuItem.Click
        Filter(DateMove.DMY.Year)
        RefreshFiles()
    End Sub


    Private Sub tmrUpdateFileList_Tick(sender As Object, e As EventArgs) Handles tmrUpdateFileList.Tick
        'If AutoLoadButtons Then
        '    AddToButtonFilesList(CurrentFolder)
        '    UpdateFileInfo()
        'End If
        ProgressBarOn(1000)

        FBH.DirectoryPath = CurrentFolder
        ProgressBarOff()
        FBH.Random = Random
        FBH.SetFirst()
        tmrUpdateFileList.Enabled = False
    End Sub

    Public Sub AddToButtonFilesList(path As String)
        Dim folder As String = Autoload(IO.Path.GetFileNameWithoutExtension(path))
        If CBXButtonFiles.Items.Contains(folder) Then
        ElseIf folder <> "" Then
            CBXButtonFiles.Items.Add(folder)
        End If
        'CBXButtonFiles.SelectedItem = folder
    End Sub
    'Private Sub ChangedTree() Handles tvMain2.DirectorySelected
    '    ChangeFolder(tvMain2.SelectedFolder)
    'End Sub
    Public Function Autoload(Foldername As String, Optional Load As Boolean = False) As String
        'If directory name exists as buttonfile, then load button file (optionally)
        'Check for button file

        Dim f As New IO.DirectoryInfo(Foldername)
        Dim file As New IO.FileInfo(Buttonfolder & "\" & Foldername & ".msb")
        Static lastasked As String
        If file.Exists Then ';And file.FullName <> lastasked Then

            If Load Then
                SpawnButtonForm(file)
                AddToButtonFilesList(file.FullName)
            End If
            Foldername = file.Name
            '     End If
            lastasked = file.FullName
        Else
            SpawnButtonForm(Nothing, name:=Buttonfolder & "\" & Foldername)
            Foldername = ""
        End If
        Return Foldername
        'Load it (optionally)
    End Function
    Private Function SpawnSubButtonsForm(Search As String) As FormButton
        Dim s = ShowInputDialog(Search)
        BH.GetSubButtons(s)


    End Function

    Private Function ShowInputDialog(search As String) As String
        Dim prompt As String = "Correct?"
        Dim title As String = "Input Box"
        Dim defaultResponse As String = search

        Dim userInput As String = InputBox(prompt, title, defaultResponse)

        If userInput <> "" Then
            ' Perform actions with userInput
            Return userInput
        Else
            ' Do nothing if the input box was canceled or empty
            Return ""
        End If

    End Function
    Private Function SpawnButtonForm(file As FileInfo, Optional name As String = "", Optional Nofile As Boolean = False) As FormButton
        Dim fname As String
        'If already a file, load it. Otherwise create a new one. 
        If file Is Nothing Then
            If Not Nofile Then

                Dim tBH As New ButtonHandler
                tBH.SaveButtonSet(name)
                fname = name
            End If
        Else
            fname = file.FullName
        End If
        Dim x As New FormButton With {
                        .ButtonFilePath = fname
                    }
        x.BH.Alpha = iCurrentAlpha
        AddHandler x.KeyDown, AddressOf HandleKeys

        x.Show()
        x.Top = Me.Height * 0.8
        Return x
        x.BH.SaveButtonSet(fname)
    End Function

    Private Sub ToolStripMenuItem8_Click(sender As Object, e As EventArgs) Handles CalendarToolStripMenuItem.Click
        DM.FilterByCalendar(New DirectoryInfo(CurrentFolder))
        RefreshFiles()
    End Sub

    Private Sub ToolStripMenuItem9_Click(sender As Object, e As EventArgs) Handles ByExtToolStripMenuItem.Click
        DM.FilterByType(CurrentFolder)
        RefreshFiles()
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
        RefreshFiles()
    End Sub


    Private Sub DashboardToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DashboardToolStripMenuItem.Click
        Dashboard.Show()
    End Sub

    Private Sub DuplicatesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DuplicatesToolStripMenuItem1.Click
        StartDuplicatesForm()
        Exit Sub
        T = New Thread(New ThreadStart(Sub() StartDuplicatesForm()))

        T.IsBackground = True
        T.SetApartmentState(ApartmentState.STA)
        T.Start()
        While T.IsAlive
            Thread.Sleep(0)
        End While



    End Sub

    Private Sub StartDuplicatesForm()
        Dim x As New FormDuplicates

        x.DB = DB
        x.FindDuplicates()
        AddHandler x.FileHighlighted, AddressOf OnThumbnailHover
        x.Show()


    End Sub

    Private Sub ResetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetToolStripMenuItem.Click
        PreferencesReset(True)
    End Sub

    Private Sub SaveToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem1.Click
        SavePreferences(Parameters)

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
                'FaveMinder.NewPath(CurrentFavesPath)
                SavePreferences(Parameters)

            End If

        Catch ex As Exception

        End Try

    End Sub

    Private Sub tmrMovieSlideShow_Tick(sender As Object, e As EventArgs) Handles tmrMovieSlideShow.Tick
        tmrSlideShow_Tick(sender, e)
        ' tmrMovieSlideShow.Interval = 500
        'Media.LargeJump(Nothing, False, True)
        'Dim x = Int(Rnd() * 100)
        'Dim FSPercent = 70
        'If x < FSPercent Then
        '    SP.Fullspeed = True
        '    Media.Position = Media.Player.Ctlcontrols.currentPosition
        '    Media.Player.settings.rate = 1
        '    Media.Player.Ctlcontrols.play()
        '    tmrSlowMo.Enabled = False
        '    SP.Fullspeed = True
        'Else
        '    SP.Speed = Int((x - FSPercent) / ((100 - FSPercent) / 2.8))
        '    SpeedChange(New KeyEventArgs(speedkeys(SP.Speed)))
        'End If

        ''  tmrMovieSlideShow.Interval = tmrMovieSlideShow.Interval * 29 / Media.Speed.FrameRates(Media.Speed.Speed)
        ' AdvanceFile(True, False)
    End Sub

    Private Sub cbxStartPoint_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbxStartPoint.SelectedIndexChanged
        Media.SPT.State = cbxStartPoint.SelectedIndex
        MSFiles.SetStartStates(Media.SPT)
        '  OnStartChanged(sender, e)
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
        If DB.Entries.Count = 0 Then
            MsgBox("Load database first")
            Exit Sub
        End If

        ProgressBarOn(100)
        Dim op2 As New OrphanFinder
        Dim lbx As New ListBox
        Dim files As New List(Of String)

        lbx = FocusControl
        For Each m In lbx.SelectedItems
            files.Add(m)
            ProgressIncrement(1)
        Next
        op2.DB = DB
        op2.OrphanList = files
        T = New Thread(New ThreadStart(Sub() op2.FindOrphans())) With {
            .IsBackground = True
        }
        T.SetApartmentState(ApartmentState.STA)

        T.Start()

        '        If op2.FindOrphans() Then FBH.FillBox()
        ProgressBarOff()
    End Sub

    Private Sub lbxGroups_DoubleClick(sender As Object, e As EventArgs) Handles lbxGroups.DoubleClick
        FNG.AdvanceOption()
        FNG.Filenames = FBH.ItemList
        FG.FilePaths = FBH.ItemList
        FG.GroupFiles(4, CurrentFolder)
    End Sub


    Private Sub SelectNonFavouritsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectNonFavouritsToolStripMenuItem.Click
        'Op.ImportLinks(Showlist)
        'Select files with links
        Dim hilist As New List(Of String)
        Dim FLBH As New ListBoxHandler()
        If FocusControl Is lbxFiles Then
            FLBH = FBH
        Else
            FLBH = LBH
        End If
        For Each f In FLBH.ItemList
            If AllFaveMinder.GetLinksOf(f).Count = 0 Then
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



    Private Sub ToolStripTextBox1_Click(sender As Object, e As EventArgs)
        BH.Autobuttons = Not BH.Autobuttons

    End Sub

    Private Sub Response_Tick(sender As Object, e As EventArgs) Handles Response.Tick
        Response.Enabled = False
    End Sub
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub DisplayedToolStripMenuItem_Click(sender As Object, e As EventArgs)
        AllFaveMinder.RedirectShortCutList(lbxFiles.SelectedItem, AllfromListbox(lbxShowList))
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
            'DirectoriesList = GetDirectoriesList(Rootpath, True)
        End If
    End Sub

    Private Sub RemoveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem6.Click
        If FocusControl Is LBH.ListBox Then
            RemoveBrackets(LBH.ItemList)
        Else
            RemoveBrackets(FBH.ItemList)
            FBH.Refresh()
        End If
    End Sub


    Private Sub lbxFiles_KeyDown(sender As Object, e As KeyEventArgs) Handles lbxFiles.KeyDown
        If e.KeyCode = KeyTraverseTree Or e.KeyCode = KeyTraverseTreeBack Or e.KeyCode = Keys.Left Or e.KeyCode = Keys.Right Then
            tvmain2.tvFiles_KeyDown(sender, e)
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



    Private Sub Scrubber_Paint(sender As Object, e As PaintEventArgs) Handles Scrubber.Paint
        'DrawScrubberMarks()

        'MsgBox("Uh-oh")
    End Sub
    Private Sub ToggleMovieScan()
        tmrJumpRandom.Interval = tbScanRate.Value
        tmrJumpRandom.Enabled = Not tmrJumpRandom.Enabled
        chbScan.Checked = tmrJumpRandom.Enabled
        UpdateLabel(InstructionLabel2, ScanInstructions(1))
    End Sub
    Private Sub chbSeparate_CheckedChanged(sender As Object, e As EventArgs) Handles chbSeparate.CheckedChanged
        separate = chbSeparate.Checked

        MSFiles.DockMedias(separate)

    End Sub

    Private Sub tvMain2_Load(sender As Object, e As EventArgs)

    End Sub

    Private Sub PromoteFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PromoteFileToolStripMenuItem.Click
        Dim folder As New IO.DirectoryInfo(FBH.DirectoryPath)
        blnSuppressCreate = True
        MoveFiles(FBH.SelectedItemsList, folder.Parent.FullName)
    End Sub

    Private Sub FilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilesToolStripMenuItem.Click
        FilterAlphabetic(False, True)

    End Sub

    Private Sub FoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FoldersToolStripMenuItem.Click
        FilterAlphabetic(True, False)
    End Sub

    Private Sub NewButtonFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewButtonFileToolStripMenuItem.Click
        If Me Is ActiveForm Then
            BH.buttons.Clear()
            BH.SaveButtonSet()
        Else
            SpawnButtonForm(Nothing)
        End If

    End Sub

    Private Sub SelectedDeepToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectedDeepToolStripMenuItem.Click
        MsgBox("Does Nothing")
    End Sub

    Private Sub ListDeadFilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListDeadFilesToolStripMenuItem.Click
        ProgressBarOn(1000)
        'FormDeadFiles.Show()
        SelectDeadLinks(FocusControl)
        Dim op2 As New OrphanFinder
        Dim lbx As New ListBox
        Dim files As New List(Of String)
        lbx = FocusControl

        For Each m In lbx.SelectedItems
            files.Add(m)
            ProgressIncrement(1)
        Next
        op2.OrphanList = files
        Dim deadfiles As New List(Of String)
        deadfiles = op2.ListOfDeadFiles
        deadfiles.Sort()

        For Each f In deadfiles
            '   FormDeadFiles.ListBox1.Items.Add(f)
            ProgressIncrement(1)

        Next
        ProgressBarOff()
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
        Application.DoEvents()
        tmrProgressBar.Interval = 100
        ProgressIncrement(1)
    End Sub

    Private Sub SelectAndBundleToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAndBundleToolStripMenuItem.Click
        SelectDeadLinks(FocusControl)
        Dim m As New BundleHandler(tvmain2, FocusControl, CurrentFolder)
        m.Bundle("Dead")

    End Sub

    Private Sub RefreshAllFilesWithLinksToolStripMenuItem_Click(sender As Object, e As EventArgs)
        SelectDeadLinks(FocusControl)
        InvertSelectionToolStripMenuItem_Click(sender, e)
        RefreshLinks(ListfromSelectedInListbox(FocusControl))
    End Sub

    Private Sub CBXButtonFiles_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CBXButtonFiles.SelectedIndexChanged
        Dim folder As String = CBXButtonFiles.SelectedItem
        Dim found As Boolean
        For Each fb In Application.OpenForms().OfType(Of FormButton)
            If fb.ButtonFilePath.Contains(CBXButtonFiles.SelectedItem) Then
                fb.BringToFront()
                found = True
            End If
        Next
        If Not found Then
            Autoload(Path.GetFileNameWithoutExtension(CBXButtonFiles.SelectedItem), False)
        End If
    End Sub

    Private Sub CHBAutoAdvance_CheckedChanged(sender As Object, e As EventArgs) Handles CHBAutoAdvance.CheckedChanged
        AutoTraverse = CHBAutoAdvance.Checked
        chbOnDir.Checked = Not AutoTraverse
    End Sub

    Private Sub chbShowAttr_CheckedChanged(sender As Object, e As EventArgs) Handles chbShowAttr.CheckedChanged
        lblAttributes.Visible = chbShowAttr.Checked

        'MsgBox("Attribute Changed")
    End Sub

    Private Sub chbAutoTrail_CheckedChanged(sender As Object, e As EventArgs) Handles chbAutoTrail.CheckedChanged

        tmrAutoTrail.Enabled = chbAutoTrail.Checked
    End Sub

    Private Sub MovieScanToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MovieScanToolStripMenuItem.Click
        ToggleMovieScan()
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles tmrJumpRandom.Tick
        tmrJumpRandom.Interval = tbScanRate.Value
        JumpRandom(False)
    End Sub

    Private Sub ResetButtonsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ResetButtonsToolStripMenuItem.Click
        CBXButtonFiles.SelectedItem = "Watch.msb" 'TODO Make general

    End Sub

    Private Sub PreferencesToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PreferencesToolStripMenuItem1.Click
        Mysettings.loadform()
    End Sub

    Private Sub LBH_ListboxChanged(sender As Object, e As EventArgs) Handles LBH.ListIndexChanged
        HighlightCurrent(lbxShowList.SelectedItem)
    End Sub

    Private Sub lbxFiles_Leave(sender As Object, e As EventArgs) Handles lbxFiles.Leave
        If ShowListVisible Then
            lbxShowList.Focus()
        Else
            tvmain2.Focus()

        End If
    End Sub

    Private Sub tbScanRate_ValueChanged(sender As Object, e As EventArgs)
        tmrJumpRandom.Interval = tbScanRate.Value
    End Sub
    Private Sub MainForm_Mousewheel(sender As Object, e As MouseEventArgs) Handles Me.MouseWheel
        If Media.PicHandler.WheelAdvance And Media.MediaType = Filetype.Movie Then
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
        If tmrMovieSlideShow.Enabled Then
            chbSlideShow.Text = "Movie Slide Show (" & Str(tbMovieSlideShowSpeed.Value / 1000) & "s.)"
        Else
            chbSlideShow.Text = "Movie Slide Show"
        End If
    End Sub

    Private Sub tbMovieSlideShowSpeed_Scroll(sender As Object, e As EventArgs) Handles tbMovieSlideShowSpeed.ValueChanged, tbMovieSlideShowSpeed.Scroll
        tmrMovieSlideShow.Interval = tbMovieSlideShowSpeed.Value
        chbSlideShow.Text = "Movie Slide Show (" & Str(tbMovieSlideShowSpeed.Value / 1000) & "s.)"
    End Sub

    Private Sub tbAutoTrail_Scroll(sender As Object, e As EventArgs)

    End Sub

    Private Sub PromoteIdenticalSubFoldersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PromoteIdenticalSubFoldersToolStripMenuItem.Click
        'Dim m As New IO.DirectoryInfo(CurrentFolder)
        MoveSameFolders(CurrentFolder)
    End Sub

    Private Sub lbxFiles_GotFocus(sender As Object, e As EventArgs) Handles lbxFiles.GotFocus
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
    End Sub

    Private Sub RenewLinksOfSelectedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RenewLinksOfSelectedToolStripMenuItem.Click
        Dim Oh As New OrphanFinder
        T = New Thread(New ThreadStart(Sub() Oh.ListReuniter(AllfromListbox(lbxShowList)))) With {
            .IsBackground = True
        }

        T.SetApartmentState(ApartmentState.STA)

        T.Start()

    End Sub

    Private Sub PopulateContextMenu(Optional Tag As String = "")
        ContextMenuStrip1.Items.Clear()
        For Each m In MenuStrip2.Items
            Dim clone As New ToolStripMenuItem
            clone = CloneMenu(m)
            If clone IsNot Nothing Then
                If Tag = "" Then
                    ContextMenuStrip1.Items.Add(clone)
                Else
                    If clone.Tag IsNot Nothing Then
                        If clone.Tag = Tag Then
                            ContextMenuStrip1.Items.Add(clone)
                        End If
                    Else
                    End If
                End If
            End If
        Next
    End Sub

    'Private Sub DatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatabaseToolStripMenuItem.Click
    '    DatabaseForm.Show()

    'End Sub

    Private Sub tmrHighlightCurrent_Tick(sender As Object, e As EventArgs) Handles tmrHighlightCurrent.Tick
        If HighlightPath = "" Then Exit Sub 'Empty
        Dim finfo As New FileInfo(HighlightPath)
        'Change the tree if it needs changing

        Dim s As String = Path.GetDirectoryName(HighlightPath)

        If tvmain2.SelectedFolder <> s Then
            tvmain2.SelectedFolder = s
            tmrUpdateFileList.Enabled = False
            FBH.DirectoryPath = s
        End If
        'Select file in filelist
        FBH.SetNamed(HighlightPath)
        tmrHighlightCurrent.Enabled = False
    End Sub

    Friend Sub PrependFolderName(list As List(Of String))
        For Each s In list
            Dim f As New IO.FileInfo(s)
            Dim name = f.Name
            Dim dir = f.Directory.Name
            ' AllFaveMinder.CheckFile(f)
            Try
                f.MoveTo(Path.GetFullPath(s).Replace(f.Name, dir & "_" & name))

            Catch ex As Exception

            End Try
        Next
    End Sub
    Private Sub PrependAllFilenamesWithFolderNameToolStripMenuItem_Click(sender As Object, e As EventArgs)
        MSFiles.CancelURL()
        Dim list As New List(Of String)
        list = FBH.ItemList
        PrependFolderName(list)
        FBH.Refresh()
    End Sub

    Private Sub RecursivelyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        T = New Thread(New ThreadStart(Sub() PrependRecursive(False)))
        T.IsBackground = True
        T.SetApartmentState(ApartmentState.STA)
        T.Start()
        T.Join()
        FBH.Refresh()
    End Sub

    Private Sub PrependRecursive(Shortonly As Boolean)
        MSFiles.CancelURL()
        Dim list As New List(Of String)
        Dim dir As New IO.DirectoryInfo(CurrentFolder)
        For Each d In dir.GetDirectories("*", IO.SearchOption.AllDirectories)
            For Each f In d.GetFiles
                If Shortonly And f.Name.Length > 8 Then
                Else
                    list.Add(f.FullName)
                End If
            Next
        Next
        For Each f In dir.GetFiles
            list.Add(f.FullName)
        Next
        PrependFolderName(list)

    End Sub

    Private Sub SelectedToolStripMenuItem1_Click(sender As Object, e As EventArgs)
        MSFiles.CancelURL()
        Dim list As New List(Of String)
        list = FBH.SelectedItemsList
        PrependFolderName(list)
        FBH.Refresh()
    End Sub

    Private Sub ShortFilenamesRecursivelyToolStripMenuItem_Click(sender As Object, e As EventArgs)
        T = New Thread(New ThreadStart(Sub() PrependRecursive(True)))
        T.IsBackground = True
        T.SetApartmentState(ApartmentState.STA)
        T.Start()
        While T.IsAlive
            Thread.Sleep(0)
        End While
        FBH.Refresh()
    End Sub

    Private Sub FlattenAllSubfolderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FlattenAllSubfolderToolStripMenuItem.Click
        FlattenAllSubFolders(New IO.DirectoryInfo(CurrentFolder))
    End Sub



    Private Sub LoadDatabase(filename As String, Append As Boolean)
        If Not Append Then
            DB.Entries.Clear()

        End If
        ProgressBarOn(1000)
        Dim list As New List(Of String)
        list = ReadListfromFile(filename, Encrypted)
        DB.Name = filename
        Dim i As Integer = 0
        For Each entry In list
            Try
                Dim x() As String
                x = entry.Split(vbTab)
                Dim ent As New DatabaseEntry
                ent.Filename = x(0)
                ent.Path = x(1)
                ent.Size = x(2)
                ent.Dt = x(3)
                DB.AddEntry(ent)
                ProgressIncrement(1)


            Catch ex As Exception

            End Try
            Application.DoEvents()
        Next
        ProgressBarOff()
        'DB.Sort()
        ShowCurrentlyLoadedToolStripMenuItem.Enabled = True
    End Sub

    Public Sub OnFBHIncrease(sender As Object, e As EventArgs) Handles FBH.ListBoxIncreased
        ProgressIncrement(1)

    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub
    Public Sub ProgressBarOn(max As Long, Optional Message As String = "")
        With TSPB
            .Value = 0

            .Maximum = max 'Math.Max(lngListSizeBytes, 100)
            .Visible = True

        End With
        If Message <> "" Then lblProgress.Text = Message
        tmrProgressBar.Enabled = True
    End Sub
    Public Sub ProgressBarOff()
        tmrProgressBar.Enabled = False
        With TSPB
            .Visible = False
        End With
        TSPB.Value = 0
        lblProgress.Text = ""
    End Sub
    Public Sub ProgressIncrement(st As Integer)
        With TSPB
            '   .Maximum = max
            .Value = (.Value + st) Mod .Maximum
        End With
        'MainForm.Update()
    End Sub

    Private Sub CreateNewToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateNewToolStripMenuItem.Click
        CreateNewDatabase()

    End Sub

    Private Sub CreateNewDatabase()
        'Open file browser
        MsgBox("Browse to folder to be included, with its subfolders in the database" & vbCrLf & "Ensure filter is selected appropriately")
        Dim filename = SaveDatabaseFileName("")
        'Create Database from current folder
        If filename <> "" Then
            ProgressBarOn(1000)
            CreateDatabase(CurrentFolder, filename)
            ProgressBarOff()
            MsgBox(filename & " has been created.")
        End If
        'Throw New NotImplementedException()
    End Sub

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadToolStripMenuItem.Click
        DatabaseLoader(False)

    End Sub


    Private Sub DatabaseLoader(Append As Boolean)
        Dim filename = LoadDatabaseFileName("")

        'Import database

        If filename <> "" Then
            LoadDatabase(filename, Append)
            ShowDatabase(DB)
        End If
    End Sub

    Private Sub PicFullScreen(sender As Object, e As EventArgs) Handles PictureBox1.DoubleClick

    End Sub

    Private Sub ShowCurrentlyLoadedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ShowCurrentlyLoadedToolStripMenuItem.Click
        ShowDatabase(DB)
    End Sub

    Private Sub ShowDatabase(dB As Database)
        ProgressBarOn(100)
        ShowDBThread(dB)
        'Dim t = New Thread(New ThreadStart(Sub() ShowDBThread(dB)))
        't.IsBackground = True
        't.SetApartmentState(ApartmentState.STA)
        't.Start()
        ProgressBarOff()

        'Throw New NotImplementedException()
    End Sub

    Private Shared Sub ShowDBThread(dB As Database)
        Dim form As New FormDatabaseShower With {.MDB = dB, .Text = .MDB.Name & " " & .MDB.ItemCount & " files loaded"}
        form.LoadEntries(dB)
        form.Show()
    End Sub

    Private Sub AppendDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AppendDatabaseToolStripMenuItem.Click
        DatabaseLoader(True)
    End Sub

    Private Sub RefreshLinksFromDatabaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshLinksFromDatabaseToolStripMenuItem.Click
        If DB.Entries.Count = 0 Then Exit Sub
        RefreshAllDatabaseLinks()
    End Sub

    Private Sub RefreshAllDatabaseLinks()
        For Each m In DB.Entries
            Dim x As List(Of String) = AllFaveMinder.GetLinksOf(m.FullPath)
            If x.Count <> 0 Then Op.ReuniteWithFile(x, m.FullPath)
        Next
    End Sub
    Private Sub RefreshLinksOnPathNameFromDatabase(ListofLinks As List(Of String))
        DB.Sort()
        Dim entry As New List(Of DatabaseEntry)
        For Each link In ListofLinks
            Dim Filename As String = Path.GetFileNameWithoutExtension(Split(link, "%")(0))
            entry = DB.FindPartEntry(Filename)
            If entry.Count = 1 Then
                Dim x As New ShortcutHandler()
                x.ReAssign(entry(0).FullPath, link)
            End If
        Next

    End Sub



    Private Sub LinksByPartOfPathToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LinksByPartOfPathToolStripMenuItem.Click
        Dim x As New GroupByLinkFolders()
        x.Folder = CurrentFolder
        x.Files = FBH.SelectedItemsList
        x.MoveFiles(2)
    End Sub



    Private Sub CreateFromCurrentShowlistToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CreateFromCurrentShowlistToolStripMenuItem.Click
        If ShowListVisible Then
            DB.Entries.Clear()
            Dim list As New List(Of String)
            list = LBH.ItemList
            Dim x As New FilterHandler With {.State = CurrentFilterState.State}
            x.FileList = list
            list = x.FileList
            For Each file In list
                Try
                    Dim f As New FileInfo(file)
                    Dim entry As New DatabaseEntry
                    entry.Filename = f.Name
                    entry.Path = f.FullName.Replace(f.Name, "")
                    entry.Size = f.Length
                    DB.Entries.Add(entry)
                    Application.DoEvents()

                Catch ex As Exception

                End Try
            Next
            ShowCurrentlyLoadedToolStripMenuItem.Enabled = True
            'Filter them
            'Then create database list of those files
        End If

    End Sub

    Private Sub SaveToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem2.Click
        SaveDatabaseFileName("")
    End Sub




    Private Sub ButtonFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ButtonFormToolStripMenuItem.Click
        SpawnButtonForm(New FileInfo(ButtonFilePath)).BH.buttons.CurrentLetter = iCurrentAlpha
    End Sub

    Public Sub DatabaseExperimentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatabaseExperimentToolStripMenuItem.Click
        Form4.Show()
        AddHandler Form4.KeyDown, AddressOf HandleKeys
    End Sub

    Private Sub RefreshSelectedFromLinkNameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshSelectedFromLinkNameToolStripMenuItem.Click
        Dim x As List(Of String) = FBH.SelectedItemsList
        RefreshLinksOnPathNameFromDatabase(x)
    End Sub


    Private Sub SelectAllWrongLinksToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllWrongLinksToolStripMenuItem.Click
        Dim list As New List(Of String)
        FBH.ListBox.SelectionMode = SelectionMode.MultiExtended
        'For each link in folder
        'Select if link doesn't point to named file
        For Each f In FBH.ItemList
            Dim file As New IO.FileInfo(f)
            If file.Extension = LinkExt Then
                If LinkTargetContainsNameFile(file.FullName) Then
                Else
                    list.Add(f)
                End If
            End If
        Next
        FBH.SelectItems = list
    End Sub

    Private Sub ListhandlerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListhandlerToolStripMenuItem.Click
        Form3.Show()
        'Form3.LBH.FillBox()

    End Sub

    Private Sub ScreenOfMoviesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ScreenOfMoviesToolStripMenuItem.Click
        Dim x As New FormMovieScreen
        x.List = FBH.ItemList
        x.Show()

    End Sub

    Private Sub NewMediaSwapperToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewMediaSwapperToolStripMenuItem.Click
        Dim x As New FormTestMSwapper
        x.Show()
    End Sub

    Private Sub chbScan_CheckStateChanged(sender As Object, e As EventArgs) Handles chbScan.CheckStateChanged

        'ToggleMovieScan()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles cbxSingleLinks.CheckedChanged
        If FocusControl Is lbxShowList Then
            LBH.SingleLinks = cbxSingleLinks.Checked
            LBH.FillBox()
        Else
            FBH.SingleLinks = cbxSingleLinks.Checked

            FBH.FillBox()
        End If
    End Sub

    Private Sub ListOfFilesFormToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListOfFilesFormToolStripMenuItem.Click
        Dim f As New FavouriteFiles
        f.Show()
    End Sub

    Private Sub tmrPumpFiles_Tick(sender As Object, e As EventArgs) Handles tmrPumpFiles.Tick

    End Sub

    Private Sub SilentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SilentToolStripMenuItem.Click, NormalSoundToolStripMenuItem.Click, SlowSoundToolStripMenuItem.Click
        SilentToolStripMenuItem.Checked = False
        NormalSoundToolStripMenuItem.Checked = False
        SlowSoundToolStripMenuItem.Checked = False
        sender.checked = True
        If sender Is SilentToolStripMenuItem Then
            SlowMoSoundOpt = SlowMoSoundOptions.Silent
        ElseIf sender Is NormalSoundToolStripMenuItem Then
            SlowMoSoundOpt = SlowMoSoundOptions.Normal
        ElseIf sender Is SlowSoundToolStripMenuItem Then
            SlowMoSoundOpt = SlowMoSoundOptions.Slow
        End If
        MSFiles.SlowSoundOption = SlowMoSoundOpt
    End Sub

    Private Sub TaggingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TaggingToolStripMenuItem.Click
        Dim x As New FormTagging
        x.Show()

    End Sub

    Private Sub lbxGroups_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbxGroups.SelectedIndexChanged

    End Sub

    Private Sub Me_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        ' DrawScrubberMarks()
    End Sub

    Private Sub ctrPicAndButtons_Resize(sender As Object, e As EventArgs) Handles ctrPicAndButtons.Resize
        ' If Media IsNot Nothing Then DrawScrubberMarks()
    End Sub

    Private Sub GroupButtonsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GroupButtonsToolStripMenuItem.Click
        BH.buttons.StringButtons(InputBox("String"))
        BH.RefreshButtons()
    End Sub

    Public Class FolderCleaner

        Private mFolderPath As String
        Private mMaxFileCount As Integer
        Public Event CleanupFinished(ByVal folderPath As String)

        Public Property FolderPath As String
            Get
                Return mFolderPath
            End Get
            Set(ByVal value As String)
                mFolderPath = value
            End Set
        End Property

        Public Property MaxFileCount As Integer
            Get
                Return mMaxFileCount
            End Get
            Set(ByVal value As Integer)
                mMaxFileCount = value
            End Set
        End Property

        Public Sub CleanupFolder()
            Dim files As String() = Directory.GetFiles(mFolderPath, "*", IO.SearchOption.AllDirectories)

            Dim subFolders = Directory.GetDirectories(mFolderPath)

            For Each folder As String In subFolders
                Dim folderFiles As String() = Directory.GetFiles(folder)
                If folderFiles.Length <= mMaxFileCount Or mMaxFileCount = 0 Then
                    For Each file As String In folderFiles
                        Dim destinationFile As String = Path.Combine(mFolderPath, Path.GetFileName(file))
                        FileSystem.MoveFile(file, destinationFile, UIOption.OnlyErrorDialogs, UICancelOption.ThrowException)
                    Next
                    FileSystem.DeleteDirectory(folder, UIOption.OnlyErrorDialogs, RecycleOption.SendToRecycleBin)
                End If
            Next

            RaiseEvent CleanupFinished(mFolderPath)
        End Sub
    End Class

    Private Sub SubButtonToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SubButtonToolStripMenuItem.Click
        SpawnSubButtonsForm(InputBox("Search for what?"))
    End Sub

    Private Sub ExtractSubFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExtractSubFileToolStripMenuItem.Click
        VideoTrim.StartTime = Media.Position
        VideoTrim.Finish = StartTime + 1000
        VideoTrim.InputFile1 = MSFiles.CurrentItem
        VideoTrim.Main()
        '        Marks.DrawConnectedCircles(Media.Position, Media.Position + 2500)

    End Sub

















#End Region
#Region "Event Responders"
    Public Sub OnSpeedParameterChanged(sender As Object, e As EventArgs) Handles SP.ParameterChanged
        MSFiles.SetSpeedHandlers(SP)
        Marks.Fractions = SP.FractionalJump
        Marks.SmallJumps = SP.AbsoluteJump
    End Sub
    Public Sub OnSplashScreenUpdated(sender As Object, e As EventArgs)
        splashScreen.lblLoading.Text = e.ToString
    End Sub
    Public Sub OnLetterChanged(sender As Object, e As EventArgs) Handles BH.LetterChanged
        BH.LetterLabel.Text = Chr(AsciifromLetterNumber(BH.buttons.CurrentLetter))
        iCurrentAlpha = BH.Alpha
    End Sub
    Public Sub OnButtonFileChanged(sender As Object, e As EventArgs) Handles BH.ButtonFileChanged
        tbButton.Text = "BUTTONFILE: " & BH.ButtonfilePath

    End Sub
    Public Sub OnFilesMoving(sender As Object, e As EventArgs)
        MSFiles.CancelURL()
    End Sub

    Public Sub OnThumbnailHover(sender As Object, e As EventArgs)
        If TypeOf (sender) Is PictureBox Then
            Dim de As DatabaseEntry = DirectCast(sender.tag, DatabaseEntry)
            HighlightCurrent(de.FullPath)
        End If
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
        Debug.Print("End reached")
        If AutoTraverse Then

            tvmain2.Traverse(False)
        End If
    End Sub
    Public Sub OnFolderRename(sender As Object, e As NodeLabelEditEventArgs) Handles tvmain2.LabelEdited
        If Not e.CancelEdit Then
            CurrentFolder = CurrentFolder.Replace(e.Node.Text, e.Label)

        End If
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
        If Not Media.DontLoad Then Media.MediaJumpToMarker() 'On Start Change

    End Sub
    Public Sub OnStateChanged(sender As Object, e As EventArgs) Handles NavigateMoveState.StateChanged, CurrentFilterState.StateChanged, PlayOrder.StateChanged
        'If StartingUpFlag Then Exit Sub
        Dim s() As String = Split(sender.ToString, ".")


        Select Case s(s.Length - 1)
            Case "StateHandler"
                ChangeLabel(lblNavigateState, NavigateMoveState.Instructions)
                tbState.Text = UCase(NavigateMoveState.Description)
                SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

            Case "FilterHandler"
                If FocusControl Is lbxShowList Then
                    LBH.Filter = CurrentFilterState
                Else
                    FBH.Filter = CurrentFilterState
                End If
                cbxFilter.BackColor = CurrentFilterState.Colour
                cbxFilter.SelectedIndex = CurrentFilterState.State
                If CurrentFilterState.State = FilterHandler.FilterState.LinkOnly Then
                    cbxSingleLinks.Visible = True
                Else
                    cbxSingleLinks.Visible = False
                    cbxSingleLinks.Checked = False
                End If
                tbFilter.Text = "FILTER:" & UCase(CurrentFilterState.Description)
                If Not Media.DontLoad Then
                    SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

                End If
            Case "SortHandler"
                If Not Media.DontLoad Then

                    If FocusControl Is lbxShowList Then
                        ' LBH.ListBox = lbxShowList
                        LBH.SortOrder = PlayOrder
                    Else
                        FBH.SortOrder = PlayOrder
                    End If
                End If
                FBH.SortOrder = PlayOrder
                cbxOrder.SelectedIndex = PlayOrder.State
                cbxOrder.BackColor = PlayOrder.Colour
                tbRandom.Text = "ORDER:" & UCase(PlayOrder.Description)
        End Select
        If Not Media.DontLoad AndAlso FocusControl Is lbxShowList Then UpdatePlayOrder(LBH)
        If Not Media.DontLoad AndAlso FocusControl Is lbxFiles Then UpdatePlayOrder(FBH)


        If sender IsNot NavigateMoveState Then
            'If Not Media.DontLoad Then
            '    If FocusControl Is lbxShowList Then

            '        UpdatePlayOrder(LBH)
            '    Else
            '        UpdatePlayOrder(FBH)
            '    End If
            'End If
        Else
            If NavigateMoveState.State = StateHandler.StateOptions.ExchangeLink Then
                CurrentFilterState.State = FilterHandler.FilterState.LinkOnly
            End If
        End If
        If Not Media.DontLoad Then SavePreferences(Parameters)

    End Sub

    Private Sub SetControlColours(MainColor As Color, FilterColor As Color)
        If Media.DontLoad Then Exit Sub
        Me.SuspendLayout()
        lbxFiles.BackColor = FilterColor
        lbxShowList.BackColor = FilterColor
        If tvmain2.Focused Then
            tvmain2.BackColor = MainColor
        Else
            tvmain2.BackColor = FilterColor
            FocusControl.BackColor = MainColor
        End If
        tvmain2.HighlightSelectedNodes()
        Me.ResumeLayout()
    End Sub



    Public Sub OnSpeedChange(sender As Object, e As EventArgs) Handles SP.SpeedChanged
        If SP.Slideshow Then
            tmrSlideShow.Interval = SP.Interval
            tbSpeed.Text = "Slide Interval=" & SP.Interval
        Else
            If SP.Speed = -1 Then
                tmrSlowMo.Enabled = False
                'MSFiles.Soundhandler.Muted = True
                If Media.MediaType = Filetype.Movie Then
                    MSFiles.Soundhandler.Slow = False

                End If

                If SP.Fullspeed Then
                    tbSpeed.Text = "Speed:FULL"
                Else
                    tbSpeed.Text = "Speed:Paused"
                End If
            Else
                MSFiles.Soundhandler.Slow = True
                'MSFiles.Soundhandler.Muted = False

                tmrSlowMo.Interval = 1000 / SP.FrameRate
                tmrSlowMo.Enabled = True
                tbSpeed.Text = "Speed:" & SP.FrameRate & "fps"
            End If
        End If
    End Sub
    Public Sub OnSlideShowChange(sender As Object, e As EventArgs) Handles SP.SlideShowChange
        tmrSlideShow.Enabled = SP.Slideshow

    End Sub
    Public Sub OnRenameFolderStart(sender As Object, e As KeyEventArgs) Handles tvmain2.KeyDown
        If e.KeyCode = Keys.F2 Then
            CancelDisplay(True)
        End If
        If e.KeyCode = KeyTraverseTree Or e.KeyCode = KeyTraverseTreeBack Or e.KeyCode = Keys.Down Or e.KeyCode = Keys.Up Then
            tvmain2.tvFiles_KeyDown(sender, e)
        End If
    End Sub
    Friend Sub OnFolderMoved(ByVal Source As String, ByVal Dest As String)
        Dim d As New IO.DirectoryInfo(tvmain2.SelectedFolder)
        'tvmain2.RemoveNode(Source)
        tvmain2.MoveNode(Source, Dest)
        ' tvMain2.RefreshTree(d.FullName)
        'Dim dir As New IO.DirectoryInfo(path)
        'Dir.Delete()
    End Sub
#End Region
#Region "Markers"

    Private Sub AddMarker()
        If Media.IsLink Then
            CreateFavourite(Media.LinkPath)
        Else
            CreateFavourite(Media.MediaPath)
        End If
        PopulateLinkList(Media)
        '       e.SuppressKeyPress = True
        'DrawScrubberMarks()
    End Sub
    Private Sub RefreshLinks(List As List(Of String))
        For Each filename In List
            Dim x As List(Of String) = AllFaveMinder.GetLinksOf(filename)
            If x.Count <> 0 Then Op.ReuniteWithFile(x, filename)
        Next
    End Sub
    Public Sub RemoveMarker(filepath As String, timecode As Double)
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

                T.Start()
            End If
        End If

        Dim i = 0
        If x.Count = 0 Then
            chbPreviewLinks.Font = New Font(chbPreviewLinks.Font, FontStyle.Regular)
            chbPreviewLinks.Text = "Preview links (None)"
            Scrubber.BackColor = FBH.ListBox.BackColor
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
        End If
        Media.Markers = Media.GetMarkersFromLinkList
        Media.Markers.Sort()
        Marks.Markers = Media.Markers
        Dim longs As New List(Of Long)
        For Each m In Marks.Markers
            longs.Add(CLng(m))
        Next
        'FileMetadataModule.WriteMetadata(Media.MediaPath, longs, New List(Of String)())
        DrawScrubberMarks()



    End Sub

    Public Sub DrawScrubberMarks()

        If Media.Duration = 0 Then Exit Sub

        Marks.Duration = Media.Duration
        Scrubber.Width = Media.Player.Width * ScrubberProportion
        Scrubber.Left = Media.Player.Width * ((1 - ScrubberProportion) / 2)
        Marks.Create()

        '  Scrubber.Image = Marks.Bitmap 'NEVER add this back.


    End Sub
#End Region
#Region "Saving and loading"

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
        'ButtonHandling.BH = BH
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
        ProgressBarOn(100, "Loading list " & DesiredFile)
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
        LBH.ListBox.Tag = path
        'LastShowList = path
        FillBoxAsync(path)
        ProgressBarOff()
        tbShowfile.Text = "SHOWFILE LOADED:" & path
    End Sub
    Private Sub FillBoxAsync(path As String)
        Task.Run(Sub() FillBox(path))
    End Sub
    Private Sub FillBox(path As String)
        If path <> "" Then
            CollapseShowlist(False)

            LBH.ItemList = ReadListfromFile(path, True)
            LBH.FillBox()
            ' Getlist(Showlist, path, lbxShowList)
        Else

        End If
    End Sub
#End Region

#Region "Core Functions inc. Key Handling"
    Public Sub CancelDisplay(SlideshowsOff As Boolean)
        MSFiles.CancelURL(Media.MediaPath)
        'MSFiles.ResettersOff()


        If SlideshowsOff Then
            tmrMovieSlideShow.Enabled = False
            'tmrSlideShow.Enabled = False
            tmrAutoTrail.Enabled = False
            SP.Slideshow = False
        End If
    End Sub
    Sub RemoveBrackets(list As List(Of String))
        Dim x As New BracketRemover
        x.ListOfFiles = list
        x.RemoveBrackets()
        Exit Sub
        For Each f In list
            Dim file As New IO.FileInfo(f)
            If file.Exists Then
                Dim n As String = file.Name
                'Need to use regex here.
                For i = 1 To 100
                    If n = n.Replace("(" & Str(i) & ")", "") Then
                    Else
                        n.Replace("(" & Str(i) & ")", "")
                    End If


                Next
                n = AppendExistingFilename(n, file)
                If file.Name <> n Then

                    Try
                        My.Computer.FileSystem.RenameFile(file.FullName, n)

                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            End If

        Next
    End Sub
    Public Sub SelectDeadLinks(lbx As ListBox)

        HighlightList(lbx, GetDeadLinks(lbx))

    End Sub

    Public Sub HighlightList(lbx As ListBox, ls As List(Of String))
        lbx.SelectedItems.Clear()
        lbx.Refresh()
        lbx.SelectionMode = SelectionMode.MultiExtended
        For Each f In ls
            Dim i = lbx.FindString(f)
            If i >= 0 Then lbx.SetSelected(i, True)
        Next

    End Sub
    Private Sub DeleteFolder(FolderName As String, tvw As FileSystemTree, blnConfirm As Boolean)

        If My.Computer.FileSystem.DirectoryExists(FolderName) Then
            Dim m As MsgBoxResult = MsgBoxResult.No
            If blnConfirm Then
                m = MsgBox("Delete folder " & FolderName & "?", MsgBoxStyle.YesNoCancel)
            End If
            If Not blnConfirm OrElse m = MsgBoxResult.Yes Then

                SafeDeleteDir(New DirectoryInfo(FolderName))
                tvw.Traverse(False)
                tvw.RemoveNode(FolderName)
                BH.SwapPath(FolderName, "")
                BH.RefreshButtons()
            End If
        End If
        Exit Sub
    End Sub


    Public Sub UpdatePlayOrder(LBHandler As ListBoxHandler)
        If LBHandler.SortOrder.Equals(PlayOrder) Then
        Else
            LBHandler.SortOrder = PlayOrder
        End If
        If LBHandler Is LBH Then
            Dim s = lbxShowList.SelectedItem
            LBH.FillBox()
            LBH.SetNamed(s)
        Else
            Dim e = New DirectoryInfo(CurrentFolder)
            '  FillFileBox(lbxFiles, e, False)
            ' Dim t1 As New Thread(Sub()
            FBH.DirectoryPath = e.FullName
            FBH.SetNamed(Media.MediaPath)
            '                     End Sub)
        End If


    End Sub

    Private Function SpeedChange(e As KeyEventArgs) As KeyEventArgs
        Dim speednumber As Int16 = e.KeyCode - KeySpeed1
        If Media.Playing Then
            ' SP.TogglePause()
            If speednumber = SP.Speed Then
                ' MsgBox("Set Fullspeed")
                If speednumber = -1 Then
                    SP.TogglePause()
                Else

                    SP.Speed = -1
                End If
            Else
                SP.Speed = speednumber '-1,0,1 or 2. Could make 0, 1,2 and 3 to be compatible with AT'
            End If
        Else 'Slideshow
            If speednumber < 0 Then
                SP.Slideshow = Not SP.Slideshow
            Else
                SP.Slideshow = True
                SP.SSSpeed = speednumber
            End If
        End If
        e.SuppressKeyPress = True
        Return e
    End Function

    Public Sub GoFullScreen(blnGo As Boolean)
        FullScreen.Changing = True
        'TogglePause()
        'Media.Speed.Paused = True
        If blnGo Then
            'Dim s As String = Media.MediaPath
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
            SP.TogglePause()
            FullScreen.Show()
            FullScreen.FSFiles = MSFiles
            FullScreen.FSFiles.ListIndex = MSFiles.ListIndex
            ' TogglePause()
        Else
            '            SplitterPlace(0.25)
            MSFiles = FullScreen.FSFiles
            MSFiles.AssignPlayers(MainWMP1, MainWMP2, MainWMP3)
            MSFiles.AssignPictures(PictureBox1, PictureBox2, PictureBox3)
            MSFiles.ListIndex = MSFiles.ListIndex


            FullScreen.Close()
            'TogglePause()
        End If
        FullScreen.Changing = False
        blnFullScreen = Not blnFullScreen
    End Sub

    Private Sub ToggleRandomStartPoint()
        Random.StartPointFlag = Not Random.StartPointFlag
    End Sub


    Private Sub ToggleButtons()
        BH.SwitchRow(BH.buttons.CurrentRow)
        blnButtonsLoaded = Not blnButtonsLoaded
        ctrPicAndButtons.Panel2Collapsed = Not blnButtonsLoaded
        BH.UpdateButtonAppearance()
    End Sub
    Dim LBHH As New ListBoxHandler()
    Public Sub AdvanceFile(blnForward As Boolean, Optional Random As Boolean = False)
        LastGoodFile = Media.MediaPath
        SavePreferences(Parameters)

        LBHH.ListBox = FocusControl
        LBHH.ListBox.SelectionMode = SelectionMode.One
        'MSFiles.Listbox = LBHH.ListBox
        MSFiles.RandomNext = Random
        If Random Then
            LBHH.ListBox.SelectedItem = MSFiles.NextItem
        Else
            LBHH.IncrementIndex(blnForward)
        End If



    End Sub

    Public Sub CollapseShowlist(Collapse As Boolean)
        ShowListVisible = Not Collapse
        MasterContainer.Panel2Collapsed = Collapse
        lbxFiles.SelectionMode = SelectionMode.One
        If Initialising = False And Collapse Then ControlSetFocus(lbxFiles)
        MasterContainer.SplitterDistance = MasterContainer.Height / 3
        'UpdateFileInfo()
    End Sub


    Public Sub MediaSmallJump(Small As Boolean, forward As Boolean)
        Media.SmallJump(Small, forward)

    End Sub

    Public Sub MediaLargeJump(Small As Boolean, Forward As Boolean)
        Media.LargeJump(Small, Forward)

    End Sub
    Public Sub JumpRandom(blnAutoTrail As Boolean)
        If Media.MediaType = Filetype.Movie Or Media.IsLink Then
            If Not blnAutoTrail Then
                Media.Position = (Rnd(1) * (Media.Duration))
                tbStartpoint.Text = "START:RANDOM"
            Else
                ToggleAutoTrail()
            End If
        ElseIf Media.MediaType = Filetype.Pic Then
            AdvanceFile(True, True)
        Else
        End If
    End Sub
    Public Sub ToggleMovieSlideShow()
        tmrMovieSlideShow.Enabled = Not tmrMovieSlideShow.Enabled
        chbSlideShow.Checked = tmrMovieSlideShow.Enabled
        UpdateLabel(InstructionLabel2, ScanInstructions(0))
    End Sub
    Public Sub ToggleAutoTrail()

        tmrAutoTrail.Enabled = Not tmrAutoTrail.Enabled
        MSFiles.Soundhandler.AutoTrail = tmrAutoTrail.Enabled
        chbAutoTrail.Checked = tmrAutoTrail.Enabled
        TrailerModeToolStripMenuItem.Checked = tmrAutoTrail.Enabled
        If tmrAutoTrail.Enabled Then
            Media.SPT.SavedState = Media.SPT.State
            Media.SPT.State = StartPointHandler.StartTypes.Random
        Else
            Media.SPT.State = Media.SPT.SavedState
            Debug.Print("Normal")
            tmrSlowMo.Enabled = False
            Media.Player.settings.rate = 1
            Media.Player.Ctlcontrols.play()
        End If
        UpdateLabel(InstructionLabel2, ScanInstructions(2))

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
        'Principal key handling
        '  KeyDownFlag = True
        Me.Cursor = Cursors.WaitCursor
        Dim nke As New Keys
        nke = NumKeyEquivalent(e)

        Select Case e.KeyCode
#Region "Function Keys"
            Case Keys.F2
                CancelDisplay(False)
            Case Keys.F4
                e = ExitApplication(e)
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                HandleFunctionKeyDown(sender, e)
                e.SuppressKeyPress = True

            Case KeyDelete
                DeleteFiles(e)
#End Region

#Region "Alpha and Numeric"

            Case Keys.Enter
                HandleEnterKey(e)
            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If sender Is FolderSearch Then
                Else
                    RespondToKey(sender, e)

                End If

#End Region

#Region "Control Keys"
            Case KeyTraverseTree, KeyTraverseTreeBack
                '   If FocusControl IsNot tvMain2 Then
                tvmain2.Traverse(e.KeyCode = KeyTraverseTreeBack)
                e.SuppressKeyPress = True
              '  End If

            Case Keys.Left, Keys.Right
                ControlSetFocus(tvmain2)
                tvmain2.tvFiles_KeyDown(sender, e)
            Case Keys.Up, Keys.Down
                If tvmain2.Focused Then
                    ControlSetFocus(tvmain2)
                    tvmain2.tvFiles_KeyDown(sender, e)

                End If
            Case Keys.Escape
                MSFiles.CancelURL()
                BH.GetSubButtons("")
                CancelDisplay(True)

            Case KeyToggleButtons
                ToggleButtons()
            Case KeyNextFile, KeyPreviousFile, LKeyNextFile, LKeyPreviousFile
                SP.Slideshow = False
                AdvanceFile(e.KeyCode = KeyNextFile, Random.NextSelect)
                e.SuppressKeyPress = True
                'e = SelectNextFile(e)


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
                        If Media.Speed.Paused Then
                            If e.KeyCode = KeySmallJumpUp Then

                                Media.Player.Ctlcontrols.step(1)
                            Else

                                Media.Player.Ctlcontrols.step(-1)
                                For i = 0 To 22
                                    Media.Player.Ctlcontrols.step(1)
                                Next

                            End If
                            Media.Speed.PausedPosition = Media.Player.Ctlcontrols.currentPosition
                        Else

                            Media.SmallJump(e.Modifiers = Keys.Control, e.KeyCode > (KeySmallJumpUp + KeySmallJumpDown) / 2)
                            If tmrMovieSlideShow.Enabled Then
                                ToggleMovieSlideShow()
                            End If
                        End If

                    End If
                End If
                e.SuppressKeyPress = True

            Case KeyBigJumpOn, KeyBigJumpBack
                If e.Alt Then
                    If e.Control Then
                        SP.ChangeJump(True, e.KeyCode = KeyBigJumpOn)
                        UpdateFileInfo()
                    Else
                    End If
                Else
                    MediaLargeJump(e.Modifiers = Keys.Control, e.KeyCode = KeyBigJumpOn)
                    If tmrMovieSlideShow.Enabled Then
                        ToggleMovieSlideShow()
                    End If
                End If
            Case KeyMarkFavourite
                If e.Control And e.Alt And e.Shift Then
                    RemoveAllFavourites(Media.MediaPath)
                ElseIf e.Control And e.Alt Then 'Update link

                    If Media.IsLink Then 'TODO: Does this work?
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
                    Dim l As Double = Media.Markers(m)
                    If Media.IsLink Then
                        RemoveMarker(Media.LinkPath, l)
                    Else
                        RemoveMarker(Media.MediaPath, l)

                    End If
                Else
                    AddMarker()
                End If

            Case KeyJumpToPoint

                Media.MediaJumpToMarker(ToEnd:=True)
                e.SuppressKeyPress = True
            Case KeyJumpToMark, LKeyMarkPoint
                'Addmarker(Media.MediaPath)
                e = JumpToMark(e)

            Case KeyMuteToggle
                Muted = Not Muted
                MSFiles.MuteAll(Muted)
                e.SuppressKeyPress = True

            Case KeyLoopToggle

                e.SuppressKeyPress = True

            Case KeyJumpRandom
                If e.Control AndAlso e.Shift Then
                    JumpRandom(True) 'Autotrail
                ElseIf e.Alt Then
                    ToggleMovieScan()
                Else
                    JumpRandom(False)
                End If

            Case KeyToggleSpeed
                'If SP.Fullspeed Then
                '    SP.TogglePause()
                'Else
                SpeedChange(e)
                'End If
            Case KeySpeed1, KeySpeed2, KeySpeed3
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
                If CtrlDown And AltDown Then
                    SpawnSubButtonsForm("")
                Else
                    SelectSubList(FocusControl, False)

                End If

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
                'Same button used for fullscreen
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
                        tvmain2.SelectedFolder = LastFolder.Pop '= CurrentFolder
                    End If
                End If

            Case KeyReStartSS
                If e.Shift Then
                    ToggleMovieSlideShow()

                Else
                    SP.Slideshow = Not SP.Slideshow
                    ' chbSlideShow.Checked = SP.Slideshow

                End If
            Case Keys.Home
                If e.Control Then FBH.SetFirst(True)
            Case Keys.End
                If e.Control Then FBH.SetIndex(FBH.ItemList.Count - 1)
        End Select
        '   End If
        Me.Cursor = Cursors.Default
        'e.SuppressKeyPress = True
        '    Response.Enabled = False
    End Sub

    Private Sub HandleEnterKey(e As KeyEventArgs)
        If e.Control Then
            'Find selected file in filetree/list (finds source of a link, rather than the link)
            If Media.IsLink Then
                'A link finds the actual file
                tmrUpdateFileList.Enabled = False
                CurrentFilterState.State = FilterHandler.FilterState.All
                HighlightCurrent(Media.LinkPath)
            ElseIf FocusControl Is lbxShowList Then
                'If focus control is the Showlist
                CurrentFilterState.State = FilterHandler.FilterState.All
                HighlightCurrent(Media.MediaPath)
            Else
                'Get link file
                'Highlight 
                If Media.Markers.Count > 0 Then
                    'CTRL is held - highlight the link for this file.
                    If CurrentFilterState.SingleLinks Then
                        CurrentFilterState.State = FilterHandler.FilterState.All
                        HighlightCurrent(Media.MediaPath)
                    Else
                        Dim s As String = AllFaveMinder.GetLinksOf(Media.MediaPath)(Media.LinkCounter)
                        CurrentFilterState.State = FilterHandler.FilterState.All
                        HighlightCurrent(s)
                    End If

                End If
            End If
        Else
            'General press return refreshes tree.
            tvmain2.RefreshTree(CurrentFolder)
        End If
    End Sub

    Private Shared Function ExitApplication(e As KeyEventArgs) As KeyEventArgs
        Exit Function
        'If e.Alt Then
        '    For Each m In MSFiles.MediaHandlers
        '        m.ResetPositionCanceller.Enabled = True
        '    Next
        '    LastTimeSuccessful = True
        '    SavePreferences(Parameters)

        '    e.SuppressKeyPress = True
        '    Application.Exit()
        'End If

        'Return e
    End Function

    Public Sub HandleFunctionKeyDown(sender As Object, e As KeyEventArgs)

        Dim mBH As New ButtonHandler
        If TypeOf (sender) Is FormButton Then
            mBH = sender.BH
        Else
            mBH = BH
        End If
        Dim buttonnumber As Byte = e.KeyCode - Keys.F5
        Dim s As StateHandler.StateOptions = NavigateMoveState.State
        '  CancelDisplay() 'Need to cancel display to prevent 'already in use' problems when moving files or deleting them. 
        Dim path = mBH.buttons.CurrentRow.Buttons(buttonnumber).Path

        'Just assign in all modes when all three control buttons held
        If (e.Shift And e.Control And e.Alt) Then
            'Assign button
            mBH.AssignButton(buttonnumber, CurrentFolder)

            Exit Sub
        ElseIf e.Alt And e.Control Then
            sender = mBH.ActualButtons(buttonnumber)
            mBH.ShowPreview(sender, e)
            Exit Sub
        ElseIf e.Alt Then
            Autoload(mBH.buttons.CurrentRow.Buttons(buttonnumber).Label, True)
            Exit Sub
        End If

        If s <> StateHandler.StateOptions.Navigate Then
            'Non navigate behaviour
            If e.Control And e.Shift Then
                'Jump to folder
                If path <> CurrentFolder Then
                    ChangeFolder(path)

                    tvmain2.SelectedFolder = CurrentFolder
                ElseIf Random.OnDirChange Then
                    AdvanceFile(True, True) 'TODO: Whaat?
                End If
            ElseIf e.Shift Then
                'Move Folder
                If path <> "" Then
                    Dim source As String = tvmain2.SelectedFolder
                    MovingFolder(source, path)
                    Dim xx As String = System.IO.Path.GetFileName(source)
                    mBH.SwapPath(source, path & "\" & xx)
                End If
            Else

                Dim flag = blnSuppressCreate
                blnSuppressCreate = True
                If ShowListVisible Then
                    Dim m = LBH.SelectedItemsList
                    LBH.RemoveItems(m)

                    FBH.RemoveItems(m)
                    MoveFiles(m, path)
                Else
                    Dim m = FBH.SelectedItemsList
                    FBH.RemoveItems(m)
                    MoveFiles(m, path)
                End If
                blnSuppressCreate = flag
            End If
        Else
            'Navigate behaviour
            If e.Shift And e.Control And path <> "" Then
                'Move folder
                MovingFolder(tvmain2.SelectedFolder, path)

            ElseIf e.Shift Then
                'Move files
                If path <> "" Then
                    MoveFiles(ListfromSelectedInListbox(lbxFiles), path)
                End If

            ElseIf e.Alt Then
                Autoload(mBH.buttons.CurrentRow.Buttons(buttonnumber).Label, True)
            Else

                'SWITCH folder
                If path <> CurrentFolder Then
                    ChangeFolder(path)
                    'CancelDisplay()
                    tvmain2.SelectedFolder = path
                    ' FmBH.DirectoryPath = path
                Else
                    'Change file if same folder
                    AdvanceFile(True, True)

                End If
            End If
        End If

        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)

    End Sub



    Private Function JumpToMark(e As KeyEventArgs) As KeyEventArgs
        If Media.Markers.Count > 0 Then
            If e.Alt Then
                Media.LinkCounter = Media.RandomCounter
            Else
                Media.IncrementLinkCounter(e.Modifiers <> Keys.Control)
            End If
            Media.MediaJumpToMarker(ToMarker:=True)
        Else
            MediaLargeJump(e.Modifiers = Keys.Control, True)
        End If
        e.Handled = True
        e.SuppressKeyPress = True
        Return e
    End Function
    Private Sub JumpToMark(forward As Boolean, Optional randomJump As Boolean = False)
        If Media.Markers.Count > 0 Then
            If randomJump Then
                ' Assign a random index from the markers list
                Media.LinkCounter = Media.RandomCounter
            Else
                ' Increment or decrement the LinkCounter based on the forward parameter
                Media.IncrementLinkCounter(forward)
            End If

            ' Perform the jump to the specified marker
            Media.MediaJumpToMarker(ToMarker:=True)
        Else
            ' If no markers are available, perform a larger jump
            ' The function MediaLargeJump and its parameters might need to be adjusted according to its definition
            ' Assuming MediaLargeJump is another method that you've implemented
            Media.LargeJump(forward, False)
        End If
    End Sub


    Private Sub RemoveAllFavourites(mediaPath As String)
        Throw New NotImplementedException()
    End Sub

    Private Sub DeleteFiles(e As KeyEventArgs)
        'Use Movefiles with current selected list, and option to delete. 
        Dim lbx As New ListBox
        CancelDisplay(False)
        If e.Shift Then
            DeleteFolder(CurrentFolder, tvmain2, NavigateMoveState.State = StateHandler.StateOptions.Navigate)
        Else
            If FocusControl Is tvmain2 Then
            Else
                lbx = FocusControl
            End If

            Dim m As List(Of String) = ListfromSelectedInListbox(lbx)
            'If lbx Is lbxShowList And NavigateMoveState.State = StateHandler.StateOptions.Navigate Then
            LBH.RemoveItems(m)
            'Else
            FBH.RemoveItems(m)
            blnSuppressCreate = True
            MoveFiles(m, "")
            'FBH.Refresh()
            'End If


        End If
    End Sub

    Private Sub SpeedIncrease(e As KeyEventArgs)
        If e.KeyCode = KeySmallJumpUp Then
            SP.IncreaseSpeed()
        Else
            SP.DecreaseSpeed()
        End If
        'tmrSlideShow.Interval = SP.Interval
        'tmrSlowMo.Interval = 1000 / SP.FrameRate

    End Sub

    Private Sub NavigateToFavourites()
        CurrentFilterState.OldState = CurrentFilterState.State
        CurrentFilterState.State = FilterHandler.FilterState.Linked
        cbxSingleLinks.Checked = False
        ChangeFolder(CurrentFavesPath)

        tvmain2.SelectedFolder = CurrentFolder
    End Sub
    Public Sub ChangeFolder(strPath As String)
        If strPath = "" Then Exit Sub
        'If strPath <> FavesFolderPath Then
        '    CurrentfilterState.State = CurrentfilterState.OldState
        'End If
        If strPath.Contains(GlobalFavesPath) Then
            CurrentFilterState.State = FilterHandler.FilterState.LinkOnly
        Else
            CurrentFilterState.State = CurrentFilterState.OldState

        End If
        If strPath = CurrentFolder Then
        Else
            ' MSFiles.CancelURL()

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

            If BH.Autobuttons Then
                BH.AssignLinear(New DirectoryInfo(CurrentFolder))
                BH.buttons.CurrentLetter = Asc(Keys.D0)
            End If
            '   My.Computer.Registry.CurrentUser.SetValue("File", Media.MediaPath)
        End If

    End Sub



    Public Sub OnTVMainReady(sender As Object, e As EventArgs) Handles tvmain2.TreeBuilt
        HighlightCurrent(Media.MediaPath)
        Me.Show()
        Me.BringToFront()
        splashScreen.Close()
    End Sub

    Public HighlightPath As String
    Friend Shared LastGoodFile As String

    Friend Sub HighlightCurrent(strPath As String)
        HighlightPath = strPath
        tmrHighlightCurrent.Enabled = True



    End Sub




#End Region

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
        Me.Hide()
        splashScreen.UpdateStatus("Initialising...")
        Randomize()
        splashScreen.UpdateStatus("Populating lists...")

        PopulateLists()
        CollapseShowlist(True)
        splashScreen.UpdateStatus("Setting up players...")

        SetupPlayers()
        PositionUpdater.Enabled = False
        Media.DontLoad = True
        splashScreen.UpdateStatus("Fetching preferences...")

        'PreferencesGet()
        GetPreferences(Parameters)
        SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
        '   LastTimeSuccessful = False
        'Media.Picture = PictureBox1
        tbPercentage.Enabled = True


        AddHandler FileHandling.FolderMoved, AddressOf OnFolderMoved
        AddHandler FileHandling.FileMoved, AddressOf OnFilesMoved
        splashScreen.UpdateStatus("Initialising buttons (please wait)...")

        InitialiseButtons()
        MSFilesNew.MediaHandlers.Add(New MediaHandler("1"))
        MSFilesNew.MediaHandlers.Add(New MediaHandler("2"))
        MSFilesNew.MediaHandlers.Add(New MediaHandler("3"))

        OnRandomChanged()
        ControlSetFocus(lbxFiles)
        Media.DontLoad = False
        PopulateContextMenu("File")
        ctrPicAndButtons.Panel1.Controls.Add(Media.Textbox)
        FBH.DirectoryPath = CurrentFolder
        FBH.ListBox = lbxFiles
        LBH.ListBox = lbxShowList
        Dim ff As String = PassFilename()

        If ff.EndsWith(".msl") Then
            splashScreen.UpdateStatus("Loading showlist (please wait)...")

            LoadShowList(ff)
        End If
        Me.Show()
    End Sub
    Private Sub GetPreferences(p As AppSettings)
        LastTimeSuccessful = p.LastTimeSuccessful
        ctrFileBoxes.SplitterDistance = p.VertSplit
        ctrMainFrame.SplitterDistance = p.HorSplit
        ctrPicAndButtons.SplitterDistance = p.ButtonSplit
        CurrentFilterState.State = p.Filter
        PlayOrder.State = p.SortOrder
        Media.SPT.State = p.Startpoint
        NavigateMoveState.State = p.State
        ButtonFilePath = p.LastButtonFile
        iCurrentAlpha = p.CurrentAlpha
        Media.MediaPath = p.LastGoodFilePath
        CurrentFavesPath = p.FavouritesPath
        chbPreviewLinks.Checked = p.PreviewLinks
        Rootpath = p.RootScanPath
        GlobalFavesPath = p.GlobalFavesPath
        chbLoadButtonFiles.Checked = p.RandomAutoLoadButtons
        chbNextFile.Checked = p.RandomNextFile
        chbOnDir.Checked = p.RandomOnDirectoryChange
        ' chbAutoTrail.Checked = p.RandomAutoTrail
        chbShowAttr.Checked = p.ShowAttributes
        chbEncrypt.Checked = p.Encrypt
        CHBAutoAdvance.Checked = p.AutoAdvance
        chbSeparate.Checked = p.Separate
        ThumbDestination = p.ThumbnailDestination
        Buttonfolder = p.ButtonFolder
        SP.AbsoluteJump = p.AbsoluteJump
        SP.FractionalJump = p.FractionalJump
        NavigateMoveState.State = p.State
        chbSlideShow.Checked = p.MovieScan
        chbScan.Checked = p.ScanBookmarks
        cbxSingleLinks.Checked = p.SingleLinks
        tbMovieSlideShowSpeed.Value = p.MovieSlideShowSpeed
        SlowMoSoundOpt = p.SlowmoSoundOption
        cbxDontLoad.Checked = Boolean.Parse(p.DontLoadVidsShowPic)
        cbxDontPreLoad.Checked = Boolean.Parse(p.DontPreLoadVids)
    End Sub
    Private Sub SavePreferences(p As AppSettings)

        ' Capture settings from form and controls
        p.LastTimeSuccessful = LastTimeSuccessful
        p.LastGoodFilePath = Media.MediaPath
        p.VertSplit = ctrFileBoxes.SplitterDistance
        p.HorSplit = ctrMainFrame.SplitterDistance
        p.Filter = CurrentFilterState.State
        p.SortOrder = PlayOrder.State
        p.Startpoint = Media.SPT.State
        p.State = NavigateMoveState.State
        p.LastButtonFile = ButtonFilePath
        p.CurrentAlpha = iCurrentAlpha
        p.FavouritesPath = CurrentFavesPath
        p.PreviewLinks = chbPreviewLinks.Checked
        p.RootScanPath = Rootpath
        p.GlobalFavesPath = GlobalFavesPath
        p.RandomAutoLoadButtons = chbLoadButtonFiles.Checked
        p.RandomNextFile = chbNextFile.Checked
        p.RandomOnDirectoryChange = chbOnDir.Checked
        'p.RandomAutoTrail = chbAutoTrail.Checked
        p.ShowAttributes = chbShowAttr.Checked
        p.Encrypt = chbEncrypt.Checked
        p.AutoAdvance = CHBAutoAdvance.Checked
        p.Separate = chbSeparate.Checked
        p.ThumbnailDestination = ThumbDestination
        p.ButtonFolder = Buttonfolder
        p.AbsoluteJump = SP.AbsoluteJump
        p.FractionalJump = SP.FractionalJump
        p.MovieScan = chbSlideShow.Checked
        p.ScanBookmarks = chbScan.Checked
        p.SingleLinks = cbxSingleLinks.Checked
        p.MovieSlideShowSpeed = tbMovieSlideShowSpeed.Value
        p.SlowmoSoundOption = SlowMoSoundOpt
        p.DontLoadVidsShowPic = cbxDontLoad.Checked.ToString()
        p.DontPreLoadVids = cbxDontPreLoad.Checked.ToString()
        p.buttonSplit = ctrPicAndButtons.SplitterDistance
        SettingsManager.SaveSettings(p)
    End Sub


    Private Sub InitialiseButtons()
        BH.Tooltip = Me.ToolTip1
        BH.ActualButtons = {btn1, btn2, btn3, btn4, btn5, btn6, btn7, btn8}
        BH.InitialiseActualButtons()
        BH.Labels = {lbl1, lbl2, lbl3, lbl4, lbl5, lbl6, lbl7, lbl8}
        BH.LetterLabel = lblAlpha
        BH.LoadButtonSetAsync(LoadButtonFileName(ButtonFilePath))

        'AddHandler BH.buttons.LetterChanged, AddressOf OnLetterChanged
        BH.Alpha = iCurrentAlpha
        BH.RowProgressBar = ProgressBar1
        BH.SwitchRow(BH.buttons.CurrentRow)
        'Try
        '    KeyAssignmentsRestore(ButtonFilePath)

        'Catch ex As FileNotFoundException
        '    ctrPicAndButtons.Panel2Collapsed = True
        '    Exit Try
        'Catch ex As DirectoryNotFoundException
        '    ctrPicAndButtons.Panel2Collapsed = True

        '    Exit Try
        'End Try
    End Sub


    Private Sub Logchange(ByVal source As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher1.Changed
        If e.ChangeType = System.IO.WatcherChangeTypes.Changed Then
            Exit Sub 'TODO This gets repeatedly called. 
            MsgBox("File " & e.FullPath & " has been modified")
        End If

    End Sub

    Private Sub cbxDontLoad_CheckedChanged(sender As Object, e As EventArgs) Handles cbxDontLoad.CheckedChanged
        MSFiles.Media1.DontLoadVids = cbxDontLoad.Checked
        MSFiles.Media2.DontLoadVids = cbxDontLoad.Checked
        MSFiles.Media3.DontLoadVids = cbxDontLoad.Checked

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Media.ExtractFrameUsingFFmpeg(Media.MediaPath, ScreenDestination & "\" & System.IO.Path.GetFileNameWithoutExtension(Media.MediaPath) & ".png", TimeSpan.FromSeconds(Media.Position).ToString("hh\:mm\:ss"))
    End Sub

    Private Sub cbxDontPreLoad_CheckedChanged(sender As Object, e As EventArgs) Handles cbxDontPreLoad.CheckedChanged
        MSFiles.DontLoad = cbxDontPreLoad.Checked
    End Sub

    Private Sub FoldersAttemptToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FoldersAttemptToolStripMenuItem.Click
        Dim f As New FolderSearch
        f.Show()
    End Sub

    Private Sub Me_MenuComplete(sender As Object, e As EventArgs) Handles Me.MenuComplete

    End Sub

    Private Sub ctrFileBoxes_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles ctrFileBoxes.SplitterMoved

    End Sub

    Private Sub ctrMainFrame_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles ctrMainFrame.SplitterMoved
    End Sub
#End Region

End Class
