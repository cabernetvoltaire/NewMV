
Friend Module Mysettings


    Public PFocus As Byte = CtrlFocus.Tree

    Public Property ZoneSize As Decimal = 0.4

    Public Const OrientationId As Integer = &H112
    Public ButtonFilePath As String

    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)

#Region "Options"

    Public iSSpeeds() As Integer = {1000, 300, 200}
    Public iPlaybackSpeed() As Integer = {3, 15, 45}

    Public Property LastShowList As String
    Public Property blnJumpToMark As Boolean = False
    Public blnLoopPlay As Boolean = True
    'Public blnChooseRandomFile As Boolean = True
    Public PlaybackSpeed As Double = 30
    Public CurrentFavesPath As String

    Public Autozoomrate As Decimal = 0.4
    Public iCurrentAlpha As Integer = 0
#End Region

#Region "Internal"
    Public Property blnSecondScreen As Boolean = True
    'Public blnAutoAdvanceFolder As Boolean = True
    Public blnRestartSlideShowFlag As Boolean = False
    Public Property blnLink As Boolean
    Public Property lastselection As String


    Public lngInterval = 10
    Public lngMediaDuration As Long
    Public lngMark As Long
    Public strExt As String
    Public FileboxContents As New List(Of String)
    Public FBCShown As Boolean()
    Public fType As Filetype

    Public Showlist As New List(Of String)
    Public Oldlist As New List(Of String)

    Public blnDontShowRepeats As Boolean = True
    Public Sublist As New List(Of String)
    Public currentPicBox As New PictureBox
    Public strVisibleButtons(8) As String
    Public NofShown As Int16
    Public blnButtonsLoaded As Boolean = False
    Public ssspeed As Integer = 200
    'Public CountSelections As Int16
    Public blnFullScreen As Boolean
    Public Property lngListSizeBytes As Long
    'Public blnTVCurrent As Boolean

    Public ChosenPlayOrder As Byte = 0
#End Region
    Public Sub PreferencesSave()

        With My.Computer.Registry.CurrentUser
            .SetValue("VertSplit", MainForm.ctrFileBoxes.SplitterDistance)
            .SetValue("HorSplit", MainForm.ctrMainFrame.SplitterDistance)
            .SetValue("File", Media.MediaPath)
            .SetValue("Filter", MainForm.CurrentFilterState.State)
            .SetValue("SortOrder", MainForm.PlayOrder.State)
            .SetValue("StartPoint", Media.StartPoint.State)
            .SetValue("State", MainForm.NavigateMoveState.State)
            .SetValue("LastButtonFile", ButtonFilePath)
            .SetValue("LastAlpha", iCurrentAlpha)
            .SetValue("Favourites", CurrentFavesPath)
            .SetValue("PreviewLinks", MainForm.chbPreviewLinks.Checked)
            .SetValue("RootScanPath", Rootpath)
            .SetValue("Directories List", DirectoriesPath)
            .SetValue("GlobalFaves", GlobalFavesPath)
        End With

    End Sub

    Public Sub Preferences(GetPrefs As Boolean)
        Dim M As New Dictionary(Of String, Object)

        M.Add("VertSplit", MainForm.ctrFileBoxes.SplitterDistance)
        M.Add("HorSplit", MainForm.ctrMainFrame.SplitterDistance)
        M.Add("File", Media.MediaPath)
        M.Add("Filter", MainForm.CurrentFilterState.State)
        M.Add("SortOrder", MainForm.PlayOrder.State)
        M.Add("StartPoint", Media.StartPoint.State)
        M.Add("State", MainForm.NavigateMoveState.State)
        M.Add("LastButtonFile", ButtonFilePath)
        M.Add("LastAlpha", iCurrentAlpha)
        M.Add("Favourites", CurrentFavesPath)
        M.Add("PreviewLinks", MainForm.chbPreviewLinks.Checked)
        M.Add("RootScanPath", Rootpath)
        M.Add("Directories List", DirectoriesPath)

        For Each s In M
            If GetPrefs Then
                Dim o As Object = GetObject(s.Value)
                o = GetterSetter(s.Key, s.Value, True)
            Else
                GetterSetter(s.Key, s.Value, False)
            End If
        Next
    End Sub
    Private Function GetterSetter(Name As String, Value As Object, Getter As Boolean) As Object
        With My.Computer.Registry.CurrentUser
            If Getter Then
                Return .GetValue(Name, Value)
            Else
                .SetValue(Name, Value)
            End If
        End With
    End Function
    Public Sub PreferencesGet()
        'Preferences(True)
        'Exit Sub
        MainForm.ctrPicAndButtons.SplitterDistance = 9 * MainForm.ctrPicAndButtons.Height / 10
        With My.Computer.Registry.CurrentUser
            MainForm.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", MainForm.ctrFileBoxes.Height / 4)
            MainForm.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", MainForm.ctrFileBoxes.Width / 2)
            MainForm.CurrentFilterState.State = .GetValue("Filter", 0)
            Media.StartPoint.State = .GetValue("StartPoint", 0)
            MainForm.NavigateMoveState.State = .GetValue("State", 0)
            MainForm.PlayOrder.State = .GetValue("SortOrder", 0)
            iCurrentAlpha = .GetValue("LastAlpha", 0)
            ButtonFilePath = .GetValue("LastButtonFile", "")
            CurrentFavesPath = .GetValue("Favourites", CurrentFolder)
            Dim fol As New IO.DirectoryInfo(CurrentFavesPath)
            DirectoriesPath = .GetValue("Directories List", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            GlobalFavesPath = .GetValue("GlobalFaves", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            Rootpath = .GetValue("Rootpath", Environment.GetFolderPath(Environment.SpecialFolder.MyComputer))

            If fol.Exists = False Then
                MainForm.FavouritesFolderToolStripMenuItem.PerformClick()
            End If

            Dim s As String = .GetValue("File", "")
            If s = "" Then s = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            'Media.Player = MainForm.MainWMP4
            Media.MediaPath = s
            MainForm.chbPreviewLinks.Checked = .GetValue("PreviewLinks", False)

            'Catch ex As Exception
            ' MsgBox(ex.Message)
            'PreferencesReset()
            'End Try


        End With
        MainForm.tssMoveCopy.Text = CurrentFolder
    End Sub
    Public Sub PreferencesReset()
        If MsgBox("Reset preferences?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then

            With My.Computer.Registry.CurrentUser
                Dim s As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                'CurrentFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                Dim fol As New IO.DirectoryInfo(CurrentFolder)


                Media.MediaPath = New IO.DirectoryInfo(CurrentFolder).EnumerateFiles("*", IO.SearchOption.AllDirectories).First.FullName
                CurrentFavesPath = s & "\Favourites\"
                '            strButtonfile=Media.MediaPath
            End With
            With MainForm
                .ctrFileBoxes.SplitterDistance = .ctrFileBoxes.Height / 4
                .ctrMainFrame.SplitterDistance = .ctrFileBoxes.Width / 2

                .CurrentFilterState.State = .0
                .PlayOrder.State = 0
                .NavigateMoveState.State = 0
                iCurrentAlpha = 0
                ButtonFilePath = ""
                '.SetValue("LastButtonFolder", strButtonfile)
            End With
            Media.StartPoint.State = 0
            PreferencesSave()
        End If
    End Sub

End Module
