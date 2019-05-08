
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
    Public FavesFolderPath As String

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
            .SetValue("Favourites", FavesFolderPath)
            .SetValue("PreviewLinks", MainForm.chbPreviewLinks.Checked)
        End With

    End Sub
    Public Sub PreferencesGet()
        MainForm.ctrPicAndButtons.SplitterDistance = 9 * MainForm.ctrPicAndButtons.Height / 10
        With My.Computer.Registry.CurrentUser
            ' Try
            MainForm.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", MainForm.ctrFileBoxes.Height / 4)
                MainForm.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", MainForm.ctrFileBoxes.Width / 2)
                ButtonFilePath = .GetValue("LastButtonFile", "")
            MainForm.CurrentFilterState.State = .GetValue("Filter", 0)
            Media.StartPoint.State = .GetValue("StartPoint", 0)
            MainForm.NavigateMoveState.State = .GetValue("State", 0)
            MainForm.PlayOrder.State = .GetValue("SortOrder", 0)
            iCurrentAlpha = .GetValue("LastAlpha", 0)
            FavesFolderPath = .GetValue("Favourites", CurrentFolder)
                Dim fol As New IO.DirectoryInfo(FavesFolderPath)
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
                FavesFolderPath = s & "\Favourites\"
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
