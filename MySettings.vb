
Friend Module Mysettings


    Public PFocus As Byte = CtrlFocus.Tree
#Region "Literals"
    Public Property ZoneSize As Decimal = 0.4
    Public Const OrientationId As Integer = &H112
    Public iSSpeeds() As Integer = {1000, 300, 200}
    Public iPlaybackSpeed() As Integer = {3, 15, 45}
    Public PlaybackSpeed As Double = 30
    Public Autozoomrate As Decimal = 0.4
    Public iCurrentAlpha As Integer = 0

#End Region

#Region "Paths"
    Public blnLoopPlay As Boolean = True
    Public CurrentFavesPath As String
    Public ButtonFilePath As String
#End Region
#Region "Stacks"
    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)

#End Region
    Public Property LastShowList As String
#Region "Internal"
    Public Property blnSecondScreen As Boolean = True
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
            .SetValue("File", CurrentFolder)
            .SetValue("Filter", MainForm.CurrentFilterState.State)
            .SetValue("SortOrder", MainForm.PlayOrder.State)
            .SetValue("StartPoint", Media.StartPoint.State)
            .SetValue("State", MainForm.NavigateMoveState.State)
            .SetValue("LastButtonFile", ButtonFilePath)
            .SetValue("LastAlpha", iCurrentAlpha)
            .SetValue("Favourites", CurrentFavesPath)
            .SetValue("PreviewLinks", MainForm.chbPreviewLinks.Checked)
            .SetValue("RootScanPath", Rootpath)
            .SetValue("Directories List", DirectoriesListFile)
            .SetValue("GlobalFaves", GlobalFavesPath)
        End With

    End Sub
    Public Structure RegistryEntry
        Dim Name As String
        Dim DefaultValue As Object
        Dim VariableName As String
    End Structure
    Private Function AssignRegEntryValues(Name As String, DefaultValue As Object, VariableName As String) As RegistryEntry
        Dim m As New RegistryEntry
        m.Name = Name
        m.DefaultValue = DefaultValue
        m.VariableName = VariableName
        Return m
    End Function
    Public Sub Preferences(GetPrefs As Boolean)
        Dim M As New List(Of RegistryEntry)

        'M.Add(AssignRegEntryValues("VertSplit", MainForm.ctrFileBoxes.Height / 4, MainForm.ctrFileBoxes.SplitterDistance))
        'M.Add(AssignRegEntryValues("HorSplit",, MainForm.ctrMainFrame.SplitterDistance))
        'M.Add(AssignRegEntryValues("File",, Media.MediaPath))
        'M.Add(AssignRegEntryValues("Filter",, MainForm.CurrentFilterState.State))
        'M.Add(AssignRegEntryValues("SortOrder",, MainForm.PlayOrder.State))
        'M.Add(AssignRegEntryValues("StartPoint",, Media.StartPoint.State))
        'M.Add(AssignRegEntryValues("State", , MainForm.NavigateMoveState.State))
        'M.Add(AssignRegEntryValues("LastButtonFile", ButtonFilePath))
        'M.Add(AssignRegEntryValues("LastAlpha", iCurrentAlpha))
        'M.Add(AssignRegEntryValues("Favourites", CurrentFavesPath))
        'M.Add(AssignRegEntryValues("PreviewLinks", MainForm.chbPreviewLinks.Checked))
        'M.Add(AssignRegEntryValues("RootScanPath", Rootpath))
        'M.Add(AssignRegEntryValues("Directories List", DirectoriesPath))


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
        '  Try

        MainForm.ctrPicAndButtons.SplitterDistance = 8.7 * MainForm.ctrPicAndButtons.Height / 10
            With My.Computer.Registry.CurrentUser
                'Appearance
                MainForm.ctrFileBoxes.SplitterDistance = .GetValue("VertSplit", MainForm.ctrFileBoxes.Height / 4)
                MainForm.ctrMainFrame.SplitterDistance = .GetValue("HorSplit", MainForm.ctrFileBoxes.Width / 2)
                'States
                MainForm.CurrentFilterState.State = .GetValue("Filter", 0)
                Media.StartPoint.State = .GetValue("StartPoint", 0)
                MainForm.NavigateMoveState.State = .GetValue("State", 0)
                MainForm.PlayOrder.State = .GetValue("SortOrder", 0)
                iCurrentAlpha = .GetValue("LastAlpha", 0)

                'Files
                ButtonFilePath = .GetValue("LastButtonFile", "")
                CurrentFavesPath = .GetValue("Favourites", CurrentFolder)

                Dim fol As New IO.DirectoryInfo(CurrentFavesPath) 'TODO This whole thing is a mess

                DirectoriesListFile = .GetValue("Directories List", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
                'All .lnk files in this hierarchy get recognised and changed when files are moved. 
                GlobalFavesPath = .GetValue("GlobalFaves", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures))
            Rootpath = .GetValue("RootScanpath", "Q:\")
            If Rootpath = "" Then Rootpath = "Q:\"
            Dim folroot As New IO.DirectoryInfo(Rootpath)
            If folroot.Exists Then
                    DirectoriesList = GetDirectoriesList(Rootpath)
                Else
                    Rootpath = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)
                End If
                If fol.Exists = False Then
                    MainForm.FavouritesFolderToolStripMenuItem.PerformClick()
                End If

                Dim s As String = .GetValue("File", "")
                If s = "" Then s = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                Media.MediaPath = s
                MainForm.chbPreviewLinks.Checked = .GetValue("PreviewLinks", False)



            End With
        'Catch ex As Exception
        'PreferencesReset()
        'End Try
        MainForm.tssMoveCopy.Text = CurrentFolder
    End Sub
    Public Sub PreferencesReset()
        If MsgBox("Reset preferences?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            With My.Computer.Registry.CurrentUser
                Dim s As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                CurrentFolder = s
                Dim fol As New IO.DirectoryInfo(s)
                'Media.MediaPath = New IO.DirectoryInfo(CurrentFolder).EnumerateFiles("*", IO.SearchOption.AllDirectories).First.FullName
                CurrentFavesPath = s & "\MVFavourites\"
                '            strButtonfile=Media.MediaPath
            End With
            With MainForm
                .ctrFileBoxes.SplitterDistance = .ctrFileBoxes.Height / 4
                .ctrMainFrame.SplitterDistance = .ctrFileBoxes.Width / 2

                .CurrentFilterState.State = 0
                .PlayOrder.State = 0
                .NavigateMoveState.State = 0
                iCurrentAlpha = 0
                ButtonFilePath = ""
            End With
            Media.StartPoint.State = 0
            PreferencesSave()
        End If
    End Sub

End Module
