Imports System.Environment
Friend Module Mysettings

    Public FirstLoad As Boolean = False
    Public PFocus As Byte = CtrlFocus.Tree
    Public IncludeRemovables As Boolean = False
    Public Settings As New List(Of Setting)

#Region "Literals"
    Public Property ZoneSize As Decimal = 0.4
    Public Const OrientationId As Integer = &H112
    Public iSSpeeds() As Integer = {1000, 300, 1}
    Public iPlaybackSpeed() As Integer = {3, 15, 45}
    Public PlaybackSpeed As Double = 30
    Public Autozoomrate As Decimal = 0.4
    Public iCurrentAlpha As Integer = 0
    Public blnLoopPlay As Boolean = True
    Public blnSuppressCreate As Boolean = False
    Public blnChooseOne As Boolean = False
    Public Muted As Boolean = False

#End Region

#Region "Paths"


    Public ButtonFilePath As String
    Public Buttonfolder As String
    Public ThumbDestination As String
    Public ListFilePath As String
    Public PrefsPath As String = GetFolderPath(SpecialFolder.ApplicationData) & "\Metavisua\Preferences\"
    Public PrefsFilePath As String = PrefsPath & "\" & DrivesScan() & "MVPrefs.txt"
    Public Rootpath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
    Public GlobalFavesPath As String = Rootpath
    Public CurrentFavesPath As String = Rootpath
    Public LastTimeSuccessful As Boolean
    Private DriveString As String = "C"


#End Region
#Region "Stacks"
    Public LastPlayed As New Stack(Of String)
    Public LastFolder As New Stack(Of String)

#End Region
    Public Property LastShowList As String
    Private Property VariablesDescriptor() = {"Vertical Split position",
        "Horizontal split position",
        "Last file loaded",
        "Current filter state",
        "Current sort order",
        "Current start point state",
        "Current navigate/move state",
        "Last button file loaded",
        "Current alphabet button",
        "Current favourites location",
        "Preview links?",
        "Root scan path",
        "Directories list:",
        "Global favourites directory",
        "Choose next file at random",
        "Always choose random file on directory change",
        "Auto trail mode",
    "Auto load button sets",
    "Show attributes",
    "Preview links",
    "Encrypt text files",
    "Auto advance",
    "Directory containing thumbnails",
    "Directory containing button files"
    }
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
        Dim PrefsList As New List(Of String)
        With PrefsList

            .Add("LastTimeSuccessful" & "$" & LastTimeSuccessful)
            .Add("VertSplit" & "$" & FormMain.ctrFileBoxes.SplitterDistance)
            .Add("HorSplit" & "$" & FormMain.ctrMainFrame.SplitterDistance)
            If LastTimeSuccessful Then
                If Media.IsLink Then
                    .Add("File" & "$" & Media.LinkPath)
                Else
                    .Add("File" & "$" & Media.MediaPath)
                End If
            Else
                .Add("File" & "$" & FormMain.LastGoodFile)

            End If
            .Add("Filter" & "$" & FormMain.CurrentFilterState.State)
            .Add("SortOrder" & "$" & FormMain.PlayOrder.State)
            .Add("StartPoint" & "$" & Media.SPT.State)
            .Add("State" & "$" & FormMain.NavigateMoveState.State)
            .Add("LastButtonFile" & "$" & ButtonFilePath)
            .Add("LastAlpha" & "$" & iCurrentAlpha)
            .Add("Favourites" & "$" & CurrentFavesPath)
            .Add("PreviewLinks" & "$" & FormMain.chbPreviewLinks.Checked)
            .Add("RootScanPath" & "$" & Rootpath)
            .Add("Directories List" & "$" & DirectoriesListFile) 'DirectoriesListFile)
            .Add("GlobalFaves" & "$" & GlobalFavesPath)
            .Add("RandomNextFile" & "$" & FormMain.chbNextFile.Checked)
            .Add("RandomOnDirectoryChange" & "$" & FormMain.chbOnDir.Checked)
            .Add("RandomAutoTrail" & "$" & FormMain.chbAutoTrail.Checked)
            .Add("RandomAutoLoadButtons" & "$" & FormMain.chbLoadButtonFiles.Checked)
            .Add("OptionsShowAttr" & "$" & FormMain.chbShowAttr.Checked)
            .Add("OptionsPreviewLinks" & "$" & FormMain.chbPreviewLinks.Checked)
            .Add("OptionsEncrypt" & "$" & FormMain.chbEncrypt.Checked)
            .Add("OptionsAutoAdvance" & "$" & FormMain.CHBAutoAdvance.Checked)
            .Add("OptionsSeparate" & "$" & FormMain.chbSeparate.Checked)
            .Add("ThumbnailDestination" & "$" & ThumbDestination)
            .Add("ButtonFolder" & "$" & Buttonfolder)
            .Add("FractionalJump" & "$" & FormMain.SP.FractionalJump)
            .Add("AbsoluteJump" & "$" & FormMain.SP.AbsoluteJump)
            .Add("Speed" & "$" & FormMain.SP.Speed)
            .Add("MovieScan" & "$" & FormMain.chbSlideShow.Checked)
            .Add("Scan Bookmarks" & "$" & FormMain.chbScan.Checked)
            .Add("SingleLinks" & "$" & FormMain.cbxSingleLinks.Checked)
            .Add("Movie Slide Show Speed" & "$" & FormMain.tbMovieSlideShowSpeed.Value)
            .Add("SlowMo Sound Option" & "$" & FormMain.SlowMoSoundOpt)


        End With

        Dim f As New IO.FileInfo(PrefsFilePath)
        If f.Exists = False Then
        Else
            f.Delete()
            ' PrefsFilePath = PrefsFilePath.Replace(".", Str(Int(Rnd() * 1000)) & ".")
        End If
        PrefsFilePath = PrefsPath & "\" & DrivesScan() & "MVPrefs.txt"
        WriteListToFile(PrefsList, PrefsFilePath, False)


    End Sub
    Public Sub PreferencesSaveNew()
        With Settings
            '        
            '    "Horizontal split position",
            '    ,
            '   
            '    ,
            '    "Current navigate/move state",
            '    ,
            '    "Current alphabet button",
            '    "Current favourites location",
            '    "Preview links?",
            '    "Root scan path",
            '    "Directories list:",
            '    "Global favourites directory",
            '    "Choose next file at random",
            '    "Always choose random file on directory change",
            '    "Auto trail mode",
            '"Auto load button sets",
            '"Show attributes",
            '"Preview links",
            '"Encrypt text files",
            '"Auto advance",
            '"Directory containing thumbnails",
            '"Directory containing button files"

            .Add(New Setting("Last close successful", LastTimeSuccessful, False))
            .Add(New Setting("Vertical Split position", FormMain.ctrFileBoxes.SplitterDistance, FormMain.ctrFileBoxes.Height / 4))
            .Add(New Setting("Horizontal split position", FormMain.ctrMainFrame.SplitterDistance, FormMain.ctrFileBoxes.Width * 0.75))
            If LastTimeSuccessful Then
                .Add(New Setting("Last file loaded", Media.MediaPath, ""))
            Else
                .Add(New Setting("Last file loaded", FormMain.LastGoodFile, ""))
            End If
            .Add(New Setting("Current filter state", FormMain.CurrentFilterState.State, 0))
            .Add(New Setting("Current sort order", FormMain.PlayOrder.State, 0))
            .Add(New Setting("Current start point state", Media.SPT.State, 0))
            .Add(New Setting("Current navigate/move state", FormMain.NavigateMoveState.State, 0))
            .Add(New Setting("Last button file loaded", ButtonFilePath, ""))
            .Add(New Setting("Current alphabet button", iCurrentAlpha, 0))
            .Add(New Setting("Favourites", CurrentFavesPath, ""))
            .Add(New Setting("PreviewLinks", FormMain.chbPreviewLinks.Checked, False))
            .Add(New Setting("RootScanPath", Rootpath, "C:"))
            .Add(New Setting("Directories List", DirectoriesListFile, "")) 'DirectoriesListFile,))
            .Add(New Setting("GlobalFaves", GlobalFavesPath, "C:"))


        End With

        ''    Dim f As New IO.FileInfo(PrefsFilePath)
        ''    If f.Exists = False Then
        ''    Else
        ''        f.Delete()
        ''        ' PrefsFilePath = PrefsFilePath.Replace(".", Str(Int(Rnd() * 1000)) & ".")
        ''    End If

        ''    WriteListToFile(PrefsList, PrefsFilePath, False)


    End Sub




    Public Sub PreferencesGet()
        If FirstLoad Then
            ResetDefaultPrefs()
            Exit Sub
        End If
        Dim prefslist As New List(Of String)
        Dim f As New IO.FileInfo(PrefsFilePath)
        Dim prefs As New IO.DirectoryInfo(PrefsPath)
        'Folder not found
        If prefs.Exists = False Then
            InitialiseFolders()
        End If
        'Replace this.
        'If prefs.GetFiles.Count > 1 Then
        '    Dim found As Boolean
        '    For Each m In prefs.GetFiles
        '        Dim drives As String = DrivesScan()
        '        If m.Name.StartsWith(drives + "MV") Then
        '            f = m
        '            found = True

        '        End If
        '    Next
        '    If Not found Then
        '        'No suitable prefs file found
        '        MsgBox("New arrangement of drives")
        '        PreferencesReset(False)
        '    End If
        'Else
        '    If prefs.GetFiles.Count = 1 Then
        '        f = prefs.GetFiles.First
        '    End If
        'End If
        'Instead use:
        'Find all path refs in prefs. 
        'If any drives don't exist, create a new preferences file. 

        If f.Exists Then
            prefslist = ReadListfromFile(f.FullName, False)
            FormMain.ctrPicAndButtons.SplitterDistance = 8.7 * FormMain.ctrPicAndButtons.Height / 10
            If CheckDrives(prefslist) Then
                For Each s In prefslist
                    If InStr(s, "$") <> 0 Then

                        Dim name As String = Split(s, "$")(0)
                        Dim value As String = Split(s, "$")(1)
                        Select Case name
                            Case "LastTimeSuccessful"
                                LastTimeSuccessful = value
                            Case "VertSplit"
                                FormMain.ctrFileBoxes.SplitterDistance = value
                            Case "HorSplit"
                                FormMain.ctrMainFrame.SplitterDistance = value
                            Case "File" 'Drive
                                If PassFilename() = "" Then
                                    Media.MediaPath = value
                                Else

                                    Media.MediaPath = PassFilename()
                                End If
                                DriveString = GetDriveString(value)

                            Case "Filter"
                                If value = "" Then value = 0
                                FormMain.CurrentFilterState.State = value

                            Case "SortOrder"
                                If value = "" Then value = 0
                                FormMain.PlayOrder.State = value
                            Case "StartPoint"
                                If value = "" Then value = 0
                                Media.SPT.State = value
                            Case "State"
                                If value = "" Then value = 0
                                FormMain.NavigateMoveState.State = value
                            Case "LastButtonFile" 'Drive
                                If PassFilename().Contains(".msb") Then
                                    value = PassFilename()
                                End If

                                If value = "" Then value = LoadButtonFileName("")
                                ButtonFilePath = value
                                DriveString = GetDriveString(value)
                            Case "LastAlpha"
                                If value = "" Then value = 0
                                iCurrentAlpha = value
                            Case "Favourites" 'Drive
                                If value = "" Then value = BrowseToFolder("Choose favourites path", PrefsPath)
                                CurrentFavesPath = value
                                DriveString = GetDriveString(value)
                            Case "PreviewLinks"
                                If value = "" Then value = False
                                FormMain.chbPreviewLinks.Checked = value
                            Case "RootScanPath" 'Drive
                                If value = "" Then value = BrowseToFolder("Choose root scanning path", PrefsPath)
                                Rootpath = value
                                DriveString = GetDriveString(value)
                            Case "Directories List" 'Drive
                                If value = "" Then value = PrefsPath & "Directories.txt"


                                DriveString = GetDriveString(value)
                            Case "GlobalFaves" 'Drive
                                Dim ff As New IO.DirectoryInfo(value)
                                If value = "" Or Not ff.Exists Then
                                    value = BrowseToFolder("Choose global favourites path", PrefsPath)
                                End If
                                GlobalFavesPath = value
                                DriveString = GetDriveString(value)
                            Case "AutoLoadButtons"
                                If value = "" Then value = False
                                FormMain.chbLoadButtonFiles.Checked = value

                            Case "RandomNextFile"
                                If value = "" Then value = False

                                FormMain.chbNextFile.Checked = value
                            Case "RandomOnDirectoryChange"
                                If value = "" Then value = False
                                FormMain.chbOnDir.Checked = value
                            Case "RandomAutoTrail"
                                Exit Select
                                value = False 'TODO causing load problems when omitted
                                If value = "" Then value = False
                                FormMain.chbAutoTrail.Checked = value
                            Case "RandomAutoLoadButtons"
                                If value = "" Then value = False
                                FormMain.chbLoadButtonFiles.Checked = value
                            Case "OptionsShowAttr"
                                If value = "" Then value = False
                                FormMain.chbShowAttr.Checked = value
                            Case "OptionsPreviewLinks"
                                If value = "" Then value = False
                                FormMain.chbPreviewLinks.Checked = value
                            Case "OptionsEncrypt"
                                If value = "" Then value = False
                                FormMain.chbEncrypt.Checked = value
                            Case "OptionsAutoAdvance"
                                If value = "" Then value = False
                                FormMain.CHBAutoAdvance.Checked = value
                            Case "OptionsSeparate"
                                If value = "" Then value = False
                                FormMain.chbSeparate.Checked = value

                            Case "ThumbnailDestination" 'Drive
                                If value = "" Then value = BrowseToFolder("Choose Thumbnail Destination", PrefsPath)
                                ThumbDestination = value
                                DriveString = GetDriveString(value)

                            Case "ButtonFolder" 'Drive
                                If value = "" Then value = BrowseToFolder("Choose Button folder", PrefsPath)
                                Buttonfolder = value
                                DriveString = GetDriveString(value)
                            Case "AbsoluteJump"
                                FormMain.SP.AbsoluteJump = value
                            Case "FractionalJump"
                                FormMain.SP.FractionalJump = value
                            Case "Speed"
                          '  FormMain.SP.Speed = value
                            Case "Navstate"
                                FormMain.NavigateMoveState.State = value
                            Case "MovieScan"
                                FormMain.chbSlideShow.Checked = value

                            Case "Scan Bookmarks"
                                FormMain.chbScan.Checked = value
                            Case "SingleLinks"
                                FormMain.cbxSingleLinks.Checked = value
                            Case "Movie Slide Show Speed"
                                FormMain.tbMovieSlideShowSpeed.Value = value
                            Case "SlowMo Sound Option"
                                FormMain.SlowMoSoundOpt = value
                        End Select

                    End If
                Next
            Else 'No relevant prefs file found.
                PreferencesReset(True)
            End If
        End If
        FormMain.tssMoveCopy.Text = CurrentFolder
        'MsgBox(DriveString)
    End Sub
    Private Function BrowseToFolder(Title As String, DefaultPath As String) As String
        Dim path As String = ""
        With New FolderBrowserDialog
            .SelectedPath = DefaultPath
            .Description = Title
            If .ShowDialog() = DialogResult.OK Then
                path = .SelectedPath
            End If
        End With

        Return path
    End Function

    Private Function CheckDrives(s As List(Of String)) As Boolean
        Dim result As Boolean = True
        For Each p In s
            Dim pos As Integer = InStr(p, ":\") - 1
            If pos > 1 Then
                Dim d As New IO.DriveInfo(Mid(p, pos, 1))
                If d.IsReady Then
                    result = True And result
                End If
            End If
        Next
        Return result
    End Function

    Private Function GetDriveString(path As String) As String
        'Adds drive of path if not already present
        If path = "" Then path = "C"
        If InStr(DriveString, path(0)) <> 0 Then
        Else
            DriveString = DriveString & path(0)
        End If
        Return DriveString
    End Function
    Public Sub PreferencesReset(Check As Boolean)
        If Check Then
            If MsgBox("Reset preferences?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ResetDefaultPrefs()
            End If
        Else
            ResetDefaultPrefs()
        End If
    End Sub

    Private Sub ResetDefaultPrefs()
        InitialiseFolders()

        With My.Computer.Registry.CurrentUser
            Dim s As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            CurrentFolder = s
            Dim fol As New IO.DirectoryInfo(s)
        End With
        With FormMain
            .ctrFileBoxes.SplitterDistance = .ctrFileBoxes.Height / 4
            .ctrMainFrame.SplitterDistance = .ctrFileBoxes.Width * 0.75
            .ctrPicAndButtons.SplitterDistance = .ctrMainFrame.Height * 8 / 9
            .CurrentFilterState.State = 0
            .PlayOrder.State = 0
            .NavigateMoveState.State = 0
            iCurrentAlpha = 0
            ButtonFilePath = ""
        End With
        Media.SPT.State = 0
        PreferencesSave()
    End Sub
    Public Sub loadform()
        MsgBox("Not implemented yet")
    End Sub
    Public Sub LoadFormOld()
        Dim prefslist As New List(Of String)
        prefslist = ReadListfromFile(PrefsFilePath, False)
        Dim screed As String = ""

        Preferences.Show()

        Dim s() As String = prefslist.ToArray()
        For i = 0 To s.GetUpperBound(0)
            Dim name As String = Split(s(i), "$")(0)
            Dim value As String = Split(s(i), "$")(1)
            screed = screed + VariablesDescriptor(i) + "$" + value + vbCrLf

        Next
        screed = screed.Replace("$", vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
        Preferences.PrefsText.Text = screed

    End Sub
    ''' <summary>
    ''' Creates five folders in Application Data/Metavisua
    ''' </summary>
    Private Sub InitialiseFolders()
        Dim f As New IO.DirectoryInfo(GetFolderPath(SpecialFolder.ApplicationData))
        f.CreateSubdirectory("Metavisua")
        Dim subdir As New IO.DirectoryInfo(f.FullName & "\Metavisua")
        Dim x() As String = {"Thumbs", "Buttons", "Lists", "Preferences", "Shortcuts"}
        For i = 0 To x.Length - 1
            subdir.CreateSubdirectory(x(i))

        Next
        For i = 0 To x.Length - 1
            Dim subsub As New IO.DirectoryInfo(subdir.FullName & "\" & x(i))
            Select Case i
                Case 0
                    ThumbDestination = subsub.FullName
                Case 1
                    Buttonfolder = subsub.FullName
                Case 2
                    ListFilePath = subsub.FullName
                Case 3
                    PrefsFilePath = subsub.FullName & "\MVPrefs.txt"
                Case 4
                    CurrentFavesPath = subsub.FullName
                    GlobalFavesPath = subsub.FullName

            End Select
        Next

    End Sub
    Private Function DrivesScan() As String
        Dim drivestring As String = ""
        For Each m In My.Computer.FileSystem.Drives
            If m.DriveType = IO.DriveType.Fixed Then
                drivestring = drivestring + m.Name
                drivestring = drivestring.Replace(":\", "")
            ElseIf IncludeRemovables Then
                drivestring = drivestring + m.Name
                drivestring = drivestring.Replace(":\", "")
            End If

        Next
        Return drivestring
    End Function
    Public Class Setting

        Public Sub New(Descriptor, Value, DefaultValue)
        End Sub
        Public Sub New()

        End Sub

        Property Descriptor As String
        Private _Value As Object
        Property Value As Object
            Get
                Return _Value
            End Get
            Set
                _Value = Value
            End Set
        End Property
        Private _Zone As String
        Public Property Zone() As String
            Get
                Return _Zone
            End Get
            Set(ByVal value As String)
                _Zone = value
            End Set
        End Property
        Property DefaultValue


    End Class

End Module

