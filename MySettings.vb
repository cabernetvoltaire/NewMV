﻿Imports System.Environment
Imports Newtonsoft.Json
Imports System.IO
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
    Public ScreenDestination As String
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
        SettingsManager.SaveSettings(FormMain.Parameters)
        Exit Sub
        'SavePreferences()
        'Exit Sub
        Dim PrefsList As New List(Of String)
        With PrefsList

            .Add("LastTimeSuccessful" & "$" & LastTimeSuccessful)
            .Add("VertSplit" & "$" & FormMain.ctrFileBoxes.SplitterDistance)
            .Add("HorSplit" & "$" & FormMain.ctrMainFrame.SplitterDistance)
            If LastTimeSuccessful Then
                'If Media.IsLink Then
                '    .Add("File" & "$" & Media.LinkPath)
                'Else
                .Add("File" & "$" & Media.MediaPath)
                'End If
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
            .Add("DontLoadVids" & "$" & FormMain.cbxDontLoad.Checked.ToString)
            .Add("DontPreLoadVids" & "$" & FormMain.cbxDontPreLoad.Checked.ToString)


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
    Public Sub PreferencesGetNew(p As AppSettings)
        ' Return and reset preferences if this is the first load

        ' Ensure the preferences directory exists, initialize folders if it doesn't

        ' Proceed only if the preferences file exists
        If IO.File.Exists(PrefsFilePath) Then
            Dim prefsList As List(Of String) = ReadListfromFile(PrefsFilePath, False)

            ' Set the splitter distance based on the form's height

            ' Load preferences from the list if all required drives are available,
            ' otherwise reset to default preferences
            If CheckDrives(prefsList) Then
                LoadPreferencesFromList(prefsList)
            Else
                PreferencesReset(True)
            End If
        End If

        ' Update the status strip with the current folder path
        FormMain.tssMoveCopy.Text = CurrentFolder
    End Sub

    Public Sub PreferencesGet()
        Exit Sub
        ' Return and reset preferences if this is the first load
        If FirstLoad Then
            ResetDefaultPrefs()
            Exit Sub
        End If

        ' Ensure the preferences directory exists, initialize folders if it doesn't
        If Not IO.Directory.Exists(PrefsPath) Then
            InitialiseFolders()
        End If

        ' Proceed only if the preferences file exists
        If IO.File.Exists(PrefsFilePath) Then
            Dim prefsList As List(Of String) = ReadListfromFile(PrefsFilePath, False)

            ' Set the splitter distance based on the form's height
            FormMain.ctrPicAndButtons.SplitterDistance = 8 * FormMain.ctrPicAndButtons.Height / 10

            ' Load preferences from the list if all required drives are available,
            ' otherwise reset to default preferences
            If CheckDrives(prefsList) Then
                LoadPreferencesFromList(prefsList)
            Else
                PreferencesReset(True)
            End If
        End If

        ' Update the status strip with the current folder path
        FormMain.tssMoveCopy.Text = CurrentFolder
    End Sub

    Private Sub LoadPreferencesFromList(prefslist As List(Of String))
        For Each s In prefslist
            If InStr(s, "$") <> 0 Then
                Dim name As String = Split(s, "$")(0)
                Dim value As String = Split(s, "$")(1)
                ProcessPreferenceItem(name, value)
            End If
        Next
    End Sub


    Private Sub ProcessPreferenceItem(name As String, value As String)
        Select Case name
            Case "LastTimeSuccessful"
                LastTimeSuccessful = value
            Case "VertSplit"
                FormMain.ctrFileBoxes.SplitterDistance = value
            Case "HorSplit"
                FormMain.ctrMainFrame.SplitterDistance = value
            Case "File" 'Drive
                If PassFilename() = "" Then
                    If LastTimeSuccessful Then
                        Media.MediaPath = value
                    Else
                        Media.MediaPath = ""
                    End If
                Else
                    Media.MediaPath = PassFilename()
                End If
            Case "Filter"
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
                ' Similar cases here ...
            Case "DontLoadVids"
                FormMain.cbxDontLoad.Checked = Boolean.Parse(value)
            Case "DontPreLoadVids"
                FormMain.cbxDontPreLoad.Checked = Boolean.Parse(value)
            Case Else
                ' Handle other cases ...
        End Select
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
        Exit Sub
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

        ' Preferences.Show()

        Dim s() As String = prefslist.ToArray()
        For i = 0 To s.GetUpperBound(0)
            Dim name As String = Split(s(i), "$")(0)
            Dim value As String = Split(s(i), "$")(1)
            screed = screed + VariablesDescriptor(i) + "$" + value + vbCrLf

        Next
        screed = screed.Replace("$", vbTab + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
        ' Preferences.PrefsText.Text = screed

    End Sub
    ''' <summary>
    ''' Creates five folders in Application Data/Metavisua
    ''' </summary>
    Private Sub InitialiseFolders()
        ' Define the base directory under Application Data
        Dim baseDir As String = IO.Path.Combine(GetFolderPath(SpecialFolder.ApplicationData), "Metavisua")

        ' Ensure the base directory exists
        IO.Directory.CreateDirectory(baseDir)

        ' Define subdirectories
        Dim subDirs As String() = {"Thumbs", "Buttons", "Lists", "Preferences", "Shortcuts", "Screenshots"}

        ' Iterate through the subdirectories and create them
        For Each subDirName As String In subDirs
            Dim subDirPath As String = IO.Path.Combine(baseDir, subDirName)
            IO.Directory.CreateDirectory(subDirPath)

            ' Assign directory paths to global variables based on the subdirectory name
            Select Case subDirName
                Case "Thumbs"
                    ThumbDestination = subDirPath
                Case "Buttons"
                    Buttonfolder = subDirPath
                Case "Lists"
                    ListFilePath = subDirPath
                Case "Preferences"
                    PrefsFilePath = IO.Path.Combine(subDirPath, "MVPrefs.txt")
                Case "Shortcuts"
                    CurrentFavesPath = subDirPath
                    GlobalFavesPath = subDirPath
                Case "Screenshots"
                    ScreenDestination = subDirPath

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
    Public Class Preference
        Private Shared _preferences As Dictionary(Of String, Object) = Nothing
        Private Shared _prefsPath As String = IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Metavisua", "Preferences")

        Public Shared ReadOnly Property Preferences() As Dictionary(Of String, Object)
            Get
                If _preferences Is Nothing Then
                    LoadPreferences()
                End If
                Return _preferences
            End Get
        End Property

        Private Shared Sub LoadPreferences()
            Dim prefsFilePath As String = GetPrefsFilePath()

            If File.Exists(prefsFilePath) Then
                Try
                    Dim prefsJson As String = File.ReadAllText(prefsFilePath)
                    _preferences = JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(prefsJson)
                Catch ex As Exception
                    MsgBox("Error loading preferences file: " & ex.Message)
                End Try
            Else
                _preferences = New Dictionary(Of String, Object)
                SavePreferences()
            End If
        End Sub

        Public Shared Sub SavePreferences()
            Dim preferencesJson As String = JsonConvert.SerializeObject(_preferences)

            Dim prefsFilePath As String = GetPrefsFilePath()

            If IO.File.Exists(prefsFilePath) Then
                IO.File.Delete(prefsFilePath)
            End If

            IO.File.WriteAllText(prefsFilePath, preferencesJson)
        End Sub

        Private Shared Function GetPrefsFilePath() As String
            Dim driveLetters As String = GetDriveLetters()

            Dim prefsFilePath As String = IO.Path.Combine(_prefsPath, driveLetters & "MVPrefs.json")

            Return prefsFilePath
        End Function

        Private Shared Function GetDriveLetters() As String
            Dim driveLetters As String = ""

            For Each drive As DriveInfo In DriveInfo.GetDrives()
                If drive.IsReady Then
                    driveLetters &= drive.Name.Substring(0, 1)
                End If
            Next

            Return driveLetters
        End Function
    End Class
End Module

