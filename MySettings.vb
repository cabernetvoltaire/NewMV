﻿Imports System.Environment
Friend Module Mysettings

    Public FirstLoad As Boolean = False
    Public PFocus As Byte = CtrlFocus.Tree
    Public IncludeRemovables As Boolean = False
#Region "Literals"
    Public Property ZoneSize As Decimal = 0.4
    Public Const OrientationId As Integer = &H112
    Public iSSpeeds() As Integer = {1000, 300, 200}
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
            .Add("VertSplit" & "$" & MainForm.ctrFileBoxes.SplitterDistance)
            .Add("HorSplit" & "$" & MainForm.ctrMainFrame.SplitterDistance)
            .Add("File" & "$" & Media.MediaPath)
            .Add("Filter" & "$" & MainForm.CurrentFilterState.State)
            .Add("SortOrder" & "$" & MainForm.PlayOrder.State)
            .Add("StartPoint" & "$" & Media.SPT.State)
            .Add("State" & "$" & MainForm.NavigateMoveState.State)
            .Add("LastButtonFile" & "$" & ButtonFilePath)
            .Add("LastAlpha" & "$" & iCurrentAlpha)
            .Add("Favourites" & "$" & CurrentFavesPath)
            .Add("PreviewLinks" & "$" & MainForm.chbPreviewLinks.Checked)
            .Add("RootScanPath" & "$" & Rootpath)
            .Add("Directories List" & "$" & DirectoriesListFile) 'DirectoriesListFile)
            .Add("GlobalFaves" & "$" & GlobalFavesPath)
            .Add("RandomNextFile" & "$" & MainForm.chbNextFile.Checked)
            .Add("RandomOnDirectoryChange" & "$" & MainForm.chbOnDir.Checked)
            .Add("RandomAutoTrail" & "$" & MainForm.chbAutoTrail.Checked)
            .Add("RandomAutoLoadButtons" & "$" & MainForm.chbLoadButtonFiles.Checked)
            .Add("OptionsShowAttr" & "$" & MainForm.chbShowAttr.Checked)
            .Add("OptionsPreviewLinks" & "$" & MainForm.chbPreviewLinks.Checked)
            .Add("OptionsEncrypt" & "$" & MainForm.chbEncrypt.Checked)
            .Add("OptionsAutoAdvance" & "$" & MainForm.CHBAutoAdvance.Checked)
            .Add("OptionsSeparate" & "$" & MainForm.chbSeparate.Checked)
            .Add("ThumbnailDestination" & "$" & ThumbDestination)
            .Add("ButtonFolder" & "$" & Buttonfolder)
            .Add("FractionalJump" & "$" & MainForm.SP.FractionalJump)
            .Add("AbsoluteJump" & "$" & MainForm.SP.AbsoluteJump)
            .Add("Speed" & "$" & MainForm.SP.Speed)
            .Add("Navmode" & "$" & MainForm.NavigateMoveState.State)

        End With

        Dim f As New IO.FileInfo(PrefsFilePath)
        If f.Exists = False Then
        Else
            f.Delete()
            ' PrefsFilePath = PrefsFilePath.Replace(".", Str(Int(Rnd() * 1000)) & ".")
        End If

        WriteListToFile(PrefsList, PrefsFilePath, False)


    End Sub


    Public Sub PreferencesGet()
        If FirstLoad Then
            ResetDefaultPrefs()
            Exit Sub
        End If
        Dim prefslist As New List(Of String)
        Dim f As New IO.FileInfo(PrefsFilePath)
        Dim prefs As New IO.DirectoryInfo(PrefsPath)

        If prefs.Exists = False Then
            InitialiseFolders()
        End If

        For Each m In prefs.GetFiles
            Dim drives As String = DrivesScan()
            If m.CreationTimeUtc > f.CreationTimeUtc AndAlso m.FullName.Contains(drives + "MV") Then
                f = m
            End If
        Next
        If f.Exists Then
            prefslist = ReadListfromFile(f.FullName, False)
            MainForm.ctrPicAndButtons.SplitterDistance = 8.7 * MainForm.ctrPicAndButtons.Height / 10
            For Each s In prefslist
                If InStr(s, "$") <> 0 Then

                    Dim name As String = Split(s, "$")(0)
                    Dim value As String = Split(s, "$")(1)
                    Select Case name
                        Case "LastTimeSuccessful"
                            LastTimeSuccessful = value
                        Case "VertSplit"
                            MainForm.ctrFileBoxes.SplitterDistance = value
                        Case "HorSplit"
                            MainForm.ctrMainFrame.SplitterDistance = value
                        Case "File" 'Drive
                            If LastTimeSuccessful Then
                                Dim st As String = value 'TODO: Don't load if last time unsuccessful
                                If st = "" Then st = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                                Media.MediaPath = st
                                DriveString = GetDriveString(st)
                            End If
                        Case "Filter"
                            If value = "" Then value = 0
                            MainForm.CurrentFilterState.State = value

                        Case "SortOrder"
                            If value = "" Then value = 0
                            MainForm.PlayOrder.State = value
                        Case "StartPoint"
                            If value = "" Then value = 0
                            Media.SPT.State = value
                        Case "State"
                            If value = "" Then value = 0
                            MainForm.NavigateMoveState.State = value
                        Case "LastButtonFile" 'Drive
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
                            MainForm.chbPreviewLinks.Checked = value
                        Case "RootScanPath" 'Drive
                            If value = "" Then value = BrowseToFolder("Choose root scanning path", PrefsPath)
                            Rootpath = value
                            DriveString = GetDriveString(value)
                        Case "Directories List" 'Drive
                            If value = "" Then value = PrefsPath & "Directories.txt"

                            'DirectoriesListFile = value
                            'Dim ff As New IO.FileInfo(value)
                            'If ff.Exists = False Then
                            '    DirectoriesList = GetDirectoriesList(Rootpath, True)
                            'Else
                            '    DirectoriesList = GetDirectoriesList(value)


                            'End If
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
                            MainForm.chbLoadButtonFiles.Checked = value

                        Case "RandomNextFile"
                            If value = "" Then value = False

                            MainForm.chbNextFile.Checked = value
                        Case "RandomOnDirectoryChange"
                            If value = "" Then value = False
                            MainForm.chbOnDir.Checked = value
                        Case "RandomAutoTrail"
                            Exit Select
                            value = False 'TODO causing load problems when omitted
                            If value = "" Then value = False
                            MainForm.chbAutoTrail.Checked = value
                        Case "RandomAutoLoadButtons"
                            If value = "" Then value = False
                            MainForm.chbLoadButtonFiles.Checked = value
                        Case "OptionsShowAttr"
                            If value = "" Then value = False
                            MainForm.chbShowAttr.Checked = value
                        Case "OptionsPreviewLinks"
                            If value = "" Then value = False
                            MainForm.chbPreviewLinks.Checked = value
                        Case "OptionsEncrypt"
                            If value = "" Then value = False
                            MainForm.chbEncrypt.Checked = value
                        Case "OptionsAutoAdvance"
                            If value = "" Then value = False
                            MainForm.CHBAutoAdvance.Checked = value
                        Case "OptionsSeparate"
                            If value = "" Then value = False
                            MainForm.chbSeparate.Checked = value

                        Case "ThumbnailDestination" 'Drive
                            If value = "" Then value = BrowseToFolder("Choose Thumbnail Destination", PrefsPath)
                            ThumbDestination = value
                            DriveString = GetDriveString(value)

                        Case "ButtonFolder" 'Drive
                            If value = "" Then value = BrowseToFolder("Choose Button folder", PrefsPath)
                            Buttonfolder = value
                            DriveString = GetDriveString(value)
                        Case "AbsoluteJump"
                            MainForm.SP.AbsoluteJump = value
                        Case "FractionalJump"
                            MainForm.SP.FractionalJump = value
                        Case "Speed"
                            MainForm.SP.Speed = value
                        Case "Navstate"
                            MainForm.NavigateMoveState.State = value
                    End Select

                End If
            Next
        Else
            PreferencesReset(False)
        End If
        MainForm.tssMoveCopy.Text = CurrentFolder
        'MsgBox(DriveString)
    End Sub
    Private Function GetDriveString(path As String) As String
        'Adds drive of path if not already present

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
        With MainForm
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

    Public Sub LoadForm()
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


End Module
