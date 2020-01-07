Imports System.Environment
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
        Dim PrefsList As New List(Of String)
        With PrefsList

            .Add("VertSplit" & "$" & MainForm.ctrFileBoxes.SplitterDistance)
            .Add("HorSplit" & "$" & MainForm.ctrMainFrame.SplitterDistance)
            .Add("File" & "$" & Media.MediaPath)
            .Add("Filter" & "$" & MainForm.CurrentFilterState.State)
            .Add("SortOrder" & "$" & MainForm.PlayOrder.State)
            .Add("StartPoint" & "$" & Media.StartPoint.State)
            .Add("State" & "$" & MainForm.NavigateMoveState.State)
            .Add("LastButtonFile" & "$" & ButtonFilePath)
            .Add("LastAlpha" & "$" & iCurrentAlpha)
            .Add("Favourites" & "$" & CurrentFavesPath)
            .Add("PreviewLinks" & "$" & MainForm.chbPreviewLinks.Checked)
            .Add("RootScanPath" & "$" & "Q:\") 'Rootpath)
            .Add("Directories List" & "$" & "Q:\Directories.txt") 'DirectoriesListFile)
            .Add("GlobalFaves" & "$" & GlobalFavesPath)
            .Add("RandomNextFile" & "$" & MainForm.chbNextFile.Checked)
            .Add("RandomOnDirectoryChange" & "$" & MainForm.chbOnDir.Checked)
            .Add("RandomAutoTrail" & "$" & MainForm.chbAutoTrail.Checked)
            .Add("RandomAutoLoadButtons" & "$" & MainForm.chbLoadButtonFiles.Checked)
            .Add("OptionsShowAttr" & "$" & MainForm.chbShowAttr.Checked)
            .Add("OptionsPreviewLinks" & "$" & MainForm.chbPreviewLinks.Checked)
            .Add("OptionsEncrypt" & "$" & MainForm.chbEncrypt.Checked)
            .Add("OptionsAutoAdvance" & "$" & MainForm.CHBAutoAdvance.Checked)




        End With
        Dim appdata As String = "C:\MVPrefs.txt" 'GetFolderPath(SpecialFolder.LocalApplicationData)
        WriteListToFile(PrefsList, appdata, True)


    End Sub


    Public Sub PreferencesGet()
        Dim prefslist As New List(Of String)
        ReadListfromFile(prefslist, "C:\MVPrefs.txt", True)
        MainForm.ctrPicAndButtons.SplitterDistance = 8.7 * MainForm.ctrPicAndButtons.Height / 10
        For Each s In prefslist
            If InStr(s, "$") <> 0 Then

                Dim name As String = Split(s, "$")(0)
                Dim value As String = Split(s, "$")(1)
                Select Case name
                    Case "VertSplit"
                        MainForm.ctrFileBoxes.SplitterDistance = value
                    Case "HorSplit"
                        MainForm.ctrMainFrame.SplitterDistance = value
                    Case "File"
                        Dim st As String = value
                        If st = "" Then st = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
                        Media.MediaPath = st
                    Case "Filter"
                        If value = "" Then value = 0
                        MainForm.CurrentFilterState.State = value

                    Case "SortOrder"
                        If value = "" Then value = 0
                        MainForm.PlayOrder.State = value
                    Case "StartPoint"
                        If value = "" Then value = 0
                        Media.StartPoint.State = value
                    Case "State"
                        If value = "" Then value = 0
                        MainForm.NavigateMoveState.State = value
                    Case "LastButtonFile"
                        If value = "" Then LoadButtonFileName("")
                        ButtonFilePath = value
                    Case "LastAlpha"
                        If value = "" Then value = 0
                        iCurrentAlpha = value
                    Case "Favourites"
                        If value = "" Then value = BrowseToFolder("Choose favourites path")
                        CurrentFavesPath = value
                    Case "PreviewLinks"
                        If value = "" Then value = False
                        MainForm.chbPreviewLinks.Checked = value
                    Case "RootScanPath"
                        If value = "" Then value = BrowseToFolder("Choose root scanning path")
                        Rootpath = value
                    Case "Directories List"
                        If value = "" Then value = "C:\Directories.txt"

                        DirectoriesListFile = value
                        GetDirectoriesList(value)
                    Case "GlobalFaves"
                        If value = "" Then value = BrowseToFolder("Choose global favourites path")
                        GlobalFavesPath = value
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
                End Select

            End If
        Next
        'With My.Computer.Registry.CurrentUser
        '    'Appearance
        '    'States
        '    'Files
        '    '   Dim fol As New IO.DirectoryInfo(CurrentFavesPath) 'TODO This whole thing is a mess
        '    'All .lnk files in this hierarchy get recognised and changed when files are moved. 

        '    Rootpath = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer)
        '    If fol.Exists = False Then
        '        MainForm.FavouritesFolderToolStripMenuItem.PerformClick()
        '    End If
        'End With
        ''Catch ex As Exception
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
