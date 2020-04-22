Imports System.IO
Imports System.Threading
Imports MasaSam.Forms.Controls
Friend Module FileHandling

    '  Public WithEvents StartPoint As New StartPointHandler

    Public CurrentfilterState As FilterHandler = MainForm.CurrentFilterState
    Public Random As RandomHandler = MainForm.Random
    Public NavigateMoveState As StateHandler = MainForm.NavigateMoveState
    ' Public FP As New FilePump
    Public Event FolderMoved(Path As String)
    Public Event FileMoved(Files As List(Of String), lbx As ListBox)
    Public WithEvents t As Thread
    Public WithEvents Media As New MediaHandler("Media")
    Public WithEvents MSFiles As New MediaSwapper(MainForm.MainWMP1, MainForm.MainWMP2, MainForm.MainWMP3, MainForm.PictureBox1, MainForm.PictureBox2, MainForm.PictureBox3)
    Public AllFaveMinder As New FavouritesMinder("Q:\Favourites")
    Public FaveMinder As New FavouritesMinder("Q:\Favourites")

    Public Sub OnZoomChanged(sender As Object, e As EventArgs) Handles Media.Zoomchanged
        MSFiles.ZoomPics(Media.PicHandler.ZoomFactor)

    End Sub
    Public Sub OnMediaFinished(sender As Object, e As EventArgs) Handles Media.MediaFinished
        MainForm.AdvanceFile(True, MainForm.Random.NextSelect)
        'Media.Playing= False
    End Sub
    '  Public WithEvents SndH As New SoundController
    Public Sub OnMediaPlaying(sender As Object, e As EventArgs) Handles Media.MediaPlaying


    End Sub

    Public Sub OnMediaStartChanged(sender As Object, e As EventArgs) Handles Media.StartChanged
        MainForm.OnStartChanged(sender, e)

    End Sub
    Public Sub OnMediaSpeedChanged(sender As Object, e As EventArgs) Handles Media.SpeedChanged
        MainForm.OnSpeedChange(sender, e)
    End Sub
    Public Sub OnMediaShown(sender As Object, e As EventArgs) Handles MSFiles.MediaShown

        Media = sender
        MainForm.UpdateFileInfo()
        If Media.MediaType = Filetype.Movie Then
            MainForm.PopulateLinkList(sender)
        End If
        Media.SetLink(0)
        MainForm.AT.Counter = Media.Markers.Count
        'If sender.MediaPath <> "" Then
        '    LastTimeSuccessful = True
        '    PreferencesSave()
        'End If
        If ShiftDown Then MainForm.HighlightCurrent(Media.LinkPath) 'Used for links only, to go to original file


    End Sub
    Private Sub DebugStartpoint(M As MediaHandler)
        Debug.Print(M.MediaPath & " loaded into " & M.Player.Name)
        Debug.Print(M.SPT.StartPoint & " startpoint")
        Debug.Print(M.SPT.State & " State")
        Debug.Print(M.SPT.Duration & " Duration")
        Debug.Print("")
    End Sub
    Public Sub OnFilesMoved(files As List(Of String), lbx1 As ListBox)
        lbx1.SelectionMode = SelectionMode.One
        Dim ind As Long = lbx1.SelectedIndex
        For Each f In files
            Select Case NavigateMoveState.State
                Case StateHandler.StateOptions.Copy, StateHandler.StateOptions.CopyLink
                Case StateHandler.StateOptions.MoveLeavingLink
                    ReplaceListboxItem(lbx1, ind, f)
                    lbx1.SelectedItem = lbx1.Items(ind)
                    MainForm.UpdatePlayOrder(MainForm.FBH)

                Case Else
                    If MainForm.LBH.ListBox IsNot Nothing Then
                        MainForm.LBH.FillBox()
                    End If
                    ' MainForm.FBH.Refresh()
                    ' RefreshListbox(lbx1, files)
            End Select
            '            MSFiles.ResettersOff()
        Next
        'MainForm.FBH.SetIndex(Math.Max(Math.Min(ind, MainForm.FBH.ListBox.Items.Count - 1), 0))
        If lbx1.Items.Count <> 0 Then
            Dim index = Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0)
            MainForm.FBH.SetIndex(index, True)
            '            lbx1.SetSelected(0,)
            'lbx1.SetSelected(index, True)
        End If


    End Sub
    Public strFilterExtensions(6) As String


    ''' <summary>
    ''' Loads Dest into List, and adds all to lbx. Any files not found are put in notlist, which can then be removed from the lbx
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="Dest"></param>
    ''' <param name="lbx"></param>
    '''
    Public Sub Getlist(list As List(Of String), Dest As String, lbx As ListBox)

        Dim notlist As New List(Of String)
        ReadListfromFile(list, Dest, Encrypted)
        lbx.DataSource = list
        'For Each s In list
        '    Try
        '        Dim f As New FileInfo(s)
        '        If f.Exists Then
        '            lbx.Items.Add(s)
        '        Else
        '            notlist.Add(s)
        '        End If
        '    Catch ex As System.IO.PathTooLongException
        '        Continue For
        '    Catch ex As System.ArgumentException
        '        ReportFault("Filehandling.Getlist", ex.Message)
        '        Exit Sub
        '    End Try
        '    ProgressIncrement(40)
        'Next



        'If lbx.Items.Count <> 0 Then lbx.TabStop = True
        lngShowlistLines = Showlist.Count

        If notlist.Count = 0 Then Exit Sub
        If MsgBox(notlist.Count & " files were not found. Remove from list?", vbYesNo, "Metavisua") = MsgBoxResult.Yes Then
            For Each s In notlist
                list.Remove(s)
            Next
            If MsgBox("Re-save list?", vbYesNo, "Metavisua") Then
                WriteListToFile(list, Dest, Encrypted)
            End If
        End If
    End Sub
    Public Sub LevelAllFolders()
        Dim x As New DirectoryInfo(CurrentFolder)
        For Each m In x.EnumerateDirectories("*", SearchOption.AllDirectories)
            If m.Parent.FullName <> CurrentFolder Then
                MoveFolder(m.FullName, CurrentFolder)
            End If
        Next
    End Sub
    Public Sub MoveFolder(Dir As String, Dest As String)
        If Dest = "" Then
            Dim SourceDir As New DirectoryInfo(Dir)
            DirectoriesList.Remove(SourceDir.FullName)
            For Each d In SourceDir.EnumerateDirectories("*", SearchOption.AllDirectories)
                MoveDirectoryContents(SourceDir, True)
            Next
            MoveDirectoryContents(SourceDir, True)
        Else


            Dim TargetDir As New DirectoryInfo(Dest)
            Dim SourceDir As New DirectoryInfo(Dir)
            DirectoriesList.Remove(SourceDir.FullName)
            DirectoriesList.Add(TargetDir.FullName)

            'Make target subdirectories.
            MoveDirectoryContents(TargetDir, SourceDir, SourceDir, True)
            For Each d In SourceDir.EnumerateDirectories("*", SearchOption.AllDirectories)
                MoveDirectoryContents(TargetDir, SourceDir, d, True)
            Next
            If SourceDir.GetFiles.Count = 0 And SourceDir.GetDirectories.Count = 0 Then
                SourceDir.Delete()
            End If
        End If

    End Sub

    Private Sub MoveDirectoryContents(TargetDir As DirectoryInfo, SourceDir As DirectoryInfo, d As DirectoryInfo, Optional Parent As Boolean = False)
        Dim flist As New List(Of String)
        Dim newpath As String = d.FullName
        If Parent Then
            newpath = newpath.Replace(SourceDir.Parent.FullName, TargetDir.FullName)
        Else
            newpath = newpath.Replace(SourceDir.FullName, TargetDir.FullName)

        End If
        Dim NewDir As New DirectoryInfo(newpath)
        If NewDir.Exists = False Then
            NewDir.Create()
        End If
        For Each f In d.EnumerateFiles("*", SearchOption.TopDirectoryOnly)
            flist.Add(f.FullName)
        Next
        blnSuppressCreate = True
        MoveFiles(flist, newpath, True)

    End Sub

    Private Sub MoveDirectoryContents(SourceDir As DirectoryInfo, Optional Parent As Boolean = False)
        Dim flist As New List(Of String)
        For Each f In SourceDir.EnumerateFiles("*", SearchOption.TopDirectoryOnly)
            flist.Add(f.FullName)
        Next
        blnSuppressCreate = True
        MoveFiles(flist, "", True)

    End Sub
    'Private Sub GetFiles(dir As DirectoryInfo, flist As List(Of String))
    '    For Each m In dir.EnumerateFiles("*", SearchOption.AllDirectories)
    '        flist.Add(m.FullName)
    '    Next
    '    Exit Sub
    '    For Each m In dir.EnumerateFiles()
    '        flist.Add(m.FullName)
    '    Next
    '    For Each x In dir.EnumerateDirectories
    '        GetFiles(x, flist)
    '    Next
    'End Sub


    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file. If strDest is empty, files are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>
    ''' <param name="lbx1"></param>
    Public Sub MoveFiles(files As List(Of String), strDest As String, Optional Folder As Boolean = False)
        'Dim Fileundo As New Undo
        'With Fileundo
        '    .FileList = files
        '    .Action = Undo.Functions.MoveFiles
        '    .Destination = strDest

        'End With
        'UndoOperations.Push(Fileundo)


        Dim s As String = strDest 'if strDest is empty then delete

        If files.Count > 0 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(MainForm.tvMain2, strDest, True) 'Attention

        End If
        Select Case NavigateMoveState.State
            Case StateHandler.StateOptions.Copy, StateHandler.StateOptions.CopyLink
            Case Else
                MainForm.CancelDisplay(False)
        End Select

        t = New Thread(New ThreadStart(Sub() MovingFiles(files, strDest, s)))


        Static i As Byte
        With buttons
            .CurrentSet.Last.Buttons(i).Path = strDest
            i = (i + 1) Mod 8
        End With
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()

        GC.Collect()

        If Folder Then

        Else

            RaiseEvent FileMoved(files, MainForm.FBH.ListBox)

        End If
    End Sub
    Private Sub Finishedmoving()

    End Sub
    'Public Sub MoveFiles(file As String, strDest As String, lbx1 As ListBox, Optional Folder As Boolean = False)
    '    Dim files As New List(Of String)
    '    files.Add(file)
    '    MoveFiles(files, strDest, Folder)

    'End Sub

    Private Sub MovingFiles(files As List(Of String), strDest As String, s As String)
        If strDest <> "" Then
            Dim dinfo As New IO.DirectoryInfo(strDest)
            If dinfo.Exists = False Then dinfo.Create()

        End If
        Dim file As String = ""

        Try
            For Each file In files
                '   If Media.Player.URL = file Then Media.Player.URL = ""


                Dim m As New IO.FileInfo(file)
                With My.Computer.FileSystem
                    'Dim i As Long = 0
                    Dim spath As String
                    If s.EndsWith("\") Or s = "" Then
                        spath = s & m.Name

                    Else
                        spath = s & "\" & m.Name

                    End If

                    Select Case NavigateMoveState.State
                        Case StateHandler.StateOptions.Copy
                            If strDest = "" Then
                                Dim f As New IO.FileInfo(m.FullName)
                                AllFaveMinder.DestinationPath = ""
                                AllFaveMinder.CheckFile(f)
                                If AllFaveMinder.OkToDelete Then

                                    Deletefile(m.FullName)
                                End If
                            Else
                                .CopyFile(m.FullName, spath)

                            End If
                        Case StateHandler.StateOptions.Move, StateHandler.StateOptions.Navigate
                            'If Not currentPicBox.Image Is Nothing Then DisposePic(currentPicBox)
                            If strDest = "" Then
                                Dim f As New IO.FileInfo(m.FullName)
                                AllFaveMinder.DestinationPath = strDest
                                AllFaveMinder.CheckFile(f)
                                If AllFaveMinder.OkToDelete Then

                                    Deletefile(m.FullName)
                                End If
                            Else
                                Dim f As New IO.FileInfo(m.FullName)
                                Dim destfile As New IO.FileInfo(spath)
                                'Check if destination same as origin
                                If f.FullName = destfile.FullName Then
                                Else
                                    spath = AppendExistingFilename(spath, m)

                                    If m.FullName <> spath Then
                                        Try
                                            AllFaveMinder.DestinationPath = spath
                                            AllFaveMinder.CheckFile(f)
                                            MSFiles.CancelURL(spath)
                                            BreakHere = True
                                            m.MoveTo(spath)


                                        Catch ex As Exception
                                            MsgBox(ex.Message)

                                        End Try
                                    End If
                                End If
                            End If
                        Case StateHandler.StateOptions.MoveLeavingLink
                            'Move, and place link here
                            Dim f As New IO.FileInfo(m.FullName)
                            AllFaveMinder.DestinationPath = spath
                            AllFaveMinder.CheckFile(f)
                            m.MoveTo(spath)
                            Dim sh As New ShortcutHandler
                            CreateLink(sh, m.FullName, CurrentFolder, f.Name, Bookmark:=Media.Position)
                            sh = Nothing
                        Case StateHandler.StateOptions.CopyLink
                            'Paste a link in destination
                            Dim fpath As New FileInfo(spath)
                            Dim sh As New ShortcutHandler

                            CreateLink(sh, m.FullName, fpath.Directory.FullName, m.Name, Bookmark:=Media.Position)

                        Case StateHandler.StateOptions.ExchangeLink
                            'Only works on links
                            Dim sh As New ShortcutHandler

                            If m.Extension = LinkExt Then
                                Dim f As New IO.FileInfo(LinkTarget(m.FullName))
                                spath = f.FullName
                                f.MoveTo(CurrentFolder & "\" & f.Name)

                                CreateLink(sh, f.FullName, New FileInfo(spath).Directory.FullName, m.Name, Media.Bookmark)
                                Deletefile(m.FullName)
                            End If

                        Case StateHandler.StateOptions.MoveOriginal
                            'Only works on links
                            Dim sh As New ShortcutHandler

                            If m.Extension = LinkExt Then
                                Dim f As New IO.FileInfo(LinkTarget(m.FullName))
                                ' spath = f.FullName
                                strDest = strDest & "\" & f.Name
                                Try
                                    f.MoveTo(strDest)

                                Catch ex As IO.IOException
                                    MsgBox("File already exists in destination")
                                End Try
                            End If



                    End Select
                End With

            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        'RaiseEvent FileMoved(files, MainForm.lbxFiles)
    End Sub
    ''' <summary>
    ''' Checks to see if f is in the favourite links, and if so, updates the link. 
    ''' </summary>
    ''' <param name="f"></param>
    ''' <param name="path"></param>

    Private Sub Movelink(f As IO.FileInfo, path As String)
        Dim links As New List(Of String)
        Dim faves As New IO.DirectoryInfo(CurrentFavesPath)
        For Each j In faves.GetFiles
            If links.Contains(j.FullName) Then
            Else
                links.Add(LinkTarget(j.FullName))
            End If
        Next
        If links.Contains(f.FullName) Then
            CreateFavourite(path)
        End If
    End Sub

    Public Function CreateNewDirectory(tv As FileSystemTree, strDest As String, blnAsk As Boolean) As String
        Dim blnCreate As Boolean = True
        Dim blnAssign As Boolean = False
        Dim s As String = ""
        If blnAsk Then
            s = InputBox("Name of folder to create? (Blank means none)", "Create sub-folder", lastselection)
            If s = "" Then blnCreate = False
            s = strDest & "\" & s
        Else
            s = strDest & "\"
        End If
        If blnCreate Then
            Try
                IO.Directory.CreateDirectory(s)
                DirectoriesList.Add(s)
                tv.RefreshTree(strDest)
            Catch ex As IO.DirectoryNotFoundException
            End Try

        End If
        Return s
    End Function
    Public Sub AddFilesToCollectionSingle(ByRef list As List(Of String), blnRecurse As Boolean)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolder)

        s = InputBox("Only include files containing? (Leave empty to add all)")
        If s = "" Then s = "*"
        MainForm.Cursor = Cursors.WaitCursor
        ProgressBarOn(1000)
        list = GetFileFromEachFolder(d, s, Random.OnDirChange)
        MainForm.Cursor = Cursors.Default
        ProgressBarOff()
    End Sub
    Public Function AddFilesToCollection(extensions As String, blnRecurse As Boolean) As List(Of String)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolder)
        Dim List As New List(Of String)
        s = InputBox("Only include files containing? (Leave empty to add all)")
        MainForm.Cursor = Cursors.WaitCursor
        ProgressBarOn(1000)
        List = FindAllFilesBelow(d, List, extensions, False, s, blnRecurse, blnChooseOne)
        MainForm.Cursor = Cursors.Default
        Return List
    End Function


    Public Sub Deletefile(s As String)


        With My.Computer.FileSystem
            If .FileExists(s) Then
                Try
                    .DeleteFile(s, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

                Catch ex As Exception

                End Try
            End If

        End With
    End Sub
    'Public Function CountSubFiles(strPath As String) As Integer
    '    Dim dir As New DirectoryInfo(strPath)
    '    Dim i As Integer
    '    If dir.EnumerateDirectories.Count <> 0 Then
    '        For Each f In dir.EnumerateDirectories
    '            i = i + CountSubFiles(f.FullName)
    '        Next
    '    End If
    '    i = i + dir.EnumerateFiles.Count
    '    Return i
    'End Function
    Public Function ListSubFiles(DirectoryPath As String) As List(Of String)
        Dim dir As New DirectoryInfo(DirectoryPath)
        Dim ls As New List(Of String)
        If dir.EnumerateDirectories.Count <> 0 Then
            For Each f In dir.EnumerateDirectories
                For Each n In f.EnumerateFiles
                    ls.Add(n.FullName)
                Next

            Next
        End If
        For Each n In dir.EnumerateFiles
            ls.Add(n.FullName)
        Next
        Return ls
    End Function

    Public Function FindAllFilesBelow(d As DirectoryInfo, list As List(Of String), extensions As String, blnRemove As Boolean, strSearch As String, blnRecurse As Boolean, blnOneOnly As Boolean) As List(Of String)
        Dim x As New List(Of String)
        If strSearch = "" Then strSearch = "*"
        For Each f In d.EnumerateFiles()
            If UCase(f.FullName).Contains(UCase(strSearch)) Or strSearch = "*" Then
                x.Add(f.FullName)
            End If
        Next

        If blnRecurse Then
            For Each f In d.EnumerateDirectories("*", SearchOption.AllDirectories)
                Try
                    For Each fil In f.EnumerateFiles()
                        If UCase(fil.FullName).Contains(UCase(strSearch)) Or strSearch = "*" Then
                            x.Add(fil.FullName)
                        End If
                    Next

                Catch ex As Exception

                End Try
            Next

        End If
        Return x
        Exit Function

    End Function
    Public Sub SubfolderNamedSame(fol As IO.DirectoryInfo)
        For Each m In fol.GetDirectories("*", SearchOption.AllDirectories)
            If m.Name = fol.Name Then
                MoveFolder(m.FullName, fol.Parent.FullName)
            End If
            SubfolderNamedSame(m)
        Next
    End Sub

    Public Sub FindAllFoldersBelow(d As DirectoryInfo, list As List(Of String), blnRecurse As Boolean, blnNonEmptyOnly As Boolean)


        For Each di In d.EnumerateDirectories
            Try
                If blnNonEmptyOnly AndAlso di.EnumerateFiles.Count > 0 Or Not blnNonEmptyOnly Then
                    list.Add(di.FullName)
                End If

            Catch ex As PathTooLongException
                ReportFault("FindAllFoldersBelow", ex.Message)
            Catch ex As System.UnauthorizedAccessException
                Continue For
            End Try

        Next
        If blnRecurse Then

            For Each di In d.EnumerateDirectories
                Try

                    FindAllFoldersBelow(di, list, blnRecurse, blnNonEmptyOnly)
                Catch ex As UnauthorizedAccessException
                    Continue For
                Catch ex As DirectoryNotFoundException
                    Continue For
                End Try
            Next
        End If
    End Sub
    Public Async Function DeleteEmptyFolders(d As DirectoryInfo, blnRecurse As Boolean) As Task(Of Boolean)
        Dim x As New BundleHandler(MainForm.tvMain2, MainForm.lbxFiles, d.FullName)
        Await x.RemoveEmptyFolders(x.Path, blnRecurse)
        Return True

    End Function

    Public Async Function HarvestBelow(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(MainForm.tvMain2, MainForm.lbxFiles, d.FullName)
        Await x.Burst(d, True)

    End Function
    'Public Async Function HarvestFolder(d As DirectoryInfo, Recurse As Boolean, Parent As Boolean) As Task
    '    If Recurse Then
    '        Try
    '            For Each di In d.EnumerateDirectories
    '                Await HarvestFolder(di, Recurse, Parent)
    '            Next
    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try
    '    End If
    '    blnSuppressCreate = True
    '    Dim l As New List(Of String)
    '    For Each f In d.EnumerateFiles
    '        l.Add(f.FullName)
    '    Next
    '    If Parent Then
    '        MoveFiles(l, d.Parent.FullName, MainForm.lbxFiles)
    '    Else
    '        MoveFiles(l, CurrentFolder, MainForm.lbxFiles)

    '        blnSuppressCreate = False
    '    End If

    '    Await DeleteEmptyFolders(d, True)
    '    RaiseEvent FolderMoved(d.FullName)
    'End Function
    Public Async Function BurstFolder(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(MainForm.tvMain2, MainForm.lbxFiles, d.FullName)


        Await x.Burst(d) 'Needs Attention
        '        HarvestFolder(d, True, True)

    End Function

    Public Sub PromoteFile(path As String)

        Dim finfo As New IO.FileInfo(path)
        finfo.MoveTo(finfo.Directory.Parent.FullName & "\" & finfo.Name)
    End Sub

End Module
