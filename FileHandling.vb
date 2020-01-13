﻿Imports System.IO
Imports System.Threading
Imports MasaSam.Forms.Controls
Friend Module FileHandling

    '  Public WithEvents StartPoint As New StartPointHandler

    Public CurrentfilterState As FilterHandler = MainForm.CurrentFilterState
    Public Random As RandomHandler = MainForm.Random
    Public NavigateMoveState As StateHandler = MainForm.NavigateMoveState
    ' Public FP As New FilePump
    Public Listbox As ListBox = MainForm.lbxFiles
    Public Event FolderMoved(Path As String)
    Public Event FileMoved(Files As List(Of String), lbx As ListBox)
    Public WithEvents t As Thread
    Public WithEvents Media As New MediaHandler("Media")
    Public WithEvents MSFiles As New MediaSwapper(MainForm.MainWMP1, MainForm.MainWMP2, MainForm.MainWMP3, MainForm.PictureBox1, MainForm.PictureBox2, MainForm.PictureBox3)
    '   Public WithEvents MSShow As New MovieSwapper(MainForm.MainWMP, MainForm.MainWMP2)
    Public AllFaveMinder As New FavouritesMinder(GlobalFavesPath)
    Public FaveMinder As New FavouritesMinder(CurrentFavesPath)
    'Public AllFaveMinder As New FavouritesMinder("Q:\Favourites")
    'Public FaveMinder As New FavouritesMinder("Q:\Favourites")
    '
    '  Public WithEvents SndH As New SoundController
    Public Sub OnMediaPlaying(sender As Object, e As EventArgs) Handles Media.MediaPlaying


    End Sub

    Public Sub OnMediaStartChanged(sender As Object, e As EventArgs) Handles Media.StartChanged
        MainForm.OnStartChanged(sender, e)

    End Sub
    Public Sub OnMediaSpeedChanged(sender As Object, e As EventArgs) Handles Media.SpeedChanged
        MainForm.OnSpeedChange(sender, e)
    End Sub
    Public Sub OnMediaShown(M As MediaHandler) Handles MSFiles.MediaShown
        Media = M
        MainForm.UpdateFileInfo()
        If M.MediaType <> Filetype.Movie Then
            currentPicBox = M.Picture
        ElseIf M.MediaType = Filetype.Movie Then
            ' MainForm.emblem.ImageLocation = MainForm.VT.GetThumbnail(M.MediaPath, M.Position)
            'Deletefile(MainForm.emblem.ImageLocation)

        End If
        If M.IsLink Then
            MainForm.PopulateLinkList(M.LinkPath, M)
        Else
            MainForm.PopulateLinkList(M.MediaPath, M)
        End If
        ' MainForm.AT.AdvanceChance = M.Markers.Count + 1
        'Media.IsCurrent = True
        Media.SetLink(0)

        If M.MediaPath <> "" Then My.Computer.Registry.CurrentUser.SetValue("File", M.MediaPath)
        If ShiftDown Then MainForm.HighlightCurrent(Media.LinkPath) 'Used for links only, to go to original file
        If MainForm.FocusControl Is MainForm.lbxShowList Then MainForm.HighlightCurrent(Media.MediaPath)
    End Sub
    Private Sub DebugStartpoint(M As MediaHandler)
        Debug.Print(M.MediaPath & " loaded into " & M.Player.Name)
        Debug.Print(M.StartPoint.StartPoint & " startpoint")
        Debug.Print(M.StartPoint.State & " State")
        Debug.Print(M.StartPoint.Duration & " Duration")
        Debug.Print("")
    End Sub
    Public Sub OnFilesMoved(files As List(Of String), lbx1 As ListBox)
        ' Exit Sub
        lbx1.SelectionMode = SelectionMode.One
        Dim ind As Long = lbx1.SelectedIndex
        For Each f In files
            Select Case NavigateMoveState.State
                Case StateHandler.StateOptions.Copy, StateHandler.StateOptions.CopyLink
                Case StateHandler.StateOptions.MoveLeavingLink
                    MainForm.UpdatePlayOrder(MainForm.FBH)
                    ReplaceListboxItem(lbx1, ind, f)
                    lbx1.SelectedItem = lbx1.Items(ind)

                    RefreshListbox(lbx1, files)
                Case Else
                    RefreshListbox(lbx1, files)
            End Select
            MSFiles.ResettersOff()
        Next

        If lbx1.Items.Count <> 0 Then lbx1.SetSelected(Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0), True)


    End Sub
    Public strFilterExtensions(6) As String
    'Public Sub AssignExtensionFilters()
    '    'strFilterExtensions(FilterHandler.FilterState.All) = ""
    '    'strFilterExtensions(FilterHandler.FilterState.Piconly) = PICEXTENSIONS
    '    'strFilterExtensions(FilterHandler.FilterState.PicVid) = PICEXTENSIONS & VIDEOEXTENSIONS
    '    'strFilterExtensions(FilterHandler.FilterState.LinkOnly) = ".lnk"
    '    'strFilterExtensions(FilterHandler.FilterState.Vidonly) = VIDEOEXTENSIONS
    '    'strFilterExtensions(FilterHandler.FilterState.NoPicVid) = PICEXTENSIONS & VIDEOEXTENSIONS & "NOT"
    'End Sub
    ''' <summary>
    ''' Create a list from e and filter it
    ''' </summary>
    ''' <param name="e"></param>
    ''' <param name="lst"></param>
    ''' <returns></returns>
    'Public Function FilterLBList(e As DirectoryInfo, ByRef lst As List(Of String)) As List(Of String)
    '    lst.Clear()
    '    For Each f In e.EnumerateFiles
    '        lst.Add(f.FullName)
    '    Next
    '    MainForm.CurrentFilterState.FileList = lst
    '    Return MainForm.CurrentFilterState.FileList
    '    Exit Function

    'End Function

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
        For Each s In list
            Try
                Dim f As New FileInfo(s)
                If f.Exists Then
                    lbx.Items.Add(s)
                Else
                    notlist.Add(s)
                End If
            Catch ex As System.IO.PathTooLongException
                Continue For
            Catch ex As System.ArgumentException
                ReportFault("Filehandling.Getlist", ex.Message)
                Exit Sub
            End Try
            ProgressIncrement(40)
        Next



        If lbx.Items.Count <> 0 Then lbx.TabStop = True
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
        MoveFiles(flist, newpath, Listbox, True)

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
    Public Sub MoveFiles(files As List(Of String), strDest As String, lbx1 As ListBox, Optional Folder As Boolean = False)
        Dim Fileundo As New Undo
        With Fileundo
            .FileList = files
            .Action = Undo.Functions.MoveFiles
            .Destination = strDest

        End With
        UndoOperations.Push(Fileundo)


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

        If Folder Then

        Else

            RaiseEvent FileMoved(files, lbx1)

        End If
    End Sub
    Public Sub MoveFiles(file As String, strDest As String, lbx1 As ListBox, Optional Folder As Boolean = False)
        Dim files As New List(Of String)
        files.Add(file)
        MoveFiles(files, strDest, lbx1, Folder)

    End Sub

    Private Sub MovingFiles(files As List(Of String), strDest As String, s As String)
        If strDest <> "" Then
            Dim dinfo As New IO.DirectoryInfo(strDest)
            If dinfo.Exists = False Then dinfo.Create()
        End If
        Dim file As String = ""

        For Each file In files
            '   If Media.Player.URL = file Then Media.Player.URL = ""


            Dim m As New IO.FileInfo(file)
            With My.Computer.FileSystem
                Dim i As Long = 0
                Dim spath As String
                If s.EndsWith("\") Or s = "" Then
                    spath = s & m.Name

                Else
                    spath = s & "\" & m.Name

                End If
                While .FileExists(spath) AndAlso m.Name = FilenameFromPath(spath, True) 'Existing path - rename
                    Dim x = m.Extension
                    Dim b = InStr(spath, "(")
                    If b = 0 Then
                        spath = Replace(spath, x, "(" & i & ")" & x)
                    Else
                        spath = Left(spath, b - 1) & "(" & i & ")" & x
                    End If

                    i += 1
                End While
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy
                        If strDest = "" Then
                            Dim f As New IO.FileInfo(m.FullName)
                            AllFaveMinder.DestinationPath = strDest
                            AllFaveMinder.CheckFile(f)
                            If AllFaveMinder.OkToDelete Then
                                'AllFaveMinder.DeleteFavourite(m.FullName)
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
                                'AllFaveMinder.DeleteFavourite(m.FullName)
                                Deletefile(m.FullName)
                            End If
                        Else
                            Dim f As New IO.FileInfo(m.FullName)
                            AllFaveMinder.DestinationPath = spath
                            Try
                                m.MoveTo(spath)
                                'Dim fil2 As New IO.FileInfo(spath)
                                'If fil2.Exists Then m.Delete()
                                AllFaveMinder.CheckFile(f)
                                'AssignSpecialButton("0", strDest)
                            Catch ex As System.IO.IOException
                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
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

                        If m.Extension = ".lnk" Then
                            Dim f As New IO.FileInfo(LinkTarget(m.FullName))
                            spath = f.FullName
                            f.MoveTo(CurrentFolder & "\" & f.Name)

                            CreateLink(sh, f.FullName, New FileInfo(spath).Directory.FullName, m.Name, Media.Bookmark)
                            m.Delete()
                        End If

                    Case StateHandler.StateOptions.MoveOriginal
                        'Only works on links
                        Dim sh As New ShortcutHandler

                        If m.Extension = ".lnk" Then
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
    Public Sub AddFilesToCollectionSingle(ByRef list As List(Of String), extensions As String, blnRecurse As Boolean)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolder)

        s = InputBox("Only include files containing? (Leave empty to add all)")
        If s = "" Then s = "*"
        MainForm.Cursor = Cursors.WaitCursor
        ProgressBarOn(1000)
        list = GetFileFromEachFolder(d, s)
        MainForm.Cursor = Cursors.Default

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
    Public Function GetFileFromEachFolder(d As DirectoryInfo, s As String) As List(Of String)

        Dim x As New List(Of String)
        For Each Di In d.EnumerateDirectories(s, SearchOption.AllDirectories)
            If Di.GetFiles.Count > 0 Then
                x.Add(Di.EnumerateFiles.First.FullName)
            End If
        Next
        Return x
    End Function
    Public Sub AddSingleFileToList(ByVal list As List(Of String), strPath As String)
        list.Add(strPath)

    End Sub
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
    Public Function ListSubFiles(strPath As String) As List(Of String)
        Dim dir As New DirectoryInfo(strPath)
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
            If f.FullName.Contains(strSearch) Or strSearch = "*" Then
                x.Add(f.FullName)
            End If
        Next

        If blnRecurse Then
            For Each f In d.EnumerateDirectories("*", SearchOption.AllDirectories)
                Try
                    For Each fil In f.EnumerateFiles()
                        If fil.FullName.Contains(strSearch) Or strSearch = "*" Then
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

    Private Sub AddRemove(list As List(Of String), blnRemove As Boolean, strSearch As String, file As FileInfo)
        If InStr(LCase(file.FullName), LCase(strSearch)) <> 0 Or strSearch = "" Then
            If blnRemove Then
                list.Remove(file.FullName)
            Else
                list.Add(file.FullName)
            End If
        End If
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
        Exit Function
        'If blnRecurse Then
        '    Try
        '        For Each di In d.EnumerateDirectories

        '            DeleteEmptyFolders(di, True)
        '        Next
        '    Catch
        '    End Try
        'End If
        'If d.EnumerateDirectories.Count = 0 And d.EnumerateFiles.Count = 0 Then
        '    Dim s As String = d.Parent.FullName
        '    Try
        '        MainForm.tvMain2.RemoveNode(d.FullName)
        '        My.Computer.FileSystem.DeleteDirectory(d.FullName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
        '        DirectoriesList.Remove(d.FullName)

        '    Catch ex As Exception
        '        Return False
        '        Exit Function
        '    End Try
        '    ' MainForm.tvMain2.RefreshTree(s)
        'End If


        'Return True
    End Function
    'Public Function FolderCount(d As DirectoryInfo, count As Integer, blnRecurse As Boolean) As Long
    '    Try
    '        count = count + d.EnumerateDirectories.Count

    '    Catch ex As UnauthorizedAccessException
    '        Return 0
    '        Exit Function
    '    End Try
    '    If blnRecurse Then
    '        For Each di In d.EnumerateDirectories
    '            count = FolderCount(di, count, True)
    '        Next
    '    End If
    '    Return count
    'End Function
    'Public Function FileCount(d As DirectoryInfo, count As Integer, blnRecurse As Boolean) As Long
    '    Try
    '        count = count + d.EnumerateFiles.Count


    '    Catch ex As Exception
    '        Return 0
    '        Exit Function
    '    End Try
    '    If blnRecurse Then
    '        For Each di In d.EnumerateDirectories
    '            count = FileCount(di, count, True)
    '        Next
    '    End If
    '    Return count
    'End Function
    Public Async Function HarvestBelow(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(MainForm.tvMain2, MainForm.lbxFiles, d.FullName)
        Await x.HarvestBelow(25)
        'For Each di In d.EnumerateDirectories
        '    BurstFolder(di)

        'Next
    End Function
    Public Async Function HarvestFolder(d As DirectoryInfo, Recurse As Boolean, Parent As Boolean) As Task
        If Recurse Then
            Try
                For Each di In d.EnumerateDirectories
                    Await HarvestFolder(di, Recurse, Parent)
                Next
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End If
        blnSuppressCreate = True
        Dim l As New List(Of String)
        For Each f In d.EnumerateFiles
            l.Add(f.FullName)
        Next
        If Parent Then
            MoveFiles(l, d.Parent.FullName, MainForm.lbxFiles)
        Else
            MoveFiles(l, CurrentFolder, MainForm.lbxFiles)

            blnSuppressCreate = False
        End If

        'For Each f In d.EnumerateFiles
        '    If Parent Then
        '        Dim m As String = d.Parent.FullName & "\" & f.Name
        '        Dim fi As New FileInfo(m)
        '        If fi.Exists Then
        '        Else
        '            'Use an encapsulated move routine
        '            f.MoveTo(m)
        '        End If
        '    Else
        '        f.MoveTo(CurrentFolder & "\" & f.Name)
        '    End If
        'Next
        Await DeleteEmptyFolders(d, True)
        RaiseEvent FolderMoved(d.FullName)
    End Function
    Public Async Function BurstFolder(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(MainForm.tvMain2, MainForm.lbxFiles, d.FullName)


        Await x.Burst() 'Needs Attention
        '        HarvestFolder(d, True, True)

    End Function

    Public Sub PromoteFile(path As String)

        Dim finfo As New IO.FileInfo(path)
        finfo.MoveTo(finfo.Directory.Parent.FullName & "\" & finfo.Name)
    End Sub

End Module
