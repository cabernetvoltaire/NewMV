Imports System.IO
Imports System.Threading
Imports MasaSam.Forms.Controls


Friend Module FileHandling

    '  Public WithEvents StartPoint As New StartPointHandler

    Public CurrentfilterState As FilterHandler = FormMain.CurrentFilterState
    Public Random As RandomHandler = FormMain.Random
    Public NavigateMoveState As StateHandler = FormMain.NavigateMoveState
    ' Public FP As New FilePump
    Public Event FolderMoved(Source As String, Dest As String)
    Public Event FileMoved(Files As List(Of String), lbx As ListBox)

    Public WithEvents t As Thread
    Public WithEvents Media As New MediaHandler("Media")
    Public WithEvents MSFiles As New MediaSwapper(FormMain.MainWMP1, FormMain.MainWMP2, FormMain.MainWMP3, FormMain.PictureBox1, FormMain.PictureBox2, FormMain.PictureBox3, FormMain.WebView21, FormMain.WebView22, FormMain.WebView23)

    Public MSFilesNew As New MediaSwapper2
    Public AllFaveMinder As New FavouritesMinder("Q:\Favourites")
    Public FaveMinder As New FavouritesMinder("Q:\Favourites")
    Public Sub OnMediaPlaying(sender As Object, e As EventArgs) Handles Media.MediaPlaying

    End Sub
    Public Sub OnZoomChanged(sender As Object, e As EventArgs) Handles Media.Zoomchanged
        MSFiles.ZoomPics(sender)

        If Media.PicHandler.State = PictureHandler.Screenstate.Fitted Then
            FormMain.tbZoom.Text = "Fitted"
        Else
            FormMain.tbZoom.Text = "Zoom: " & Str(Media.PicHandler.ZoomFactor) & "%"
        End If
    End Sub
    Public Sub OnMediaFinished(sender As Object, e As EventArgs) Handles Media.MediaFinished
        FormMain.AdvanceFile(True, FormMain.Random.NextSelect)
        FormMain.AT.InitialMomentsCount = 3
        'Media.Playing= False
    End Sub
    '  Public WithEvents SndH As New SoundController


    Public Sub OnMediaStartChanged(sender As Object, e As EventArgs) Handles Media.StartChanged
        FormMain.OnStartChanged(sender, e)

    End Sub
    Public Sub OnMediaSpeedChanged(sender As Object, e As EventArgs) Handles Media.SpeedChanged
        FormMain.OnSpeedChange(sender, e)
    End Sub
    'Friend Sub SetDataformEntry()
    '    If Media.IsLink Then
    '        DatabaseForm.ShortFilepath = Path.GetFileName(Media.LinkPath)
    '    Else
    '        DatabaseForm.ShortFilepath = Path.GetFileName(Media.MediaPath)

    '    End If
    'End Sub
    Public Sub OnMediaShown(sender As Object, e As EventArgs) Handles MSFiles.MediaShown

        Media = sender
        FormMain.UpdateFileInfo()


        'Media.SetLink(0)
        With FormMain
            '.AT.Counter = Media.Markers.Count
            .Att.DestinationLabel = .lblAttributes
            If Not .tmrSlideShow.Enabled And .chbShowAttr.Checked Then
                .TextBox1.Text = Media.Metadata
            Else
                .Att.Text = ""
            End If
        End With

        If ShiftDown Then FormMain.HighlightCurrent(Media.LinkPath) 'Used for links only, to go to original file
        FormMain.Length.Value = Math.Min(Media.Duration / 60, 200)
        FormMain.LengthLabel.Text = New TimeSpan(0, 0, Media.Duration).ToString("hh\:mm\:ss")
        If Media.MediaType = Filetype.Movie Then
            FormMain.PopulateLinkList(sender)
        Else
            FormMain.Marks.Clear()
        End If
    End Sub
    Public Sub DebugStartpoint(M As MediaHandler)
        Debug.Print(M.MediaPath & " loaded into " & M.Player.Name)
        Debug.Print(M.SPT.Markers.Count & " markers")
        If M.SPT.Markers.Count > 0 Then Debug.Print("Current marker " & LongAsTimeCode(M.SPT.CurrentMarker))
        Debug.Print(M.SPT.Description & " State")
        Debug.Print(LongAsTimeCode(M.SPT.Duration) & " Duration")
        Debug.Print(LongAsTimeCode(M.SPT.StartPoint) & " startpoint")
        Debug.Print("")
    End Sub
    Public Sub ReplaceListboxItem(lbx As ListBox, index As Integer, newitem As String)
        lbx.Items.RemoveAt(index)
        lbx.Items.Insert(index, newitem)

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
                    FormMain.UpdatePlayOrder(FormMain.FBH)
            End Select
        Next
        'If ShowListVisible Then
        '    FormMain.LBH.FillBox()
        'End If
        'FormMain.FBH.FillBox()
        'DeleteEmptyFolders(New IO.DirectoryInfo(CurrentFolder), True)
        If lbx1.Items.Count <> 0 Then
            Dim index = Math.Max(Math.Min(ind, lbx1.Items.Count - 1), 0)
            FormMain.FBH.SetIndex(index, True)
            '            lbx1.SetSelected(0,)
            'lbx1.SetSelected(index, True)
        End If


    End Sub
    Public strFilterExtensions(6) As String




    Public Sub LevelAllFolders()
        Dim x As New DirectoryInfo(CurrentFolder)
        For Each m In x.EnumerateDirectories("*", SearchOption.AllDirectories)
            If m.Parent.FullName <> CurrentFolder Then
                MoveFolder(m.FullName, CurrentFolder)
            End If
        Next
    End Sub

    Public Sub MoveFolder(ByVal Dir As String, ByVal Dest As String)
        ' Cancel any related operations that might interfere with moving the directory
        MSFiles.CancelURL()
        ' Define DirectoryInfo objects for source and destination directories
        Dim sourceDir As New DirectoryInfo(Dir)

        Dim targetDir As New DirectoryInfo(Dest)

        ' Construct the new path where the source directory will be moved
        ' This creates a path like "C:\second\first"
        Dim newDestinationPath As String = Path.Combine(targetDir.FullName, sourceDir.Name)

        ' Check if the target "second\first" directory already exists
        If Directory.Exists(newDestinationPath) Then
            ' Handle the situation where the target directory already exists
            ' You might want to merge contents, or inform the user, etc.
            ' For simplicity, let's assume you want to delete it (be cautious with this approach)
            'SafeDeleteDir(New DirectoryInfo(newDestinationPath))
            '            Directory.Delete(newDestinationPath, True) ' True for recursive deletion
        End If

        ' Ensure the parent directory "C:\second" exists before moving
        If Not targetDir.Exists Then
            targetDir.Create()
        End If

        Try
            ' Move the directory
            Directory.Move(sourceDir.FullName, newDestinationPath)

        Catch ex As IOException
            ' Handle exceptions, such as permission issues, here
            Console.WriteLine($"An error occurred: {ex.Message}")
        End Try
    End Sub





    Private Sub MergeDirectoryContents(sourceDir As DirectoryInfo, targetDir As DirectoryInfo)
        ' Ensure the target directory exists
        If Not targetDir.Exists Then
            targetDir.Create()
        End If

        ' Copy each file to the target directory
        For Each file In sourceDir.GetFiles()
            Dim targetFilePath = Path.Combine(targetDir.FullName, file.Name)

            ' Handle file conflicts (e.g., overwrite, skip, or rename)
            If IO.File.Exists(targetFilePath) Then
                ' Example: Rename instead of overwrite
                targetFilePath = Path.Combine(targetDir.FullName, $"{Path.GetFileNameWithoutExtension(file.Name)}_{Guid.NewGuid()}{file.Extension}")
            End If

            file.CopyTo(targetFilePath, False)
        Next

        ' Recursively copy subdirectories
        For Each subdir In sourceDir.GetDirectories()
            Dim nextTargetSubdir = targetDir.CreateSubdirectory(subdir.Name)
            MergeDirectoryContents(subdir, nextTargetSubdir)
        Next
    End Sub


    Public Sub SafeDeleteDir(SourceDir As DirectoryInfo, Optional Parent As Boolean = False)
        ''To delete all the files in SourceDir
        'Dim flist As New List(Of String)
        'For Each f In SourceDir.EnumerateFiles("*", SearchOption.TopDirectoryOnly)
        '    flist.Add(f.FullName)
        'Next
        'blnSuppressCreate = True
        'MoveFiles(flist, "", True)
        Try

            My.Computer.FileSystem.DeleteDirectory(SourceDir.FullName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)

        Catch ex As Exception

        End Try
    End Sub


    ''' <summary>
    ''' Moves files to strDest, and removes them from lbx1 asking for a subfolder in destination if more than one file. If strDest is empty, files are deleted.
    ''' </summary>
    ''' <param name="files"></param>
    ''' <param name="strDest"></param>

    Public Sub MoveFiles(files As List(Of String), strDest As String, Optional Folder As Boolean = False)

        Dim s As String = strDest 'if strDest is empty then delete

        If files.Count > 0 And strDest <> "" Then
            If Not blnSuppressCreate Then s = CreateNewDirectory(FormMain.tvmain2, strDest, True) 'Attention

        End If
        Select Case NavigateMoveState.State
            Case StateHandler.StateOptions.Copy, StateHandler.StateOptions.CopyLink
            Case Else
                FormMain.CancelDisplay(False)
        End Select

        t = New Thread(New ThreadStart(Sub() MovingFiles(files, strDest, s)))
        t.IsBackground = True
        t.SetApartmentState(ApartmentState.STA)
        t.Start()




        GC.Collect()


        RaiseEvent FileMoved(files, FormMain.FBH.ListBox)

    End Sub


    'Public Sub MoveFiles(file As String, strDest As String, lbx1 As ListBox, Optional Folder As Boolean = False)
    '    Dim files As New List(Of String)
    '    files.Add(file)
    '    MoveFiles(files, strDest, Folder)

    'End Sub

    Private Sub MovingFiles(files As List(Of String), strDest As String, s As String)
        ' Ensure the destination directory exists
        If Not String.IsNullOrEmpty(strDest) Then
            Dim dinfo As New IO.DirectoryInfo(strDest)
            If Not dinfo.Exists Then dinfo.Create()
        End If

        For Each file In files
            Dim m As New IO.FileInfo(file)
            Dim spath As String = If(s.EndsWith("\") OrElse String.IsNullOrEmpty(s), s & m.Name, s & "\" & m.Name)

            Try
                Select Case NavigateMoveState.State
                    Case StateHandler.StateOptions.Copy
                        HandleCopyState(m, spath, strDest)
                    Case StateHandler.StateOptions.Move, StateHandler.StateOptions.Navigate
                        HandleMoveNavigateState(m, spath, strDest)
                    Case StateHandler.StateOptions.MoveLeavingLink
                        HandleMoveLeavingLinkState(m, spath)
                    Case StateHandler.StateOptions.CopyLink
                        HandleCopyLinkState(m, spath)
                    Case StateHandler.StateOptions.ExchangeLink
                        HandleExchangeLinkState(m, spath)
                    Case StateHandler.StateOptions.MoveOriginal
                        HandleMoveOriginalState(m, strDest)
                End Select
            Catch ex As Exception

                MsgBox(ex.Message)
            End Try
        Next

        ' Optionally, raise an event here to indicate that the file operation(s) have completed
        ' RaiseEvent FileMoved(files, MainForm.lbxFiles)
    End Sub


    Private Sub HandleCopyState(m As IO.FileInfo, spath As String, strDest As String)
        With My.Computer.FileSystem
            If String.IsNullOrEmpty(strDest) Then
                ' Custom logic for handling copying when destination is not specified
                AllFaveMinder.DestinationPath = ""
                AllFaveMinder.CheckFile(m)
                If AllFaveMinder.OkToDelete Then
                    Deletefile(m.FullName)
                End If
            Else
                .CopyFile(m.FullName, spath)
            End If
        End With
    End Sub

    Private Sub HandleMoveNavigateState(m As IO.FileInfo, spath As String, strDest As String)
        If String.IsNullOrEmpty(strDest) Then
            Deletefile(m.FullName)
        Else
            ' Ensure the destination is different from the source
            If Not m.FullName.Equals(spath, StringComparison.OrdinalIgnoreCase) Then
                spath = AppendExistingFilename(spath, m)
                If m.FullName <> spath Then
                    Try
                        AllFaveMinder.DestinationPath = spath
                        AllFaveMinder.CheckFile(m)
                        MSFiles.CancelURL(m.FullName)
                        MSFiles.CancelURL(spath)
                        m.MoveTo(spath)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If
        End If
    End Sub

    Private Sub HandleMoveLeavingLinkState(m As IO.FileInfo, spath As String)
        ' Move the file and place a link in the original location
        AllFaveMinder.DestinationPath = spath
        AllFaveMinder.CheckFile(m)
        m.MoveTo(spath)
        CreateShortcutToMovedFile(m, spath)
    End Sub

    Private Sub HandleCopyLinkState(m As IO.FileInfo, spath As String)
        ' Paste a link in the destination
        CreateShortcutToMovedFile(m, spath)
    End Sub

    Private Sub HandleExchangeLinkState(m As IO.FileInfo, spath As String)
        ' Only works on links
        If m.Extension.Equals(LinkExt, StringComparison.OrdinalIgnoreCase) Then
            Dim linkedFile As New IO.FileInfo(LinkTarget(m.FullName))
            linkedFile.MoveTo(CurrentFolder & "\" & linkedFile.Name)
            CreateShortcutToMovedFile(linkedFile, spath)
            Deletefile(m.FullName)
        End If
    End Sub

    Private Sub HandleMoveOriginalState(m As IO.FileInfo, strDest As String)
        ' Only works on links
        If m.Extension.Equals(LinkExt, StringComparison.OrdinalIgnoreCase) Then
            Dim linkedFile As New IO.FileInfo(LinkTarget(m.FullName))
            Dim destinationPath As String = IO.Path.Combine(strDest, linkedFile.Name)
            Try
                linkedFile.MoveTo(destinationPath)
            Catch ex As IO.IOException
                MsgBox("File already exists in destination")
            End Try
        End If
    End Sub

    Private Sub CreateShortcutToMovedFile(m As IO.FileInfo, spath As String)
        ' Placeholder for creating a shortcut. Implement the actual shortcut creation logic here.
        ' This might involve interacting with the Windows Shell or using a third-party library.
    End Sub

    Private Function AppendExistingFilename(spath As String, m As IO.FileInfo) As String
        ' Placeholder for logic to append to filename if it already exists in the destination.
        ' This function should return a modified path if the file exists, ensuring no overwrite occurs.
        Return spath ' Modify this to implement actual logic
    End Function

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
                'DirectoriesList.Add(s)
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
        list = GetFileFromEachFolder(d, s, Random.OnDirChange)
    End Sub
    Public Function AddFilesToCollection(extensions As String, blnRecurse As Boolean) As List(Of String)
        Dim s As String
        Dim d As New DirectoryInfo(CurrentFolder)
        Dim List As New List(Of String)
        s = InputBox("Only include files containing? (Leave empty to add all)")
        List = FindAllFilesBelow(d, List, extensions, False, s, blnRecurse, blnChooseOne)
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
    Public Event DeleteEmptyFoldersCompleted(sender As Object, e As EventArgs)

    Public Sub DeleteEmptyFolders(ByVal rootFolderPath As String)
        ' Recursively delete all empty subfolders
        For Each folderPath As String In IO.Directory.GetDirectories(rootFolderPath)
            DeleteEmptyFolders(folderPath)
        Next

        ' After recursion, check if the current folder is empty and delete if so
        If IO.Directory.GetFiles(rootFolderPath).Length = 0 AndAlso IO.Directory.GetDirectories(rootFolderPath).Length = 0 Then
            Try
                IO.Directory.Delete(rootFolderPath)
            Catch ex As Exception
                ' Consider logging the exception or notifying the user
                ' This avoids crashing the application if the deletion fails (e.g., due to permission issues)
            End Try
        End If

        ' Optionally, invoke the event handler if needed to perform additional clean-up
        ' CleanUpTree(Nothing, Nothing)
    End Sub

    Public Sub CleanUpTree(sender As Object, e As EventArgs)
        FormMain.tvmain2.Refresh()
    End Sub
    'Public Function DeleteEmptyFolders(d As DirectoryInfo, blnRecurse As Boolean) As Boolean
    '    Dim x As New FileOperations
    '    x.CurrentFolder = New DirectoryInfo(CurrentFolder)
    '    x.RemoveEmptySubfolders()
    '    Return True

    'End Function

    Public Async Function HarvestBelow(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(FormMain.tvmain2, FormMain.lbxFiles, d.FullName)
        Await x.HarvestFolder(d)

    End Function

    Public Function BurstFolder(d As DirectoryInfo) As Task
        Dim x As New BundleHandler(FormMain.tvmain2, FormMain.lbxFiles, d.FullName)


        x.Burst(d) 'Needs Attention
        '        HarvestFolder(d, True, True)

    End Function

    Public Sub PromoteFile(path As String)

        Dim finfo As New IO.FileInfo(path)
        finfo.MoveTo(finfo.Directory.Parent.FullName & "\" & finfo.Name)
    End Sub




End Module
