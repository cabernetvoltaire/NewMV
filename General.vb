Option Explicit On
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Media
Imports System.Threading

Public Module General

    Public Property ForbiddenPaths As New List(Of String)
    Public Enum ExifOrientations As Byte
        Unknown = 0
        TopLeft = 1
        TopRight = 2
        BottomRight = 3
        BottomLeft = 4
        LeftTop = 5
        RightTop = 6
        RightBottom = 7
        LeftBottom = 8
    End Enum
    Public BreakHere As Boolean = False
    Friend ShowListVisible As Boolean = False
    Public Declare Function SearchTreeForFile Lib "imagehlp" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long
    Public Const LinkExt As String = ".mvl"
    Public Property bImageDimensionState As Byte
    Public Property ShiftDown As Boolean
    Public Property CtrlDown As Boolean
    Public Property AltDown As Boolean
    Public Property KeyDownFlag As Boolean


    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.f4v.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif.jfif"
    Public DirectoriesListFile
    Public separate As Boolean = False
    Public DebugOn As Boolean = True
    Public t As Threading.Thread
    Public CurrentFolder As String
    ' Public DirectoriesList As New List(Of String)

    Public Encrypted As Boolean = False
    Public WithEvents Encrypter As New Encryption("Spunky")
    Public UndoOperations As New Stack(Of Undo)


    Public Enum CtrlFocus As Byte
        Tree = 0
        Files = 1
        ShowList = 2
    End Enum

    Public lngShowlistLines As Long = 0
    Public ReadOnly Property Asterisk As SystemSound
    Public Orientation() As String = {"Unknown", "TopLeft", "TopRight", "BottomRight", "BottomLeft", "LeftTop", "RightTop", "RightBottom", "LeftBottom"}
    Public Enum Filetype As Byte
        Pic
        Movie
        Doc
        Gif
        Xcel
        Browsable
        Link
        Unknown
    End Enum


#Region "Controls"





#End Region

#Region "Links"
    Public Sub SelectDeadLinks(lbx As ListBox)

        HighlightList(lbx, GetDeadLinks(lbx))

    End Sub
    Public Function GetDeadLinks(lbx As ListBox) As List(Of String)
        Dim ls As New List(Of String)
        ls = AllfromListbox(lbx)
        'ls = SelectFromListbox(lbx, ".lnk", False)
        'lbx.SelectedItems.Clear()
        'For Each fl In lbx.Items
        '    If InStr(fl, ".lnk") <> 0 Then
        '        ls.Add(fl)
        '    End If
        'Next
        Dim deadlinks As New List(Of String)

        For Each f In ls
            If f.EndsWith(LinkExt) Then
                If Not LinkTargetExists(f) Then
                    deadlinks.Add(f)
                End If
            End If
        Next
        Return deadlinks
    End Function
    Public Sub CreateFavourite(Filepath As String)
        Dim sh As New ShortcutHandler
        CreateLink(sh, Filepath, CurrentFavesPath, "", False, Bookmark:=Media.Position)
        AllFaveMinder.NewPath(GlobalFavesPath)
        sh = Nothing
    End Sub
    Public Function LinkTargetExists(Linkfile As String) As Boolean
        Dim f As String
        f = LinkTarget(Linkfile)
        If f = "" Then
            Return False
            Exit Function
        End If
        Dim Finfo = New FileInfo(f)
        If Finfo.Exists Then
            Return True
        Else
            Return False
        End If

    End Function
    Public Sub CreateLink(Handler As ShortcutHandler, Filepath As String, DestinationDirectory As String, Name As String, Optional Update As Boolean = True, Optional Bookmark As Long = -1)
        Dim f As New FileInfo(Filepath)
        Dim dt As Date = f.CreationTime
        Handler.TargetPath = Filepath
        Handler.ShortcutPath = DestinationDirectory
        If Name = "" Then
            Handler.ShortcutName = f.Name
        Else
            Handler.ShortcutName = Name
        End If
        Handler.New_Create_ShortCut(Bookmark)

        If DestinationDirectory = CurrentFolder And Update Then FormMain.UpdatePlayOrder(FormMain.FBH)
    End Sub
    Public Sub ConvertLink(OldLinkPath As String)
        Dim bk As Long
        Dim target As String
        target = LinkTarget(OldLinkPath)
        Dim oldfile As New IO.FileInfo(OldLinkPath)
        Dim dt = oldfile.LastWriteTime
        bk = BookmarkFromLinkName(OldLinkPath)
        Dim file As New List(Of String) From {target, bk}
        Dim newname = Replace(OldLinkPath, ".lnk", LinkExt)
        WriteListToFile(file, newname, False)
        Dim newfile As New IO.FileInfo(newname)
        newfile.LastWriteTime = dt

    End Sub
    Public Function FilenameFromLink(n As String) As String
        If n = "" Then
            Return ""
        Else
            Dim currentlink As String = n
            Dim parts() = currentlink.Split("\")
            Dim filename = parts(parts.Length - 1)
            Dim fparts() = filename.Split("%")
            filename = fparts(0)
            Return filename
        End If
    End Function
    Public Sub DisposePic(box As PictureBox)
        If box.Image IsNot Nothing Then
            box.Image.Dispose()
            GC.SuppressFinalize(box)
            box.Image = Nothing
        End If
    End Sub
    Public Sub BreakExecution()
        If BreakHere Then System.Diagnostics.Debugger.Break()
        BreakHere = False
    End Sub
    Public Function FilenameFromPath(n As String, WithExtension As Boolean, Optional WithoutBrackets As Boolean = False) As String
        Dim currentpath = n
        Dim parts() = currentpath.Split("\")
        Dim filename = parts(parts.Length - 1)
        If Not WithExtension Then
            Dim fparts() = filename.Split(".")
            'Could sometimes have . as part of filename
            Dim i = 1
            filename = fparts(0)
            While i < fparts.Length - 1
                filename = filename + "." + fparts(i)
                i += 1
            End While
            '            filename = fparts(0)
        End If
        If WithoutBrackets Then
            filename = Left(filename, InStr(filename, "("))
        End If
        Return filename
    End Function

    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function LinkTarget(str As String) As String

        Dim xl As New IO.FileInfo(str)
        If xl.Extension = LinkExt Then
            Return NewLinkTarget(str)
            Exit Function
        End If
        Try
            str = CreateObject("WScript.Shell").CreateShortcut(str).TargetPath
            Dim f As New IO.FileInfo(str)
            If f.Exists Then
            Else
                Dim x = str
                str = TryOtherDriveLetters(str)
                If str = x Then
                    'Report(str & "target not found", 1, False)
                End If
                'TODO TrywithoutBrackets
            End If
            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function
    ''' <summary>
    ''' Returns true if target contains filename contained in its name
    ''' </summary>
    ''' <param name="LinkPath"></param>
    ''' <returns></returns>
    Public Function LinkTargetContainsNameFile(LinkPath As String) As Boolean
        Dim LinkTarget As String = NewLinkTarget(LinkPath)
        Dim Filename As String = Path.GetFileName(Split(LinkPath, "%")(0))
        If Filename = Path.GetFileName(LinkTarget) Then
            Return True
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function NewLinkTarget(str As String) As String
        'Dim parts() = Split(Path.GetFileName(str), "%")
        'Return parts(0)
        'Exit Function
        Try
            Dim items As New List(Of String)
            items = ReadListfromFile(str, False)
            str = items(0)
            Dim f As New IO.FileInfo(items(0))
            If f.Exists Then

            Else
                Dim x = str
                str = TryOtherDriveLetters(str)
                If str = x Then
                    'Report(str & "target not found", 1, False)
                End If
                'TODO TrywithoutBrackets
            End If
            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function
    ''' <summary>
    ''' Get all folders beneath folder, which are safe.
    ''' </summary>
    ''' <param name="folder"></param>
    ''' <returns></returns>
    Public Function GenerateSafeFolderList(ByVal folder As String, Optional TopOnly As Boolean = False) _
            As List(Of String)

        ' -------------------------------------------
        ' Based On A Function By John Wein As Posted:
        '
        ' http://social.msdn.microsoft.com/Forums/en-US/vbgeneral/thread/d6e64558-395b-4b48-8b64-0f5a7e3a7623
        '
        ' Thanks John!
        ' -------------------------------------------

        Dim retVal As New List(Of String)

        Dim Dirs As New Stack(Of String)
        Dirs.Push(folder)

        While Dirs.Count > 0
            If False Then
                Exit While
            Else
                Dim Dir As String = Dirs.Pop

                Try
                    If TopOnly Then

                        For Each D As String In System.IO.Directory.GetDirectories(Dir, "*", SearchOption.TopDirectoryOnly)
                            ' Do not include any that are either system or hidden

                            Dim dirInfo As New System.IO.DirectoryInfo(D)
                            If (((dirInfo.Attributes And System.IO.FileAttributes.Hidden) = 0) AndAlso
                                ((dirInfo.Attributes And System.IO.FileAttributes.System) = 0)) Then

                                If Not retVal.Contains(D) Then
                                    retVal.Add(D)
                                End If
                            End If

                            Dirs.Push(D)
                        Next
                    Else

                        For Each D As String In System.IO.Directory.GetDirectories(Dir, "*", SearchOption.AllDirectories)
                            ' Do not include any that are either system or hidden

                            Dim dirInfo As New System.IO.DirectoryInfo(D)
                            If (((dirInfo.Attributes And System.IO.FileAttributes.Hidden) = 0) AndAlso
                                ((dirInfo.Attributes And System.IO.FileAttributes.System) = 0)) Then

                                If Not retVal.Contains(D) Then
                                    retVal.Add(D)
                                End If
                            End If

                            Dirs.Push(D)
                        Next
                    End If

                Catch ex As Exception

                    If retVal.Contains(Dir) Then
                        Dim indexToRemove As Integer = 0

                        For i As Integer = 0 To retVal.Count - 1
                            If retVal(i) = Dir Then
                                indexToRemove = i
                                Exit For
                            End If
                        Next

                        retVal.RemoveAt(indexToRemove)
                    End If
                    Continue While
                End Try
            End If
        End While

        Return retVal

    End Function

    Public Function GetDirectoriesList(path As String, Optional Force As Boolean = False) As List(Of String)
        Dim list As New List(Of String)
        Dim pathfile As New IO.FileInfo(DirectoriesListFile)
        If Not Force AndAlso pathfile.Exists Then
            list = ReadListfromFile(DirectoriesListFile, Encrypted)
            '   MsgBox(list.Count)
        Else
            '   DirectoriesLister(path, list)
            t = New Thread(New ThreadStart(Sub() DirectoriesLister(path, list)))
            t.IsBackground = True
            t.SetApartmentState(ApartmentState.STA)
            t.Start()
            While t.IsAlive
                Thread.Sleep(100)
            End While
            WriteListToFile(list, DirectoriesListFile, Encrypted)
            MsgBox("Directories Loaded")
        End If
        Return list
    End Function

    Private Sub DirectoriesLister(path As String, list As List(Of String))
        Try
            Dim root As New IO.DirectoryInfo(path)

            For Each m In GenerateSafeFolderList(path)
                Dim fol As New IO.DirectoryInfo(m)
                list.Add(fol.FullName)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Public Sub FolderChooser(sender As Object, e As MouseEventArgs)
        Dim x As New FormFolderSelect
        x.BH = sender.findform.BH
        x.ButtonNumber = sender.tag

        x.Show()
        x.Text = "Choose folder"
        'If CurrentFolder <> "" Then
        '    x.Folder = CurrentFolder
        'Else
        '    x.Folder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
        'End If
        x.TextBox1.Text = x.Folder
    End Sub
    Public Function LoadButtonFileName(path As String) As String
        If path = "" Then
            With New OpenFileDialog
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"

                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        Return path

    End Function
    Public Function SaveButtonFileName(path As String) As String
        If path = "" Then
            With New SaveFileDialog
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"

                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        Return path

    End Function
    Public Function LoadDatabaseFileName(path As String) As String
        If path = "" Then
            With New OpenFileDialog
                .InitialDirectory = ListFilePath
                .DefaultExt = "msd"
                .Filter = "Metavisua database files|*.msd|All files|*.*"

                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        Return path

    End Function
    Public Function SaveDatabaseFileName(path As String) As String
        If path = "" Then
            With New SaveFileDialog
                .InitialDirectory = ListFilePath
                .DefaultExt = "msd"
                .Filter = "Metavisua database files|*.msd|All files|*.*"

                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        Return path

    End Function
    Public Function BrowseToFolder(Title As String, DefaultPath As String) As String
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
    Public Function BrowseToFile(Title As String) As String
        Dim path As String
        With New OpenFileDialog
            .Title = Title
            If .ShowDialog() = DialogResult.OK Then
                path = .FileName
            Else
                path = ""
            End If
        End With
        Return path
    End Function
    Public Function TryOtherDriveLetters(str As String) As String
        Dim original = str
        Dim m As New List(Of DriveInfo)
        For Each x In IO.DriveInfo.GetDrives
            m.Add(x)
        Next
        If Len(str) <> 0 Then
            For Each drive In m

                Dim driveletter As String = drive.Name
                str = str.Replace(Left(str, 3), driveletter)
                If My.Computer.FileSystem.FileExists(str) Then
                    Return str
                    Exit Function
                End If
            Next
            str = original
        End If
        Return str

    End Function
#End Region


#Region "Rotation Functions"

    Public Function ImageOrientation(ByRef img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function
#End Region


#Region "Metadata Functions"
    Public Sub ExtractMetaData(theImage As Image)

        ' Try
        'Create an Image object. 

        'Get the PropertyItems property from image.
        Dim propItems As PropertyItem() = theImage.PropertyItems

        'Set up the display.
        Dim font As New Font("Arial", 10)
        Dim blackBrush As New SolidBrush(Color.Black)
        Dim X As Integer = 0
        Dim Y As Integer = 0

        'For each PropertyItem in the array, display the id, type, and length.
        Dim count As Integer = 0
        Dim propItem As PropertyItem
        Dim des As String = ""

        For Each propItem In propItems
            des = des + vbCrLf & "Property Item " + count.ToString()
            des = des & vbTab & "iD: 0x" & propItem.Id.ToString("x")
            des = des & vbTab & "  type" & propItem.Type.ToString()
            des = des & vbTab & "Length" & propItem.Len.ToString()


            count += 1
        Next propItem
        MsgBox(des)
        'MsgBox(PropertyItems(theImage))
        'Catch ex As ArgumentException
        'MessageBox.Show("There was an error. Make sure the path to the image file is valid.")
        'End Try

    End Sub
#End Region
#Region "List functions"

    ''' <summary>
    ''' Copies list from a sorted list2
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="list2"></param>
    Private Sub CopyList(ByVal list As List(Of String), ByVal list2 As SortedList(Of String, String))
        list.Clear()
        For Each m As KeyValuePair(Of String, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Public Function Duplicatelist(ByRef inList As List(Of String)) As List(Of String)
        Dim out As New List(Of String)
        For Each i In inList
            out.Add(i)
        Next
        Return out
    End Function
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Long, String))
        list.Clear()
        For Each m As KeyValuePair(Of Long, String) In list2
            list.Add(m.Value)
        Next
    End Sub
    Public Function ListfromSelectedInListbox(lbx As ListBox, Optional All As Boolean = False) As List(Of String)
        Dim s As New List(Of String)
        If All Then
            For Each l In lbx.Items
                s.Add(l)
            Next
        Else
            For Each l In lbx.SelectedItems
                s.Add(l)
            Next
        End If
        Return s
    End Function


    Public Sub RefreshListbox(lbx As ListBox, list As List(Of String))

        For Each m In list
            lbx.Items.Remove(m)
        Next
    End Sub


    Public Function AllfromListbox(lbx As ListBox) As List(Of String)
        Dim s As New List(Of String)
        For Each l In lbx.Items
            s.Add(l)
        Next
        Return s
    End Function
    ''' <summary>
    ''' Given a list of links, finds all the targets and places in Showbox
    ''' </summary>
    ''' <param name="list"></param>
    ''' <returns></returns>
    Public Function ListFromLinks(list As List(Of String)) As List(Of String)
        Dim s As New List(Of String)
        For Each l In list
            Dim tgt As String = LinkTarget(l)
            If s.Contains(tgt) Then
            Else
                s.Add(tgt)
            End If

        Next
        FormMain.FillShowbox(FormMain.lbxShowList, 0, s)
        Return s
    End Function
#Region "File Access Functions"

    Public Sub WriteListToFile(list As List(Of String), filepath As String, Encrypted As Boolean)
        Dim fs As New StreamWriter(New FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Write))

        For Each m In list
            If Encrypted Then

                fs.WriteLine(Encrypter.EncryptData(m))
            Else
                fs.WriteLine(m)
            End If
        Next
        fs.Close()
    End Sub
    Public Function ReadListfromFile(filepath As String, Encrypted As Boolean) As List(Of String)
        Dim fs As New StreamReader(New FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read))
        Dim list As New List(Of String)
        Dim line As String
        Do While fs.Peek <> -1
            line = fs.ReadLine
            Dim x As String = line

            If Encrypted Then
                line = Encrypter.DecryptData(line)
            End If
            ' Application.DoEvents()
            list.Add(line)
        Loop

        fs.Close()
        Return list
    End Function
#End Region
#Region "Database Functions"

    Friend Sub CreateDatabase(Path As String, Filename As String)
        Dim exclude As String = ""
        exclude = LCase(InputBox("String to exclude from folders?", ""))
        With My.Computer.FileSystem
            Dim list As New List(Of String)
            Dim dirs As New List(Of String)

            dirs = GenerateSafeFolderList(Path)
            dirs.Add(Path)
            '            list.Add("Name" & vbTab & "Path" & vbTab & "Size in bytes" & vbTab & "Date")
            For Each d In dirs
                Dim dir As New DirectoryInfo(d)
                Dim paths As New List(Of String)
                'Create list of filepaths only
                For Each f In dir.GetFiles
                    paths.Add(f.FullName)
                Next
                Dim x As New FilterHandler With {.State = CurrentfilterState.State}
                x.FileList = paths
                paths = x.FileList
                For Each file In paths
                    Try

                        Dim f As New FileInfo(file)
                        If exclude <> "" And LCase(file).Contains(exclude) Then
                        Else
                            list.Add(f.Name & vbTab & f.FullName.Replace(f.Name, "") & vbTab & f.Length & vbTab & GetDate(f))
                        End If

                    Catch ex As Exception

                    End Try
                Next
                'Filter them
                'Then create database list of those files
                Application.DoEvents()
            Next
            WriteListToFile(list, Filename, Encrypted)
        End With
    End Sub
#End Region
    Public Function GetFileFromEachFolder(d As DirectoryInfo, s As String, Optional Random As Boolean = True) As List(Of String)

        Dim x As New List(Of String)
        For Each Di In d.EnumerateDirectories(s, SearchOption.AllDirectories)
            Dim dirs() As FileInfo
            dirs = Di.GetFiles
            If dirs.Count > 0 Then
                If Random Then
                    Dim i = Int(Rnd() * dirs.Count)
                    x.Add(dirs(i).FullName)

                Else
                    x.Add(Di.EnumerateFiles.First.FullName)

                End If
            End If
        Next
        Return x
    End Function
    Public Function GetAllFolders(d As DirectoryInfo, s As String, Optional Random As Boolean = True) As List(Of String)

        Dim x As New List(Of String)
        For Each Di In d.EnumerateDirectories(s, SearchOption.AllDirectories)
            x.Add(Di.FullName)
        Next
        Return x
    End Function
#End Region

#Region "Debugging functions"

    Public Sub ReportFault(routinename As String, msg As String, Optional box As Boolean = True)
        If box Then
            MsgBox("Exception in " & routinename & vbCrLf & msg)
        Else
            Console.Write("Exception in " & routinename & vbCrLf & msg)

        End If
    End Sub
    Public Sub Report(str As String, Optional gaps As Integer = 0, Optional Sound As Boolean = False)
        Exit Sub
        If DebugOn Then
            If Sound Then SystemSounds.Asterisk.Play()

            For i = 1 To gaps
                Console.WriteLine()
            Next
            Console.WriteLine(str)
            For i = 1 To gaps
                Console.WriteLine()
            Next
        End If

    End Sub
    Public Sub LabelStartPoint(ByRef MH As MediaHandler)
        If MH.MediaPath = "" Then Exit Sub
        Dim s As String = ""
        Dim sh As StartPointHandler = MH.SPT
        s = s & MH.Name & vbCrLf
        s = s & MH.MediaPath & vbCrLf
        s = s & "Duration: " & sh.Duration & vbCrLf & "Percentage:" & sh.Percentage & vbCrLf & " Absolute:" & sh.Absolute & vbCrLf & " Startpoint:" & sh.StartPoint & vbCrLf & " Player:" & Media.Player.Name
        s = s & vbCrLf & sh.Description
        Debug.Print(s)
        FormMain.lblNavigateState.Text = s
    End Sub
    Public Sub ReportTime(str As String)
        Debug.Print(Int(Now().Second) & "." & Int(Now().Millisecond) & " " & str)
    End Sub
    Public Function TimeOperation(blnStart As Boolean) As TimeSpan
        Static StartTime As Date
        If blnStart Then
            StartTime = Now
            Return Now - StartTime
        Else
            Return Now - StartTime
        End If
    End Function
#End Region

    Public Function FindType(file As String) As Filetype
        Try
            Dim info As New IO.FileInfo(file)
            Select Case LCase(info.Extension)
                Case ""
                    Return Filetype.Unknown
                Case ".lnk", LinkExt
                    Return Filetype.Link
            End Select

            Dim strExt = LCase(info.Extension)
            If InStr(VIDEOEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(PICEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Pic
            ElseIf InStr(LCase(".txt.prn.sty.doc.csv.html"), LCase(strExt)) <> 0 Then
                Return Filetype.Doc
            Else
                Return Filetype.Unknown


            End If

        Catch ex As Exception
            Return Filetype.Unknown
        End Try

    End Function

    Public Sub ChangeWatcherPath(path As String)
        Dim d As New DirectoryInfo(path)
        If d.Parent Is Nothing Then
        Else

            FormMain.WatchStart(d.Parent.FullName)
        End If


    End Sub

    Public Function FileLengthCheck(file As String) As Boolean
        Return True
        Exit Function
        If Len(file) > 247 Then
            If MsgBox("Filename too long - truncate?", MsgBoxStyle.YesNo, "Filename too long") = MsgBoxResult.Yes Then
                Dim m As New FileInfo(file)
                Dim i As Integer = Len(m.FullName)
                Dim l As Integer = Len(m.Directory.FullName)
                If l > 247 Then
                    ReportFault("FileLengthCheck", "Unsuccessful - folder name alone is too long")
                    Return False
                    Exit Function
                Else
                    Dim str As String = Right(m.Name, 247 - l)
                    m.MoveTo(m.Directory.FullName & "\" & str)
                    Return True
                End If
            End If
        End If
        Return False
    End Function


    Public Function SetPlayOrder(Order As Byte, List As List(Of String), Reverse As Boolean) As List(Of String)
        Try
            Select Case Order
                Case SortHandler.Order.Name
                    List.Sort(New CompareByEndNumber)
                Case SortHandler.Order.Size
                    Dim cpr As New CompareByFilesize
                    List.Sort(cpr)
                Case SortHandler.Order.DateTime
                    Dim cpr As New CompareByDate
                    List.Sort(cpr)
                Case SortHandler.Order.PathName
                    List.Sort()
                Case SortHandler.Order.Random
                    List = Randomise(Of String)(List)

                Case SortHandler.Order.Type
                    Dim cpr As New CompareByType
                    List.Sort(cpr)


            End Select
        Catch ex As Exception

        End Try
        If Reverse Then
            List.Reverse()
        End If
        Return List
    End Function

    Friend Function ExtractParagraphs(Story As String, MH As MediaHandler) As String
        Dim txtan As New TextAnalyzer With {.RawText = Story}

        Return txtan.CleanText

        'Return Story
        Exit Function

        Dim lines() As String = Story.Split(vbLf)
        If lines.Length = 1 Then
            Return lines(0)
            Exit Function
        End If
        Dim paras As New List(Of String)
        Dim currentpara As String = ""
        For i = 0 To lines.Length - 1
            Dim line As String = lines(i)
            line = line.Replace(vbCr, "").Replace(vbLf, " ").Replace(vbCrLf, "")
            If line.Length <> 0 And line.Replace(" ", "").Length <> 0 Then
                If line(0) = " " Then
                    paras.Add(currentpara)
                    currentpara = ""
                    currentpara = " " & line
                Else
                    currentpara = currentpara & " " & line
                End If
            End If
        Next
        paras.Add(currentpara)
        Dim returnstring As String = ""
        For Each p In paras
            returnstring = returnstring & vbCrLf & vbCrLf & p
        Next
        Return returnstring
    End Function


    Function GetDate(f As FileInfo) As DateTime
        Dim time As DateTime = f.CreationTime
        Dim time2 As DateTime = f.LastAccessTime
        Dim time3 As DateTime = f.LastWriteTime
        If time2 < time Then time = time2
        If time3 < time Then time = time3
        Return time
    End Function

    Public Sub ReplaceListboxItem(lbx As ListBox, index As Integer, newitem As String)
        lbx.Items.RemoveAt(index)
        lbx.Items.Insert(index, newitem)

    End Sub
    Public Function GetDirSizeString(Rootfolder As String) As String
        Return "##"
        Exit Function
        Dim size As Long = GetDirSize(Rootfolder, 0)
        If size = 0 Then
            Return "0"
            Exit Function
        End If
        Dim sizestring As String
        'This could be done in almost one line. 
        'eg
        Dim log As Integer = Int(Math.Log(size) / Math.Log(1024))
        Dim suffix() As String = {" bytes", " Kb", " Mb", " Gb", " Tb"}
        sizestring = Str(Format(size / (1024) ^ Int(log), "###,###,###,###,###.#")) & suffix(log)
        Return sizestring

    End Function

    Public Function GetDirSize(RootFolder As String, TotalSize As Long, Optional NoSubs As Boolean = False) As Long

        ' Exit Function
        If RootFolder.EndsWith(":\") Then
            Return 0
            Exit Function
        End If

        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim dir As New IO.DirectoryInfo(RootFolder)
        Dim filessize As Long
        If Not NoSubs Then

            If dir.Exists Then
                Dim profile = fso.GetFolder(RootFolder)

                TotalSize = profile.size
            End If
        Else
            Try

                For Each f In dir.GetFiles
                    Dim profile = fso.getfile(f.FullName)
                    filessize = filessize + profile.size
                Next
            Catch ex As Exception

            End Try
            TotalSize = filessize
        End If
        'Dim FolderInfo = New IO.DirectoryInfo(RootFolder)
        'For Each File In FolderInfo.GetFiles : TotalSize += File.Length
        'Next
        'For Each SubFolderInfo In FolderInfo.GetDirectories : GetDirSize(SubFolderInfo.FullName, TotalSize)
        'Next
        Return TotalSize
    End Function

    Public Sub OnNotEncrypted() Handles Encrypter.NotEncrypted
        Encrypted = False
    End Sub

    Public Function LoadImage(fname As String) As Image
        Try
            Dim FileStream1 As New System.IO.FileStream(fname, IO.FileMode.Open, IO.FileAccess.Read)


            Try
                Dim MyImage As Image = Image.FromStream(FileStream1)
                FileStream1.Close()
                '  FileStream1.Dispose()
                Return MyImage
            Catch ex As System.ArgumentException
                FileStream1.Close()
                FileStream1.Dispose()
                Return Nothing
            End Try
        Catch ex As Exception

            Return Nothing
        End Try


    End Function
    Public Function SelectFromListbox(lbx As ListBox, s As String, blnRegex As Boolean) As List(Of String)
        Dim ls As New List(Of String)
        Dim i As Long
        lbx.SelectionMode = SelectionMode.MultiExtended
        lbx.SelectedItem = Nothing
        For i = 0 To lbx.Items.Count - 1
            If blnRegex Then
                Try

                    Dim r As New System.Text.RegularExpressions.Regex(s)
                    If r.Matches(lbx.Items(i)).Count > 0 Then
                        lbx.SetSelected(i, True)
                        ls.Add(lbx.Items(i))
                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            Else
                If InStr(UCase(lbx.Items(i)), UCase(s)) <> 0 Then
                    lbx.SuspendLayout()
                    lbx.SetSelected(i, True)
                    ls.Add(lbx.Items(i))
                End If
            End If
        Next
        lbx.ResumeLayout()
        Return ls

    End Function

    Public Sub HighlightList(lbx As ListBox, ls As List(Of String))
        lbx.SelectedItems.Clear()
        lbx.Refresh()
        lbx.SelectionMode = SelectionMode.MultiExtended
        For Each f In ls
            Dim i = lbx.FindString(f)
            If i >= 0 Then lbx.SetSelected(i, True)
        Next

    End Sub
    Public Function ReverseListOrder(m As List(Of String)) As List(Of String)
        Dim k As New List(Of String)
        For Each x In m
            k.Insert(0, x)

        Next
        Return k
    End Function
    ''' <summary>
    ''' Removes the brackets from a list of filenames, optionally only retaining the original
    ''' </summary>
    ''' <param name="list"></param>
    Sub RemoveBrackets(list As List(Of String))
        Dim x As New BracketRemover
        x.ListOfFiles = list
        x.RemoveBrackets()
        Exit Sub
        For Each f In list
            Dim file As New IO.FileInfo(f)
            If file.Exists Then
                Dim n As String = file.Name
                'Need to use regex here.
                For i = 1 To 100
                    If n = n.Replace("(" & Str(i) & ")", "") Then
                    Else
                        n.Replace("(" & Str(i) & ")", "")
                    End If


                Next
                n = AppendExistingFilename(n, file)
                If file.Name <> n Then

                    Try
                        My.Computer.FileSystem.RenameFile(file.FullName, n)

                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If

            End If

        Next
    End Sub
    ''' <summary>
    ''' Randomizes the contents of the list using Fisher–Yates shuffle (a.k.a. Knuth shuffle).
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="list"></param>
    ''' <returns>Randomized result</returns>
    ''' <remarks></remarks>

    Function Randomise(Of T)(ByVal list As List(Of T)) As List(Of T)
        Dim rand As New Random()
        Dim temp As T
        Dim indexRand As Integer
        Dim indexLast As Integer = list.Count - 1
        For index As Integer = 0 To indexLast
            indexRand = rand.Next(index, indexLast)
            temp = list(indexRand)
            list(indexRand) = list(index)
            list(index) = temp
        Next index
        Return list
    End Function
    Public Function BookmarkFromLinkName(path As String) As Long
        Dim s() = path.Split("%")
        If s.Length > 2 Then
            Return s(s.Length - 2)
        Else
            Return 0
        End If

    End Function

    Public Function AdvanceArrayIndexModular(a As Array, i As Integer, Forward As Boolean)
        If Forward Then
            i = (i + 1) Mod a.Length
        Else
            i = (i - 1)
            If i < 0 Then i = i + a.Length
        End If
        Return i
    End Function
    ''' <summary>
    ''' Takes the Ascii for the letter
    ''' translates it into an integer with A=0 and 9=36
    ''' (0 is 27)
    ''' </summary>
    ''' <param name="asc"></param>
    ''' <returns></returns>
    Public Function LetterNumberFromAscii(asc As Integer) As Integer
        Dim n As Integer
        If asc <= 57 Then
            n = asc - 48 + 26

        Else
            n = asc - 65
        End If
        If n < 0 Or n > 36 Then n = 27
        Return n
    End Function
    Public Function AsciifromLetterNumber(button As Integer) As Integer
        Dim asc As Integer
        If button <= 25 Then
            asc = button + 65
        Else
            asc = button - 26 + 48
        End If
        Return asc
    End Function
    Friend Function AreSameImage(ByVal I1 As Image, ByVal I2 As Image, Optional EmptyTrue As Boolean = False) As Boolean
        Dim BM1 As Bitmap = I1
        Dim BM2 As Bitmap = I2
        If BM1 Is Nothing Or BM2 Is Nothing Then
            Return EmptyTrue
        Else
            'Check the overlapping regions of the pics. This should be good enough
            Dim xw As Integer = Math.Min(BM1.Width, BM2.Width)
            Dim yh As Integer = Math.Min(BM1.Height, BM2.Height)
            For x = 1 To xw - 1
                For y = 1 To yh - 1
                    If BM1.GetPixel(x, y) <> BM2.GetPixel(x, y) Then
                        Return False
                        Exit Function
                    End If
                Next
            Next
            Return True
        End If
    End Function

    Friend Function ThumbnailName(Filename As String) As String
        Dim th As String
        Dim f As New IO.FileInfo(Filename)
        th = ThumbDestination & "\" & f.Name.Replace(".", "") & "thn.png"
        Return th
    End Function


    Friend Function FolderPathFromPath(path As String) As String

        Dim parts() = path.Split("\")
        Dim s As String

        For i = 0 To parts.Length - 2
            s = s + parts(i) & "\"
        Next
        Return s
    End Function
    Friend Function CloneMenu(m As ToolStripMenuItem) As ToolStripMenuItem
        Dim newitem As New ToolStripMenuItem()
        With newitem
            .Name = m.Name
            .Text = m.Text
            .ToolTipText = m.ToolTipText
            .Enabled = m.Enabled
            .Tag = m.Tag
            'Add any other things you want to preserve, like .Image
            'Not much point preserving Keyboard Shortcuts on a context menu
            'as it's designed for a mouse click.
            If m.HasDropDownItems Then
                For Each sm In m.DropDownItems
                    If TypeOf (sm) Is ToolStripMenuItem Then
                        newitem.DropDownItems.Add(CloneMenu(sm))
                    End If
                Next
            End If
        End With
        AddHandler newitem.Click, AddressOf m.PerformClick
        Return newitem
    End Function


    'Friend Function GetAllMenuItems(t As ToolStripMenuItem) As List(Of ToolStripMenuItem)
    '    Dim list As New List(Of ToolStripMenuItem)
    '    list.Add(t)
    '    If t.Tag Then
    '        For Each m In t.DropDownItems
    '            list.AddRange(GetAllMenuItems(m))
    '        Next
    '    End If
    '    Return list
    'End Function

    Public Sub MovietoPic(pic As PictureBox, img As Image)
        pic.Image = img
        'PreparePic(pic, img)
        'SndH.Muted = True


    End Sub
    Public Sub OrientPic(img As Image)

        Select Case ImageOrientation(img)
            Case ExifOrientations.BottomRight
                img.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case ExifOrientations.RightTop
                img.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case ExifOrientations.LeftBottom
                img.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select
    End Sub
    Public Function GetImage(strPath As String) As Image
        If strPath = "" Then Return Nothing
        'If strPath.EndsWith(".gif") = 0 Then
        '    Return LoadImage(strPath)

        '    ' Exit Function 'This Causes problems if extension is .gif
        'Else
        Try
            Dim img As Image = Image.FromFile(strPath)
            Return img
        Catch ex As Exception
            'Reportfault("",ex.message)
            Return Nothing

        End Try
        'End If
    End Function
    Public Sub StoreList(list As List(Of String), Dest As String)
        If Dest = "" Then Exit Sub
        WriteListToFile(list, Dest, Encrypted)

    End Sub
    ''' <summary>
    ''' If spath exists, then returns an appropriately suffixed new name
    ''' </summary>
    ''' <param name="spath"></param>
    ''' <param name="m"></param>
    ''' <returns></returns>
    Public Function AppendExistingFilename(spath As String, m As IO.FileInfo) As String
        Dim i = 0
        If m.Name = FilenameFromPath(spath, True) Then 'Existing path - rename
            While My.Computer.FileSystem.FileExists(spath)
                Dim x = m.Extension
                Dim b = InStr(spath, "(")
                If b = 0 Then
                    spath = Replace(spath, x, "(" & i & ")" & x)
                Else
                    spath = Left(spath, b - 1) & "(" & i & ")" & x
                End If

                i += 1
            End While
        End If
        Return spath
    End Function
    Friend Sub PrependFolderName(list As List(Of String))
        For Each s In list
            Dim f As New IO.FileInfo(s)
            Dim name = f.Name
            Dim dir = f.Directory.Name
            ' AllFaveMinder.CheckFile(f)
            Try
                f.MoveTo(Path.GetFullPath(s).Replace(f.Name, dir & "_" & name))

            Catch ex As Exception

            End Try
        Next
    End Sub
    Friend Sub AddToolTip(ctl As Control, tp As ToolTip, text As String)
        tp.SetToolTip(ctl, text)
    End Sub

    Friend Sub OutlineControl(ctl As Control, ByRef outliner As PictureBox)
        ctl.Parent.Controls.Add(outliner)
        Dim r As Rectangle = ctl.Bounds
        outliner.SetBounds(r.Left - 5, r.Top - 5, r.Width + 10, r.Height + 10)
        outliner.Visible = True
    End Sub
    ''' <summary>
    ''' Returns all subfolders of path which have the same name as name
    ''' </summary>
    ''' <param name="path"></param>
    ''' <returns></returns>
    Friend Function SameNameFoldersBelow(path As String, name As String) As List(Of DirectoryInfo)
        Dim folder As New IO.DirectoryInfo(path)
        Dim list As New List(Of IO.DirectoryInfo)
        For Each f In folder.GetDirectories("*", SearchOption.AllDirectories)
            If LCase(f.Name) = LCase(name) Then
                list.Add(f)
            End If
        Next
        Return list

    End Function
    Friend Sub MoveSameFolders(path As String)
        Dim folder As New IO.DirectoryInfo(path)
        Dim fcl As New FoldersAndTheirMatches
        For Each f In folder.GetDirectories("*", SearchOption.TopDirectoryOnly)
            fcl.Folder = f
            fcl.MoveFolders()
        Next


    End Sub
    Friend Sub FlattenAllSubFolders(fol As IO.DirectoryInfo)
        For Each f In fol.GetDirectories("*", SearchOption.AllDirectories)
            If f.Parent.FullName <> fol.FullName Then
                MoveFolder(f.FullName, fol.FullName)
            End If
        Next

    End Sub

    Public Class FoldersAndTheirMatches
        Private Property _Folder As IO.DirectoryInfo
        WriteOnly Property Folder As IO.DirectoryInfo
            Set(value As IO.DirectoryInfo)
                _Folder = value
                Matches = SameNameFoldersBelow(value.Parent.FullName, value.Name)
            End Set
        End Property

        Friend Matches As List(Of IO.DirectoryInfo)
        Friend Sub MoveFolders()
            For Each m In Matches
                If m.Parent.FullName <> _Folder.Parent.FullName Then MoveFolder(m.FullName, _Folder.Parent.FullName)
            Next
        End Sub
    End Class
    Public Class MetaInfo
        Public Mdir As MetadataExtractor.Directory
        Public Directory As IReadOnlyList(Of MetadataExtractor.Directory)
        Public FileName As String
        Public CreateDate As String
        Public FileSize As Long
        Public Duration As Long
        Public Sub GetData()
            For Each m In Directory
                'If m.GetTagName() = "File Name" Then

                'End If
            Next
        End Sub
    End Class
    Public Class DatabaseEntry
        Implements IComparable

        Private Property mFullPath As String
        Public Property Filename As String
        Public ReadOnly Property FullPath As String
            Get
                Return mFullPath
            End Get
        End Property

        Private Property mPath As String
        Public Property Path As String
            Set(value As String)
                mPath = value
                mFullPath = String.Concat(mPath, Filename)
            End Set
            Get
                Return mPath
            End Get
        End Property
        Public Size As Long
        Public Dt As DateTime

        Public Property Mark As Boolean = False
        'Public Thumbnail As New Bitmap(100, 100)
        'Public Function GetThumbnail()
        '    Thumbnail = LoadImage(mFullPath).GetThumbnailImage(100, 100, Nothing, Nothing)
        'End Function
        Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
            If obj.size = Size Then
                Return 0
            ElseIf obj.size < Size Then
                Return 1
            Else
                Return -1


            End If
        End Function

    End Class
    Public Class Database
        Public Entries As New List(Of DatabaseEntry)
        Public Name As String
        Private Property mItemCount As Long = 0
        Public ReadOnly Property ItemCount As Long
            Get
                Return mItemCount
            End Get
        End Property


        Public Sub AddEntry(entry As DatabaseEntry)
            Dim f As New FileInfo(entry.FullPath)
            If f.Exists Then
                Entries.Add(entry)
                mItemCount += 1
            End If
        End Sub
        Public Sub ConvertDB()
            For Each e In Entries

            Next
        End Sub
        Private mSorted As Boolean = False
        Public Sub Sort(Optional comparer As IComparer(Of DatabaseEntry) = Nothing)
            Entries.Sort(comparer)
            mSorted = True
        End Sub
        Public Function FindEntry(filename As String) As DatabaseEntry
            If Not mSorted Then Sort(New CompareDBByFilename)

            Dim n As New List(Of DatabaseEntry)
            n = Entries.FindAll(Function(x) x.Filename.StartsWith(filename))
            Select Case n.Count
                Case 0
                    n = Entries.FindAll(Function(x) x.Filename.Contains(Path.GetFileNameWithoutExtension(filename)))
                    If n.Count = 0 Then
                        Return Nothing
                    Else

                    End If
                Case 1
                    Return n(0)
                Case Else
                    Return n(0)
                    'Decide which one to return
            End Select

        End Function
        Public Function FindPartEntry(Search As String) As List(Of DatabaseEntry)
            If Not mSorted Then Sort()

            Dim n As New List(Of DatabaseEntry)
            n = Entries.FindAll(Function(x) x.FullPath.Contains(Search))
            Return n
        End Function


    End Class
    Public Class BracketRemover
        Public ListOfFiles As New List(Of String)
        Public RemoveDuplicates As Boolean = True
        Private Sub RemoveNonBracketed()
            Dim tList As New List(Of String)
            For Each m In ListOfFiles
                Dim x As New IO.FileInfo(m)
                If m.EndsWith(")" & x.Extension) Then
                    tList.Add(m)
                End If
            Next
            ListOfFiles = tList
        End Sub
        Public Sub RemoveBrackets()
            RemoveNonBracketed()

            For Each f In ListOfFiles
                Dim file As New IO.FileInfo(f)
                If file.Exists Then
                    'Remove any bracketed numbers at the end of the filename
                    Dim n As String = file.FullName
                    'Need to use regex here.
                    For i = 0 To 100
                        Dim p As String = Str(i).Substring(1)
                        p = n.Replace("(" & p & ")", "")
                        If n <> p Then
                            n = p
                            Exit For
                        End If
                    Next
                    'If it was successful, then rename the file without the brackets
                    'but if this causes a clash with an existing file of that name
                    'delete this file if permitted, and if their lengths are the same
                    If file.FullName <> n Then
                        Dim nfile As New IO.FileInfo(n)
                        Try
                            My.Computer.FileSystem.RenameFile(file.FullName, nfile.Name)
                        Catch ex As IOException
                            If RemoveDuplicates Then
                                If nfile.Exists Then
                                    If file.Length = nfile.Length Then
                                        file.Delete()
                                    End If
                                End If
                            End If
                        End Try
                    End If

                End If

            Next
        End Sub
    End Class
    Public Function AreSameFile(path As String, qath As String) As Boolean
        Dim same As Boolean = False
        Dim pbytes, qbytes As Byte()
        pbytes = ReadBinary(path, 1000, 5000)
        qbytes = ReadBinary(qath, 1000, 5000)
        Dim k As Integer
        While k < 1000

            If pbytes(k) = qbytes(k) Then
                same = True
            Else
                same = False
                Exit While
            End If
            k += 1
        End While

        Return same
    End Function

    Public Function ReadBinary(path As String, size As Long, start As Long) As Byte()
        Dim bytesRead As Long
        Dim bytes(size) As Byte
        Using strm As New FileStream(path, FileMode.Open, FileAccess.Read)
            strm.Position = start
            Using rdr As New BinaryReader(strm)

                bytesRead = rdr.Read(bytes, 0, size)


            End Using
        End Using
        Return bytes
        Exit Function
    End Function

    Public Sub ChangeLabel(lbl As Label, Text As String)
        lbl.Text = Text

    End Sub
    Public Sub AppendLabel(lbl As Label, Text As String)
        lbl.Text = lbl.Text & vbCrLf & Text

    End Sub
    Friend Sub UpdateLabel(lbl As Label, TextToAdd As String)
        If lbl.Text.Contains(TextToAdd) Then
            lbl.Text = lbl.Text.Replace(vbCrLf & TextToAdd, "")
        Else
            lbl.Text = lbl.Text & vbCrLf & TextToAdd
        End If

    End Sub
    Friend Function ListFromDatabase(DB As Database) As List(Of String)
        Dim s As New List(Of String)
        For Each e In DB.Entries
            s.Add(e.FullPath)
        Next
        Return s

    End Function
End Module
