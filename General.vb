Option Explicit On
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Media
Imports System.Threading
Imports Microsoft.VisualBasic.Devices
Public Module General

    Public Property ForbiddenPaths As New List(Of String)
#Region "Enumerations"
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
    Public Enum SlowMoSoundOptions
        Silent
        Normal
        Slow
    End Enum
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
    Public Enum CtrlFocus As Byte
        Tree = 0
        Files = 1
        ShowList = 2
    End Enum
#End Region
#Region "Global Variables"
    Public RandomInstance As New Random()
    Public BreakHere As Boolean = True
    Friend ShowListVisible As Boolean = False
    Public Declare Function SearchTreeForFile Lib "imagehlp" (ByVal RootPath As String, ByVal InputPathName As String, ByVal OutputPathBuffer As String) As Long
    Public Const LinkExt As String = ".mvl"
    Public Property bImageDimensionState As Byte
    Public Property ShiftDown As Boolean
    Public Property CtrlDown As Boolean
    Public Property AltDown As Boolean
    Public Property KeyDownFlag As Boolean


    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.f4v.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif.jfif.webp"
    Public DirectoriesListFile
    Public separate As Boolean = False
    Public DebugOn As Boolean = True
    Public t As Threading.Thread
    Public CurrentFolder As String = "C:\"
    ' Public DirectoriesList As New List(Of String)

    Public Encrypted As Boolean = False
    Public WithEvents Encrypter As New Encryption("Spunky")
    Public UndoOperations As New Stack(Of Undo)




    Public lngShowlistLines As Long = 0
    Public ReadOnly Property Asterisk As SystemSound
    Public Orientation() As String = {"Unknown", "TopLeft", "TopRight", "BottomRight", "BottomLeft", "LeftTop", "RightTop", "RightBottom", "LeftBottom"}
#End Region

#Region "Links"

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
            ' GC.SuppressFinalize(box)
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

    Public Function LinkTargetContainsNameFile(LinkPath As String) As Boolean
        Dim LinkTarget As String = LinkTarget(LinkPath)
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
    Public Function LinkTarget(str As String) As String
        'Dim parts() = Split(Path.GetFileName(str), "%")
        'Return parts(0)
        'Exit Function
        Try
            Dim items As New List(Of String)
            items = ReadListfromFile(str, False)
            If items.Count > 0 Then

                str = items(0)
            Else
                str = ""
            End If

            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function
    Public Function GenerateSafeFolderList(folder As String, Depth As Integer) As List(Of String)
        Dim D As New DirectoryLister
        D.Depth = Depth
        D.Path = folder
        D.GenerateDirs()
        Return D.DirectoryList

    End Function
    Public Sub CollectFoldersRecursively(folder As String, ByRef allFolders As List(Of String))

        Try
            ' Add the current folder to the list
            allFolders.Add(folder)

            ' Recursively collect all subfolders
            For Each subFolder As String In Directory.GetDirectories(folder)
                CollectFoldersRecursively(subFolder, allFolders)
            Next

        Catch ex As UnauthorizedAccessException
            ' Log the error or ignore it
            'MsgBox($"Access denied to {folder}. Skipping...")
        Catch ex As Exception
            ' Handle other types of exceptions if necessary
            MsgBox($"An error occurred accessing {folder}: {ex.Message}")
        End Try
    End Sub
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
    Public Function GetRandomLong() As Long
        Dim rnd As New Random()
        Dim buffer As Byte() = New Byte(7) {} ' Long is 8 bytes
        rnd.NextBytes(buffer)
        Return BitConverter.ToInt64(buffer, 0)
    End Function
    Public Function GetRandomInt32() As Integer
        Dim rnd As New Random()
        Return rnd.Next(0, Integer.MaxValue)
    End Function

    Public Function getAvailableRAM() As ULong
        Dim CI As New ComputerInfo()
        Dim avl, used As String
        Dim mem As ULong = ULong.Parse(CI.AvailablePhysicalMemory.ToString())
        Dim mem1 As ULong = ULong.Parse(CI.TotalPhysicalMemory.ToString()) - ULong.Parse(CI.AvailablePhysicalMemory.ToString())
        avl = (mem / (1024 * 1024) & " MB").ToString() 'changed + to &
        used = (mem1 / (1024 * 1024) & " MB").ToString() 'changed + to &
        Return mem
    End Function

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
    Public Function SaveButtonFileName(path As String, Optional Force As Boolean = False) As String
        If path = "" Or Force Then
            With New SaveFileDialog
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"
                If Force Then
                    .FileName = path
                    Return path
                    Exit Function
                End If
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
#End Region




#Region "List functions"


    Public Function Duplicatelist(ByRef inList As List(Of String)) As List(Of String)
        Dim out As New List(Of String)
        For Each i In inList
            out.Add(i)
        Next
        Return out
    End Function
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

            CollectFoldersRecursively(Path, dirs)

            '            list.Add("Name" & vbTab & "Path" & vbTab & "Size in bytes" & vbTab & "Date")
            For Each d In dirs
                Dim dir As New DirectoryInfo(d)
                Dim paths As New List(Of String)
                'Create list of filepaths only
                For Each f In dir.GetFiles
                    Try
                        paths.Add(f.FullName)

                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
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
    Public Function GetAllFolders(d As DirectoryInfo, s As String) As List(Of String)

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
        s &= "Markers Count:" & MH.Markers.Count & vbCrLf

        '  s = s & MH.MediaPath & vbCrLf
        s = s & "Duration: " & sh.Duration & vbCrLf & "Percentage:" & sh.Percentage & vbCrLf & " Absolute:" & sh.Absolute & vbCrLf & " Startpoint:" & sh.StartPoint & vbCrLf & " Player:" & MH.Player.Name
        s = s & vbCrLf & sh.Description
        Debug.Print(s)
        FormMain.TextBox1.Text = s
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
                Case LinkExt
                    Return Filetype.Link
            End Select

            Dim strExt = LCase(info.Extension)
            If InStr(VIDEOEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(PICEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Pic
            ElseIf InStr(LCase(".txt.prn.sty.doc.csv.html.rtf.mht"), LCase(strExt)) <> 0 Then
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

            '       FormMain.WatchStart(d.Parent.FullName)
        End If


    End Sub








    Function GetDate(f As FileInfo) As DateTime
        Dim time As DateTime = f.CreationTime
        Dim time2 As DateTime = f.LastAccessTime
        Dim time3 As DateTime = f.LastWriteTime
        If time2 < time Then time = time2
        If time3 < time Then time = time3
        Return time
    End Function





    Public Sub OnNotEncrypted() Handles Encrypter.NotEncrypted
        Encrypted = False
    End Sub


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




    Public Function BookmarkFromLinkName(path As String) As Long
        Dim s() = path.Split("%")
        If s.Length > 2 Then
            Return s(s.Length - 2)
        Else
            Return 0
        End If

    End Function
    Public Function GetFileNameFromLinkName(Link As String) As String
        Dim s() As String
        s = Link.Split("%")
        Return s(0)
        '        Throw New NotImplementedException()
    End Function
    Public Function AdvanceModular(Count As Integer, i As Integer, Forward As Boolean, Optional Inc As Integer = 1)
        If Forward Then
            i = (i + Inc) Mod Count
        Else
            i = (i - Inc)
            If i < 0 Then i = i + Count
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
        Dim s As String = ""

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
    Public Function LongAsTimeCode(n As Long) As String
        Return New TimeSpan(0, 0, n).ToString("hh\:mm\:ss")
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



    Public Sub ChangeLabel(lbl As Label, Text As String)
        lbl.Text = Text

    End Sub

    Friend Sub UpdateLabel(lbl As Label, TextToAdd As String)
        If lbl.Text.Contains(TextToAdd) Then
            lbl.Text = lbl.Text.Replace(vbCrLf & TextToAdd, "")
        Else
            lbl.Text = lbl.Text & vbCrLf & TextToAdd
        End If

    End Sub

    Friend Function PassFilename() As String
        Dim TheFileName As String = ""
        Dim Vars As String()
        Vars = Environment.GetCommandLineArgs
        If Vars.Length > 1 Then
            TheFileName = Environment.GetCommandLineArgs(1)

        End If
        Return TheFileName
    End Function

End Module
