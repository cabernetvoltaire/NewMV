﻿Option Explicit On
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Media
Imports System.Threading

Public Module General
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
    Public ThumbDestination As String = "Q:\Thumbs\"

    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.f4v.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif"
    Public DirectoriesListFile
    Public separate As Boolean = False
    Public t As Threading.Thread
    Public CurrentFolder As String
    Public DirectoriesList As New List(Of String)
    Public Rootpath As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
    Public GlobalFavesPath As String
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

    Public Sub ProgressBarOn(max As Long)
        With MainForm.TSPB
            .Value = 0
            .Maximum = max 'Math.Max(lngListSizeBytes, 100)
            .Visible = True
        End With

    End Sub
    Public Sub ProgressIncrement(st As Integer)
        With MainForm.TSPB
            '   .Maximum = max
            .Value = (.Value + st) Mod .Maximum
        End With
        'MainForm.Update()
    End Sub
    Public Sub ProgressBarOff()
        With MainForm.TSPB
            .Visible = False
        End With

    End Sub
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
            If f.EndsWith(".lnk") Then
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
        Handler.Create_ShortCut(Bookmark)

        If DestinationDirectory = CurrentFolder And Update Then MainForm.UpdatePlayOrder(MainForm.FBH)
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
    Public Function FilenameFromPath(n As String, WithExtension As Boolean) As String
        Dim currentpath = n
        Dim parts() = currentpath.Split("\")
        Dim filename = parts(parts.Length - 1)
        If Not WithExtension Then
            Dim fparts() = filename.Split(".")
            filename = fparts(0)
        End If
        Return filename
    End Function

    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function LinkTarget(str As String) As String

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
            End If
            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function
    Private Function GenerateSafeFolderList(ByVal folder As String) _
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
                    For Each D As String In System.IO.Directory.GetDirectories(Dir)
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
            ReadListfromFile(list, DirectoriesListFile, Encrypted)
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

    Public Sub FolderChooser(Message As String, DefaultFolderName As String)
        Dim x As New FolderSelect

        x.Show()
        x.Text = Message
        If CurrentFolder <> "" Then
            x.Folder = CurrentFolder
        Else
            x.Folder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
        End If
        x.TextBox1.Text = DefaultFolderName
    End Sub

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

    Public Function ImageOrientation(ByVal img As Image) As ExifOrientations
        ' Get the index of the orientation property.
        Dim orientation_index As Integer = Array.IndexOf(img.PropertyIdList, OrientationId)

        ' If there is no such property, return Unknown.
        If (orientation_index < 0) Then Return ExifOrientations.Unknown

        ' Return the orientation value.
        Return DirectCast(img.GetPropertyItem(OrientationId).Value(0), ExifOrientations)
    End Function
#End Region




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
    Public Sub TransferURLS(MS1 As MediaSwapper, MS2 As MediaSwapper)
        MS1.Media1 = MS2.Media1
        MS1.Media2 = MS2.Media2
        MS1.Media3 = MS2.Media3
        MS1.ListIndex = MS2.ListIndex
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
    Public Function ListfromSelectedInListbox(lbx As ListBox) As List(Of String)
        Dim s As New List(Of String)
        For Each l In lbx.SelectedItems
            s.Add(l)
        Next
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

    Public Function ListFromLinks(list As List(Of String)) As List(Of String)
        Dim s As New List(Of String)
        For Each l In list
            Dim tgt As String = LinkTarget(l)
            If s.Contains(tgt) Then
            Else
                s.Add(LinkTarget(tgt))
            End If

        Next
        MainForm.FillShowbox(MainForm.lbxShowList, 0, s)
        Return s
    End Function
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
    Public Sub ReadListfromFile(list As List(Of String), filepath As String, Encrypted As Boolean)
        Dim fs As New StreamReader(New FileStream(filepath, FileMode.OpenOrCreate, FileAccess.Read))
        Dim line As String
        Do While fs.Peek <> -1
            line = fs.ReadLine
            If Encrypted Then
                line = Encrypter.DecryptData(line)
            End If
            list.Add(line)
        Loop
        fs.Close()
    End Sub

    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub
#End Region

#Region "Debugging functions"

    Public Sub ReportFault(routinename As String, msg As String, Optional box As Boolean = True)
        If box Then
            MsgBox("Exception in " & routinename & vbCrLf & msg)
        Else
            Console.Write("Exception in " & routinename & vbCrLf & msg)

        End If
    End Sub
    Public Sub Report(str As String, gaps As Integer, Optional Sound As Boolean = False)
        If Sound Then SystemSounds.Asterisk.Play()

        For i = 1 To gaps
            Console.WriteLine()
        Next
        Console.WriteLine(str)
        For i = 1 To gaps
            Console.WriteLine()
        Next

    End Sub
    Public Sub LabelStartPoint(ByRef MH As MediaHandler)
        If MH.MediaPath = "" Then Exit Sub
        Dim s As String = ""
        Dim sh As StartPointHandler = MH.StartPoint
        s = s & MH.Name & vbCrLf
        s = s & MH.MediaPath & vbCrLf
        s = s & "Duration: " & sh.Duration & vbCrLf & "Percentage:" & sh.Percentage & vbCrLf & " Absolute:" & sh.Absolute & vbCrLf & " Startpoint:" & sh.StartPoint & vbCrLf & " Player:" & Media.Player.Name
        s = s & vbCrLf & sh.Description
        Debug.Print(s)
        MainForm.lblNavigateState.Text = s
    End Sub
    Public Sub ReportTime(str As String)
        Debug.Print(Int(Now().Second) & "." & Int(Now().Millisecond) & " " & str)
    End Sub
#End Region
    Public Sub DeleteThumbs()
        Exit Sub
        If Application.OpenForms.Count = 1 Then
            Dim d As New IO.DirectoryInfo("Q:\Thumbs")
            For Each f In d.GetFiles
                Try
                    f.Delete()

                Catch ex As Exception

                End Try
            Next
        End If
    End Sub
    Public Function FindType(file As String) As Filetype
        Try
            Dim info As New IO.FileInfo(file)
            Select Case LCase(info.Extension)
                Case ""
                    Return Filetype.Unknown
                Case ".lnk"
                    Return Filetype.Link
            End Select

            Dim strExt = LCase(info.Extension)
            If InStr(VIDEOEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Movie
            ElseIf InStr(PICEXTENSIONS, strExt) <> 0 Then
                Return Filetype.Pic
            ElseIf InStr(".txt.prn.sty.doc", strExt) <> 0 Then
                Return Filetype.Doc
            Else
                Return Filetype.Unknown


            End If

        Catch ex As Exception
            Return Filetype.Unknown
        End Try

    End Function
    Public Sub ChangeFolder(strPath As String)
        'If strPath <> FavesFolderPath Then
        '    CurrentfilterState.State = CurrentfilterState.OldState
        'End If
        If strPath = CurrentFolder Then
        Else
            If Not LastFolder.Contains(CurrentFolder) Then
                LastFolder.Push(CurrentFolder)

            End If
            MainForm.FNG.Clear()
            ' MainForm.tvMain2.SelectedFolder = strPath
            ChangeWatcherPath(strPath)
            CurrentFolder = strPath
            ReDim FBCShown(0)
            NofShown = 0

            If AutoButtons Then
                AssignLinear(CurrentFolder, LetterNumberFromAscii(Asc("0")), True)
                ChangeButtonLetter(New KeyEventArgs(Keys.D0))
            End If
            '   My.Computer.Registry.CurrentUser.SetValue("File", Media.MediaPath)
        End If

    End Sub
    Public Sub ChangeWatcherPath(path As String)
        Dim d As New DirectoryInfo(path)
        If d.Parent Is Nothing Then
        Else

            '     MainForm.WatchStart(d.Parent.FullName)
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


    Public Function SetPlayOrder(Order As Byte, List As List(Of String)) As List(Of String)
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
                    List = Randomize(Of String)(List)

                Case SortHandler.Order.Type
                    Dim cpr As New CompareByType
                    List.Sort(cpr)


            End Select
        Catch ex As Exception

        End Try
        If MainForm.PlayOrder.ReverseOrder Then
            List.Reverse()
        End If
        Return List
    End Function


    Public Function SetPlayOrderOld(Order As Byte, ByVal List As List(Of String)) As List(Of String)

        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of DateTime, String)
        'frmMain.ListBox1.BringToFront()

        Try
            Select Case Order
                Case SortHandler.Order.Name
                    List.Sort(New CompareByEndNumber)
                    'SortByName(List, NewListS)
                Case SortHandler.Order.Size
                    List.Sort(New CompareByFilesize)
                    'SortBySize(List, NewListL)

                Case SortHandler.Order.DateTime
                    List.Sort(New CompareByDate)
                   ' SortByDate(List, NewListD)
                Case SortHandler.Order.PathName
                    List.Sort()
'                    SortbyPathName(List, NewListS)

                Case SortHandler.Order.Type
                    SortbyType(List, NewListS)

                Case SortHandler.Order.Random
                    SortbyRandom(List, NewListS)
                Case Else

            End Select
        Catch ex As System.ArgumentException 'TODO could do better than this. 
            ReportFault("General.SetPlayOrder", ex.Message)
        End Try

        If NewListD.Count <> 0 Then
            CopyList(List, NewListD)
        ElseIf NewListS.Count <> 0 Then
            CopyList(List, NewListS)
        ElseIf NewListL.Count <> 0 Then
            CopyList(List, NewListL)
        End If

        If MainForm.PlayOrder.ReverseOrder Then
            List = ReverseListOrder(List)
        End If

        Return List




    End Function

    Private Sub SortbyRandom(List As List(Of String), NewListS As SortedList(Of String, String))
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)

            Dim l As Long
            l = Int(Rnd() * (100 * List.Count))
            While NewListS.ContainsKey(Str(l))
                l = Int(Rnd() * (100 * List.Count))
                '                       frmMain.ListBox1.Items.Add(l)

            End While
            NewListS.Add(Str(l), file.FullName)
        Next
    End Sub

    Private Sub SortbyType(List As List(Of String), NewListS As SortedList(Of String, String))
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            NewListS.Add(file.Extension & file.Name & Str(Rnd() * (100)), file.FullName)
        Next
    End Sub

    Private Sub SortbyPathName(List As List(Of String), NewListS As SortedList(Of String, String))
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            Dim l As Long = 0
            Dim s As String
            s = file.FullName & Str(l)
            While NewListS.ContainsKey(s)
                l += 1
                s = file.FullName & Str(l)
                '               frmMain.ListBox1.Items.Add(s)

            End While
            '                        MsgBox(file.FullName)
            NewListS.Add(s, file.FullName)
        Next
    End Sub

    Private Sub SortByDate(List As List(Of String), NewListD As SortedList(Of Date, String))
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            'MsgBox(time)
            Dim time = GetDate(file)
            While NewListD.ContainsKey(time)
                time = time.AddSeconds(1)
            End While
            NewListD.Add(time, file.FullName)
        Next
    End Sub

    Private Sub SortBySize(List As List(Of String), NewListL As SortedList(Of Long, String))
        For Each f In List
            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)
            Try
                Dim l As Long
                l = file.Length
                While NewListL.ContainsKey(l)
                    l += 1
                    'MsgBox(l)
                End While
                NewListL.Add(l, file.FullName)

            Catch ex As FileNotFoundException
            Catch ex As Exception
                MsgBox("Unhandled Error in SetPlayOrder" & vbCrLf & ex.Message)


            End Try
        Next
    End Sub

    Private Sub SortByName(List As List(Of String), NewListS As SortedList(Of String, String))

        For Each f In List

            If Len(f) > 247 Then Continue For
            Dim file As New FileInfo(f)

            Dim l As Long = 0
            Dim s As String
            s = file.Name & Str(l)
            While NewListS.ContainsKey(s)
                l += 1
                s = file.Name & Str(l)
                '               frmMain.ListBox1.Items.Add(s)

            End While
            NewListS.Add(s, file.FullName)
        Next
    End Sub

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
        Dim size As Long = GetDirSize(Rootfolder, 0).Result
        Dim sizestring As String
        If size > 10 ^ 12 Then
            sizestring = Str(Format(size / 1024 / 1024 / 1024 / 1024, "###,###,###,###,###.#")) & " Tb"
        ElseIf size > 10 ^ 9 Then
            sizestring = Str(Format(size / 1024 / 1024 / 1024, "###,###,###,###,###.#")) & " Mb"

        ElseIf size > 10 ^ 6 Then
            sizestring = Str(Format(size / 1024 / 1024, "###,###,###,###,###.#")) & " Mb"
        ElseIf size > 10 ^ 3 Then
            sizestring = Str(Format(size / 1024, "###,###,###,###,###.#")) & " Kb"
        Else
            If size = 0 Then
                sizestring = "0 bytes"
            Else
                sizestring = Str(Format(size, "###,###,###,###,###.#")) & " bytes"

            End If

        End If
        Return sizestring
    End Function

    Public Async Function GetDirSize(RootFolder As String, TotalSize As Long) As Task(Of Long)
        ' Exit Function
        If RootFolder.EndsWith(":\") Then
            Return 0
            Exit Function
        End If

        Dim fso = CreateObject("Scripting.FileSystemObject")
        Dim dir As New IO.DirectoryInfo(RootFolder)
        If dir.Exists Then
            Dim profile = fso.GetFolder(RootFolder)

            TotalSize = profile.size
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
                FileStream1.Dispose()
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
                If InStr(UCase(FilenameFromPath(lbx.Items(i), False)), UCase(s)) <> 0 Then
                    lbx.SetSelected(i, True)
                    ls.Add(lbx.Items(i))
                End If
            End If
        Next

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
    ''' Randomizes the contents of the list using Fisher–Yates shuffle (a.k.a. Knuth shuffle).
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="list"></param>
    ''' <returns>Randomized result</returns>
    ''' <remarks></remarks>

    Function Randomize(Of T)(ByVal list As List(Of T)) As List(Of T)
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
        th = ThumbDestination & f.Name.Replace(".", "") & "thn.png"
        Return th
    End Function

    Friend Function FolderNameFromPath(path As String) As String
        Dim parts() = path.Split("\")
        Return parts(parts.Length - 1)
    End Function

End Module
