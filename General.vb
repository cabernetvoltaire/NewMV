Option Explicit On
Imports System.IO
Imports System.Drawing.Imaging
Imports System.Media
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
    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.f4v.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif"
    Public DirectoriesListFile
    Public separate As Boolean = False

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
            If Not LinkTargetExists(f) Then
                deadlinks.Add(f)
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
        Dim currentlink = n
        Dim parts() = currentlink.Split("\")
        Dim filename = parts(parts.Length - 1)
        Dim fparts() = filename.Split("%")
        filename = fparts(0)
        Return filename
    End Function
    'Public Function GetAllFilesBelow(DirectoryPath As String, ByVal FileList As List(Of String))
    '    If DirectoryPath.Contains("RECYCLE") Then
    '        Return FileList
    '        Exit Function
    '    End If
    '    Dim m As New DirectoryInfo(DirectoryPath)
    '    Try
    '        For Each k In m.EnumerateDirectories
    '            FileList = GetAllFilesBelow(k.FullName, FileList)
    '        Next

    '    Catch ex As System.UnauthorizedAccessException
    '        Return FileList
    '        Exit Function
    '    End Try
    '    For Each f In m.EnumerateFiles
    '        FileList.Add(f.FullName)
    '    Next
    '    Return FileList
    'End Function

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
                    Report(str & "target not found", 1, True)
                End If
            End If
            Return str
        Catch ex As Exception
            Return str
        End Try

    End Function


    Public Function GetDirectoriesList(path As String, Optional Force As Boolean = False) As List(Of String)
        Dim list As New List(Of String)
        Dim pathfile As New IO.FileInfo(DirectoriesListFile)
        If Not Force AndAlso pathfile.Exists Then
            ReadListfromFile(list, DirectoriesListFile, Encrypted)
        Else
            Try
                Dim root As New IO.DirectoryInfo(path)
                For Each m In root.GetDirectories("*", SearchOption.AllDirectories)
                    list.Add(m.FullName)
                Next
            Catch ex As Exception

            End Try
            WriteListToFile(list, DirectoriesListFile, Encrypted)
        End If
        Return list
    End Function

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

        If Len(str) <> 0 Then
            Dim driveletter As String = "A"
            For i = Asc("A") To Asc("Z")
                driveletter = Chr(i)
                str = str.Replace(Left(str, 2), driveletter & ":")
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
    Public Function Duplicatelist(ByVal inList As List(Of String)) As List(Of String)
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


    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub
#End Region


    Public Sub ReportFault(routinename As String, msg As String, Optional box As Boolean = True)
        If box Then
            MsgBox("Exception in " & routinename & vbCrLf & msg)
        Else
            Console.Write("Exception in " & routinename & vbCrLf & msg)

        End If
    End Sub
    Public Sub Report(str As String, gaps As Integer, Optional Sound As Boolean = False)
        If Sound Then SystemSounds.Asterisk.Play()

        For i = 0 To gaps
            Debug.WriteLine("")

        Next
        Debug.WriteLine(str)
        For i = 0 To gaps
            Debug.Print("")
        Next

    End Sub
    Public Sub ReportTime(str As String)
        Debug.Print(Int(Now().Second) & "." & Int(Now().Millisecond) & " " & str)
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




    Public Function SetPlayOrder(Order As Byte, ByVal List As List(Of String)) As List(Of String)
        Dim NewListS As New SortedList(Of String, String)
        Dim NewListL As New SortedList(Of Long, String)
        Dim NewListD As New SortedList(Of DateTime, String)
        'frmMain.ListBox1.BringToFront()

        Try
            Select Case Order
                Case SortHandler.Order.Name
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
                Case SortHandler.Order.Size
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

                Case SortHandler.Order.DateTime
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
                Case SortHandler.Order.PathName
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

                Case SortHandler.Order.Type
                    For Each f In List
                        If Len(f) > 247 Then Continue For
                        Dim file As New FileInfo(f)
                        NewListS.Add(file.Extension & file.Name & Str(Rnd() * (100)), file.FullName)
                    Next

                Case SortHandler.Order.Random
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
    Public Function GetDirSize(RootFolder As String, TotalSize As Long) As Long
        Dim FolderInfo = New IO.DirectoryInfo(RootFolder)
        For Each File In FolderInfo.GetFiles : TotalSize += File.Length
        Next
        For Each SubFolderInfo In FolderInfo.GetDirectories : GetDirSize(SubFolderInfo.FullName, TotalSize)
        Next
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
                Dim r As New System.Text.RegularExpressions.Regex(s)
                If r.Matches(lbx.Items(i)).Count > 0 Then
                    lbx.SetSelected(i, True)
                    ls.Add(lbx.Items(i))
                End If
            Else
                If InStr(UCase(lbx.Items(i)), UCase(s)) <> 0 Then
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
    Private Function Encrypt(ByVal strInput As String, ByVal strKey As String) As String
        Dim icount As Long
        Dim lngPtr As Long
        For icount = 1 To Len(strInput)
            Mid(strInput, icount, 1) = Chr((Asc(Mid(strInput, icount, 1))) Xor (Asc(Mid(strKey, lngPtr + 1, 1))))
            lngPtr = ((lngPtr + 1) Mod Len(strKey))
        Next icount
        Return strInput
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
End Module
