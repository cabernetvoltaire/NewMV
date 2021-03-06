﻿Option Explicit On
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
    Public VIDEOEXTENSIONS = ".divx.vob.webm.avi.flv.mov.m4p.mpeg.f4v.mpg.m4a.m4v.mkv.mp4.rm.ram.wmv.wav.mp3.3gp .lnk"
    Public PICEXTENSIONS = "arw.jpeg.png.jpg.bmp.gif.lnk"
    Public separate As Boolean = False
    Public CurrentFolder As String
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
        'ls = SelectFromListbox(lbx, ".lnk", False)
        lbx.SelectedItems.Clear()
        For Each fl In lbx.Items
            If InStr(fl, ".lnk") <> 0 Then
                ls.Add(fl)
            End If
        Next
        Dim deadlinks As New List(Of String)
        For Each f In ls
            If Not LinkTargetExists(f) Then
                deadlinks.Add(f)
            End If
        Next
        Return deadlinks
   End Function
    Public Sub CreateFavourite(Filepath As String)
        CreateLink(Filepath, FavesFolderPath, "", Bookmark:=Media.Position)
        Exit Sub


    End Sub
    Public Sub CreateLink(Filepath As String, DestinationDirectory As String, Name As String, Optional Update As Boolean = True, Optional Bookmark As Long = -1)
        Dim sh As New ShortcutHandler
        Dim f As New FileInfo(Filepath)
        Dim dt As Date = f.CreationTime

        'sh.Bookmark = Media.Position

        sh.TargetPath = Filepath
        sh.ShortcutPath = DestinationDirectory
        If Name = "" Then
            sh.ShortcutName = f.Name
        Else
            sh.ShortcutName = Name
        End If
        sh.Create_ShortCut(Bookmark)

        If DestinationDirectory = CurrentFolder And Update Then MainForm.UpdatePlayOrder(False)
    End Sub

    Public Function GetAllFilesBelow(DirectoryPath As String, ByVal FileList As List(Of String))
        If DirectoryPath.Contains("RECYCLE") Then
            Return FileList
            Exit Function
        End If
        Dim m As New DirectoryInfo(DirectoryPath)
        Try
            For Each k In m.EnumerateDirectories
                FileList = GetAllFilesBelow(k.FullName, FileList)
            Next

        Catch ex As System.UnauthorizedAccessException
            Return FileList
            Exit Function
        End Try
        For Each f In m.EnumerateFiles
            FileList.Add(f.FullName)
        Next
        Return FileList
    End Function

    ''' <summary>
    ''' Returns the path of the link defined in str
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Public Function LinkTarget(str As String) As String
        'Dim f As IO.File(str)

        Try
            str = CreateObject("WScript.Shell").CreateShortcut(str).TargetPath
            Return str
        Catch ex As Exception
            Return str
        End Try

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

                ' e.Graphics.DrawString("Property Item " + count.ToString(),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   iD: 0x" & propItem.Id.ToString("x"),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   type: " & propItem.Type.ToString(),
                'font, blackBrush, X, Y)
                ' Y += font.Height

                ' e.Graphics.DrawString("   length: " & propItem.Len.ToString() &
                ' " bytes", font, blackBrush, X, Y)
                ' Y += font.Height

                count += 1
            Next propItem
            MsgBox(des)
        'MsgBox(PropertyItems(theImage))
        'Catch ex As ArgumentException
        'MessageBox.Show("There was an error. Make sure the path to the image file is valid.")
        'End Try

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
    ''' Copies list from a lbx
    ''' </summary>
    ''' <param name="list"></param>
    ''' <param name="lbx"></param>
    Public Sub CopyList(ByVal list As List(Of String), lbx As ListBox)
        list.Clear()
        For Each m In lbx.Items
            list.Add(m)
        Next
    End Sub
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
    Public Function ListfromListbox(lbx As ListBox) As List(Of String)
        Dim s As New List(Of String)
        For Each l In lbx.SelectedItems
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
        FillShowbox(MainForm.lbxShowList, 0, s)
        Return s
    End Function
    ''' <summary>
    ''' Fill a listbox with a list, ignores the filter - dunno why
    ''' </summary>
    ''' <param name="lbx"></param>
    ''' <param name="Filter"></param>
    ''' <param name="lst"></param>
    '''
    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))


        If lst.Count = 0 Then Exit Sub
        If lst.Count > 1000 Then
            ProgressBarOn(lst.Count)
        End If


        lbx.Items.Clear()
        MainForm.CurrentFilterState.FileList = lst
        lst = MainForm.CurrentFilterState.FileList
        For Each s In lst
            lbx.Items.Add(s)

            'ProgressIncrement(1)
        Next

        If lbx.Name = "lbxShowList" Then
            MainForm.CollapseShowlist(False)
            'Dim m As New ShowListForm
            ShowListForm.Show()

            ShowListForm.ItemList = lst

        End If
        '  lbx.Refresh()
        '        lbx.TabStop = True
        ProgressBarOff()



        'MainForm.UpdateFileInfo()
    End Sub
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
            Console.WriteLine()
        Next
        Console.WriteLine(str)
        For i = 0 To gaps
            Console.WriteLine()
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

    Public Sub ChangeFolder(strPath As String)
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
                AssignLinear(CurrentFolder, ButtfromAsc(Asc("0")), True)
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
        Return True
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

                        Catch ex As Exception
                            MsgBox("Fail")

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

    Public Function ReturnListOfDirectories(ByVal list As List(Of String), strPath As String) As List(Of String)
        Dim d As New DirectoryInfo(strPath)
        For Each di In d.EnumerateDirectories
            list.Add(di.Name)
            list = ReturnListOfDirectories(list, di.Name)
        Next
        Return list
    End Function
    Public Function ButtfromAsc(asc As Integer) As Integer
        Dim n As Integer
        If asc <= 57 Then
            n = asc - 48 + 26

        Else
            n = asc - 65
        End If
        Return n
    End Function
    Public Function AscfromButt(button As Integer) As Integer
        Dim asc As Integer
        If button <= 25 Then
            asc = button + 65
        Else
            asc = button - 26 + 48
        End If
        Return asc
    End Function
End Module
