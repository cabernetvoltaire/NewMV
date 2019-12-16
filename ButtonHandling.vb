Imports System.IO

Module ButtonHandling
    Public buttons As New ButtonSet
    Public AutoButtons As Boolean = False
    Public nletts As Int16 = 36
    '   Public layers(nletts) As Byte
    ' Public currentlayer As Byte
    Public strButtonFilePath(8, nletts, 1) As String
    ' Public fButtonDests(8, nletts, 3) As IO.DirectoryInfo
    Private ReadOnly iAlphaCount = nletts
    Public strButtonCaptions(8, nletts, 1)
    Public Buttonfolder As String = "Q:\.msb"


    Public btnDest() As Button = {MainForm.btn1, MainForm.btn2, MainForm.btn3, MainForm.btn4, MainForm.btn5, MainForm.btn6, MainForm.btn7, MainForm.btn8}



    Public lblDest() As Label = {MainForm.lbl1, MainForm.lbl2, MainForm.lbl3, MainForm.lbl4, MainForm.lbl5, MainForm.lbl6, MainForm.lbl7, MainForm.lbl8}
    Public Sub Buttons_Load()
        buttons.CurrentLetter = LetterNumberFromAscii(Asc("A"))
        For i As Byte = 0 To 7
            lblDest(i).Font = New Font(lblDest(i).Font, FontStyle.Bold)
        Next
        ' blnButtonsLoaded = True
        InitialiseButtons()
    End Sub

    Public Sub InitialiseButtons()
        For i As Byte = 0 To 7
            With btnDest(i)
                .Text = buttons.CurrentRow.Buttons(i).FaceText
                'AddHandler .Click, AddressOf ButtonClick
                AddHandler .Click, AddressOf ShowPreview
                'AddHandler .MouseLeave, AddressOf hidePreview
                '  AddHandler .MouseHover, AddressOf ChangePreviewMedia

            End With

            lblDest(i).Text = buttons.CurrentRow.Buttons(i).Label
        Next
    End Sub


    Private Sub ShowPreview(Sender As Object, e As EventArgs)


        ' Exit Sub
        Dim index As Byte = Val(Sender.Name.ToString(3))

        If FolderSelect.Visible Then
            If index - 1 <> FolderSelect.ButtonNumber Then
            Else
                FolderSelect.Hide()

                Exit Sub
            End If

        End If
        FolderSelect.ButtonNumber = index - 1
        FolderSelect.Alpha = iCurrentAlpha
        FolderSelect.Show()
        Dim s As String = buttons.CurrentRow.Buttons(index - 1).Path
        If s = "" Then s = CurrentFolder
        FolderSelect.Folder = s
        Debug.Print("New Folder" & s)
        '        x.Show()
        Dim control As Control = CType(Sender, Control)
        Dim startpoint As Point
        startpoint.X = control.Left
        startpoint.Y = control.Top
        startpoint = control.PointToScreen(startpoint)
        FolderSelect.Left = startpoint.X - FolderSelect.Width / 2
        FolderSelect.Top = startpoint.Y - FolderSelect.Height
        FolderSelect.BringToFront()
        'FolderSelect.UpdateFolder()
    End Sub




    ''' <summary>
    ''' Assigns all the buttons in a generation, beneath sPath, to the letter iAlpha
    ''' </summary>
    Public Sub AssignLinear(sPath As String, iAlpha As Integer, blnNext As Boolean)
        If AutoButtons Then
        Else
            If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        End If

        Dim d As New DirectoryInfo(sPath)
        Dim i As Byte
        Dim di() As DirectoryInfo
        di = d.GetDirectories
        Dim n = d.GetDirectories.Count - 1
        If n > 0 Then
            For i = 0 To n
                Dim k As Byte
                k = i Mod 8

                AssignButton(k, iAlpha + Int(i / 8), 1, di(i).FullName)
            Next
        End If
        If AutoButtons Then
        Else
            KeyAssignmentsStore(ButtonFilePath)

        End If

    End Sub
    Public Sub AssignAlphabetic(blntest As Boolean)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        Dim dlist As New List(Of String)
        Dim d As New DirectoryInfo(Media.MediaDirectory)
        FindAllFoldersBelow(d, dlist, True, True)
        ' dlist = SetPlayOrder(PlayOrder.Name, dlist)
        dlist.Sort()
        Dim n(nletts) As Integer
        Dim layer As Int16 = 1
        For i = 0 To dlist.Count - 1
            Dim s As String = dlist.Item(i)
            Dim sht As String = New DirectoryInfo(s).Name
            Dim l As String = UCase(sht(0))
            Dim k As Int16 = LetterNumberFromAscii(Asc(l))
            If k >= 0 AndAlso k < nletts Then
                If (n(k) Mod 8) = 0 Then
                    layer += 1
                    ReDim Preserve strButtonFilePath(8, nletts, layer)
                    ReDim Preserve strButtonCaptions(8, nletts, layer)
                Else
                End If
                AssignButton(n(k), k, layer, s)
                n(k) += 1

            End If

        Next
        KeyAssignmentsStore(ButtonFilePath)
    End Sub
    Public Sub AutoButtonsToggle()
        AutoButtons = Not AutoButtons

    End Sub
    Public Sub AssignSize(Start As String)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub
        Dim dlist As New SortedList(Of String, DirectoryInfo)
        Dim plist As New SortedList(Of Integer, DirectoryInfo)
        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        Dim d As New DirectoryInfo(Start)
        For Each di In d.EnumerateDirectories
            If InStr(di.Name, exclude) = 0 Then
                plist.Add(GetDirSize(di.Name, 0), di)
            End If
        Next

    End Sub



    Public Sub AssignTree(strStart As String)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        Dim dlist As New SortedList(Of String, DirectoryInfo)
        Dim plist As New SortedList(Of String, DirectoryInfo)
        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        Dim d As New DirectoryInfo(strStart)
        For Each di In d.EnumerateDirectories
            If exclude = "" Or InStr(di.Name, exclude) = 0 Then
                'MsgBox(di.Name & "is " & Format(GetDirSize(di.FullName, 0), "###,###,###,###,###.#"))
                dlist.Add(di.Name, di)
            End If

        Next
        Dim i As Int16 = 0



        Dim n(nletts) As Integer
        While i < 8 AndAlso (n(0) < 8 OrElse n(1) < 8 OrElse n(2) < 8 OrElse n(3) < 8 OrElse n(4) < 8 OrElse n(5) < 8 OrElse n(6) < 8 OrElse n(7) < 8 OrElse n(8) < 8 OrElse n(9) < 8 OrElse n(10) < 8 OrElse n(11) < 8 OrElse n(12) < 8 OrElse n(1) < 8 OrElse n(13) < 8 OrElse n(14) < 8 OrElse n(15) < 8 OrElse n(17) < 8 OrElse n(18) < 8 OrElse n(19) < 8 OrElse n(20) < 8 OrElse n(21) < 8 OrElse n(22) < 8 OrElse n(23) < 8 OrElse n(24) < 8)

            For Each di In dlist.Values
                Dim l As String = UCase(di.Name(0))
                Dim ButtonNumber As Int16 = LetterNumberFromAscii(Asc(l))
                If ButtonNumber >= 0 AndAlso ButtonNumber < nletts Then
                    If exclude <> "" AndAlso InStr(di.Name, exclude) <> 0 Then
                    Else
                        If n(ButtonNumber) < 8 Then
                            AssignButton(n(ButtonNumber), ButtonNumber, 1, di.FullName)
                            n(ButtonNumber) += 1
                        End If
                    End If
                End If


            Next

            plist = FindNextTier(dlist)
            dlist = plist
            i += 1
        End While


        KeyAssignmentsStore(ButtonFilePath)
    End Sub
    Public Sub AssignTreeNew(StartingFolder As String, SizeMagnitude As Byte)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        ClearCurrentButtons()
        Dim d As New DirectoryInfo(StartingFolder)

        Dim icomp As New MyComparer
        Dim dlist As New SortedList(Of Long, DirectoryInfo)(icomp)

        For Each di In d.EnumerateDirectories("*", searchOption:=IO.SearchOption.AllDirectories)

            Dim disize = GetDirSize(di.FullName, 0)
            If (exclude = "" Or Not di.Name.Contains(exclude)) And disize > 10 ^ SizeMagnitude Then
                'MsgBox(di.Name & " is " & Format(GetDirSize(di.FullName, 0), "###,###,###,###,###.#"))
                While dlist.Keys.Contains(disize)
                    disize += 1
                End While
                dlist.Add(disize, di)
            End If

        Next
        Dim i As Int16 = 0
        'dlist.Reverse
        '   dlist.Reverse

        Dim n(nletts) As Integer

        For Each dx In dlist
            Dim di As IO.DirectoryInfo
            di = dx.Value
            Dim l As String = UCase(di.Name(0))
            Dim ButtonNumber As Integer = LetterNumberFromAscii(Asc(l))
            buttons.CurrentLetter = ButtonNumber
            Dim firstbtn As New MVButton()
            firstbtn.Letter = ButtonNumber
            firstbtn.Position = buttons.CurrentRow.GetFirstFree()
            firstbtn.Path = di.FullName
            If firstbtn.Position < 8 Then AssignButton(firstbtn.Position, firstbtn.Letter, 1, di.FullName)


        Next

        ButtonFilePath = Buttonfolder & "\" & d.Name & ".msb"

        KeyAssignmentsStore(ButtonFilePath)

    End Sub
    Public Function FindNextTier(tier As SortedList(Of String, DirectoryInfo)) As SortedList(Of String, DirectoryInfo)
        Dim i As Int16 = 0
        Dim nexttier As New SortedList(Of String, DirectoryInfo)
        For Each d In tier.Values
            Try
                For Each di In d.EnumerateDirectories
                    nexttier.Add(di.Name + Str(i), di)
                    i += 1

                Next
            Catch ex As Exception
                Continue For
            End Try
        Next
        Return nexttier
    End Function


    Public Sub AssignButton(ByVal ButtonNumber As Byte, ByVal ButtonLetter As Integer, ByVal Layer As Byte, ByVal Path As String, Optional Store As Boolean = False)
        Dim f As New DirectoryInfo(Path)
        'Static AskAgain As Boolean
        'If strVisibleButtons(i) <> "" And NavigateMoveState.State = StateHandler.StateOptions.Navigate Then
        '    If AskAgain Then
        '        If MsgBox("Replace button assignment for F" & i + 5 & " with " & f.Name & "?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Cancel Then
        '            Exit Sub
        '        End If
        '    End If
        'End If
        With buttons.CurrentRow.Buttons(ButtonNumber)
            Try
                .Path = Path
                .Label = f.Name

            Catch ex As Exception

            End Try
        End With
        strVisibleButtons(ButtonNumber) = Path
        strButtonFilePath(ButtonNumber, ButtonLetter, Layer) = Path

        lblDest(ButtonNumber).Text = f.Name
        strButtonCaptions(ButtonNumber, ButtonLetter, Layer) = f.Name
        UpdateButtonAppearance()
        If Store Then
            KeyAssignmentsStore(ButtonFilePath)
        End If
    End Sub

    Public Sub AssignButton(i As Byte, Path As String, ask As Boolean)
        Dim f As New DirectoryInfo(Path)
        With buttons.CurrentRow.Buttons(i)
            If .Path <> "" And ask Then
                If Not MsgBox("Replace button assignment for F" & i + 5 & " with " & f.Name & "?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then Exit Sub
            End If
            .Path = Path
            .Label = f.Name
        End With
    End Sub
    Public Sub AssignSpecialButton(Letter As Char, path As String)
        'Exit Sub
        'Dim f As New DirectoryInfo(path)
        'Dim oldletter As Integer
        'oldletter = AsciifromLetterNumber(iCurrentAlpha)
        ''        buttons.CurrentLetter = buttons.CurrentSet.Count - 1
        '' buttons.FirstFree(LetterNumberFromAscii(Asc(Letter))).Path = path
        ''AssignButton(buttonnumber, 35, 0, path)
        'ChangeButtonLetter(New KeyEventArgs(Keys.D0))
        'AssignButton(0, iCurrentAlpha, 0, path, True)
        'ChangeButtonLetter(New KeyEventArgs(oldletter))
        ''       buttons.CurrentLetter = oldletter
    End Sub

    Public Sub ChangeButtonLetter(e As KeyEventArgs)

        MainForm.lblAlpha.Text = e.KeyCode.ToString
        iCurrentAlpha = LetterNumberFromAscii(e.KeyCode)
        UpdateButtonAppearance()

    End Sub
    Public Sub SaveButtonlist()
        KeyAssignmentsStore("")
    End Sub


    ''' <summary>
    ''' This is for when we hold down CTRL or SHIFT and the appearance of the buttons schanges. 
    ''' </summary>
    Public Sub UpdateButtonAppearance()
        MainForm.lblAlpha.Text = Chr(AsciifromLetterNumber(iCurrentAlpha)).ToString
        For i = 0 To 7
            'buttons.CurrentRow.Buttons(i).Path = strButtonFilePath(i, iCurrentAlpha, 1)
            Try
                MainForm.ToolTip1.SetToolTip(btnDest(i), buttons.CurrentRow.Buttons(i).Path)

            Catch ex As Exception

            End Try
        Next
        For i = 0 To 7
            MainForm.lblAlpha.Text = Chr(AsciifromLetterNumber(iCurrentAlpha)).ToString
            Dim s As String
            Dim f As String = strButtonFilePath(i, iCurrentAlpha, 1)
            MainForm.ToolTip1.SetToolTip(btnDest(i), f)

            strVisibleButtons(i) = f
            s = strButtonCaptions(i, iCurrentAlpha, 1)
            If s <> "" Then
                lblDest(i).Text = s
                If My.Computer.FileSystem.DirectoryExists(f) Then
                    If MainForm.NavigateMoveState.State = StateHandler.StateOptions.Move Xor CtrlDown Then
                        lblDest(i).ForeColor = Color.Red
                    ElseIf ShiftDown Then
                        lblDest(i).ForeColor = Color.Blue
                    Else
                        lblDest(i).ForeColor = Color.Black
                    End If

                Else
                    Dim m = DirectoriesList.Find(Function(x) x.Contains(s))
                    If m IsNot Nothing Then
                        strButtonCaptions(i, iCurrentAlpha, 1) = s
                        strButtonFilePath(i, iCurrentAlpha, 1) = m
                    Else
                        lblDest(i).ForeColor = Color.Gray

                    End If

                End If

            Else
                lblDest(i).Text = "ABCDEFGH"(i)

            End If
        Next
    End Sub

    'Public Sub InitialiseButtons(row As ButtonRow)
    '    For i As Byte = 0 To 7
    '        With btnDest(i)
    '            .Text = "F" & Str(i + 5)

    '            AddHandler .Click, AddressOf ButtonClick
    '        End With

    '        lblDest(i).Text = row.Row(i).Label
    '    Next
    'End Sub
    Private Sub ButtonClick(sender As Object, e As MouseEventArgs)
        FolderSelect.Show()
        '   showPreview(sender, e)
    End Sub
    Public Function Autoload(Foldername As String) As String
        'If directory name exists as buttonfile, then load button file (optionally)

        'Check for button file
        Dim f As New IO.DirectoryInfo(Buttonfolder)
        Dim file As New IO.FileInfo(f.FullName & "\" & Foldername & ".msb")
        Static lastasked As String
        If file.Exists And file.FullName <> lastasked Then
            '     If MsgBox("Do you want to load the buttons?", MsgBoxStyle.YesNoCancel) = MsgBoxResult.Yes Then
            KeyAssignmentsRestore(file.FullName)
            Foldername = file.Name
            '     End If
            lastasked = file.FullName
        Else
            Foldername = ""
        End If
        Return Foldername
        'Load it (optionally)
    End Function
    Public Sub NewButtonList()
        ClearCurrentButtons()
        SaveButtonlist()

    End Sub
    Public Sub ClearCurrentButtons()
        Array.Clear(strButtonFilePath, 0, strButtonFilePath.Length)
        Array.Clear(strButtonCaptions, 0, strButtonCaptions.Length)
        Array.Clear(strVisibleButtons, 0, strVisibleButtons.Length)
        For i = 0 To 7
            lblDest(i).Text = ""

        Next
        ClearCurrentButtons(buttons)
    End Sub
    Public Sub ClearCurrentButtons(b As ButtonSet)
        For Each row In b.CurrentSet
            For Each btn In row.Buttons
                btn.Clear()
            Next
        Next
    End Sub
    Public Sub KeyAssignmentsRestore(Optional filename As String = "")

        Dim intIndex, intLetter As Integer
        Dim path As String
        If filename = "" Then
            path = MainForm.LoadButtonList()
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        ClearCurrentButtons()
        'Get the file path
        Dim fs As New List(Of String)
        ReadListfromFile(fs, path, Encrypted)
        Dim subs As String()
        For Each s In fs

            subs = s.Split("|")

            If subs.Length <> 4 Then
                'MsgBox("Not a button file")
                ' Exit Sub
            Else
                intIndex = Val(subs(0))
                intLetter = Val(subs(1))
                strButtonFilePath(intIndex, intLetter, 1) = (subs(2))
                strButtonCaptions(intIndex, intLetter, 1) = (subs(3))
                Dim m As New MVButton
                m.Path = (subs(2))
                m.Label = (subs(3))
            End If
        Next
        ButtonFilePath = path

        UpdateButtonAppearance()
        MainForm.UpdateFileInfo()
    End Sub
    Public Sub KeyAssignmentsStore(path As String)
        Dim intLoop As Integer
        Dim iLetter As Integer
        Dim Okay As Boolean
        If path = "" Then
            Okay = False
        Else
            Dim f As New FileInfo(path)
            If f.Exists Then Okay = True
        End If
        If Not Okay Then

            With MainForm.SaveFileDialog1
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"
                .FileName = path
                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                    ButtonFilePath = path
                Else
                    Exit Sub

                End If
            End With
        End If
        Dim fs As New List(Of String)
        Try
            Dim s As String
            For iLetter = 0 To iAlphaCount - 1
                For intLoop = 0 To 7

                    If strButtonFilePath(intLoop, iLetter, 1) <> "" Then
                        s = intLoop & "|" & iLetter & "|" & strButtonFilePath(intLoop, iLetter, 1) & "|" & strButtonCaptions(intLoop, iLetter, 1)
                        fs.Add(s)
                    End If
                Next
            Next
            WriteListToFile(fs, path, Encrypted)
        Catch ex As Exception

        End Try
        ButtonFilePath = path

        MainForm.UpdateFileInfo()
        PreferencesSave()

    End Sub

    Private Sub SaveASButton(path As String)
        Throw New NotImplementedException()
    End Sub

End Module

