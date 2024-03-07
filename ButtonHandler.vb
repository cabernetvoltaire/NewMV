Imports System.Threading
Imports System.IO
Imports System.Threading.Tasks

Imports System.Windows.Forms
Public Class ButtonHandler
    Public isOverSpawnedForm As Boolean
    Public SubFormVisible As Boolean
    Public WithEvents Oldbuttons As New Stack(Of ButtonSet)
    Public WithEvents buttons As New ButtonSet

    Public RowProgressBar As New ProgressBar
    Public ButtonsLoaded As Boolean = False
    Public Autobuttons As Boolean
    Public Labels(8) As Label
    Public LetterLabel As Label
    Public Tooltip As ToolTip
    Public ButtonfilePath As String
    Public NMS As New StateHandler
    Private mFSOpen As Boolean = False
    Public Event ButtonFileChanged(sender As Object, e As EventArgs)
    Public Event LetterChanged(sender As Object, e As EventArgs)
    Public Event ButtonsAltered(sender As Object, e As EventArgs)
    Private mListOfButtonFiles As New List(Of String)
    Private mAlpha As String
    Public ActualButtons(8) As Button

    Public Property Alpha() As String
        Get
            Return mAlpha
        End Get
        Set(ByVal value As String)
            mAlpha = value
            buttons.CurrentLetter = mAlpha
        End Set
    End Property
    Private Sub OnButtonsAltered(sender As Object, e As EventArgs) Handles Me.ButtonsAltered
        SaveButtonSet(ButtonfilePath)
        SwitchRow(buttons.CurrentRow)
    End Sub
    Sub OnLetterChanged(sender As Object, e As EventArgs) Handles buttons.LetterChanged
        'RaiseEvent LetterChanged(sender, e)
    End Sub
    'Public Sub LoadButtonSet(Optional filename As String = "")
    '    Try

    '        Dim path As String
    '        Dim file As New IO.FileInfo(filename)

    '        If filename = "" Or Not file.Exists Then
    '            path = LoadButtonFileName("")
    '        Else
    '            path = filename
    '        End If
    '        If path = "" Then Exit Sub

    '        buttons.Clear()
    '        Dim btnList As New List(Of String)
    '        btnList = ReadListfromFile(path, True)
    '        LoadListIn(btnList)
    '        ButtonfilePath = path
    '        Dim dir As New IO.DirectoryInfo(Buttonfolder)
    '        For Each f In dir.EnumerateDirectories
    '            If f.FullName.EndsWith(".msb") Then
    '                mListOfButtonFiles.Add(f.FullName)
    '            End If
    '        Next
    '        RaiseEvent ButtonFileChanged(Me, Nothing)
    '        UpdateButtonAppearance()
    '    Catch ex As Exception
    '        MsgBox("Button load failed")
    '    End Try

    'End Sub

    Public Sub LoadButtonSetAsync(Optional filename As String = "")
        Task.Run(Sub() LoadButtonSet(filename))
    End Sub

    Private Sub LoadButtonSet(filename As String)
        Try
            Dim path As String
            Dim file As New IO.FileInfo(filename)

            If filename = "" Or Not file.Exists Then
                path = LoadButtonFileName("")
            Else
                path = filename
            End If
            If path = "" Then Exit Sub
            buttons.Clear()
            Dim btnList As New List(Of String)
            btnList = ReadListfromFile(path, True)
            LoadListIn(btnList)
            ButtonfilePath = path

            Dim dir As New IO.DirectoryInfo(Buttonfolder)
            For Each f In dir.EnumerateDirectories
                If f.FullName.EndsWith(".msb") Then
                    mListOfButtonFiles.Add(f.FullName)
                End If
            Next

            RaiseEvent ButtonFileChanged(Me, Nothing)
            UpdateButtonAppearance()

        Catch ex As Exception
            ' Marshalling the exception handling to the UI thread
            ' MsgBox("Button load failed")
        End Try
    End Sub


    Public Sub New()

    End Sub
    Private Sub LoadListIn(btnList As List(Of String))
        Dim subs As String()
        For Each s In btnList
            subs = s.Split("|")
            If subs.Length <> 4 Then
                'MsgBox("Not a button file")
                'Exit Sub
            Else
                Try

                    Dim m As New MVButton
                    m.Position = (subs(0))
                    m.Letter = (subs(1))
                    m.Path = (subs(2))
                    m.Label = (subs(3))

                    AddLoadedButtonToSet(m)


                Catch ex As Exception
                    'MsgBox("Button skipped")
                End Try
            End If
        Next
    End Sub
    ''' <summary>
    ''' Saves current button set to file.
    ''' </summary>
    ''' <param name="filename"> Filename to save under </param>
    ''' <param name="NewFile"> Option to force new file </param>
    Public Sub SaveButtonSet(Optional filename As String = "", Optional NewFile As Boolean = False)
        If NewFile Then buttons.Clear()

        Dim path As String
        If filename = "" Then
            path = SaveButtonFileName("")
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        'Get the file path
        Dim output As New List(Of String)

        For Each r In buttons.CurrentSet
            For Each b In r.Buttons
                output.Add(String.Format("{0}|{1}|{2}|{3}", b.Position, b.Letter, b.Path, b.Label))
            Next
        Next
        WriteListToFile(output, path, True)
        If ButtonfilePath <> path Then
            ButtonfilePath = path
            RaiseEvent ButtonFileChanged(Me, Nothing)
        End If
    End Sub

    Public Sub UpdateButtonAppearance()
        For Each btn In buttons.CurrentRow.Buttons
            If NMS.State = StateHandler.StateOptions.Move Xor CtrlDown Then
                btn.Colour = Color.Red
            ElseIf ShiftDown Then
                btn.Colour = Color.Blue
                'The following should make a list of all the files in the button directories (at the start)
                'and then just load it in as a list. Checking would then take no time at all. Too much overhead here. 
                'ElseIf My.Computer.FileSystem.DirectoryExists(btn.Path) Then
                '    If btn.Colour <> Color.Purple Then btn.Colour = Color.Black
                '    End If
            Else
                'btn.Colour = Color.Gray
            End If
            If mListOfButtonFiles.Contains(btn.Path) Then
                btn.Colour = Color.Purple
            End If
        Next

        LetterLabel.Text = Chr(AsciifromLetterNumber(buttons.CurrentRow.Letter))
    End Sub
    Private Sub AddLoadedButtonToSet(btn As MVButton)
        ButtonsLoaded = True
        With buttons
            'Switch to appropriate row
            buttons.CurrentLetter = btn.Letter
            buttons.CurrentRow = buttons.CurrentSet.FindLast(Function(x) x.Letter = btn.Letter) 'If there is another row with this btn.Letter, need to find it.
            'btn.FaceText = buttons.CurrentRow.Buttons(btn.Position).FaceText
            'If this row is full, insert one.
            If buttons.CurrentRow.GetFirstFree = 8 Then
                buttons.InsertRow(btn.Letter)
            End If
            'Put the button in its correct place, or, if not possible in the first free space. 
            If buttons.CurrentRow.Buttons(btn.Position).Empty Then
                buttons.CurrentRow.Buttons(btn.Position) = btn
            Else
                buttons.CurrentRow.Buttons(buttons.CurrentRow.GetFirstFree) = btn

            End If

        End With
    End Sub
    ''' <summary>
    ''' Assigns all subdirectories linearly to btnset
    ''' </summary>
    ''' <param name="e"></param>

    Public Sub AssignLinear(e As DirectoryInfo)
        Dim i = 0
        For Each d In e.EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
            Dim btn As MVButton = buttons.CurrentRow.Buttons(i)
            If Not btn.Empty Then
                buttons.InsertRow(buttons.CurrentLetter)
            End If
            buttons.CurrentRow.Buttons(i).Path = d.FullName
            i = (i + 1) Mod 8
        Next
        RaiseEvent ButtonsAltered(Me, Nothing)
    End Sub
    Public Function AssignAlphabetical(b As ButtonSet, findstring As String) As ButtonSet
        Dim lst As New Dictionary(Of String, String)
        Dim i = 0
        For Each rw In b.CurrentSet
            For i = 0 To 7
                Dim s = rw.Buttons(i).Path
                Dim l = rw.Buttons(i).Label
                If UCase(l).Contains(UCase(findstring)) Then
                    If Not lst.ContainsKey(s) Then
                        lst.Add(s, rw.Buttons(i).Label)
                    End If
                End If

            Next
        Next
        Dim newbuttons As New ButtonSet
        lst = lst.OrderBy(Function(x) x.Key.Count(Function(c As Char) c = "\")).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        i = 0
        For Each d In lst
            Dim btn As MVButton = newbuttons.CurrentRow.Buttons(i)
            If Not btn.Empty Then
                newbuttons.InsertRow(newbuttons.CurrentLetter)
            End If
            newbuttons.CurrentRow.Buttons(i).Path = d.Key
            newbuttons.CurrentRow.Buttons(i).Label = d.Value

            i = (i + 1) Mod 8

        Next
        'RaiseEvent ButtonsAltered(Me, Nothing)
        Return newbuttons
    End Function
    ''' <summary>
    ''' Assigns a list of paths to the buttons, as if it were assign linear. 
    ''' </summary>
    ''' <param name="s"></param>
    Public Sub AssignList(s As List(Of String))
        Dim i = 0
        For Each d In s
            Dim btn As MVButton = buttons.CurrentRow.Buttons(i)
            If Not btn.Empty Then
                buttons.InsertRow(buttons.CurrentLetter)
            End If
            buttons.CurrentRow.Buttons(i).Path = d
            i = (i + 1) Mod 8
        Next
        RaiseEvent ButtonsAltered(Me, Nothing)
    End Sub

    Public Sub AssignAlphabetical(e As DirectoryInfo)
        buttons.Clear()
        Dim i = 0
        Dim lst As New Dictionary(Of String, String)
        Dim ls As New List(Of String)
        CollectFoldersRecursively(e.FullName, ls)
        For Each dx In ls
            'For Each dx In GetAllFolders(e, "*")
            Dim d As New DirectoryInfo(dx)
            lst.Add(d.FullName, d.Name)
        Next
        lst = lst.OrderBy(Function(x) x.Key.Count(Function(c As Char) c = "\")).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        For Each d In lst

            Dim _fstletter As Char = UCase(d.Value(0))
            _fstletter = UCase(_fstletter)
            buttons.CurrentLetter = LetterNumberFromAscii(Asc(_fstletter))
            Dim btn As MVButton
            btn = buttons.FirstFree(buttons.CurrentLetter)
            btn.Path = d.Key
        Next
        RaiseEvent ButtonsAltered(Me, Nothing)

    End Sub


    Public Sub AssignTreeNew(StartingFolder As String, SizeMagnitude As Byte)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub

        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        buttons.Clear()
        Dim d As New DirectoryInfo(StartingFolder)

        Dim icomp As New MyComparer
        Dim dlist As New List(Of String)
        dlist = GenerateSafeFolderList(d.FullName, SizeMagnitude)

        Dim i As Int16 = 0
        For Each n In dlist

            Dim de As New IO.DirectoryInfo(n)
            If de.EnumerateFiles.Count <> 0 Then
                Dim m As Char = UCase(de.Name(0))
                m = UCase(m)
                buttons.CurrentLetter = LetterNumberFromAscii(Asc(m))
                Dim btn As MVButton
                btn = buttons.FirstFree(buttons.CurrentLetter)
                btn.Path = de.FullName
            End If
        Next


        RaiseEvent ButtonsAltered(Me, Nothing)


    End Sub
    Public Sub AssignAlphaHierarchical(StartingFolder As String, Depth As Byte, Optional NoEmpties As Boolean = False)
        If MsgBox("This will replace a large number of button assignments. Are you sure?", MsgBoxStyle.OkCancel) = MsgBoxResult.Cancel Then Exit Sub
        If MsgBox("Include empty folders?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then NoEmpties = True
        Dim exclude As String = ""
        exclude = InputBox("String to exclude from folders?", "")
        buttons.Clear()
        Dim d As New DirectoryInfo(StartingFolder)

        Dim icomp As New MyComparer
        Dim dlist As New List(Of FolderInfo)

        Dim sortedFolders = FolderSorter.GetSortedFolders(CurrentFolder, Depth)
        For Each n In sortedFolders
            If NoEmpties AndAlso n.Files = 0 Then
            Else
                Dim de As String = n.ShortName
                Dim m As Char = UCase(de(0))
                m = UCase(m)
                buttons.CurrentLetter = LetterNumberFromAscii(Asc(m))
                Dim btn As MVButton
                btn = buttons.FirstFree(buttons.CurrentLetter)
                btn.Path = n.Path
            End If
        Next

        RaiseEvent ButtonsAltered(Me, Nothing)


    End Sub
    ''' <summary>
    ''' Creates a dlist of directories whose size is larger than 10^SizeMagnitude under d
    ''' </summary>
    ''' <param name="SizeMagnitude"></param>
    ''' <param name="exclude"></param>
    ''' <param name="d"></param>
    ''' <param name="dlist"></param>
    Private Shared Sub CreateListOfLargeDirectories(SizeMagnitude As Byte, exclude As String, d As DirectoryInfo, dlist As SortedList(Of Long, DirectoryInfo))
        Dim dirlist As New List(Of String)
        dirlist = GenerateSafeFolderList(d.FullName, SizeMagnitude)

    End Sub
    Public Sub AssignButton(ByVal ButtonNumber As Byte, ByVal Path As String)
        Dim f As New DirectoryInfo(Path)

        With buttons.CurrentRow.Buttons(ButtonNumber)
            Try
                .Path = Path
                .Label = f.Name
                .Position = ButtonNumber
                .Letter = buttons.CurrentRow.Letter
                RaiseEvent ButtonsAltered(Me, Nothing)
            Catch ex As Exception

            End Try
        End With

    End Sub
    Public Sub SwitchRow(m As ButtonRow, Optional Backwards As Boolean = False)
        For i = 0 To 7
            'ActualButtons(i).Text = m.Buttons(i).FaceText
            Labels(i).Text = m.Buttons(i).Label
            Tooltip.SetToolTip(ActualButtons(i), buttons.CurrentRow.Buttons(i).Path)
        Next
        RowProgressBar.Maximum = buttons.RowIndexCount
        RowProgressBar.Value = buttons.RowIndex + 1
    End Sub
    Friend Sub RefreshButtons()
        SwitchRow(buttons.CurrentRow)
    End Sub
    Friend Sub FindRowContaining(s As String)
        For Each m In Me.buttons.CurrentSet
            For Each f In m.Buttons
                If Path.GetFileName(f.Path).IndexOf(s, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    buttons.CurrentRow = m
                    buttons.CurrentLetter = m.Letter
                    SwitchRow(buttons.CurrentRow)
                End If
            Next

        Next
    End Sub
    Public Sub InitialiseActualButtons()
        For i = 0 To 7
            ActualButtons(i).Tag = i
            ActualButtons(i).Text = "f" & Str(i + 5)
            '            AddHandler ActualButtons(i).MouseClick, AddressOf ShowPreview
            AddHandler ActualButtons(i).MouseHover, AddressOf ShowPreview
            AddHandler ActualButtons(i).MouseEnter, AddressOf ShowPreview

            AddHandler ActualButtons(i).MouseLeave, AddressOf HideFS
            AddHandler ActualButtons(i).MouseMove, AddressOf ChangeFS



        Next
    End Sub



    Private Sub ChangeFS(sender As Object, e As MouseEventArgs)
        Static mousepos As Point
        If Math.Abs(mousepos.X - e.X) > sender.width / 30 Then
            For Each f In Application.OpenForms
                If f.name = "FS" Then
                    f.UPdatePic
                    mousepos = e.Location
                    Exit For
                End If
            Next
        End If

    End Sub
    Private Sub HideFS(Sender As Object, e As EventArgs)

        For Each f In Application.OpenForms
            If f.name = "FS" Then
                f.timer1.enabled = True


                Exit For
            End If
        Next
    End Sub

    Public Sub ShowPreview(Sender As Object, e As EventArgs)
        Dim folderselect As FormFolderSelect

        SubFormVisible = True
        If mFSOpen Then
            For Each f In Application.OpenForms
                If f.name = "FS" Then
                    folderselect = AssignFolderSelect(Sender, f)
                    f.show
                    f.Timer1.enabled = False
                    Exit For
                End If
            Next
        Else
            folderselect = SpawnFolderSelect(Sender)



        End If
        Dim control As Control = CType(Sender, Control)
        Dim startpoint As Point
        startpoint.X = control.Left
        startpoint.Y = control.Top
        startpoint = control.PointToScreen(startpoint)
        folderselect.Show()
        folderselect.TextBox1.Text = ""
        folderselect.TextBox1.Focus()
        folderselect.BringToFront()
        folderselect.Left = startpoint.X - folderselect.Width / 2
        folderselect.Top = startpoint.Y - folderselect.Height + control.Height / 10
    End Sub



    Private Function SpawnFolderSelect(Sender As Object) As FormFolderSelect
        Dim folderselect As New FormFolderSelect With {.Name = "FS"}

        ' Exit Sub
        folderselect = AssignFolderSelect(Sender, folderselect)
        mFSOpen = True
        Return folderselect
    End Function

    Private Function AssignFolderSelect(Sender As Object, folderselect As FormFolderSelect) As FormFolderSelect
        Dim index As Byte = Val(Sender.tag)

        folderselect.ButtonNumber = index
        folderselect.Alpha = iCurrentAlpha
        Dim s As String = buttons.CurrentRow.Buttons(index).Path
        If s = "" Then s = CurrentFolder
        folderselect.Folder = s
        Return folderselect
    End Function

    Public Sub SwapPath(Oldpath As String, Newpath As String)
        Dim changed As Boolean = False
        For Each r In buttons.CurrentSet
            For Each b In r.Buttons
                If b.Path = Oldpath Then b.Path = Newpath
                changed = True
            Next
        Next
        SaveButtonSet(ButtonfilePath)
    End Sub
    Public Sub GetSubButtons(Findstring As String)
        '
        If Findstring = "" AndAlso Oldbuttons.Count <> 0 Then
            RestoreButtons()
        Else
            Oldbuttons.Push(buttons)
            buttons = AssignAlphabetical(buttons, Findstring)
        End If

    End Sub
    Public Sub RestoreButtons()
        buttons = Oldbuttons.Pop
    End Sub
End Class




Module FolderSorter
    Private Sub TraverseFolders(ByVal Folder As DirectoryInfo, ByVal SortedFolders As List(Of FolderInfo), ByVal MaxDepth As Integer)
        If MaxDepth < 0 Then Exit Sub ' Exit recursion if max depth is reached
        Dim NewFolder As New FolderInfo(Folder.FullName)
        SortedFolders.Add(NewFolder)
        For Each Subfolder As DirectoryInfo In Folder.GetDirectories()
            TraverseFolders(Subfolder, SortedFolders, MaxDepth - 1)
        Next
    End Sub

    Public Function GetSortedFolders(ByVal RootFolder As String, ByVal MaxDepth As Integer) As List(Of FolderInfo)
        Dim SortedFolders As New List(Of FolderInfo)
        Dim RootDir As New DirectoryInfo(RootFolder)
        TraverseFolders(RootDir, SortedFolders, MaxDepth)
        SortedFolders.Sort(Function(x, y)
                               Dim result = x.Path.Count(Function(c) c = "\"c) - y.Path.Count(Function(c) c = "\"c)
                               If result = 0 Then
                                   result = String.Compare(x.ShortName, y.ShortName)
                               End If
                               Return result
                           End Function)

        Return SortedFolders
    End Function
End Module

Public Class FolderInfo
    Public Property Path As String
    Public Property ShortName As String
    Public Property Depth As Integer
    Public Property Files As Integer

    Public Sub New(ByVal FolderPath As String)
        Path = FolderPath
        Dim d = New DirectoryInfo(FolderPath)
        ShortName = d.Name
        Depth = Path.Split(IO.Path.DirectorySeparatorChar).Count - 1
        Files = d.GetFiles.Length
    End Sub
End Class
