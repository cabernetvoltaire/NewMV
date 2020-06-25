Imports System.IO
Imports System.Threading
Public Class ButtonHandler
    Public WithEvents buttons As New ButtonSet
    Public RowProgressBar As New ProgressBar

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
        RaiseEvent LetterChanged(sender, e)
    End Sub
    Public Sub LoadButtonSet(Optional filename As String = "")

        Dim path As String
        If filename = "" Then
            path = LoadButtonFileName("")
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        buttons.Clear()
        'Get the file path

        Dim btnList As New List(Of String)
        btnList = ReadListfromFile(path, True)
        LoadListIn(btnList)
        'Dim t = New Thread(New ThreadStart(Sub() LoadListIn(btnList)))
        't.IsBackground = True
        't.SetApartmentState(ApartmentState.STA)
        't.Start()
        'While t.IsAlive
        '    Thread.Sleep(100)
        'End While
        ButtonfilePath = path
        RaiseEvent ButtonFileChanged(Me, Nothing)

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
                Dim m As New MVButton
                m.Position = (subs(0))
                m.Letter = (subs(1))
                m.Path = (subs(2))
                m.Label = (subs(3))

                AddLoadedButtonToSet(m)

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
            If My.Computer.FileSystem.DirectoryExists(btn.Path) Then
                If NMS.State = StateHandler.StateOptions.Move Xor CtrlDown Then
                    btn.Colour = Color.Red
                ElseIf ShiftDown Then
                    btn.Colour = Color.Blue
                Else
                    btn.Colour = Color.Black
                End If
            Else
                btn.Colour = Color.Gray
            End If
        Next
        LetterLabel.Text = Chr(AsciifromLetterNumber(buttons.CurrentRow.Letter))
    End Sub
    Private Sub AddLoadedButtonToSet(btn As MVButton)
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
    ''' <param name="buttons"></param>
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
    Public Sub AssignAlphabetical(e As DirectoryInfo)
        buttons.Clear()
        Dim i = 0
        Dim lst As New Dictionary(Of String, String)
        For Each dx In GenerateSafeFolderList(e.FullName)
            Dim d As New DirectoryInfo(dx)
            lst.Add(d.FullName, d.Name)
        Next
        lst = lst.OrderBy(Function(x) x.Value).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        For Each d In lst

            Dim m As Char = UCase(d.Value(0))
            m = UCase(m)
            buttons.CurrentLetter = LetterNumberFromAscii(Asc(m))
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
        Dim dlist As New SortedList(Of Long, DirectoryInfo)(icomp)

        For Each di In d.EnumerateDirectories("*", searchOption:=IO.SearchOption.AllDirectories)

            Dim disize = GetDirSize(di.FullName, 0, True)
            If (exclude = "" Or Not di.Name.Contains(exclude)) And disize > 10 ^ SizeMagnitude Then
                'MsgBox(di.Name & " is " & Format(GetDirSize(di.FullName, 0), "###,###,###,###,###.#"))
                While dlist.Keys.Contains(disize)
                    disize += 1
                End While
                dlist.Add(disize, di)
            End If

        Next
        Dim i As Int16 = 0
        For Each de In dlist

            Dim m As Char = UCase(de.Value.Name(0))
            m = UCase(m)
            buttons.CurrentLetter = LetterNumberFromAscii(Asc(m))
            Dim btn As MVButton
            btn = buttons.FirstFree(buttons.CurrentLetter)
            btn.Path = de.Value.FullName
        Next
        'For Each dx In dlist
        '    Dim di As IO.DirectoryInfo
        '    di = dx.Value
        '    Dim l As String = UCase(di.Name(0))
        '    Dim ButtonNumber As Integer = LetterNumberFromAscii(Asc(l))
        '    buttons.CurrentLetter = ButtonNumber
        '    Dim firstbtn As New MVButton()
        '    firstbtn.Letter = ButtonNumber
        '    firstbtn.Position = buttons.CurrentRow.GetFirstFree()
        '    firstbtn.Path = di.FullName
        '    If firstbtn.Position < 8 Then AssignButton(firstbtn.Position, firstbtn.Letter, 1, di.FullName)


        'Next
        RaiseEvent ButtonsAltered(Me, Nothing)


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
    Public Sub SwitchRow(m As ButtonRow)
        For i = 0 To 7
            ActualButtons(i).Text = m.Buttons(i).FaceText
            Labels(i).Text = m.Buttons(i).Label
            Tooltip.SetToolTip(ActualButtons(i), buttons.CurrentRow.Buttons(i).Path)
        Next
        RowProgressBar.Maximum = buttons.RowIndexCount
        RowProgressBar.Value = buttons.RowIndex + 1
    End Sub
    Friend Sub RefreshButtons()
        SwitchRow(buttons.CurrentRow)
    End Sub

    Public Sub InitialiseActualButtons()
        For i = 0 To 7
            ActualButtons(i).Tag = i
            '            AddHandler ActualButtons(i).MouseClick, AddressOf ShowPreview
            AddHandler ActualButtons(i).MouseHover, AddressOf ShowPreview
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
                f.close
                Exit For
            End If
        Next
    End Sub

    Public Sub ShowPreview(Sender As Object, e As EventArgs)
        Dim folderselect As New FormFolderSelect With {.Name = "FS"}

        ' Exit Sub
        Dim index As Byte = Val(Sender.tag)

        folderselect.ButtonNumber = index
        folderselect.Alpha = iCurrentAlpha
        Dim s As String = buttons.CurrentRow.Buttons(index).Path
        If s = "" Then s = CurrentFolder
        folderselect.Folder = s
        Dim control As Control = CType(Sender, Control)
        Dim startpoint As Point
        startpoint.X = control.Left
        startpoint.Y = control.Top
        startpoint = control.PointToScreen(startpoint)
        folderselect.Show()
        folderselect.BringToFront()
        folderselect.Left = startpoint.X - folderselect.Width / 2
        folderselect.Top = startpoint.Y - folderselect.Height + control.Height / 10
    End Sub
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

End Class
