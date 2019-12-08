Imports System.IO
Public Class ButtonHandler
    Public Sub LoadButtonSet(Optional filename As String = "")

        Dim path As String
        If filename = "" Then
            path = LoadButtonFileName("")
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        ClearCurrentButtons(buttons)
        'Get the file path

        Dim btnList As New List(Of String)
        ReadListfromFile(btnList, path, True)
        Dim subs As String()
        For Each s In btnList
            subs = s.Split("|")
            If subs.Length <> 4 Then
                MsgBox("Not a button file")
                Exit Sub
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
    Public Sub SaveButtonSet(Optional filename As String = "")
        Dim path As String
        If filename = "" Then
            path = LoadButtonFileName("")
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

    End Sub
    Public Sub AddLoadedButtonToSet(btn As MVButton)
        With buttons
            buttons.CurrentLetter = btn.Letter
            buttons.CurrentSet(btn.Letter).Current = True
            btn.FaceText = buttons.CurrentRow.Buttons(btn.Position).FaceText
            If Not buttons.CurrentRow.Buttons(btn.Position).Empty Then
                buttons.InsertRow(btn.Letter)
            End If
            buttons.CurrentRow.Buttons(btn.Position) = btn

        End With
    End Sub
    Public Sub AssignLinear(e As DirectoryInfo, btnset As ButtonSet)
        Dim i = 0
        For Each d In e.EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
            Dim btn As MVButton = btnset.CurrentRow.Buttons(i)
            If Not btn.Empty Then
                btnset.InsertRow(btnset.CurrentLetter)
            End If
            btnset.CurrentRow.Buttons(i).Path = d.FullName
            i = (i + 1) Mod 8
        Next
    End Sub
    Public Sub AssignAlphabetical(e As DirectoryInfo, btnset As ButtonSet)
        btnset.Clear()
        Dim i = 0
        Dim lst As New Dictionary(Of String, String)
        For Each d In e.EnumerateDirectories("*", SearchOption.AllDirectories)
            lst.Add(d.FullName, d.Name)
        Next
        lst = lst.OrderBy(Function(x) x.Value).ToDictionary(Function(x) x.Key, Function(x) x.Value)
        For Each d In lst

            Dim m As Char = UCase(d.Value(0))
            m = UCase(m)
            btnset.CurrentLetter = LetterNumberFromAscii(Asc(m))
            Dim btn As MVButton
            btn = btnset.FirstFree(btnset.CurrentLetter)
            btn.Path = d.Key
        Next
    End Sub
    Public Sub AssignTree(e As DirectoryInfo, btnset As ButtonSet)
        Dim i = 0
        Dim lst As New Dictionary(Of String, String)
        For Each d In e.EnumerateDirectories("*", SearchOption.TopDirectoryOnly).Where(Function(x) x.Attributes <> FileAttributes.System)

            lst.Add(d.FullName, d.Name)
        Next
        For Each d In lst

            Dim m As Char = UCase(d.Value(0))
            m = UCase(m)
            btnset.CurrentLetter = LetterNumberFromAscii(Asc(m))
            Dim btn As MVButton
            btn = btnset.FirstFree(btnset.CurrentLetter)
            btn.Path = d.Key
        Next
        For Each d In e.EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
            AssignTree(d, btnset)
        Next
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
End Class
