Imports System.IO
Public Class ButtonForm
    Public WithEvents buttons As New ButtonSet
    Private ofd As New OpenFileDialog
    Private sfd As New SaveFileDialog
    Public sbtns As Button()
    Public lbls As Label()
    Public Function LoadButtonFileName(path As String) As String
        If path = "" Then
            With ofd
                .DefaultExt = "msb"
                .Filter = "Metavisua button files|*.msb|All files|*.*"

                If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                    path = .FileName
                End If
            End With
        End If
        Return path

    End Function
    Public Function SaveButtonFileName() As String
        Dim path As String = ""

        With sfd
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            End If
        End With
        Return path

    End Function

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub ButtonForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sbtns = {Me.Button4, Me.Button2, Me.btn1, Me.Button1, Me.Button9, Me.Button7, Me.Button5, Me.Button6}
        lbls = {Me.Label4, Me.Label2, Me.lbl1, Me.Label1, Me.Label9, Me.Label7, Me.Label5, Me.Label6}
        LoadButtonSet(LoadButtonFileName(ButtonFilePath))
        ' buttons.CurrentLetter = LetterNumberFromAscii(Asc("A"))
        'TranscribeButtons(buttons.CurrentRow)
    End Sub
    Public Sub TranscribeButtons(m As ButtonRow)
        For i = 0 To 7
            sbtns(i).Text = m.Buttons(i).FaceText
            lbls(i).Text = m.Buttons(i).Label
            lbls(i).ForeColor = m.Buttons(i).Colour
            ToolTip1.SetToolTip(sbtns(i), buttons.CurrentRow.Buttons(i).Path)
        Next

    End Sub
    Private Sub OnLetterChanged(m As Integer) Handles buttons.LetterChanged
        lblAlpha.Text = Chr(AsciifromLetterNumber(m))
        TranscribeButtons(buttons.CurrentRow)

    End Sub

    Private Sub RespondToKey(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                Debug.Print(e.KeyCode.ToString)
                If e.Shift AndAlso e.Alt AndAlso e.Control Then
                    buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                    TranscribeButtons(buttons.CurrentRow)

                ElseIf e.Shift Then
                    MainForm.HandleFunctionKeyDown(sender, e)
                Else

                    Dim s As String = buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path
                    If s <> "" Then
                        CurrentFolder = s
                    Else
                        Exit Sub
                    End If
                    MainForm.tvMain2.SelectedFolder = CurrentFolder
                End If

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If e.Control AndAlso e.Alt Then
                    Select Case e.KeyCode
                        Case Keys.L
                            AssignLinear(New DirectoryInfo(CurrentFolder), buttons)
                        Case Keys.A
                            AssignAlphabetical(New DirectoryInfo(CurrentFolder), buttons)
                        Case Keys.T
                            buttons.Clear()
                            AssignTree(New DirectoryInfo(CurrentFolder), buttons)
                    End Select
                End If

                If e.Control AndAlso Not e.Alt Then
                    Select Case e.KeyCode
                        Case Keys.L
                            LoadButtonSet(LoadButtonFileName(""))
                        Case Keys.S
                            SaveButtonSet(SaveButtonFileName)
                    End Select
                Else
                    If buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode) Then
                        buttons.NextRow(LetterNumberFromAscii(e.KeyCode))
                    Else
                        buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode)
                    End If
                    OnLetterChanged(buttons.CurrentLetter)

                End If
                e.Handled = True
            Case Else
                MainForm.frmMain_KeyDown(sender, e)

        End Select
    End Sub
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
                'MsgBox("Not a button file")
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
    Private Sub AddLoadedButtonToSet(btn As MVButton)
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
    Private Sub AssignLinear(e As DirectoryInfo, btnset As ButtonSet)
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
    Private Sub AssignAlphabetical(e As DirectoryInfo, btnset As ButtonSet)
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
    Private Sub AssignTree(e As DirectoryInfo, btnset As ButtonSet)
        Dim i = 0
        Dim lst As New Dictionary(Of String, String)
        For Each d In e.EnumerateDirectories("*", SearchOption.TopDirectoryOnly)
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
End Class