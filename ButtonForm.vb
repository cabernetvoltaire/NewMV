Imports System.IO
Public Class ButtonForm

    '    Public WithEvents buttons As New ButtonSet
    Public WithEvents handler As New ButtonHandler
    Private ofd As New OpenFileDialog
    Private sfd As New SaveFileDialog
    Public sbtns As Button()
    Public lbls As Label()

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
        ' FocusOnMain()
        handler.LoadButtonSet(LoadButtonFileName(ButtonFilePath))
        ' buttons = handler.buttons
        handler.buttons.CurrentLetter = LetterNumberFromAscii(Asc("A"))
        TranscribeButtons(buttons.CurrentRow)
        For i = 0 To sbtns.Length - 1
            AddHandler sbtns(i).KeyDown, AddressOf MainForm.HandleFunctionKeyDown
        Next
    End Sub
    Private Sub FocusOnMain()
        With MainForm
            sbtns = { .btn1, .btn2, .btn3, .btn4, .btn5, .btn6, .btn7, .btn8}
            lbls = { .lbl1, .lbl2, .lbl3, .lbl4, .lbl5, .lbl6, .lbl7, .lbl8}
        End With

    End Sub


    Public Sub TranscribeButtons(m As ButtonRow)
        For i = 0 To 7
            sbtns(i).Text = m.Buttons(i).FaceText
            lbls(i).Text = m.Buttons(i).Label
            lbls(i).ForeColor = m.Buttons(i).Colour
            ToolTip1.SetToolTip(sbtns(i), buttons.CurrentRow.Buttons(i).Path)
        Next

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
                    handler.buttons.CurrentLetter = buttons.CurrentLetter

                    Select Case e.KeyCode
                        Case Keys.L
                            handler.AssignLinear(New DirectoryInfo(CurrentFolder), buttons)
                        Case Keys.A
                            handler.AssignAlphabetical(New DirectoryInfo(CurrentFolder), buttons)
                        Case Keys.T
                            handler.AssignTreeNew(CurrentFolder, 5)
                    End Select
                End If

                If e.Control AndAlso Not e.Alt Then

                    Select Case e.KeyCode
                        Case Keys.L
                            handler.LoadButtonSet(LoadButtonFileName(""))
                        Case Keys.S
                            handler.SaveButtonSet(SaveButtonFileName)
                    End Select
                Else
                    If buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode) Then
                        buttons.NextRow(LetterNumberFromAscii(e.KeyCode))
                    Else
                        buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode)
                    End If
                    lblAlpha.Text = Chr(AsciifromLetterNumber(handler.buttons.CurrentLetter))
                    TranscribeButtons(buttons.CurrentRow)
                    pbrButtons.Maximum = buttons.RowIndexCount
                    pbrButtons.Value = buttons.RowIndex + 1
                    'OnLetterChanged(buttons.CurrentLetter, buttons.RowIndex)

                End If
                buttons = handler.buttons

                MainForm.Main_KeyDown(sender, e)
                e.Handled = True
            Case Else
                MainForm.Main_KeyDown(sender, e)

        End Select
    End Sub
    'Private Sub OnLetterChanged(letter As Integer, count As Integer) Handles handler.buttons.LetterChanged
    '    lblAlpha.Text = Chr(AsciifromLetterNumber(letter))
    '    TranscribeButtons(buttons.CurrentRow)
    '    pbrButtons.Maximum = buttons.RowIndexCount
    '    pbrButtons.Value = buttons.RowIndex + 1
    'End Sub


End Class

