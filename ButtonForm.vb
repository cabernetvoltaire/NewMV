Imports System.IO
Public Class ButtonForm

    Public WithEvents buttons As New ButtonSet
    Public WithEvents handler As New ButtonHandler

    Public ButtonFilePath As String
    Private ofd As New OpenFileDialog
    Private sfd As New SaveFileDialog
    Public Sub OnLetterChanged(sender As Object, e As EventArgs)
        lblAlpha.Text = Chr(AsciifromLetterNumber(buttons.CurrentLetter))
        RefreshAppearance
        UpdateLabelColours()
    End Sub

    Private Sub RefreshAppearance()
        UpdateTinyButtons()
        UpdateLabelColours()
        handler.UpdateButtonAppearance()
        ' handler.SaveButtonSet(handler.ButtonfilePath)
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        handler.Tooltip = Me.ToolTip1
        handler.ActualButtons = {Button1, Button2, Button3a, Button4, Button5, Button6, Button7, Button8a}
        handler.Labels = {Label1, Label2, Label3, Label4, Label5, Label6, Label7, Label8}
        handler.RowProgressBar = pbrButtons
    End Sub
    Private Sub UpdateTinyButtons()
        For Each btn In FlowLayoutPanel1.Controls
            Dim b As Button
            b = DirectCast(btn, Button)
            If buttons.CurrentSet(b.Tag).Active Then
                b.Font = New Font(b.Font, FontStyle.Bold)
                'make font bold
            Else
                b.Font = New Font(b.Font, FontStyle.Regular)
            End If
        Next

    End Sub

    Private Sub DrawTinyButtons()
        For i = 0 To 35
            Dim btn As New Button With {.Width = 20, .Height = 20, .Margin = New Padding(0), .Text = UCase(Chr(AsciifromLetterNumber(i)))}
            If buttons.CurrentSet(i).Active Then
                btn.Font = New Font(btn.Font, FontStyle.Bold)
                'make font bold
            Else
                btn.Font = New Font(btn.Font, FontStyle.Regular)
            End If
            btn.Tag = i
            AddHandler btn.MouseClick, AddressOf ChangeLetter
            FlowLayoutPanel1.Controls.Add(btn)
        Next
        ' Throw New NotImplementedException()
    End Sub

    Private Sub ChangeLetter(sender As Object, e As MouseEventArgs)
        sender = DirectCast(sender, Button)
        buttons.CurrentLetter = sender.tag
        handler.SwitchRow(buttons.CurrentRow)
        RefreshAppearance()
    End Sub

    Private Sub ButtonForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' FocusOnMain()
        handler.LoadButtonSet(LoadButtonFileName(ButtonFilePath))
        buttons = handler.buttons
        buttons.CurrentLetter = iCurrentAlpha
        handler.SwitchRow(buttons.CurrentRow)
        Me.Text = handler.ButtonfilePath & " - " & buttons.CurrentSet.Count & " button rows. Use alpha keys to change letter, function keys to access."
        DrawTinyButtons()
        UpdateLabelColours
        ' AddHandler Me.KeyDown, AddressOf MainForm.HandleKeys
        AddHandler buttons.LetterChanged, AddressOf OnLetterChanged


    End Sub

    Private Sub UpdateLabelColours()
        For i = 0 To 7
            handler.Labels(i).ForeColor = buttons.CurrentRow.Buttons(i).Colour
            'make font bold
        Next
        'Throw New NotImplementedException()
    End Sub

    Private Sub OnButtonFileChanged(sender As Object, e As EventArgs) Handles handler.ButtonFileChanged
        Me.Text = handler.ButtonfilePath & " - " & buttons.CurrentSet.Count & " button rows."
        UpdateTinyButtons()
    End Sub

    Private Sub RespondToKey(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                Debug.Print(e.KeyCode.ToString)
                If e.Shift AndAlso e.Alt AndAlso e.Control Then
                    'Assign button
                    buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                    handler.SwitchRow(buttons.CurrentRow)
                ElseIf e.Shift Then
                    MainForm.HandleFunctionKeyDown(sender, e)
                Else

                    Dim s As String = buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path
                    If s <> "" Then
                        CurrentFolder = s
                    Else
                        buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                        handler.SwitchRow(buttons.CurrentRow)
                    End If
                    MainForm.tvmain2.SelectedFolder = CurrentFolder
                End If

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9

                If e.Control AndAlso e.Alt Then
                    handler.buttons.CurrentLetter = buttons.CurrentLetter

                    Select Case e.KeyCode
                        Case Keys.L
                            handler.AssignLinear(New DirectoryInfo(CurrentFolder))
                        Case Keys.A
                            handler.AssignAlphabetical(New DirectoryInfo(CurrentFolder))
                        Case Keys.T
                            handler.AssignTreeNew(CurrentFolder, 5)
                    End Select
                End If

                If e.Control AndAlso Not e.Alt Then

                    Select Case e.KeyCode
                        Case Keys.L
                            handler.LoadButtonSet(LoadButtonFileName(""))
                        Case Keys.S
                            handler.SaveButtonSet(SaveButtonFileName(ButtonFilePath))
                        Case Keys.N
                            handler.SaveButtonSet("", True)

                    End Select
                Else
                    'Change letter
                    If buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode) Then
                        buttons.NextRow(LetterNumberFromAscii(e.KeyCode))
                    Else
                        buttons.CurrentLetter = LetterNumberFromAscii(e.KeyCode)
                    End If
                    lblAlpha.Text = Chr(AsciifromLetterNumber(handler.buttons.CurrentLetter))
                    handler.SwitchRow(buttons.CurrentRow)
                    'pbrButtons.Maximum = buttons.RowIndexCount
                    'pbrButtons.Value = buttons.RowIndex + 1
                    'OnLetterChanged(buttons.CurrentLetter, buttons.RowIndex)

                End If

            Case Else
                MainForm.Main_KeyDown(sender, e)


        End Select
        RefreshAppearance()
        e.Handled = True
    End Sub


End Class

