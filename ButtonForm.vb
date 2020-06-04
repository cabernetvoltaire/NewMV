﻿Imports System.IO
Public Class ButtonForm

    '    Public WithEvents buttons As New ButtonSet
    Public WithEvents handler As New ButtonHandler
    Private ofd As New OpenFileDialog
    Private sfd As New SaveFileDialog


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
        handler.Tooltip = Me.ToolTip1
        handler.ActualButtons = {Button1, Button2, Button3a, Button4, Button5, Button6, Button7, Button8a}
        handler.Labels = {Label1, Label2, Label3, Label4, Label5, Label6, Label7, Label8}
    End Sub
    Private Sub ButtonForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' FocusOnMain()
        handler.LoadButtonSet(LoadButtonFileName(ButtonFilePath))
        ' buttons = handler.buttons
        handler.buttons.CurrentLetter = LetterNumberFromAscii(Asc("A"))
        handler.TranscribeButtons(buttons.CurrentRow)
        ' AddHandler Me.KeyDown, AddressOf MainForm.HandleKeys

    End Sub
    Private Sub FocusOnMain()
        With MainForm
            'sbtns = { .btn1, .btn2, .btn3, .btn4, .btn5, .btn6, .btn7, .btn8}
            'lbls = { .lbl1, .lbl2, .lbl3, .lbl4, .lbl5, .lbl6, .lbl7, .lbl8}
        End With

    End Sub




    Private Sub RespondToKey(sender As Object, e As KeyEventArgs) Handles MyBase.KeyDown
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                Debug.Print(e.KeyCode.ToString)
                If e.Shift AndAlso e.Alt AndAlso e.Control Then
                    buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                    handler.TranscribeButtons(buttons.CurrentRow)

                ElseIf e.Shift Then
                    MainForm.HandleFunctionKeyDown(sender, e)
                Else

                    Dim s As String = buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path
                    If s <> "" Then
                        CurrentFolder = s
                    Else
                        buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                        handler.TranscribeButtons(buttons.CurrentRow)
                    End If
                    MainForm.tvMain2.SelectedFolder = CurrentFolder
                End If

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                buttons = handler.buttons
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
                    handler.TranscribeButtons(buttons.CurrentRow)
                    pbrButtons.Maximum = buttons.RowIndexCount
                    pbrButtons.Value = buttons.RowIndex + 1
                    'OnLetterChanged(buttons.CurrentLetter, buttons.RowIndex)

                End If

                'MainForm.Main_KeyDown(sender, e)
                e.Handled = True
            Case Else
                MainForm.Main_KeyDown(sender, e)

        End Select
    End Sub


End Class

