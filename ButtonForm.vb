Imports System.IO
Public Class ButtonForm
    Public WithEvents buttons As New ButtonSet
    Private ofd As New OpenFileDialog
    Public sbtns As Button()
    Public lbls As Label()

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Function LoadButtonFileName() As String
        Dim path As String = ""
        With ofd
            .DefaultExt = "msb"
            .Filter = "Metavisua button files|*.msb|All files|*.*"

            If .ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                path = .FileName
            End If
        End With
        Return path

    End Function
    Private Sub ButtonForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sbtns = {Me.Button4, Me.Button2, Me.btn1, Me.Button1, Me.Button9, Me.Button7, Me.Button5, Me.Button6}
        lbls = {Me.Label4, Me.Label2, Me.lbl1, Me.Label1, Me.Label9, Me.Label7, Me.Label5, Me.Label6}
        TranscribeButtons(buttons.CurrentRow)
    End Sub
    Private Sub TranscribeButtons(m As ButtonRow)
        For i = 0 To 7
            sbtns(i).Text = m.Buttons(i).FaceText
            lbls(i).Text = m.Buttons(i).Label

        Next

    End Sub
    Private Sub OnLetterChanged(m As Integer) Handles buttons.LetterChanged
        lblAlpha.Text = Chr(AscfromButt(m))
        TranscribeButtons(buttons.CurrentRow)

    End Sub

    Private Sub RespondToKey(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.F5, Keys.F6, Keys.F7, Keys.F8, Keys.F9, Keys.F10, Keys.F11, Keys.F12
                If e.Shift Then
                    buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path = CurrentFolder
                    TranscribeButtons(buttons.CurrentRow)

                Else
                    Dim s As String = buttons.CurrentRow.Buttons(e.KeyCode - Keys.F5).Path
                    If s <> "" Then
                        CurrentFolder = s
                    Else
                        Exit Sub
                    End If
                    ChangeFolder(s)
                    'CancelDisplay()
                    MainForm.tvMain2.SelectedFolder = CurrentFolder
                End If

            Case Keys.A To Keys.Z, Keys.D0 To Keys.D9
                If e.Control Then
                    Select Case e.KeyCode
                        Case Keys.L
                            KeyAssignmentsRestore(LoadButtonFileName)
                    End Select
                Else

                    buttons.CurrentLetter = ButtfromAsc(e.KeyCode)
                End If
            Case Else
                MainForm.frmMain_KeyDown(sender, e)

        End Select
    End Sub
    Public Sub KeyAssignmentsRestore(Optional filename As String = "")

        Dim path As String
        If filename = "" Then
            path = LoadButtonFileName()
        Else
            path = filename
        End If
        If path = "" Then Exit Sub
        ClearCurrentButtons(buttons)
        'Get the file path
        Dim fs As New StreamReader(New FileStream(path, FileMode.OpenOrCreate, FileAccess.Read))
        Dim s As String
        Dim btnList As New List(Of MVButton)
        Dim subs As String()
        Do While fs.Peek <> -1
            s = fs.ReadLine
            subs = s.Split("|")

            If subs.Length <> 4 Then
                'MsgBox("Not a button file")
                Exit Sub
            Else
                Dim m As New MVButton

                m.Position = (subs(0))
                m.Letter = (subs(1))
                'buttons.CurrentLetter

                m.Path = (subs(2))
                m.Label = (subs(3))
                AddLoadedButtonToSet(m)
                '    btnList.Add(m)
                '    row.Row(m.Position) = m
            End If
        Loop


        fs.Close()


    End Sub
    Private Sub AddLoadedButtonToSet(btn As MVButton)
        With buttons
            buttons.CurrentLetter = btn.Letter
            buttons.CurrentSet(btn.Letter).Current = True
            btn.FaceText = buttons.CurrentRow.Buttons(btn.Position).FaceText
            buttons.CurrentRow.Buttons(btn.Position) = btn

        End With
    End Sub
    Private Sub lblAlpha_Click(sender As Object, e As EventArgs) Handles lblAlpha.Click
        'buttons.CurrentSet.Insert(m, New ButtonRow)
    End Sub
End Class