Public Class Dashboard

    Private Sub AddButton(name As String, size As Integer, shortcut As String, keycode As Integer, Col As Integer, Row As Integer)
        Dim d As New Button
        d.Width = size
        d.Height = size
        d.Text = name

        d.Tag = keycode
        d.ForeColor = Color.AntiqueWhite
        d.BackColor = Color.Black
        ToolTip1.SetToolTip(d, shortcut)
        TableLayoutPanel1.Controls.Add(d, Col, Row)
        AddHandler d.Click, AddressOf ButtonPress
    End Sub
    Private Sub ButtonPress(sender As Object, e As EventArgs)
        sender.backcolor = Color.AntiqueWhite
        MainForm.HandleKeys(sender, New KeyEventArgs(sender.tag))
        sender.backcolor = Color.Black
    End Sub
    Private Sub LoadButtons()
        Dim size As Integer = 75
        Dim numleft As Byte = 7
        Dim insert As Byte = 4
        Dim start As Byte = 0

        For i = 0 To 254
            Select Case i
                Case Keys.F2
                    AddButton("Edit folder name", size, "[F2]", Keys.F2, 1, 0)
                Case KeyDelete
                    AddButton("Delete file/folder", size, "[Delete]", KeyDelete, insert, 2)
                Case KeyTraverseTree
                    AddButton("Traverse FileTree", size, "[+]", KeyTraverseTree, 9, 3)
                Case KeyTraverseTreeBack
                    AddButton("Traverse FileTree" & vbCrLf & " (Back)", size, "[-]", KeyTraverseTreeBack, 9, 2)
                Case Keys.Left
                    AddButton("Collapse FileTree", size, "[+]/[-]", Keys.Left, insert, 5)
                Case Keys.Right
                    AddButton("Expand FileTree", size, "[Right]", Keys.Right, insert + 2, 5)
                Case Keys.Up
                    AddButton("Move up FileTree", size, "[Up]", Keys.Up, insert + 1, 4)
                Case Keys.Down
                    AddButton("Move down FileTree", size, "[Down]", Keys.Down, insert + 1, 5)
                Case KeyEscape
                    AddButton("Cancel display", size, "[ESC]", i, 0, 0)
                Case KeyNextFile
                    AddButton("Next file", size, "[PGDN]", i, numleft - 1, 2)
                Case KeyPreviousFile
                    AddButton("Previous file", size, "[PGUP]", i, numleft - 1, 1)
                Case KeySmallJumpUp
                    AddButton("Small Jump FWD", size, "[NUMPAD 6]", i, numleft + 2, 3)
                Case KeySmallJumpDown
                    AddButton("Small Jump BCK", size, "[NUMPAD 5]", i, numleft + 1, 3)
                Case KeyBigJumpOn
                    AddButton("Big Jump FWD", size, "[NUMPAD 9]", i, numleft + 2, 2)
                Case KeyBigJumpBack
                    AddButton("Big Jump BCK", size, "[NUMPAD 8]", i, numleft + 1, 2)
                Case KeyMarkFavourite
                    AddButton("Make Favourite", size, "[INSERT]", i, insert, 1)
                Case KeyJumpToPoint
                    AddButton("Jump to Marker", size, "[NUMPAD 7]", i, numleft, 2)
                Case KeyMarkPoint
                    AddButton("Make Marker", size, "[NUMPAD 4]", i, numleft, 3)
                Case KeyMuteToggle
                    AddButton("Toggle Mute", size, "[NUMPAD .]", i, numleft + 2, 5)
                Case KeyJumpAutoT
                    AddButton("Jump Random", size, "[NUMPAD /]", i, numleft + 1, 1)
                Case KeyToggleSpeed
                    AddButton("Toggle Speed", size, "[NUMPAD 0]", i, numleft, 5)
                Case KeySpeed1
                    AddButton("Slowest Speed", size, "[NUMPAD 1]", i, numleft, 4)
                Case KeySpeed2
                    AddButton("Medium Slow Speed", size, "[NUMPAD 2]", i, numleft + 1, 4)
                Case KeySpeed3
                    AddButton("Fast Slow Speed", size, "[NUMPAD 3]", i, numleft + 2, 4)
                Case KeyRotateBack
                    AddButton("Rotate Pic CCW", size, "[{]", i, 2, 2)
                Case KeyRotate
                    AddButton("Rotate Pic CW", size, "[}]", i, 3, 2)
                Case KeyRandomize
                    AddButton("Increment Play Order", size, "[PAUSE]", i, 5, 0)
                Case KeyFilter 'Cycle through listbox filters
                    AddButton("Increment Filter", size, "[?]", i, 2, 4)
                Case KeySelect
                    AddButton("Select Files", size, "[F3]", i, 2, 0)
                Case KeyFullscreen
                    AddButton("Go Fullscreen", size, "[SCROLL]", i, 5, 0)
                Case KeyMoveToggle
                    AddButton("Mode Toggle", size, "[#]", i, 3, 3)
                Case KeyTrueSize
                    AddButton("Pics Truesize", size, "[=+]", i, 2, 1)

                Case KeyBackUndo
                    AddButton("Back/Undo", size, "[Backspace]", i, 3, 1)
                Case KeyReStartSS
                    AddButton("SlideShow Toggle", size, "[SPACE]", i, 1, 5)

            End Select
            '            Me.Cursor = Cursors.Default
            '            ' e.suppresskeypress = True
        Next
    End Sub

    Private Sub Dashboard_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.Opacity = 0.4

        LoadButtons()
    End Sub

    Private Sub Dashboard_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        MainForm.HandleKeys(sender, e)
    End Sub
End Class