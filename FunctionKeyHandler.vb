Public Class FunctionKeyHandler
    Public e As KeyEventArgs
    Public FileList As New List(Of String)
    Public KeyNumber As Byte = e.KeyCode = Keys.F5
    Public NavState As StateHandler
    Public WithEvents ButtonRow As New ButtonRow
    Public ButtonFilePath As String
    Public Rand As RandomHandler
    Public ListBoxHandler As ListBoxHandler

    Public Event Changefolder(path As String)

    Private mCurrentPath As String
    Public Property CurrentFolder As String
        Get
            CurrentFolder = mCurrentPath
        End Get
        Set(value As String)
            mCurrentPath = value
        End Set
    End Property



#Region "Methods"
    Friend Sub Assign()
        ButtonRow.Buttons(KeyNumber).Path = mCurrentPath
        KeyAssignmentsStore(ButtonFilePath)
    End Sub

    Friend Sub Switch()
        Dim s As String = ButtonRow.Buttons(KeyNumber).Path
        If mCurrentPath <> s Then
            mCurrentPath = ButtonRow.Buttons(KeyNumber).Path
        Else
            If Rand.OnDirChange Then

            End If
        End If
    End Sub

    Friend Sub FilesMove()
        Dim s As String = ButtonRow.Buttons(KeyNumber).Path
        MoveFiles(FileList, s, ListBoxHandler.ListBox)
    End Sub
    Friend Sub FolderMove()
        Dim s As String = ButtonRow.Buttons(KeyNumber).Path
        MoveFolder(mCurrentPath, s)
    End Sub

#End Region

#Region "Events"

#End Region


    'Dim i As Byte = e.KeyCode - Keys.F5
    'Dim s As StateHandler.StateOptions = NavigateMoveState.State
    ''  CancelDisplay() 'Need to cancel display to prevent 'already in use' problems when moving files or deleting them. 
    'If (e.Shift And e.Control And e.Alt) Or strVisibleButtons(i) = "" Then
    '        'Assign button
    '        AssignButton(i, iCurrentAlpha, 1, CurrentFolder, True) 'Just assign in all modes when all three control buttons held
    '        'Always update the button file. 
    '        If My.Computer.FileSystem.FileExists(ButtonFilePath) Then
    '            KeyAssignmentsStore(ButtonFilePath)
    '        Else
    '            SaveButtonlist()
    '        End If
    'Exit Sub
    'End If
    'If s <> StateHandler.StateOptions.Navigate Then
    ''Non navigate behaviour

    'If e.Control And e.Shift Then
    ''Jump to folder
    'If strVisibleButtons(i) <> CurrentFolder Then
    '                ChangeFolder(strVisibleButtons(i))
    '                tvMain2.SelectedFolder = CurrentFolder
    '            ElseIf Random.OnDirChange Then
    '                AdvanceFile(True, True) 'TODO: Whaat?
    '            End If
    'ElseIf e.Shift Then
    '            MovingFolder(tvMain2.SelectedFolder, strVisibleButtons(i))
    '        Else

    'If lbxShowList.Visible Then
    '                MoveFiles(ListfromSelectedInListbox(lbxShowList), strVisibleButtons(i), lbxShowList)
    '            Else
    '                MoveFiles(ListfromSelectedInListbox(lbxFiles), strVisibleButtons(i), lbxFiles)
    '            End If
    'End If
    'Else
    ''Navigate behaviour
    'If e.Shift And e.Control And strVisibleButtons(i) <> "" Then
    '            MovingFolder(tvMain2.SelectedFolder, strVisibleButtons(i))

    '        ElseIf e.Shift Then
    '            MoveFiles(ListfromSelectedInListbox(lbxFiles), strVisibleButtons(i), lbxFiles)
    '        Else
    ''SWITCH folder
    'If strVisibleButtons(i) <> CurrentFolder Then
    '                ChangeFolder(strVisibleButtons(i))
    '                'CancelDisplay()
    '                tvMain2.SelectedFolder = strVisibleButtons(i)

    '            ElseIf Random.OnDirChange Then
    '                AdvanceFile(True, True)

    '            End If
    'End If
    'End If

    '    SetControlColours(NavigateMoveState.Colour, CurrentFilterState.Colour)
End Class
