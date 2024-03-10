Public Class UndoNew

End Class
Public Interface ICommand
    Sub Execute()
    Sub Undo()
End Interface
Public Class FileMove
    Implements ICommand

    Public Sub Execute() Implements ICommand.Execute
        'Move Source to Dest
        Throw New NotImplementedException()
    End Sub

    Public Sub Undo() Implements ICommand.Undo
        'Move Dest to Source
        Throw New NotImplementedException()
    End Sub
End Class
