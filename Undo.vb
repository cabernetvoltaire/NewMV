Public Class Undo
    Public Enum Functions
        MoveFiles
        MoveFolder
        DeleteFile
        DeleteFolder


    End Enum

    Private Structure Filesource
        Dim filename As String
        Dim source As String

    End Structure


    Private mSource As String
    Public Property Source() As String
        Get
            Return mSource
        End Get
        Set(ByVal value As String)
            mSource = value

        End Set
    End Property

    Private mDestination As String
    Public Property Destination() As String
        Get
            Return mDestination
        End Get
        Set(ByVal value As String)
            mDestination = value
        End Set
    End Property
    Private mList As New List(Of String)
    Public Property FileList() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))
            mList = value
        End Set
    End Property
    Private mAction As Functions
    Public Property Action() As Functions
        Get
            Return mAction
        End Get
        Set(ByVal value As Functions)
            mAction = value
        End Set
    End Property

    Public Sub Undo()
        Select Case mAction
            Case Functions.MoveFiles
                'Movefiles back from where moved to
                'Refresh listbox
            Case Functions.MoveFolder
                'Movefolder back
                'Refresh treeview
            Case Functions.DeleteFile

            Case Functions.DeleteFolder


        End Select
    End Sub
End Class
