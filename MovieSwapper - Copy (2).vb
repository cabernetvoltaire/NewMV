Imports AxWMPLib
Public Class MediaSwapper2

    Public MediaHandlers As New List(Of MediaHandler)
    Private mListofFiles As List(Of String)
    Private _Current As String
    Private _NextOne As String
    Private _Previous As String
    Private _CurrentIndex As Integer = 0

    Public Property ListofFiles() As List(Of String)
        Get
            Return mListofFiles
        End Get
        Set(ByVal value As List(Of String))
            mListofFiles = value
        End Set
    End Property
    Public Property CurrentIndex As Integer
        Get
            Return _CurrentIndex
        End Get
        Set
            _CurrentIndex = Value
            _Current = mListofFiles(Value)
        End Set
    End Property

    Private Property _IndexOfNextHandler As Integer = 1
    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        '  HideMedias(MHX)

        MHX.PlaceResetter(False) 'Starts the video playing
        MHX.Visible = False
        With MHX.Player
            .Visible = True
            .BringToFront()
            .settings.mute = Muted
        End With


    End Sub
    Public Sub ShowFirst(Forward As Boolean)
        MediaHandlers(0).MediaPath = mListofFiles(CurrentIndex)
        MediaHandlers(1).MediaPath = mListofFiles(AdvanceModular(mListofFiles.Count, CurrentIndex, Forward))
        ShowPlayer(MediaHandlers(0))
    End Sub
    Public Sub AdvanceList(Forward As Boolean)
        'Advance current index

        CurrentIndex = AdvanceModular(mListofFiles.Count, CurrentIndex, Forward)
        'Get handler to put it in
        MediaHandlers(_IndexOfNextHandler).PlaceResetter(True)
        _IndexOfNextHandler = AdvanceModular(MediaHandlers.Count, _IndexOfNextHandler, Forward)
        'LoadNext
        With MediaHandlers(_IndexOfNextHandler)
            .MediaPath = NextOne
            .PlaceResetter(True)
        End With

        'ShowCurrent

        ShowPlayer(MediaHandlers(AdvanceModular(MediaHandlers.Count, _IndexOfNextHandler, False)))
    End Sub

    Public ReadOnly Property Previous As String
        Get
            _Previous = mListofFiles(AdvanceModular(mListofFiles.Count, CurrentIndex, False))
            Return _Previous
        End Get
    End Property

    Public Property NextOne As String
        Get
            _NextOne = mListofFiles(AdvanceModular(mListofFiles.Count, CurrentIndex, True))
            Return _NextOne
        End Get
        Set
            _NextOne = Value
        End Set
    End Property

    Public Property Current As String
        Get
            Return _Current
        End Get
        Set
            _Current = Value
        End Set
    End Property
End Class
