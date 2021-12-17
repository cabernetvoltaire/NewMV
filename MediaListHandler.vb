Public Class MediaListHandler
    Private _List As New List(Of String)
    Private _ListLength As Integer
    Public Previous As String
    Public Current As String
    Public NextOne As String
    Private _Index As Integer

    Property Index As Integer
        Get
            Return _Index
        End Get
        Set
            _Index = Value

        End Set
    End Property

    Public Property List As List(Of String)
        Get
            Return _List
        End Get
        Set
            _List = Value
            _ListLength = _List.Count
        End Set
    End Property


    Public Function AdvanceList(Forward As Boolean) As String
        _Index = AdvanceModular(_ListLength, 1, Forward)
        Return List(_Index)

    End Function
    Public Function SetOfThree() As List(Of String)
        Dim ls As New List(Of String)
        Current = List(_Index)
        Previous = List(AdvanceModular(_ListLength, _Index, False))
        NextOne = List(AdvanceModular(_ListLength, _Index, True))
        ls.Add(Previous)
        ls.Add(Current)
        ls.Add(NextOne)
        Return ls
    End Function
End Class
