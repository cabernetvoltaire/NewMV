Public Class MediaInfo
    Public Property FullName As String
    Public Property Name As String

    Public Property Type As Filetype
    Public Property Duration As Long = 0
    Public Property Links As New List(Of String)

    Public Sub AddLinks(SentLinks As List(Of String))
        For Each l In SentLinks
            Links.Add(l)

        Next

    End Sub

End Class
