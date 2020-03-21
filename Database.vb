Public Class Database
    Property Filename As String = "Q:\Database.txt"
    Property Encrypted As Boolean = False
    Property CurrentList As New List(Of String)
    Public Sub AddItems(list As List(Of String))
        CurrentList.Concat(list)
        WriteListToFile(CurrentList, Filename, Encrypted)

    End Sub
End Class
