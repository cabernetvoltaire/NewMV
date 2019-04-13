Public Class DuplicateHolder
    Private mName As String
    Private mdups As List(Of String)
    Public Property FilePath() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value
        End Set
    End Property
    Public Property DuplicatesList() As List(Of String)
        Get
            Return mdups
        End Get
        Set(ByVal value As List(Of String))
            mdups = value
        End Set
    End Property
    Public Sub New(filePath As String)
        Me.FilePath = filePath

    End Sub
    Public Sub AddDuplicate(fpath As String)
        mdups.Add(fpath)
    End Sub
    Public Sub RemoveDuplicate(fpath As String)
        mdups.Remove(fpath)
    End Sub
End Class
