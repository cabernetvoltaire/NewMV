Imports System.IO
Public Class FileHandler
    Private _FileList As List(Of String)
    Private _Destination As String

    Public Property FileList() As List(Of String)
        Get
            Return _FileList
        End Get
        Set(ByVal value As List(Of String))
            _FileList = value
        End Set
    End Property
    Public Property FileBox As FileboxHandler
    Public Property FaveMinder As FavouritesMinder
    Public Property State As StateHandler

    Public Property Destination As String
        Get
            Return _Destination
        End Get
        Set
            _Destination = Value
            Dim x As New IO.DirectoryInfo(_Destination)
            If x.Exists Then
            Else
                x.Create()
            End If
        End Set
    End Property

    Public Sub MoveFiles()
        For Each f In _FileList
            Dim file As New IO.FileInfo(f)
            FaveMinder.DestinationPath = Destination
            FaveMinder.CheckFile(file)
            Select Case State.State
                Case StateHandler.StateOptions.Copy
                    Try
                        file.CopyTo(CombinePathAndName(Destination, file))
                    Catch ex As IOException


                    End Try
                Case StateHandler.StateOptions.Move
                    If Destination = "" Then
                        file.Delete()
                    Else
                        file.MoveTo(CombinePathAndName(Destination, file))
                    End If
                Case StateHandler.StateOptions.Navigate
            End Select
        Next
        FileBox.RemoveItems(_FileList)



    End Sub

    Private Function CombinePathAndName(s As String, m As FileInfo) As String
        If s.EndsWith("\") Or s = "" Then
            s = s & m.Name

        Else
            s = s & "\" & m.Name

        End If
        Return s
    End Function
End Class
