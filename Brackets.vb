Imports System.IO

Public Class BracketRemover
    Public ListOfFiles As New List(Of String)
    Public RemoveDuplicates As Boolean = True
    Private Sub RemoveNonBracketed()
        Dim tList As New List(Of String)
        For Each m In ListOfFiles
            Dim x As New IO.FileInfo(m)
            If m.EndsWith(")" & x.Extension) Then
                tList.Add(m)
            End If
        Next
        ListOfFiles = tList
    End Sub
    Public Sub RemoveBrackets()
        RemoveNonBracketed()

        For Each f In ListOfFiles
            Dim file As New IO.FileInfo(f)
            If file.Exists Then
                'Remove any bracketed numbers at the end of the filename
                Dim n As String = file.FullName
                'Need to use regex here.
                For i = 0 To 100
                    Dim p As String = Str(i).Substring(1)
                    p = n.Replace("(" & p & ")", "")
                    If n <> p Then
                        n = p
                        Exit For
                    End If
                Next
                'If it was successful, then rename the file without the brackets
                'but if this causes a clash with an existing file of that name
                'delete this file if permitted, and if their lengths are the same
                If file.FullName <> n Then
                    Dim nfile As New IO.FileInfo(n)
                    Try
                        My.Computer.FileSystem.RenameFile(file.FullName, nfile.Name)
                    Catch ex As IOException
                        If RemoveDuplicates Then
                            If nfile.Exists Then
                                If file.Length = nfile.Length Then
                                    file.Delete()
                                End If
                            End If
                        End If
                    End Try
                End If

            End If

        Next
    End Sub
End Class