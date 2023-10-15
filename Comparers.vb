Imports System.IO
Public Class MyComparer
        Implements Generic.IComparer(Of Long)

        ''' <returns>
        ''' Zero if x is equal to y;
        ''' A value less than zero if x is greater than y;
        ''' A value greater than zero if x is less than y.
        ''' </returns>
        ''' <remarks></remarks>
        Public Function Compare(ByVal x As Long, ByVal y As Long) As Integer Implements System.Collections.Generic.IComparer(Of Long).Compare
            If x = y Then
                Return 0
            ElseIf x > y Then
                Return -1
            Else
                Return 1

            End If


        End Function


    End Class
Public Class CompareDBByFilename
    Implements Generic.IComparer(Of DatabaseEntry)

    Public Function Compare(x As DatabaseEntry, y As DatabaseEntry) As Integer Implements IComparer(Of DatabaseEntry).Compare
        If x.Filename = y.Filename Then
            Return 0
        ElseIf x.Filename < y.Filename Then
            Return 1
        Else
            Return -1
        End If


    End Function
End Class
Public Class CompareDBByFilesize
    Implements Generic.IComparer(Of DatabaseEntry)

    Public Function Compare(x As DatabaseEntry, y As DatabaseEntry) As Integer Implements IComparer(Of DatabaseEntry).Compare
        If x.Size = y.Size Then
            Return 0
        ElseIf x.Size < y.Size Then
            Return 1
        Else
            Return -1
        End If


    End Function
End Class
Public Class CompareByFilesize
        Implements Generic.IComparer(Of String)

        Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        Try

            Dim xf As New IO.FileInfo(x)
            Dim yf As New IO.FileInfo(y)
            If xf.Exists And yf.Exists Then
                If xf.Length = yf.Length Then
                    Return 0
                ElseIf xf.Length < yf.Length Then
                    Return 1
                Else
                    Return -1
                End If
            Else
                Return 0
            End If
        Catch ex As Exception
            Return 0
        End Try


    End Function
    End Class
    Public Class CompareByDate
        Implements Generic.IComparer(Of String)

        Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
            Dim xf As New IO.FileInfo(x)
            Dim yf As New IO.FileInfo(y)
            If xf.Exists And yf.Exists Then
                If GetDate(xf) = GetDate(yf) Then
                    Return 0
                ElseIf GetDate(xf) < GetDate(yf) Then
                    Return -1
                Else
                    Return 1
                End If


            Else
                Return 0
            End If

        End Function
    End Class
    Public Class CompareByType
        Implements Generic.IComparer(Of String)

        Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
            Dim xf As New IO.FileInfo(x)
            Dim yf As New IO.FileInfo(y)
            If xf.Exists And yf.Exists Then
                If xf.Extension = yf.Extension Then
                    Return 0
                ElseIf xf.Extension < yf.Extension Then
                    Return -1
                Else
                    Return 1
                End If


            Else
                Return 0
            End If

        End Function
    End Class

Public Class CompareByEndNumber
    Implements Generic.IComparer(Of String)

    'Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
    '    x = FilenameFromPath(x, False)
    '    y = FilenameFromPath(y, False)
    '    'Counting from end, find the trailing numerics
    '    Dim i = 0
    '    Dim xnum, ynum As String
    '    xnum = ""
    '    ynum = ""
    '    For i = 0 To x.Length - 1
    '        Dim m = x.Length - 1 - i
    '        If Instr("0123456789", x(m)) <> 0 Then
    '            xnum = x(x.Length - i - 1) & xnum
    '        Else
    '            If xnum = "" Then
    '            Else
    '                Exit For

    '            End If
    '        End If
    '    Next
    '    For i = 0 To y.Length - 1
    '        Dim m = y.Length - 1 - i
    '        If InStr("0123456789", y(m)) <> 0 Then
    '            ynum = y(y.Length - i - 1) & ynum
    '        Else
    '            If ynum = "" Then
    '            Else
    '                Exit For

    '            End If
    '        End If
    '    Next
    '    'If same Then order in normal way
    '    If ynum.Length = xnum.Length Or ynum = "" Or xnum = "" Then
    '        If y < x Then
    '            Return 1
    '        ElseIf x < y Then
    '            Return -1
    '        Else
    '            Return 0
    '        End If
    '    Else
    '        'Otherwise, order according to those numbers
    '        Dim ynumnum = Val(ynum)
    '        Dim xnumnum = Val(xnum)
    '        If ynumnum < xnumnum Then
    '            Return 1
    '        ElseIf xnumnum < ynumnum Then
    '            Return -1
    '        Else
    '            Return 0
    '        End If
    '    End If

    'End Function

    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        ' Extract filenames without paths
        Dim fileNameX As String = Path.GetFileName(x)
        Dim fileNameY As String = Path.GetFileName(y)

        ' Find trailing numeric parts
        Dim xNumericPart As String = GetTrailingNumericPart(fileNameX)
        Dim yNumericPart As String = GetTrailingNumericPart(fileNameY)

        ' If both have numeric parts, compare them numerically
        If Not String.IsNullOrEmpty(xNumericPart) AndAlso Not String.IsNullOrEmpty(yNumericPart) Then
            Dim xNum As Integer
            Dim yNum As Integer
            If Integer.TryParse(xNumericPart, xNum) AndAlso Integer.TryParse(yNumericPart, yNum) Then
                Return xNum.CompareTo(yNum)
            End If
        End If

        ' If one or both don't have numeric parts, compare them as strings
        Return String.Compare(fileNameX, fileNameY)
    End Function

    Private Function GetTrailingNumericPart(fileName As String) As String
        ' Iterate in reverse to find the trailing numeric part
        Dim numericPart As String = ""
        For i As Integer = fileName.Length - 1 To 0 Step -1
            If Char.IsDigit(fileName(i)) Then
                numericPart = fileName(i) & numericPart
            Else
                Exit For
            End If
        Next
        Return numericPart
    End Function

End Class

Public Class FolderComparer
    Implements IComparer(Of String)

    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        Dim xArray = x.Split("\")
        Dim yArray = y.Split("\")
        Dim minLength = Math.Min(xArray.Length, yArray.Length)

        For i As Integer = 0 To minLength - 1
            Dim result = String.Compare(xArray(i), yArray(i), StringComparison.OrdinalIgnoreCase)
            If result <> 0 Then
                Return result
            End If
        Next

        Return xArray.Length.CompareTo(yArray.Length)
    End Function
End Class

