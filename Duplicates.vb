Imports System.IO
Class Duplicates
    Private Enum KeepPref As Byte
        Deepest
        Longest
        Shortest
        Shallowest
    End Enum
    Private mInputList As List(Of String)
    Public Property InputList() As List(Of String)
        Get
            Return mInputList
        End Get
        Set(ByVal value As List(Of String))
            mInputList = value
            mInputList = SetPlayOrder(SortHandler.Order.Size, value)
        End Set
    End Property

    Private mDuplicates As List(Of List(Of String))
    Public Property Duplicates() As List(Of List(Of String))
        Get
            mDuplicates = ExtractDups(mInputList)
            Return mDuplicates
        End Get
        Set(ByVal value As List(Of List(Of String)))
            mDuplicates = value
        End Set
    End Property



    ''' <summary>
    ''' Takes size-ordered List1 and produces List of String lists, each containing duplicates of the first indexed. 
    ''' </summary>
    Private Function ExtractDups(List1 As List(Of String)) As List(Of List(Of String))
        Dim mCurrentRow As New List(Of String)
        Dim mDuplicateArray As New List(Of List(Of String))
        Dim lastlength As Long = 0
        Dim lastinfo As FileInfo = Nothing
        Dim finfo3 As FileInfo = Nothing
        For Each file In List1
            'if file is same length as previous, add to current row
            'Otherwise add current row to array, scrap current row
            finfo3 = New FileInfo(file)
            Dim currlength As Long = finfo3.Length
            If currlength <> lastlength Then 'Start new row
                If mCurrentRow.Count > 1 Then 'Only add row if it has more than one member
                    Dim l As New List(Of String) '= Nothing
                    For Each m In mCurrentRow
                        l.Add(m)
                    Next
                    mDuplicateArray.Add(l)
                    Debug.Print(mDuplicateArray(0).Item(0))
                End If
                mCurrentRow.Clear()
                mCurrentRow.Add(finfo3.FullName)

            Else
                mCurrentRow.Add(finfo3.FullName)
            End If
            lastinfo = finfo3
            lastlength = lastinfo.Length
        Next
        If mCurrentRow.Count > 1 Then
            Dim l As New List(Of String) '= Nothing
            For Each m In mCurrentRow
                l.Add(m)
            Next
            mDuplicateArray.Add(l)

        End If
        Return mDuplicateArray

    End Function

    Private Sub SortList(l As List(Of String))
        Throw New NotImplementedException()
    End Sub
End Class
