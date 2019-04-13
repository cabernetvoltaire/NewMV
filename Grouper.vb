
Public Class Grouper
    Private mPath As String
    Private mList As List(Of String)
    Public Sub New(Path As String)
        mPath = Path
        mFilesPerList = 2
        InputList = IO.Directory.GetFiles(mPath).ToList
    End Sub
    Public Property InputList() As List(Of String)
        Get
            Return mList
        End Get
        Set(ByVal value As List(Of String))

            mList = value
            mList = SorttheListbyLength(mList)
        End Set
    End Property

    Private mFilesPerList As Integer
    Public Property FilesPerList() As Integer
        Get
            Return mFilesPerList
        End Get
        Set(ByVal value As Integer)
            mFilesPerList = value
        End Set
    End Property

    Private mSublists As List(Of List(Of String))
    Public ReadOnly Property Sublists() As List(Of List(Of String))
        Get
            Return GroupByName()
        End Get
    End Property
    Private Function SorttheListbyLength(lst As List(Of String)) As List(Of String)
        Dim out As New List(Of String)
        Dim m As New SortedDictionary(Of Integer, String)
        For Each r In lst
            Dim k As Integer = 10000 * Len(r) + Int(Rnd() * 10000)
            While m.ContainsKey(k)
                k = 10000 * Len(r) + Int(Rnd() * 10000)

            End While
            m.Add(k, r)
        Next
        m.Reverse
        For Each r In m
            out.Insert(0, r.Value)
        Next
        Return out
    End Function

    Private Function GroupByName() As List(Of List(Of String))
        Dim mSublists As New List(Of List(Of String))
        Dim n = Len(mList(0)) 'Get the length of this string.
        While n > 1
            Dim i As Int16 = 0
            While i <= mList.Count - 1  'Work through the input list
                Dim s As String = mList(i)
                Dim trylist As List(Of String)
                Dim c As String = Left(s, n)
                trylist = mList.FindAll(Function(x) x.Contains(c)) 'trylist contains all names containing c

                If trylist.Count >= mFilesPerList And trylist.Count < mList.Count Then
                    mSublists.Add(trylist)
                    Debug.Print("")

                    For Each j In trylist
                        Debug.Print("Adding {0} to sublist", j)
                        mList.Remove(j)
                    Next
                    'trylist.Clear()
                Else
                End If
                i += 1
            End While
            n = n - 1
        End While
        Return mSublists
    End Function

End Class
