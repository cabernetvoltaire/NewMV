Imports System.IO

Public Class Database
    Public Entries As New List(Of DatabaseEntry)
    Public Name As String

    Private Property mItemCount As Long = 0
    Public ReadOnly Property ItemCount As Long
        Get
            Return mItemCount
        End Get
    End Property


    Public Sub AddEntry(entry As DatabaseEntry)
        Dim f As New FileInfo(entry.FullPath)
        If f.Exists Then
            Entries.Add(entry)
            mItemCount += 1
        End If
    End Sub

    Private mSorted As Boolean = False
    Public Sub Sort(Optional comparer As IComparer(Of DatabaseEntry) = Nothing)
        Entries.Sort(comparer)
        mSorted = True
    End Sub
    Public Function FindEntry(filename As String) As DatabaseEntry
        If Not mSorted Then Sort(New CompareDBByFilename)

        Dim n As New List(Of DatabaseEntry)
        n = Entries.FindAll(Function(x) x.Filename.StartsWith(filename))
        Select Case n.Count
            Case 0
                n = Entries.FindAll(Function(x) x.Filename.Contains(Path.GetFileNameWithoutExtension(filename)))
                If n.Count = 0 Then
                    Return Nothing
                End If
            Case Else
                Return n(0)
                'Decide which one to return
        End Select

    End Function
    Public Function FindPartEntry(Search As String) As List(Of DatabaseEntry)
        If Not mSorted Then Sort()

        Dim n As New List(Of DatabaseEntry)
        n = Entries.FindAll(Function(x) x.FullPath.Contains(Search))
        Return n
    End Function


End Class

Public Class DatabaseEntry
    Implements IComparable
    Public Rand As Long = RandomInstance.Next(0, Integer.MaxValue)
    Private Property mFullPath As String
    Public Property Filename As String
    Public ReadOnly Property FullPath As String
        Get
            Return mFullPath
        End Get
    End Property

    Private Property mPath As String
    Public Property Path As String
        Set(value As String)
            mPath = value
            mFullPath = String.Concat(mPath, Filename)
        End Set
        Get
            Return mPath
        End Get
    End Property
    Public Size As Long
    Public Dt As DateTime

    Public Sub New()

    End Sub

    Public Property Mark As Boolean = False

    Public Function CompareTo(obj As Object) As Integer Implements IComparable.CompareTo
        If obj.size = Size Then
            Return 0
        ElseIf obj.size < Size Then
            Return 1
        Else
            Return -1


        End If
    End Function

End Class