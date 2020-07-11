
Public Class ButtonSet
    Public WithEvents CurrentSet As New List(Of ButtonRow)
    Public Event LetterChanged(sender As Object, e As EventArgs)
    Public Event NewSetThisLetter(index As Integer, total As Integer)

    Private alph As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Private mRowIndex As Integer
    Public Property RowIndex() As Integer
        Get
            Return mRowIndex
        End Get
        Set(ByVal value As Integer)
            mRowIndex = value
        End Set
    End Property

    Private mRowIndexCount As Integer
    Public Property RowIndexCount() As Integer
        Get
            Return mRowIndexCount
        End Get
        Set(ByVal value As Integer)
            mRowIndexCount = value
        End Set
    End Property
    Private mCurrentRow As New ButtonRow
    ''' <summary>
    ''' The row which is current and displayed, of all the rows of the button set. 
    ''' </summary>
    ''' <returns></returns>
    Public Property CurrentRow() As ButtonRow
        Get
            Return mCurrentRow
        End Get
        Set(ByVal value As ButtonRow)
            If Not mCurrentRow.Equals(value) Then
                mCurrentRow = value
            End If
        End Set
    End Property
    Private mCurrentLetter As Integer

    ''' <summary>
    ''' The current letter, as an integer
    ''' </summary>
    ''' <returns></returns>

    Public Property CurrentLetter() As Integer
        Get
            Return mCurrentLetter
        End Get
        Set(ByVal value As Integer)

            mCurrentLetter = value
            mCurrentRow = CurrentSet.Find(Function(x) x.Letter = value)
                Dim tl As List(Of ButtonRow) = Nothing
                mRowIndex = 0
                CountRowIndices(value, tl, mRowIndexCount)
            RaiseEvent LetterChanged(Me, Nothing)

        End Set
    End Property
    ''' <summary>
    ''' Sets the current row to be the next one having given letter
    ''' Does nothing if only one such row.
    ''' </summary>
    ''' <param name="letter"></param>
    Public Function NextRow(letter As Integer, Optional Backward As Boolean = False) As ButtonRow
        Dim x As List(Of ButtonRow) = Nothing
        Dim count As Integer = Nothing
        CountRowIndices(letter, x, count)

        If count = 1 Then
            Return mCurrentRow
            Exit Function
        End If
        mRowIndex = x.IndexOf(mCurrentRow)
        If Backward Then
            If mRowIndex = 0 Then
                mRowIndex = count - 1
            Else
                mRowIndex = (mRowIndex - 1) Mod count
            End If
        Else
            mRowIndex = (mRowIndex + 1) Mod count
        End If
        mCurrentRow = x(mRowIndex)
        RaiseEvent LetterChanged(Me, Nothing)

        Return mCurrentRow
    End Function

    Private Sub CountRowIndices(letter As Integer, ByRef x As List(Of ButtonRow), ByRef count As Integer)
        x = CurrentSet.FindAll(Function(m) m.Letter = letter)
        count = x.Count
    End Sub
    ''' <summary>
    ''' Returns first free button in a whole buttonset
    ''' </summary>
    ''' <param name="letter"></param>
    ''' <returns></returns>
    Public Function FirstFree(letter As Integer) As MVButton
        Dim x As New List(Of ButtonRow)
        x = CurrentSet.FindAll(Function(m) m.Letter = letter)
        Dim count As Integer = x.Count
        Dim i = 0
        For Each row In x
            i = row.GetFirstFree
            If i < 8 Then
                mCurrentRow = row
                Exit For
            End If
        Next
        If i = 8 Then
            InsertRow(letter)
            i = 0
        End If
        mCurrentRow.Buttons(i).Position = i
        mCurrentRow.Buttons(i).Letter = letter

        Return mCurrentRow.Buttons(i)
    End Function
    Public Sub New()
        Initialise()
    End Sub
    Public Sub Initialise()
        Dim Rows(Len(alph)) As ButtonRow
        For i = 0 To Len(alph) - 1
            Rows(i) = New ButtonRow
            Rows(i).Letter = LetterNumberFromAscii(Asc(alph(i)))
            CurrentSet.Add(Rows(i))
        Next
        mCurrentRow = CurrentSet(0)
    End Sub

    Public Sub Clear()
        CurrentSet.Clear()
        Initialise()
    End Sub
    ''' <summary>
    ''' Inserts a row at the end of all the alphs
    ''' Makes it the current row.
    ''' </summary>
    ''' <param name="alph"></param>
    Public Sub InsertRow(alph As Integer)
        Dim index As Integer
        index = CurrentSet.FindIndex(Function(x) x.Letter = alph + 1)
        mCurrentRow = New ButtonRow(alph)
        If index = -1 Then
            CurrentSet.Add(mCurrentRow)
        Else
            CurrentSet.Insert(index, mCurrentRow)
        End If

        mRowIndexCount = mRowIndexCount + 1
    End Sub
    Private Function LetterFromNumber(Num As Integer) As Char
        Dim ch As Char = alph(Num)
        Return ch
    End Function

    Private Function NumberFromLetter(ch As Char) As Integer
        Dim num As Integer
        num = InStr(alph, ch)
        Return num
    End Function
End Class

