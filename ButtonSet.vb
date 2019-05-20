
Public Class ButtonSet
    Public WithEvents CurrentSet As New List(Of ButtonRow)
    Public Event LetterChanged(l As Keys)
    Private alph As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

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
            Dim b = mCurrentLetter
            If b <> value Then
                Dim c = value
                mCurrentLetter = value
                mCurrentRow = CurrentSet.Find(Function(x) x.Letter = value)
                'RaiseEvent LetterChanged(value)
            Else
                '                NextRow(mCurrentLetter)
            End If

        End Set
    End Property
    ''' <summary>
    ''' Sets the current row to be the next one having given letter
    ''' Does nothing if only one such row.
    ''' </summary>
    ''' <param name="letter"></param>
    Public Function NextRow(letter As Integer) As ButtonRow
        Dim x As New List(Of ButtonRow)
        x = CurrentSet.FindAll(Function(m) m.Letter = letter)
        Dim count As Integer = x.Count
        If count = 1 Then Exit Function
        Dim nextindex As Integer
        nextindex = x.IndexOf(mCurrentRow)
        nextindex = (nextindex + 1) Mod count
        mCurrentRow = x(nextindex)
        CurrentRow = mCurrentRow
        Return mCurrentRow
    End Function
    Public Function FirstFree(letter As Integer) As MVButton
        Dim x As New List(Of ButtonRow)
        x = CurrentSet.FindAll(Function(m) m.Letter = letter)
        Dim count As Integer = x.Count
        Dim i = 0
        For Each row In x
            i = row.GetFirstFree
            If i < 8 Then
                CurrentRow = row
                Exit For
            End If
        Next
        If i = 8 Then
            InsertRow(letter)
            i = 0
        End If
        Return CurrentRow.Buttons(i)
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
        CurrentRow = CurrentSet(0)
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
        CurrentRow = mCurrentRow
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

