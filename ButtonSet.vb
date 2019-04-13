
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

                Dim c = InStr(alph, ButtfromAsc(value))
                If c = 0 Then
                Else
                    mCurrentLetter = value
                    mCurrentRow = CurrentSet(c - 1)
                    RaiseEvent LetterChanged(value)
                End If
            Else
            End If

        End Set
    End Property

    Public Sub New()
        Dim Rows(Len(alph)) As ButtonRow
        For i = 0 To Len(alph) - 1
            Rows(i) = New ButtonRow
            Rows(i).Letter = CType(Asc(alph(i)), Keys)
            CurrentSet.Add(Rows(i))
        Next
        CurrentRow = CurrentSet(0)
    End Sub
    Private Function LetterFromNumber(Num As Integer) As Char
        Dim ch As Char = alph(Num)
        'If Num < Asc("A") Then
        '    Num = Num -
        'End If
        Return ch
    End Function

    Private Function NumberFromLetter(ch As Char) As Integer
        Dim num As Integer
        num = InStr(alph, ch)
        'If num < 26 Then
        '    num = num + Asc("A")
        'Else
        '    num = num + Asc("0")
        'End If
        Return num
    End Function
End Class

