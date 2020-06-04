Public Class ButtonRow
    Public Buttons(7) As MVButton
    Public Event CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Sub New()
        Dim S As String = "ABCDEFGH"
        For i = 0 To 7
            Buttons(i) = New MVButton With {
                .FaceText = "f" & Str(i + 5),
                .Label = S(i),
                .Position = i
            }
        Next

        'Current = False
        'Letter = "A"
    End Sub
    Public Sub New(letter As Integer)
        Dim S As String = "ABCDEFGH"
        For i = 0 To 7
            Buttons(i) = New MVButton With {
                .FaceText = "f" & Str(i + 5),
                .Label = S(i),
                .Letter = letter,
                .Position = i
            }
        Next
        mLetter = letter
        'Current = False
        'Letter = "A"
    End Sub
    Private mCurrent As Boolean
    Public Property Current() As Boolean
        Get
            Return mCurrent
        End Get
        Set(ByVal value As Boolean)
            If mCurrent <> value Then
                mCurrent = value
                RaiseEvent CurrentChanged(Me, New EventArgs)
            End If

        End Set
    End Property
    Private mLetter As Integer
    Public Property Letter() As Integer
        Get
            Return mLetter
        End Get
        Set(ByVal value As Integer)

            mLetter = value
        End Set
    End Property
    Public Function GetFirstFree() As Byte
        For i = 0 To 7
            '  Report(Buttons(i).Path, 1)
            If Buttons(i).Empty Then
                Return i
                Exit Function
            End If
        Next
        Return 8
    End Function
    Private Function GetKey(value As Char) As Integer
        Dim alpha As String = "ABCDEFGHIJKLMNOPWRSTUVWXYZ"
        Dim i As Byte = InStr(alpha, value)
        If i = 0 Then
            i = InStr("0123456789", value)
            Return i + Keys.D0
        Else
            Return i + Keys.A
        End If
        Throw New NotImplementedException()
    End Function
    Private mLastChanged As MVButton
    Public Property LastChanged() As MVButton
        Get
            Return mLastChanged
        End Get
        Set(ByVal value As MVButton)
            mLastChanged = value
        End Set
    End Property

    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            Buttons(i).FaceText = "f" & Str(i + 5)
            Buttons(i).Label = alph(i)
        Next
        End Sub
    End Class
