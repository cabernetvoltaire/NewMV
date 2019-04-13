Public Class ButtonRow
    Public Buttons(7) As MVButton
    Public Event CurrentChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Sub New()
        Dim S As String = "ABCDEFGH"
        For i = 0 To 7
            Buttons(i) = New MVButton
            Buttons(i).FaceText = "f" & Str(i + 5)
            Buttons(i).Label = S(i)
        Next

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
    Private mLetter As Keys
    Public Property Letter() As Keys
        Get
            Return mLetter
        End Get
        Set(ByVal value As Keys)

            mLetter = value
        End Set
    End Property

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




    Public Sub InitialiseButtons()
        Dim alph As String = "ABCDEFGH"
        For i As Byte = 0 To 7
            Buttons(i).FaceText = "F" & Str(i + 5)
            Buttons(i).Label = alph(i)
        Next
        End Sub
    End Class
