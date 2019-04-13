Public Class AutoTrailer
    Private probs(3) As Decimal
    Private mDuration As Byte
    Private mProb As Decimal
    Public Property RandomTimes As Byte() = {5, 5, 10, 5}
    Public Property SelectionWeights As Decimal() = {0.25, 0.5, 0.75, 1}
    ''' <summary>
    ''' Framerates for each of the 3 different slow speeds
    ''' </summary>
    ''' <returns></returns>
    Public Property Framerates As Byte() = {5, 12, 20}
    Public Property AdvanceChance As Integer = 15

    Public ReadOnly Property Duration() As Byte
        Get
            Dim x As Decimal = Rnd()
            mDuration = IndexFromNumber(x)
            Return mDuration
        End Get
    End Property
    Private mSpeedIndex As Byte
    Public ReadOnly Property SpeedIndex() As Byte
        Get
            Dim x As Decimal = Rnd()
            mSpeedIndex = IndexFromNumber(x)
            Return mSpeedIndex
        End Get
    End Property
    Public Sub New()
        SetProbs()
    End Sub
    Private Sub SetProbs()
        probs = SelectionWeights

        'For i = 0 To 3
        '    probs(i) = CumBin(3, mProb, i)
        'Next
    End Sub

    Private Function IndexFromNumber(x As Decimal) As Byte
        Dim ind As Byte
        If x < probs(0) Then
            ind = 0
        ElseIf x < probs(1) Then
            ind = 1
        ElseIf x < probs(2) Then
            ind = 2
        Else
            ind = 3
        End If
        Return ind
    End Function
    Public WriteOnly Property Probability() As Decimal
        Set(ByVal value As Decimal)
            mProb = value
            SetProbs()

        End Set
    End Property
    Private Function CumBin(n As Integer, p As Decimal, x As Integer)
        Dim prob As Decimal
        For i = 0 To x
            prob = prob + Comb(n, i) * p ^ i * (1 - p) ^ (n - i)
        Next
        Return prob
    End Function
    Private Function Comb(n As Integer, r As Integer) As Long
        Return Factorial(n) / (Factorial(r) * Factorial(n - r))
    End Function
    Private Function Factorial(n As Long) As Long
        Dim number As Long = 1
        For i = n To 1 Step -1
            number = number * i
        Next

        Return number
    End Function


End Class
