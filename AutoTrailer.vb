Public Class AutoTrailer
    Private ReadOnly rand As New Random()
    ' Change RandomTimes to Single to support fractional seconds.
    Public Property RandomTimes As Single()
    Public Property SelectionWeights As Decimal() = {0.1D, 0.3D, 0.6D, 1D}
    Public Property Framerates As Integer() = {5, 12, 20}
    Public Property AdvanceChance As Integer = 8

    ' Counter for tracking conditions.
    Public Property Counter As Int16 = 0

    ' Simplified way to get a weighted random index for speed.
    Public ReadOnly Property SpeedIndex As Byte
        Get
            Return GetWeightedIndex(SelectionWeights)
        End Get
    End Property

    ' Adjusted Duration property to return a Single, allowing for fractional seconds.
    Public ReadOnly Property Duration As Single
        Get
            Return RandomTimes(rand.Next(RandomTimes.Length))
        End Get
    End Property

    ' Utility function to get a weighted random index.
    Private Function GetWeightedIndex(weights As Decimal()) As Byte
        Dim value As Decimal = rand.NextDouble()
        Dim cumulative As Decimal = 0
        For i As Integer = 0 To weights.Length - 1
            cumulative += weights(i)
            If value < cumulative Then
                Return i
            End If
        Next
        Return weights.Length - 1 ' Fallback to the last index
    End Function
End Class
Public Class AutoTrailerOld

    Private WithEvents timer As New Timer
    Private probs(3) As Decimal
    Private mDuration As Byte
    Private mProb As Decimal
    Public Property Counter As Int16
    ''' <summary>
    ''' Seeds of the base durations of different selections (varied by randomiser)
    ''' </summary>
    ''' <returns></returns>
    Public Property RandomTimes As Byte()
    ''' <summary>
    ''' Cumulative probabilities of selecting each of the 4 speeds
    ''' </summary>
    ''' <returns></returns>
    Public Property SelectionWeights As Decimal() = {0.1, 0.3, 0.6, 1}
    ''' <summary>
    ''' Framerates for each of the 3 different slow speeds
    ''' </summary>
    ''' <returns></returns>
    Public Property Framerates As Integer() = {5, 12, 20}
    ''' <summary>
    ''' Reciprocal of probability of changing
    ''' </summary>
    ''' <returns></returns>
    Public Property AdvanceChance As Integer = 8

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
            prob = prob + Bin(n, p, i)
        Next
        Return prob
    End Function
    Private Function Bin(n As Integer, p As Decimal, x As Integer)
        Dim prob As Decimal
        For i = 0 To x
            prob = Comb(n, i) * p ^ i * (1 - p) ^ (n - i)
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
