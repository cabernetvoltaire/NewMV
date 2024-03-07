Public Class AutoTrailer
    Private ReadOnly rand As New Random()
    ' Change RandomTimes to Single to support fractional seconds.
    Public Property RandomTimes As Single() = {0.7R, 1.2R, 2.0R, 2.5R}
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
