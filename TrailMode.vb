Public Class TrailMode
    Private mActive As Boolean


    ''' <summary>
    ''' Maximum number of half seconds at each speed (in addition to default half second)
    ''' </summary>
    ''' <returns></returns>
    Public Property RandomTimes As Byte() = {4, 3, 5, 2}

    ''' <summary>
    ''' Proportions of each speed, out of 3, cumulative. Last is normal speed
    ''' </summary>
    ''' <returns></returns>
    Public Property SelectionWeights As Single() = {0.2, 0.4, 0.75}
    ''' <summary>
    ''' Framerates for each of the 3 different slow speeds
    ''' </summary>
    ''' <returns></returns>
    Public Property Speeds As Integer() = {5, 10, 15}
    ''' <summary>
    ''' Probability that file will be advanced, as a denominator of a fraction. (eg 15 means probability of 1/15) 
    ''' </summary>
    ''' <returns></returns>
    Public Property AdvanceChance As Integer = 15
    ''' <summary>
    ''' Returns an integer given a single between 0 and 1, according to SelectionWeights
    ''' </summary>
    ''' <param name="n"></param>
    ''' <returns></returns>
    Public Function ChosenSpeed(n As Single) As Byte
        If n < SelectionWeights(0) Then
            Return 0
        ElseIf n < SelectionWeights(1) Then
            Return 1
        ElseIf n < SelectionWeights(2) Then
            Return 2
        Else
            Return 3
        End If
    End Function

    Public Sub EqualiseSpeeds(framerate As Integer, duration As Byte)
        SelectionWeights = {1, 1, 1}
        Speeds = {framerate, framerate, framerate}
        RandomTimes = {duration, duration, duration, duration}
    End Sub

End Class
