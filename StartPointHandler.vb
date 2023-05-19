Public Class StartPointHandler
    Public Enum StartTypes As Byte
        NearBeginning
        FirstMarker
        NearEnd
        ParticularAbsolute
        ParticularPercentage
        Random
        Beginning

    End Enum
    Public Timer As New Timer
    Public WMP As AxWindowsMediaPlayer
    Public Event StateChanged(sender As Object, e As EventArgs)
    Public Event StartPointChanged(sender As Object, e As EventArgs)
    Public Event JumpKey()
    Private ReadOnly mOrder = {
        "Near Beginning",
        "First Marker",
        "Near End",
        "Particular(s)",
        "Particular(%)",
        "Random",
        "Beginning"
    }
    Private mDescList As New List(Of String)
    Public Sub New(Optional ByVal StartPercentage As Byte = 50, Optional ByVal StartAbsolute As Byte = 65)
        mState = StartTypes.Beginning
        mPercentage = 40
        mDuration = 100
        mAbsolute = 120
        mDistance = 65

    End Sub
#Region "Properties"

    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To 6
                mDescList.Add(mOrder(i))
            Next
            Descriptions = mDescList
        End Get
    End Property
    Public ReadOnly Property Description() As String
        Get
            Return mOrder(mState)
        End Get
    End Property
    Private mDuration As Long
    Public Property Duration() As Long
        Get
            Return mDuration

        End Get
        Set(ByVal value As Long)
            mDuration = value
            SetStartPoint()
            ' ReportTime("Duration:" & mDuration)
        End Set
    End Property
    Private mDistance As Long
    Public Property Distance() As Long
        Get
            Return mDistance
        End Get
        Set(ByVal value As Long)
            mDistance = value
        End Set
    End Property
    Private mMarkers As New List(Of Double)
    Public Property Markers() As List(Of Double)
        Get
            Return mMarkers
        End Get
        Set(ByVal value As List(Of Double))

            mMarkers = value
        End Set
    End Property
    Private mCurrentMarker As Long
    Public ReadOnly Property CurrentMarker() As Long
        Get
            mCurrentMarker = mMarkers(mMarkCounter)
            Return mCurrentMarker
        End Get
    End Property
    Private mAbsolute As Long
    Public Property Absolute() As Long
        Get
            Return mAbsolute
        End Get
        Set(ByVal value As Long)


            If value <= Duration Then
                mAbsolute = value
                mPercentage = mAbsolute / mDuration * 100
                mStartPoint = mAbsolute
                RaiseEvent StartPointChanged(Me, Nothing)
            End If
        End Set
    End Property
    Private mPercentage As Byte
    Public Property Percentage() As Byte
        Get
            Return mPercentage
        End Get
        Set(ByVal value As Byte)
            Dim b As Byte = mPercentage
            mPercentage = value
            mAbsolute = mPercentage / 100 * mDuration
            If b <> mPercentage Then RaiseEvent StartPointChanged(Me, Nothing)
        End Set
    End Property
    Private mStartPoint As Long
    Public ReadOnly Property StartPoint() As Long
        Get
            ' SetStartPoint()
            Return mStartPoint
        End Get

    End Property
    Private mState As Byte
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            Dim b As Byte = mState
            mState = value
            If b <> mState Then
                SetStartPoint()
                RaiseEvent StateChanged(Me, New EventArgs)
            End If
        End Set
    End Property
    Private mSavedState As Byte
    Public Property SavedState() As Byte
        Get

            Return mSavedState
            mSavedState = 0
        End Get
        Set(ByVal value As Byte)

            mSavedState = value

        End Set
    End Property

    Private mMarker As Long
    Public Property Marker() As Long
        Get
            Return mMarker
        End Get
        Set(ByVal value As Long)
            mMarker = value
            mStartPoint = mMarker
        End Set
    End Property
#End Region
#Region "Methods"
    Public Sub IncrementState(max As Byte)
        State = (State + 1) Mod max
    End Sub
    Private Property mMarkCounter = 0

    Public Sub Reset()
        mMarkers.Clear()
        SetStartPoint()
    End Sub
    Private Function SetStartPoint() As Long
        Select Case mState
            Case StartTypes.FirstMarker
                mStartPoint = If(mMarkers.Any(), mMarkers(0), GetAdjustedStartPoint(mDistance))
            Case StartTypes.Beginning
                mStartPoint = 0
            Case StartTypes.NearBeginning, StartTypes.NearEnd
                mStartPoint = GetAdjustedStartPoint(If(mState = StartTypes.NearBeginning, mDuration * 0.05, mDuration * 0.95))
            Case StartTypes.ParticularAbsolute
                mStartPoint = mAbsolute
            Case StartTypes.ParticularPercentage
                mStartPoint = mPercentage / 100 * mDuration
            Case StartTypes.Random
                mStartPoint = Rnd() * mDuration
        End Select

        ' Apply limits
        mStartPoint = ApplyLimitsToStartPoint(mStartPoint)
        If mStartPoint > mDuration Then MsgBox("Too far")

        Return mStartPoint
        RaiseEvent StartPointChanged(Me, Nothing)
    End Function

    Private Function GetAdjustedStartPoint(value As Long) As Long
        If value > mDuration / 2 Then
            Return mDuration * If(mState = StartTypes.NearBeginning, 0.1, 0.9)
        End If
        Return value
    End Function

    Private Function ApplyLimitsToStartPoint(value As Long) As Long
        If mDuration - value < 5 Then
            value = mDuration - 5
        End If
        If value < 0 Then
            value = 0
        End If
        If value > mDuration Then
            value = mDuration / 2
        End If
        Return value
    End Function

#End Region
End Class
