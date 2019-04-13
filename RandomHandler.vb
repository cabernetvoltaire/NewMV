Public Class RandomHandler
    Public Event RandomChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Sub New()
        mOnDirChange = True
        mNextSelect = False
        mStartPoint = False
        mAll = False
    End Sub
    Private mOnDirChange As Boolean
    Public Property OnDirChange() As Boolean
        Get
            Return mOnDirChange
        End Get
        Set(ByVal value As Boolean)
            mOnDirChange = value
            RaiseEvent RandomChanged(Me, New EventArgs)
        End Set
    End Property
    Private mNextSelect As Boolean
    Public Property NextSelect() As Boolean
        Get
            Return mNextSelect
        End Get
        Set(ByVal value As Boolean)
            mNextSelect = value
            RaiseEvent RandomChanged(Me, New EventArgs)

        End Set
    End Property
    Private mStartPoint As Boolean
    Public Property StartPointFlag() As Boolean
        Get
            Return mStartPoint
        End Get
        Set(ByVal value As Boolean)
            mStartPoint = value
            RaiseEvent RandomChanged(Me, New EventArgs)

        End Set
    End Property
    Private Property mAll As Boolean
    Public Property All() As Boolean
        Get
            Return mAll
        End Get
        Set(ByVal value As Boolean)
            OnDirChange = value
            NextSelect = value
            StartPointFlag = value
            mAll = value
            RaiseEvent RandomChanged(Me, New EventArgs)

        End Set
    End Property
End Class
