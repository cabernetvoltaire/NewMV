Imports MasaSam.Forms.Sample

Public Class UserPrefs
    Private mState As New StateHandler
    Public Property State() As StateHandler
        Get
            Return mState
        End Get
        Set(ByVal value As StateHandler)
            mState = value
        End Set
    End Property
    Private mSort As New SortHandler
    Public Property Sort() As SortHandler
        Get
            Return mSort
        End Get
        Set(ByVal value As SortHandler)
            mSort = value
        End Set
    End Property

    Private mFilter As New FilterHandler
    Public Property Filter() As FilterHandler
        Get
            Return mFilter
        End Get
        Set(ByVal value As FilterHandler)
            mFilter = value
        End Set
    End Property
    Private mStart As New StartPointHandler

    Public Sub New()
        Start = mStart
        Filter = mFilter
        Sort = mSort
        State = mState
        Media = mMedia
    End Sub
    Private mMedia As New MediaHandler("mMedia")
    Public Property Media() As MediaHandler
        Get
            Return mMedia
        End Get
        Set(ByVal value As MediaHandler)
            mMedia = value
        End Set
    End Property
    Public Property Start() As StartPointHandler
        Get
            Return mStart
        End Get
        Set(ByVal value As StartPointHandler)
            mStart = value
        End Set
    End Property


End Class
