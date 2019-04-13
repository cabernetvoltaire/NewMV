Public Class SortHandler
    Public Enum Order As Byte
        DateTime
        Random
        Name
        Size
        Original
        PathName
        Type
    End Enum

    Public Event StateChanged(sender As Object, e As EventArgs)
    Private mOrder = {"DateTime", "Random", "Name", "Size", "Original", "PathName", "Type"}
    Private mColour As Color() = {Color.IndianRed, Color.OrangeRed, Color.LightYellow, Color.LightGreen, Color.LightBlue, Color.LightCoral, Color.PaleVioletRed}

    Private mDescList As New List(Of String)
    Private mReverseOrder As Boolean

    Private mColor As Color
    Public Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
        Set(ByVal value As Color)
            mColor = value

        End Set
    End Property

    Public Property ReverseOrder() As Boolean
        Get
            Return mReverseOrder
        End Get
        Set(ByVal value As Boolean)
            Dim b = mReverseOrder
            mReverseOrder = value
            If b <> mReverseOrder Then RaiseEvent StateChanged(Me, Nothing)
        End Set
    End Property
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
    Private mState As Byte

    Public Sub New()
        Me.State = Order.Original
    End Sub

    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            Dim b As Byte = mState
            mState = value
            If b <> value Then RaiseEvent StateChanged(Me, New EventArgs)
        End Set
    End Property
    Public Sub IncrementState()
        State = (State + 1) Mod (Order.Type + 1)

    End Sub

    Public Sub Toggle()
        If Me.State = Order.Original Then
            Me.State = Order.Random
        Else
            Me.State = Order.Original
        End If
    End Sub
End Class
