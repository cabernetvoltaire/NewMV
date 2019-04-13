Public Class PlayListHandler
    Public Property FH As FilterHandler
    Public Property RH As RandomHandler
    Private mLB As ListBox
    Public Property LB() As ListBox
        Get
            Return mLB
        End Get
        Set(ByVal value As ListBox)
            mLB = value
            mCurrentItem = mLB.SelectedItem
            mIndex = mLB.SelectedIndex
        End Set
    End Property
    Private Property mIndex As Integer
    Private mPlaylist As List(Of String)
    Public Property Direction As Integer = 1
    Public Property PlayList() As List(Of String)
        Get
            Return mPlaylist
        End Get
        Set(ByVal value As List(Of String))
            mPlaylist = value
        End Set
    End Property
    Private mNextItem As String
    Public Property NextItem() As String
        Get
            If RH.NextSelect Then

            End If
            Return mNextItem
        End Get
        Set(ByVal value As String)
            mNextItem = value
        End Set
    End Property
    Private mCurrentItem As String
    Public Property CurrentItem() As String
        Get
            Return mCurrentItem
        End Get
        Set(ByVal value As String)
            mCurrentItem = value

        End Set
    End Property

    Private Sub GetNextFile()
        If RH.NextSelect Then
            mNextItem = mPlaylist.Item(Int(Rnd() * mPlaylist.Count) - 1)
        Else
            mNextItem = mPlaylist.Item((mIndex + Direction) Mod (mPlaylist.Count - 1))
        End If
    End Sub

End Class
