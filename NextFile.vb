Public Class NextFile
    Private mListCount As Integer
    Private mListbox As New ListBox
    Public Event RandomChanged(sender As Object, e As EventArgs)
    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
            mListCount = mListbox.Items.Count
        End Set
    End Property

    Public Property CurrentItem As String
    Public Property Forwards As Boolean = True
    Private mCurrentIndex As Integer
    Public Property CurrentIndex() As Integer
        Get
            Return mCurrentIndex
        End Get
        Set(ByVal value As Integer)
            mCurrentIndex = value
            If value >= 0 And value < mListCount Then
                CurrentItem = mListbox.Items(mCurrentIndex)
            End If

        End Set
    End Property
    Private Property mNextItem As String
    Public ReadOnly Property RandomItem As String
        Get
            If mListCount > 0 Then

                RandomItem = mListbox.Items(Int(Rnd() * (mListCount)))
            Else
                RandomItem = ""
            End If

        End Get
    End Property



    Public Property NextItem As String
        Get
            'If Randomised Then

            'Else
            If mListCount > 1 Then
                If Forwards Then
                    mNextItem = mListbox.Items((mCurrentIndex + 1) Mod (mListCount))
                Else
                    If mCurrentIndex = 0 Then
                        mCurrentIndex = mListCount
                    Else
                        mNextItem = mListbox.Items((mCurrentIndex - 1) Mod (mListCount))
                    End If
                End If
            Else
                mNextItem = CurrentItem

            End If
            'End If
            Return mNextItem
        End Get
        Set(value As String)
            mNextItem = value
        End Set
    End Property

    Private Property mPreviousItem As String
    Public Property PreviousItem As String
        Get
            Try

                'If Randomised Then
                '    mPreviousItem = Listbox.Items(Int(Rnd() * (mListCount - 1)))
                'Else
                If mListCount > 1 Then
                        If Not Forwards Then
                            mPreviousItem = Listbox.Items((mCurrentIndex + 1) Mod (mListCount))
                        Else
                            If mCurrentIndex = 0 Then
                                mCurrentIndex = mListCount
                            End If
                            mPreviousItem = Listbox.Items((mCurrentIndex - 1) Mod (mListCount))
                        End If
                    Else
                        mPreviousItem = Listbox.Items(0)

                    End If
                ' End If
                Return mPreviousItem
            Catch ex As Exception
                MsgBox("Exception in nextfile - previous item")
                Return 0
            End Try
        End Get
        Set(value As String)

        End Set
    End Property

    Private Property mRandomised As Boolean
    Public Property Randomised As Boolean
        Set(value As Boolean)
            mRandomised = value
            'If mRandomised Then
            '    mNextItem = NextItem
            '    RaiseEvent RandomChanged(Me, Nothing)
            'End If
        End Set
        Get
            Return mRandomised
        End Get
    End Property


End Class
