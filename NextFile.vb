Public Class NextFile
    Private mListCount As Integer
    ' Private mListbox As New ListBox
    Public Event RandomChanged(sender As Object, e As EventArgs)
    'Public Property Listbox() As ListBox
    '    Get
    '        Return mListbox
    '    End Get
    '    Set(ByVal value As ListBox)
    '        mListbox = value
    '        List = ListfromSelectedInListbox(mListbox)
    '        mListCount = mListbox.Items.Count
    '    End Set
    'End Property
    Public Property List As List(Of String)
        Get
            Return _list
        End Get
        Set(ByVal value As List(Of String))
            _list = value
            mListCount = If(value IsNot Nothing, value.Count, 0)

        End Set
    End Property
    Private _list As List(Of String)

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

                CurrentItem = List(mCurrentIndex)

            End If

        End Set
    End Property
    Private Property mNextItem As String
    Public ReadOnly Property RandomItem As String
        Get
            If mListCount > 0 Then

                RandomItem = List(Int(Rnd() * (mListCount)))
            Else
                RandomItem = ""
            End If

        End Get
    End Property



    Public ReadOnly Property NextItem As String
        Get
            If mListCount > 1 Then
                If Forwards Then
                    Return List((mCurrentIndex + 1) Mod mListCount)
                Else
                    ' Correct calculation for going backwards
                    Dim prevIndex As Integer = (mCurrentIndex - 1 + mListCount) Mod mListCount
                    Return List(prevIndex)
                End If
            Else
                Return CurrentItem
            End If
        End Get
    End Property

    Public ReadOnly Property PreviousItem As String
        Get
            If mListCount > 1 Then
                If Not Forwards Then
                    Return List((mCurrentIndex + 1) Mod mListCount)
                Else
                    ' Correct calculation for going backwards
                    Dim prevIndex As Integer = (mCurrentIndex - 1 + mListCount) Mod mListCount
                    Return List(prevIndex)
                End If
            Else
                Return CurrentItem
            End If
        End Get
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
