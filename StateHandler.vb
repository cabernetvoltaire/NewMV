Public Class StateHandler
    Public Enum StateOptions As Byte
        Navigate
        Move
        MoveLeavingLink
        Copy
        CopyLink
        ExchangeLink
        MoveOriginal

    End Enum
    Public Event StateChanged(sender As Object, e As EventArgs)
    Public mColour() As Color = ({Color.Aqua, Color.Orange, Color.LightPink, Color.LightGreen, Color.Tomato, Color.BurlyWood, Color.BlueViolet})

    Private mDescription = {"Navigate Mode", "Move Mode", "Move (Leave link)", "Copy", "Copy Link", "Exchange Link", "Move original"}
    Private mInstructions = {"Function keys navigate to folder." & vbCrLf & "[SHIFT] + Fn moves file. " & vbCrLf & "[CTRL] + [SHIFT] +Fn moves folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys move file to folder." & vbCrLf & "[SHIFT] + Fn moves current folder to folder. " & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys moves file, leaving link." & vbCrLf & "[SHIFT] + Fn does same for folder. " & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys copies file." & vbCrLf & "[SHIFT] + Fn does same for folder" & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys creates a link in the destination." & vbCrLf & "[SHIFT] + Fn does same for folder" & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys exchanges link with file." & vbCrLf & "[SHIFT] + Fn does same for folder" & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key",
        "Function keys moves original file of a link." & vbCrLf & "[SHIFT] + Fn does nothing" & vbCrLf & "[CTRL] + [SHIFT] +Fn navigates to folder" & vbCrLf & "[ALT]+[CTRL]+[SHIFT] + Fn redefines key"
    }
    Public ReadOnly Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
    End Property
    Public ReadOnly Property Description() As String
        Get
            Return mDescription(mState)
        End Get
    End Property
    Private mState As Byte
    Public ReadOnly Property Instructions() As String
        Get
            Return mInstructions(mState)
        End Get
    End Property




    Public Sub New()
        Me.State = StateOptions.Navigate
    End Sub
    Private Property mNonNavState As Byte = 1
    Private mOldState As Byte
    Public Property OldState() As Byte
        Get
            Return mOldState
        End Get
        Set(ByVal value As Byte)
            mOldState = value
        End Set
    End Property
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            mOldState = mState
            mState = value

            If mState <> StateOptions.Navigate Then
                mNonNavState = value
                mState = value

            End If
            RaiseEvent StateChanged(Me, New EventArgs)
        End Set
    End Property
    Public Sub ToggleState()
        Select Case Me.State
            Case StateOptions.Navigate
                Me.State = mNonNavState
            Case Else
                Me.State = StateOptions.Navigate
        End Select

    End Sub
    Public Sub IncrementState()
        Me.State = (Me.State + 1) Mod (StateOptions.MoveOriginal + 1)
    End Sub
End Class

