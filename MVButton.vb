Imports System.IO
Public Class MVButton
    'Inherits Button
    Public Event ButtonChanged(sender As Object, e As EventArgs)
    Private mPath As String
    Public Sub New()
        mPath = ""
        'mFaceText = ""
        mlblText = ""
        mActive = True
        mColour = Color.Black
        Empty = True
    End Sub
    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            If value = "" Then
                'Label = m.Name
                Empty = True
                mColour = Color.Gray
                Active = False

            Else

                Dim m As New IO.DirectoryInfo(value)
                If m.Exists Then
                    mPath = value
                    Label = m.Name
                    Empty = False
                    Active = True
                Else
                    'Label = m.Name
                    Empty = True
                    mColour = Color.Gray
                    Active = False
                End If
            End If
            RaiseEvent ButtonChanged(Me, Nothing)
        End Set
    End Property
    ''' <summary>
    ''' What is on the face of the button. Usually f(n)
    ''' </summary>
    Private mFaceText As String
    Public Property FaceText() As String
        Get
            Return mFaceText
        End Get
        Set(ByVal value As String)
            mFaceText = value
        End Set
    End Property
    Private mlblText As String
    ''' <summary>
    ''' The text shown in the button label
    ''' </summary>
    Public Property Label() As String
        Get
            Return mlblText
        End Get
        Set(ByVal value As String)
            If value <> mlblText Then
                mlblText = value
                RaiseEvent ButtonChanged(Me, Nothing)
            End If

        End Set
    End Property
    Private mActive As Boolean
    Public Property Active() As Boolean
        Get
            Return mActive
        End Get
        Set(ByVal value As Boolean)
            mActive = value
            If mActive Then
                mColour = Color.Black
            Else
                mColour = Color.Gray
            End If
        End Set
    End Property
    Private Property mPosition
    Public Property Position As Byte
        Set(value As Byte)
            mPosition = value
            mFaceText = "f" & Str(value + 5)
        End Set
        Get
            Return mPosition
        End Get
    End Property

    ''' <summary>
    ''' This Letter is an integer representing the position from A-Z
    ''' and then from 1 to 0
    ''' </summary>
    ''' <returns></returns>
    Public Property Letter As Integer
    Private mColour As Color
    Public Property Colour() As Color
        Get
            Return mColour
        End Get
        Set(ByVal value As Color)
            mColour = value
        End Set
    End Property
    Public Property Empty As Boolean = True
    Public Sub Clear()
        Empty = True
        mPath = ""
        mlblText = ""
        mColour = Color.Black
        Active = False
        RaiseEvent ButtonChanged(Me, Nothing)

    End Sub

End Class
