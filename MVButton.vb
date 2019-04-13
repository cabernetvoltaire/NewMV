Imports System.IO
Public Class MVButton
    Public Event LabelChanged(ByVal label As String)
    Private mPath As String
    Public Sub New()
        mPath = ""
        'mFaceText = ""
        mlblText = ""
        mActive = True
        mColour = Color.Black
    End Sub
    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            If value <> mPath And value <> "" Then
                Dim m As New IO.DirectoryInfo(value)
                If m.Exists Then
                    mPath = value
                    Label = m.Name

                End If
            End If
        End Set
    End Property
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
    Public Property Label() As String
        Get
            Return mlblText
        End Get
        Set(ByVal value As String)
            If value <> mlblText Then RaiseEvent LabelChanged(value)
            mlblText = value
        End Set
    End Property
    Private mActive As Boolean
    Public Property Active() As Boolean
        Get
            Return mActive
        End Get
        Set(ByVal value As Boolean)
            mActive = value
        End Set
    End Property
    Public Property Position As Byte
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


End Class
