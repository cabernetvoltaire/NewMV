Imports System.IO
Imports Shell32
Public Class MetaData
    Private _filepath As String
    Private _Properties As Dictionary(Of Integer, KeyValuePair(Of String, String))
    Private _PropertyString As String

    Public Property Filepath As String
        Get
            Return _filepath
        End Get
        Set(value As String)
            _filepath = value
            Properties = GetFileProperties(_filepath)
        End Set
    End Property
    Public Sub New(Pathname As String)

        ' This call is required by the designer.
        InitializeComponent()
        Filepath = Pathname
        ' Add any initialization after the InitializeComponent() call.
    End Sub


    Public Property Properties As Dictionary(Of Integer, KeyValuePair(Of String, String))
        Get
            Return _Properties
        End Get
        Set(value As Dictionary(Of Integer, KeyValuePair(Of String, String)))
            _Properties = value
        End Set
    End Property
    Public ReadOnly Property PropertyString As String
        Get
            Return StringFromProperties()
        End Get

    End Property

    Private Function GetFileProperties(ByVal FileName As String) As Dictionary(Of Integer, KeyValuePair(Of String, String))
        Dim Shell As New Shell
        Dim Folder As Folder = Shell.[NameSpace](Path.GetDirectoryName(FileName))
        Dim File As FolderItem = Folder.ParseName(Path.GetFileName(FileName))
        Dim Properties As New Dictionary(Of Integer, KeyValuePair(Of String, String))()
        Dim Index As Integer
        Dim Keys As Integer = Folder.GetDetailsOf(File, 0).Count
        For Index = 0 To Keys - 1
            Dim CurrentKey As String = Folder.GetDetailsOf(Nothing, Index)
            Dim CurrentValue As String = Folder.GetDetailsOf(File, Index)
            If CurrentValue <> "" Then
                Properties.Add(Index, New KeyValuePair(Of String, String)(CurrentKey, CurrentValue))
            End If
        Next
        Return Properties
    End Function

    Private Function StringFromProperties() As String
        Dim s As String = ""
        If _Properties.Count > 0 Then

            For i = 0 To _Properties.Count - 1
                Try
                    s &= vbCrLf & Properties(i).ToString

                Catch ex As Exception
                    Continue For
                End Try
            Next
        End If
        Return s
    End Function
End Class