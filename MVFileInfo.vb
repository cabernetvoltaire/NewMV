Imports System.Xml.Serialization
Imports System.IO
Public Class MVFileInfo

    Private _Filename As String
    Private _Fullpath As String
    Private _Markers As List(Of Double)
    Private _Tags As New List(Of String)
    Private Property _Storepath As String

    Public Property Filename As String
        Get
            Return _Filename
        End Get
        Set
            _Filename = Value
        End Set
    End Property

    Public Property Fullpath As String
        Get
            Return _Fullpath
        End Get
        Set
            _Fullpath = Value
            Dim x() As String = Split(Value, "\")
            For i = 0 To x.Length - 1
                Tags.Add(x(i))
            Next
        End Set
    End Property

    Public Property Markers As List(Of Double)
        Get
            Return _Markers
        End Get
        Set
            _Markers = Value
            Save(_Storepath)
        End Set
    End Property

    Public Property Tags As List(Of String)
        Get
            Return _Tags
        End Get
        Set
            _Tags = Value
        End Set
    End Property
    Public Sub New()

    End Sub
    Public Sub New(filepath As String, storepath As String)
        Dim f As New IO.FileInfo(filepath)
        If f.Exists Then
            Fullpath = f.FullName
            _Filename = f.Name
            _Storepath = storepath & "\" & _Filename & ".mv2"

        Else
            Throw New IO.FileNotFoundException
        End If
        Save(_Storepath)
    End Sub
    Public Sub New(sh As ShortcutHandler)
        Dim f As New IO.FileInfo(sh.TargetPath)
        If f.Exists Then
            _Fullpath = f.FullName
            _Filename = f.Name
            _Storepath = sh.ShortcutPath & "\" & _Filename & ".mv2"
        Else
            Throw New IO.FileNotFoundException
        End If
        Save(_Storepath)



    End Sub

    Public Sub AddTag(tag As String)
        If _Tags.Contains(tag) Then
        Else
            _Tags.Add(tag)
        End If
    End Sub

    Public Sub AddMarker(marker As Long)
        If _Markers.Contains(marker) Then
        Else
            _Markers.Add(marker)
        End If
    End Sub

    Public Sub NewPath(path As String)
        _Fullpath = path
    End Sub


    Private Sub Save(path As String)
        Dim serializer As New XmlSerializer(GetType(MVFileInfo))
        Dim writer As New StreamWriter(path)
        serializer.Serialize(writer, Me)
        writer.Close()
    End Sub


End Class
