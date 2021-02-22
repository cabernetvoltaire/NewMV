Imports System.IO
Imports IWshRuntimeLibrary
Public Class NewShortcutHandler
    Public ShortcutDate As Date
    Public PlaybackOffset As Long = 0
    Public MarkOffset As Long = 1
    Private mName As String
    ''' <summary>
    ''' Short name of the link
    ''' </summary>
    Public Property Name() As String
        Get
            Return mName
        End Get
        Set(ByVal value As String)
            mName = value
        End Set
    End Property
    Private mTarget As String
    ''' <summary>
    ''' Full path of target file
    ''' </summary>
    Public Property Target() As String
        Get
            Return mTarget
        End Get
        Set(ByVal value As String)
            mTarget = value
        End Set
    End Property
    Private mBookmark As Integer
    Public Property Bookmark() As Integer
        Get
            Return mBookmark
        End Get
        Set(ByVal value As Integer)
            mBookmark = value
        End Set
    End Property
    Private mPath As String
    Public Property Path() As String
        Get
            Return mPath
        End Get
        Set(ByVal value As String)
            mPath = value
        End Set
    End Property
    Private mTags As List(Of String)
    Public Property Tags() As List(Of String)
        Get
            Return mTags
        End Get
        Set(ByVal value As List(Of String))
            mTags = value
        End Set
    End Property
    Public Sub Save()

    End Sub
    Public Sub Read()

    End Sub
    Private Function Create_ShortCut(Optional bkmk As Long = -1) As String

        Dim sName As String
        Dim f As IO.FileInfo
        f = New IO.FileInfo(mPath & "\" & mName)
        If f.Extension <> LinkExt Then
            sName = mPath & "\" & mName & LinkExt
        Else
            sName = f.FullName
        End If
        Dim dt As Date
        Dim exf As New IO.FileInfo(sName)
        If exf.Exists Then
            dt = exf.LastWriteTime
            exf.Delete()
        End If

        Dim d As New IO.DirectoryInfo(mPath)
        If d.Exists Then
        Else
            d.Create()
        End If
        Try
            sName = ShortCutToText(bkmk)
            Dim newshtcut As New IO.FileInfo(sName)
            newshtcut.LastWriteTime = dt
        Catch ex As Exception
        End Try

        Return sName
    End Function
    Private Function ShortCutToText(Optional bkmk As Long = -1) As String
        If bkmk <> -1 Then
            If InStr(mName, "%") <> 0 Then
                Dim m() As String = mName.Split("%")
                mName = m(0) & "%" & Str(bkmk - MarkOffset) & "%" & m(m.Length - 1)
            Else
                Dim x As New IO.FileInfo(mPath & "\" & mName)
                mName = mName.Replace(x.Extension, x.Extension & "%" & Str(bkmk - MarkOffset) & "%" & LinkExt)
            End If
        End If
        Dim file As New List(Of String) From {mTarget, bkmk - MarkOffset}
        Dim name = mPath & "\" & mName
        WriteListToFile(file, name, False)
        Return name
    End Function
End Class
Public Class ShortcutHandler
    Public Sub New()
        Dim TargetPath As String = ""
        Dim ShortCutPath As String = ""
        Dim ShortCutName As String = ""

    End Sub
    Public ShortcutDate As Date
    Public PlaybackOffset As Long = 0
    Public MarkOffset As Long = 1
    Public Sub New(TargetPath, ShortCutPath, shortCutName)
        sTargetPath = TargetPath
        sShortcutPath = ShortCutPath
        sShortcutName = shortCutName
    End Sub
    Private oShell As WshShell
    Private oShortcut As WshShortcut
    Private sTargetPath As String
    Private sShortcutPath As String
    Private sShortcutName As String
    Public Property TargetPath() As String
        Get

            Return sTargetPath
        End Get
        Set(ByVal value As String)
            If value.EndsWith(".lnk") Or value.EndsWith(LinkExt) Then
                value = LinkTarget(value)

            End If
            sTargetPath = value
        End Set
    End Property
    Public Property ShortcutPath() As String
        Get
            Return sShortcutPath
        End Get
        Set(ByVal value As String)
            sShortcutPath = value
        End Set
    End Property
    Public Property ShortcutName() As String
        Get
            Return sShortcutName
        End Get
        Set(ByVal value As String)
            Dim m() As String = value.Split("%")
            If m.Length > 1 Then
                mBookmark = m(m.Length - 2)

            End If
            sShortcutName = value
        End Set
    End Property

    Private mBookmark As Long
    Public Property Bookmark() As Long
        Get
            Return mBookmark + PlaybackOffset
        End Get
        Set(ByVal value As Long)
            mBookmark = value
        End Set
    End Property
    Public Function New_Create_ShortCut(Optional bkmk As Long = -1) As String

        Dim sName As String
        Dim f As IO.FileInfo
        f = New IO.FileInfo(sShortcutPath & "\" & sShortcutName)
        If f.Extension <> LinkExt Then
            sName = sShortcutPath & "\" & sShortcutName & LinkExt
        Else
            sName = f.FullName

        End If
        Dim dt As Date
        Dim exf As New IO.FileInfo(sName)
        If exf.Exists Then
            dt = exf.LastWriteTime
            exf.Delete()
        End If

        Dim d As New IO.DirectoryInfo(sShortcutPath)
        If d.Exists Then
        Else
            d.Create()
        End If
        Try
            sName = ShortCutToText(bkmk)
            Dim newshtcut As New IO.FileInfo(sName)
            newshtcut.LastWriteTime = dt
        Catch ex As Exception
        End Try

        Return sName
    End Function




    Private Function ShortCutToText(Optional bkmk As Long = -1) As String
        If bkmk <> -1 Then
            If InStr(sShortcutName, "%") <> 0 Then
                Dim m() As String = sShortcutName.Split("%")
                sShortcutName = m(0) & "%" & Str(bkmk - MarkOffset) & "%" & m(m.Length - 1)
            Else
                Dim x As New IO.FileInfo(sShortcutName)
                sShortcutName = sShortcutName.Replace(x.Extension, x.Extension & "%" & Str(bkmk - MarkOffset) & "%" & LinkExt)
            End If
        End If
        Dim file As New List(Of String) From {sTargetPath, bkmk - MarkOffset}
        Dim name = sShortcutPath & "\" & sShortcutName
        WriteListToFile(file, name, False)
        Return name
    End Function



    Public Sub ReAssign(ByVal sTargetPath As String, sShortCutPath As String)
        Dim d As New FileInfo(sShortCutPath)

        CreateLink(Me, sTargetPath, d.Directory.FullName, d.Name, False, Bookmark)

    End Sub


End Class

