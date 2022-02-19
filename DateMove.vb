''' <summary>
''' Given a folder, move all files whose names contain the name of a subfolder to that subfolder
''' </summary>

Class DateMove
    Public Event FilesMoved(sender As Object, e As EventArgs)
    Private mFolder As String

    Private ReadOnly mFolders As New List(Of IO.DirectoryInfo)
    Private mYears As String = ""
    Public Enum DMY As Byte
        Year
        Month
        Day
        Hour
        Minute
        Calendar

    End Enum
    Private ReadOnly MonthNames As String() = {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
    Private ReadOnly SizeNames As String() = {"Size 0- Tiny", "Size 1-Small", "Size 2-Medium", "Size 3-Large", "Size 4-Very Large", "Size 5-Gigantic"}
    Private ReadOnly AlphaName As String() = {"A-F", "G-L", "M-R", "S-Z"}
    Public Property Folder() As String
        Get
            Return mFolder
        End Get
        Set(ByVal value As String)
            mFolder = value
        End Set
    End Property

    Public Sub New()
        mRecursive = False
    End Sub

    Private mRecursive As Boolean
    Public Property Recursive() As Boolean
        Get
            Return mRecursive
        End Get
        Set(ByVal value As Boolean)
            mRecursive = value
        End Set
    End Property
    Public Sub FilterByCalendar(Folder As IO.DirectoryInfo)
        FilterByDate(Folder.FullName, False, DMY.Year)
        For Each f In Folder.GetDirectories
            If InStr(mYears, f.Name) <> 0 Then
                FilterByDate(f.FullName, False, DMY.Month)

            End If
            For Each m In f.GetDirectories
                For i = 0 To 11
                    If m.Name = MonthNames(i) Then
                        FilterByDate(m.FullName, False, DMY.Day)
                    End If
                Next
            Next

        Next
    End Sub
    Public Sub FilterByDate(FolderName As String, Recurse As Boolean, Choice As DMY)

        Dim s As New IO.DirectoryInfo(FolderName)
        Dim i As Integer
        Dim files As New Dictionary(Of IO.FileInfo, String)

        For Each f In s.GetFiles

            Dim folname As String = ""
            With GetDate(f)
                Select Case Choice
                    Case DMY.Year
                        i = .Year
                        folname = i
                        mYears = mYears & folname
                    Case DMY.Month
                        i = .Month
                        folname = i
                        If i < 10 Then folname = "0" & folname

                        folname = folname & " - " & MonthName(i)
                    Case DMY.Day
                        i = .Day
                        folname = i
                        If i < 10 Then folname = "0" & folname
                    Case DMY.Hour
                        i = .Hour
                        folname = i
                        If i < 10 Then folname = "0" & folname

                    Case DMY.Minute
                        i = .Minute
                        folname = i
                        If i < 10 Then folname = "0" & folname
                End Select
            End With
            files.Add(f, folname)
            'If f.Directory.EnumerateDirectories(Str(i)) Is Nothing Then
            'End If
            'Try
            '    f.Directory.CreateSubdirectory(folname & "\")
            '    f.MoveTo(f.DirectoryName & "\" & folname & "\" & f.Name)

            'Catch ex As Exception

            'End Try
        Next
        For Each f In files
            s.CreateSubdirectory(f.Value & "\")
            Try
                f.Key.MoveTo(s.FullName & "\" & f.Value & "\" & f.Key.Name)

            Catch ex As Exception
                Continue For
            End Try
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)

    End Sub
    Public Sub FilterBySize(FolderName As String, Recurse As Boolean)

        Dim s As New IO.DirectoryInfo(FolderName)
        For Each f In s.GetFiles
            For m = 5 To 10
                If f.Length < 10 ^ m And f.Length >= 10 ^ (m - 1) Then
                    f.Directory.CreateSubdirectory(SizeNames(m - 5) & "\")
                    f.MoveTo(f.DirectoryName & "\" & SizeNames(m - 5) & "\" & f.Name)
                End If
            Next

            'If f.Directory.EnumerateDirectories(Str(i)) Is Nothing Then
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)

    End Sub

    Public Sub FilterByLinkFolder(FolderName As String)
        Dim s As New IO.DirectoryInfo(FolderName)
        For Each f In s.GetFiles
            If f.Extension = LinkExt Then
                Try

                    Dim tgt As New IO.FileInfo(LinkTarget(f.FullName))
                    Dim parentname As String = tgt.Directory.Name
                    If parentname = tgt.Directory.Root.FullName Then parentname = ""
                    Dim m As New IO.DirectoryInfo(f.Directory.FullName & "\" & parentname)
                    If Not m.Exists Then
                        m.Create()
                    End If
                    f.MoveTo(m.FullName & "\" & f.Name)
                Catch ex As Exception

                End Try

            End If
        Next
        RaiseEvent FilesMoved(Nothing, Nothing)
    End Sub
    Public Sub FilterByType(FolderName As String)
        Dim s As New IO.DirectoryInfo(FolderName)


        For Each f In s.GetFiles
            If f.Directory.GetDirectories(f.Extension).Length > 0 Then
            Else
                f.Directory.CreateSubdirectory(f.Extension & "\")
            End If

            Dim destfile As String = f.DirectoryName & "\" & f.Extension
            blnSuppressCreate = True
            Dim l As New List(Of String) From {f.FullName}
            MoveFiles(l, destfile)
        Next

        RaiseEvent FilesMoved(Nothing, Nothing)
    End Sub

    Public Sub FilterByAlpha(FolderName As String, Optional Files As Boolean = False, Optional Folders As Boolean = False, Optional Letters As Boolean = False)
        Dim s As New IO.DirectoryInfo(FolderName)
        If Letters Then
            Dim letter As Char
            Dim dest As String = ""
            Dim list As New List(Of String)
            For Each f In s.GetFiles
                letter = UCase(f.Name(0))
                If s.GetDirectories(letter).Length > 0 Then
                Else
                    s.CreateSubdirectory(UCase(letter) & "\")
                End If
            Next
            'Move each file into it. 
            For Each d In s.GetDirectories
                If d.Name.Length = 1 Then
                    For Each f In s.GetFiles
                        If UCase(f.Name(0)) = d.Name Then
                            list.Add(f.FullName)
                        End If
                    Next
                    blnSuppressCreate = True
                    MoveFiles(list, d.FullName)
                    list.Clear()
                End If
            Next
        Else

            If Files Then
                'For each group
                For i = 0 To 3
                    Dim Alpha As String = "Alpha_" & AlphaName(i)
                    Dim list As New List(Of String)
                    Dim dest As String = ""
                    For Each f In s.GetFiles
                        'If first name precedes last letter in the group (or equal to)
                        If LCase(f.Name(0)) <= LCase(AlphaName(i)(Len(AlphaName(i)) - 1)) Then
                            'Create any directories which don't exist
                            If s.GetDirectories(Alpha).Length > 0 Then
                            Else
                                s.CreateSubdirectory(Alpha & "\")
                            End If
                            dest = f.DirectoryName & "\" & Alpha
                            'Move each file into it.
                            list.Add(f.FullName)
                        End If
                    Next
                    blnSuppressCreate = True
                    MoveFiles(list, dest)
                Next
            End If
            If Folders Then
                For i = 0 To 3
                    Dim dest As String
                    Dim Alpha As String = "Alpha_" & AlphaName(i)
                    For Each f In s.EnumerateDirectories("*", IO.SearchOption.TopDirectoryOnly)
                        If LCase(f.Name(0)) <= LCase(AlphaName(i)(Len(AlphaName(i)) - 1)) And LCase(f.Name(0)) >= LCase(AlphaName(i)(0)) Then

                            If s.GetDirectories(Alpha).Length > 0 Then
                            Else
                                If s.FullName.Contains(Alpha) Then
                                Else
                                    s.CreateSubdirectory(Alpha & "\")

                                End If
                            End If
                            dest = s.FullName & "\" & Alpha
                            If Not f.Name.Contains("Alpha") Then MoveFolder(f.FullName, dest)
                        End If
                    Next
                Next
            End If
        End If
        RaiseEvent FilesMoved(Nothing, Nothing)
    End Sub
    'Private Sub GetFoldersBelow(folderpath As String, recurse As Boolean)

    '    Dim s As New IO.DirectoryInfo(folderpath)
    '    For Each m In s.GetDirectories
    '        If recurse Then GetFoldersBelow(m.FullName, recurse)
    '        mFolders.Add(m)
    '    Next

    'End Sub
End Class
Public Class GroupByLinkFolders
    Private Property mFolder As IO.DirectoryInfo
    Public Property Folder As String
        Set(value As String)
            mFolder = New IO.DirectoryInfo(value)
        End Set
        Get
            Return mFolder.FullName
        End Get
    End Property

    Public Files As List(Of String)

    Public Sub MoveFiles(PathPart As Byte)
        For Each f In Files
            If FindType(f) = Filetype.Link Then
                Dim parts() As String = Split(LinkTarget(f), "\")
                If parts.Length >= PathPart Then
                    'Create directory of parts(pathpart) if it doesn't already exist
                    'Move f into that directory
                    Dim file As New IO.FileInfo(f)
                    file.MoveTo(CreateDirectory(parts(PathPart)).FullName & "\" & file.Name)
                End If
            Else

            End If
        Next

    End Sub

    Private Function CreateDirectory(v As String) As IO.DirectoryInfo
        Dim fol As New IO.DirectoryInfo(Folder)
        Return fol.CreateSubdirectory(v)
    End Function
End Class
