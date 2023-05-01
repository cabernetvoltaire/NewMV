Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Threading
Imports System.Text

Public Class FileGrouper
    Private ReadOnly _filePaths As List(Of String) = New List(Of String)
    Private ReadOnly _groupSizes As List(Of Integer) = New List(Of Integer)
    Private _groupThread As Thread
    Public Event FilesMoved()

    Public WriteOnly Property FilePaths As List(Of String)
        Set(value As List(Of String))
            _filePaths.Clear()
            _filePaths.AddRange(value)
        End Set
    End Property

    Public Sub GroupFiles(groupSize As Integer, targetFolder As String)
        _groupSizes.Add(groupSize)
        _groupThread = New Thread(Sub() MoveFilesToSubfolders(targetFolder))
        _groupThread.Start()
    End Sub
    Function GroupFilesBySubstring(filePaths As List(Of String)) As List(Of List(Of String))
        Dim uniqueFileNames As New List(Of String)(filePaths.Select(Function(path) IO.Path.GetFileNameWithoutExtension(path)))
        Dim groups As New List(Of List(Of String))

        For Each fileName In uniqueFileNames
            Dim matchedGroup As List(Of String) = Nothing
            Dim maxCommonSubstringLength As Integer = 0

            For Each group In groups
                Dim commonSubstringLength As Integer = GetLongestCommonSubstringLength(group(0), fileName)
                If commonSubstringLength > maxCommonSubstringLength AndAlso commonSubstringLength > 1 Then
                    matchedGroup = group
                    maxCommonSubstringLength = commonSubstringLength
                End If
            Next

            If matchedGroup IsNot Nothing Then
                matchedGroup.Add(fileName)
            Else
                groups.Add(New List(Of String) From {fileName})
            End If
        Next

        ' Replace unique file names with full paths
        For i As Integer = 0 To groups.Count - 1
            For j As Integer = 0 To groups(i).Count - 1
                groups(i)(j) = filePaths.First(Function(path) IO.Path.GetFileNameWithoutExtension(path) = groups(i)(j))
            Next
        Next

        Return groups
    End Function
    Function GetLongestCommonSubstringLength(str1 As String, str2 As String) As Integer
        Dim lengths = New Integer(str1.Length, str2.Length) {}
        Dim maxLength As Integer = 0

        For i As Integer = 0 To str1.Length - 1
            For j As Integer = 0 To str2.Length - 1
                If str1(i) = str2(j) Then
                    lengths(i + 1, j + 1) = lengths(i, j) + 1
                    maxLength = Math.Max(maxLength, lengths(i + 1, j + 1))
                Else
                    lengths(i + 1, j + 1) = 0
                End If
            Next
        Next

        Return maxLength
    End Function



    Function FindCommonSubstring(str1 As String, str2 As String) As String
        Dim common As New StringBuilder
        Dim filename1WithoutExtension As String = Path.GetFileNameWithoutExtension(str1)
        Dim filename2WithoutExtension As String = Path.GetFileNameWithoutExtension(str2)

        For i As Integer = 0 To Math.Min(filename1WithoutExtension.Length, filename2WithoutExtension.Length) - 1
            If filename1WithoutExtension(i) = filename2WithoutExtension(i) Then
                common.Append(filename1WithoutExtension(i))
            Else
                Exit For
            End If
        Next

        Return common.ToString()
    End Function

    Private Sub MoveFilesToSubfolders(targetFolder As String)
        Dim groups = GroupFilesBySubstring(_filePaths)
        For Each group In groups
            Dim groupFolder = Path.Combine(targetFolder, $"Like {Path.GetFileNameWithoutExtension(group(0))}")
            Directory.CreateDirectory(groupFolder)
            For Each filePath In group
                Dim fileName = Path.GetFileName(filePath)
                Dim targetPath = Path.Combine(groupFolder, fileName)
                File.Move(filePath, targetPath)
            Next
        Next
        RaiseEvent FilesMoved()
    End Sub

    Private Function GetGroups() As List(Of List(Of String))
        Dim groups = New List(Of List(Of String))
        For Each groupSize In _groupSizes
            Dim currentGroups = _filePaths.GroupBy(Function(p) Path.GetFileNameWithoutExtension(p).
                                                        Substring(0, Math.Min(groupSize, Path.GetFileNameWithoutExtension(p).Length))).
                                                        Where(Function(g) g.Count() >= 2).
                                                        Select(Function(g) g.ToList()).
                                                        ToList()
            groups.AddRange(currentGroups)
        Next
        Return groups
    End Function
End Class

Public Class FilenamesGrouper
    Public Event WordsParsed()
    Private RegexOptions() As String = {"[A-Z|a-z]*[ ]", "[[A-Z|a-z]*[0-9]{3,}]|[\w]*", "(([A-Z|a-z]*[AEIOUY|aeiouy][A-Z|a-z]*[ |_|-|,]*)*)", "([A-Z|a-z]*[ |_|-]*)", "([A-Z|a-z]+[ |_|-]*)+", "([0-9]+){3}"}
    Private mWordlist As New SortedList(Of String, Integer)
    Private mOptionNumber As Integer = 0
    Private mFilenames As New List(Of String)
    Private mGroups As New List(Of List(Of String))
    Private mGroupNames As New List(Of String)
    Private mMaxDistance As Integer = 8

    Public Property WordList() As SortedList(Of String, Integer)
        Get
            Return mWordlist
        End Get
        Set(ByVal value As SortedList(Of String, Integer))
            mWordlist = value
        End Set
    End Property

    Public Property Count As Integer = 2

    Public Property MaxDistance As Integer
        Get
            Return mMaxDistance
        End Get
        Set(ByVal value As Integer)
            mMaxDistance = value
        End Set
    End Property

    Public Property Filenames() As List(Of String)
        Get
            Return mFilenames
        End Get
        Set(ByVal value As List(Of String))
            mFilenames = value
            mGroups.Clear()
            mGroupNames.Clear()
            mWordlist.Clear()
            ParseNames()
        End Set
    End Property

    Public Property Groups As List(Of List(Of String))
        Get
            Return mGroups
        End Get
        Set(ByVal value As List(Of List(Of String)))
            mGroups = value
        End Set
    End Property

    Public Property GroupNames() As List(Of String)
        Get
            Return mGroupNames
        End Get
        Set(ByVal value As List(Of String))
            mGroupNames = value
        End Set
    End Property

    Private Sub ParseNames()
        For Each f In mFilenames
            For i = 0 To RegexOptions.Length - 1
                CreateWordList(f, RegexOptions(i))
            Next
        Next
        UpdateWordlist()
        GetGroups()
        If mWordlist.Count <> 0 Then RaiseEvent WordsParsed()
    End Sub

    Private Sub CreateWordList(str As String, rgx As String)
        Dim r As New Regex(rgx)
        If r.Matches(str).Count > 0 Then
            For Each s In r.Matches(str)
                Dim match As String = s.ToString
                If mWordlist.ContainsKey(match) Then
                    Dim k As Integer = mWordlist.Item(match)
                    k += 1
                    mWordlist.Remove(match)
                    mWordlist.Add(match, k)
                Else
                    mWordlist.Add(match, 1)
                End If
            Next
        End If
    End Sub
    Private Sub UpdateWordlist()
        Dim x As New SortedList(Of String, Integer)

        For Each m In mWordlist
            For i = mFilenames.Count - 1 To Count Step -1
                If m.Value = mFilenames.Count Or m.Value < i Or InStr(PICEXTENSIONS & VIDEOEXTENSIONS & LinkExt, m.Key) <> 0 Then
                Else
                    If x.Keys.Contains(m.Key) Then
                    Else
                        x.Add(m.Key, m.Value)
                    End If
                End If
            Next
        Next
        mWordlist = x
    End Sub



    Private Function ComputeLevenshteinDistance(s As String, t As String) As Integer
        Dim n As Integer = s.Length
        Dim m As Integer = t.Length
        Dim d(n + 1, m + 1) As Integer

        If n = 0 Then
            Return m
        End If

        If m = 0 Then
            Return n
        End If

        For i = 0 To n
            d(i, 0) = i
        Next

        For j = 0 To m
            d(0, j) = j
        Next

        For i = 1 To n
            For j = 1 To m
                Dim cost As Integer = If(t(j - 1) = s(i - 1), 0, 1)
                d(i, j) = Math.Min(Math.Min(d(i - 1, j) + 1, d(i, j - 1) + 1), d(i - 1, j - 1) + cost)
            Next
        Next

        Return d(n, m)
    End Function

    Private Function GetCommonSubstring(group As List(Of String)) As String
        Dim maxLength As Integer = group.Min(Function(s) s.Length)
        Dim result As String = ""
        For i = 1 To maxLength
            Dim substrings As New List(Of String)
            For Each s In group
                substrings.Add(IO.Path.GetFileName(IO.Path.GetDirectoryName(s)))
            Next
            If substrings.Distinct().Count() = 1 Then
                result = substrings.First()
            Else
                Exit For
            End If
        Next
        Return result
    End Function

    Private Sub GetGroups()
        mGroups.Clear()
        Dim usedIndexes As New List(Of Integer)
        For i = 0 To mFilenames.Count - 1
            If Not usedIndexes.Contains(i) Then
                Dim group As New List(Of String)
                For j = i To mFilenames.Count - 1
                    If Not usedIndexes.Contains(j) Then
                        Dim distance = ComputeLevenshteinDistance(IO.Path.GetFileName(mFilenames(i)), IO.Path.GetFileName(mFilenames(j)))
                        If distance <= MaxDistance Then
                            group.Add(mFilenames(j))
                            usedIndexes.Add(j)
                        End If
                    End If
                Next
                If group.Count >= Count Then
                    mGroups.Add(group)
                    mGroupNames.Add(IO.Path.GetFileNameWithoutExtension(IO.Path.GetFileName(mFilenames(i))))
                End If
            End If
        Next
    End Sub



    Public Sub Clear()
        mFilenames.Clear()
        mGroupNames.Clear()
        mGroups.Clear()
    End Sub

    Public Sub AdvanceOption()
        mOptionNumber = (mOptionNumber + 1) Mod RegexOptions.Length
    End Sub
End Class


'Public Class FileNamesGrouper

'    Public Event WordsParsed()
'    Private RegexOptions() As String = {"[A-Z|a-z]*[ ]", "[[A-Z|a-z]*[0-9]{3,}]|[\w]*", "(([A-Z|a-z]*[AEIOUY|aeiouy][A-Z|a-z]*[ |_|-|,]*)*)", "([A-Z|a-z]*[ |_|-]*)", "([A-Z|a-z]+[ |_|-]*)+", "([0-9]+){3}"}
'    Private mWordlist As New SortedList(Of String, Integer)
'    Private mOptionNumber
'    Public Property WordList() As SortedList(Of String, Integer)
'        Get
'            Return mWordlist
'        End Get
'        Set(ByVal value As SortedList(Of String, Integer))
'            mWordlist = value
'        End Set
'    End Property

'#Region "Properties"
'    Public Property Count As Integer = 2
'    Private mFilenames As New List(Of String)
'    Public Property Filenames() As List(Of String)
'        Get
'            Return mFilenames
'        End Get
'        Set(ByVal value As List(Of String))
'            mFilenames = value
'            mGroups.Clear()
'            mGroupNames.Clear()
'            mWordlist.Clear()
'            ParseNames()
'        End Set
'    End Property
'    Private mGroups As New List(Of List(Of String))
'    Public Property Groups As List(Of List(Of String))
'        Get
'            Return mGroups

'        End Get
'        Set(ByVal value As List(Of List(Of String)))
'            mGroups = value

'        End Set
'    End Property
'    Private mGroupNames As New List(Of String)
'    Public Property GroupNames() As List(Of String)
'        Get
'            Return mGroupNames
'        End Get
'        Set(ByVal value As List(Of String))
'            mGroupNames = value
'        End Set
'    End Property

'#End Region
'#Region "Methods"
'    Private Sub ParseNames()
'        OutputList("ParseNames", mFilenames)
'        For Each f In mFilenames
'            ' For i = RegexOptions.Length - 1 To 0 Step -1
'            CreateWordList(f, RegexOptions(mOptionNumber))
'            '   StringParser(f, RegexOptions(2))

'            'Next
'        Next
'        OutputList("mWordlist", mWordlist)
'        UpdateWordlist()
'        OutputList("mWordList after Update", mWordlist)

'        GetGroups()
'        '  REcurseGroups()

'        If mWordlist.Count <> 0 Then RaiseEvent WordsParsed()
'    End Sub

'    Public Sub Clear()
'        mFilenames.Clear()
'        mGroupNames.Clear()
'        mGroups.Clear()
'    End Sub

'    Private Sub OutputList(s As String, x As Object)
'        'Exit Sub
'        Debug.Print(" ")
'        Debug.Print(UCase(s))
'        For Each m In x
'            Debug.Print(m.ToString)

'        Next
'    End Sub
'    Private Sub CreateWordList(str As String, rgx As String)
'        Dim r As New System.Text.RegularExpressions.Regex(rgx) '("([A-Z]^[0-9])\w+")
'        ' Console.WriteLine(str)
'        If r.Matches(str).Count > 0 Then
'            Dim localwords As New List(Of String)
'            'OutputList("Matches", r.Matches(str))
'            For Each s In r.Matches(str)
'                Dim match As String = s.ToString
'                'match = LCase(match)
'                If mWordlist.ContainsKey(match) Then
'                    If localwords.Contains(match) Then
'                    Else
'                        Dim k As Integer = mWordlist.Item(match) 'currently counts multiple occurrences in same str
'                        k = k + 1
'                        mWordlist.Remove(match)
'                        mWordlist.Add(match, k)
'                        localwords.Add(match)
'                    End If
'                Else
'                    mWordlist.Add(match, 1)
'                    localwords.Add(match)

'                End If

'            Next
'            OutputList("Local words", localwords)
'        End If

'    End Sub

'    ''' <summary>
'    ''' For each filename, adds to a lists which pair each name with a target group and its associated count
'    ''' </summary>
'    Private Sub GetGroups()
'        'Add each file to a countofgroup and a targetgroup
'        mGroups.Clear()
'        Dim CountOfGroup As New SortedList(Of String, Integer) 'Filenames, Count
'        Dim TargetGroup As New SortedList(Of String, String) 'Filenames, Target
'        For Each filefromlist In mFilenames
'            CountOfGroup.Add(filefromlist, 0)

'            TargetGroup.Add(filefromlist, "")

'            For Each foundword In mWordlist 'Go through wordlist and allocate files which contain it

'                If InStr(filefromlist, foundword.Key) <> 0 And CountOfGroup(filefromlist) < foundword.Value Then 'This is where the problem lies


'                    If foundword.Value > 0 And foundword.Key <> "" Then
'                        CountOfGroup(filefromlist) = foundword.Value
'                        TargetGroup(filefromlist) = foundword.Key

'                    End If
'                End If

'            Next
'        Next
'        OutputList("Count Of Group", CountOfGroup)
'        OutputList("Target Group", TargetGroup)
'        Dim targets As New List(Of String)
'        'Construct a new file list which only 
'        For Each filefromtargetgroup In TargetGroup
'            If targets.Contains(filefromtargetgroup.Value) Or filefromtargetgroup.Value = "" Then 'Don't repeat, don't add empties
'            Else
'                targets.Add(filefromtargetgroup.Value)
'            End If
'        Next
'        OutputList("targets", targets)
'        For Each tgt In targets
'            Dim flst As New List(Of String)
'            For Each f In TargetGroup
'                If f.Value = tgt Then
'                    flst.Add(f.Key)
'                End If
'            Next
'            If flst.Count >= Count Then
'                mGroups.Add(flst)
'                mGroupNames.Add(tgt)
'            End If
'        Next

'    End Sub



'    ''' <summary>
'    ''' Goes through the Wordlist and doesn't copy any entries which are singletons, or universal
'    ''' </summary>
'    ''' 
'    Private Sub UpdateWordlist()
'        'm.value is the count of files containing m.key
'        'Remove words which are universal
'        'Also remove file extensions
'        'Rank the names by m.value
'        '
'        Dim x As New SortedList(Of String, Integer)

'        For Each m In mWordlist
'            For i = mFilenames.Count - 1 To Count Step -1

'                If m.Value = mFilenames.Count Or m.Value < i Or InStr(PICEXTENSIONS & VIDEOEXTENSIONS & LinkExt, m.Key) <> 0 Then
'                Else
'                    If x.Keys.Contains(m.Key) Then
'                    Else
'                        x.Add(m.Key, m.Value)

'                    End If
'                End If
'            Next
'        Next
'        mWordlist = x
'    End Sub
'    Public Sub AdvanceOption()
'        mOptionNumber = (mOptionNumber + 1) Mod RegexOptions.Length

'    End Sub
'#End Region

'End Class




