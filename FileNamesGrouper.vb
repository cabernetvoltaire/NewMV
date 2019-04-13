Public Class FileNamesGrouper
    Public Event WordsParsed()
    Private RegexOptions() As String = {"[[A-Z|a-z]*[0-9]{3,}]|[\w]*", "(([A-Z|a-z]*[AEIOUY|aeiouy][A-Z|a-z]*[ |_|-|,]*)*)", "([A-Z|a-z]*[ |_|-]*)", "([A-Z|a-z]+[ |_|-]*)+"}
    Private mWordlist As New SortedList(Of String, Integer)
    Public Property WordList() As SortedList(Of String, Integer)
        Get
            Return mWordlist
        End Get
        Set(ByVal value As SortedList(Of String, Integer))
            mWordlist = value
        End Set
    End Property

#Region "Properties"
    Public Property Count As Integer = 2
    Private mFilenames As New List(Of String)
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
    Private mGroups As New List(Of List(Of String))
    Public Property Groups As List(Of List(Of String))
        Get
            Return mGroups

        End Get
        Set(ByVal value As List(Of List(Of String)))
            mGroups = value

        End Set
    End Property
    Private mGroupNames As New List(Of String)
    Public Property GroupNames() As List(Of String)
        Get
            Return mGroupNames
        End Get
        Set(ByVal value As List(Of String))
            mGroupNames = value
        End Set
    End Property

#End Region
#Region "Methods"
    Private Sub ParseNames()
        OutputList("ParseNames", mFilenames)
        For Each f In mFilenames
            ' For i = RegexOptions.Length - 1 To 0 Step -1
            CreateWordList(f, RegexOptions(1))
            '   StringParser(f, RegexOptions(2))

            'Next
        Next
        OutputList("mWordlist", mWordlist)
        UpdateWordlist()
        OutputList("mWordList after Update", mWordlist)

        GetGroups()
        '  REcurseGroups()

        If mWordlist.Count <> 0 Then RaiseEvent WordsParsed()
    End Sub

    Public Sub Clear()
        mFilenames.Clear()
        mGroupNames.Clear()
        mGroups.Clear()
    End Sub

    Private Sub OutputList(s As String, x As Object)
        Exit Sub
        Debug.Print(" ")
        Debug.Print(UCase(s))
        For Each m In x
            Debug.Print(m.ToString)

        Next
    End Sub
    Private Sub CreateWordList(str As String, rgx As String)
        Dim r As New System.Text.RegularExpressions.Regex(rgx) '("([A-Z]^[0-9])\w+")
        ' Console.WriteLine(str)
        If r.Matches(str).Count > 0 Then
            Dim localwords As New List(Of String)
            'OutputList("Matches", r.Matches(str))
            For Each s In r.Matches(str)
                Dim match As String = s.ToString
                'match = LCase(match)
                If mWordlist.ContainsKey(match) Then
                    If localwords.Contains(match) Then
                    Else
                        Dim k As Integer = mWordlist.Item(match) 'currently counts multiple occurrences in same str
                        k = k + 1
                        mWordlist.Remove(match)
                        mWordlist.Add(match, k)
                        localwords.Add(match)
                    End If
                Else
                    mWordlist.Add(match, 1)
                    localwords.Add(match)

                End If

            Next
            OutputList("Local words", localwords)
        End If

    End Sub

    ''' <summary>
    ''' For each filename, adds to a lists which pair each name with a target group and its associated count
    ''' </summary>
    Private Sub GetGroups()
        'Add each file to a countofgroup and a targetgroup
        mGroups.Clear()
        Dim CountOfGroup As New SortedList(Of String, Integer) 'Filenames, Count
        Dim TargetGroup As New SortedList(Of String, String) 'Filenames, Target
        For Each filefromlist In mFilenames
            CountOfGroup.Add(filefromlist, 0)

            TargetGroup.Add(filefromlist, "")

            For Each foundword In mWordlist 'Go through wordlist and allocate files which contain it

                If InStr(filefromlist, foundword.Key) <> 0 And CountOfGroup(filefromlist) < foundword.Value Then


                    If foundword.Value > 0 And foundword.Key <> "" Then
                        CountOfGroup(filefromlist) = foundword.Value
                        TargetGroup(filefromlist) = foundword.Key

                    End If
                End If

            Next
        Next
        OutputList("Count Of Group", CountOfGroup)
        OutputList("Target Group", TargetGroup)
        Dim targets As New List(Of String)
        'Construct a new file list which only 
        For Each filefromtargetgroup In TargetGroup
            If targets.Contains(filefromtargetgroup.Value) Or filefromtargetgroup.Value = "" Then 'Don't repeat, don't add empties
            Else
                targets.Add(filefromtargetgroup.Value)
            End If
        Next
        OutputList("targets", targets)
        For Each tgt In targets
            Dim flst As New List(Of String)
            For Each f In TargetGroup
                If f.Value = tgt Then
                    flst.Add(f.Key)
                End If
            Next
            If flst.Count >= Count Then
                mGroups.Add(flst)
                mGroupNames.Add(tgt)
            End If
        Next

    End Sub



    ''' <summary>
    ''' Goes through the Wordlist and doesn't copy any entries which are singletons, or universal
    ''' </summary>
    ''' 
    Private Sub UpdateWordlist()
        'm.value is the count of files containing m.key
        'Remove words which are universal
        'Also remove file extensions
        'Rank the names by m.value
        '
        Dim x As New SortedList(Of String, Integer)

        For Each m In mWordlist
            For i = mFilenames.Count - 1 To Count Step -1

                If m.Value = mFilenames.Count Or m.Value < i Or InStr(PICEXTENSIONS & VIDEOEXTENSIONS & ".lnk", m.Key) <> 0 Then
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
#End Region

End Class
