
Public Class TextAnalyzer


    Private Enum DocType
        AlternateLinesSpaces
        LinesIndented
        LinesStartSpaces
        ParasIndented
        ParasDefinedbyBlankLine
        HTML
        LeaveRaw
    End Enum
    Private _Lines() As String
    Private _RawText As String
    Private _MeanLineLength As Decimal
    Private _MinLineLength As Integer
    Private _MaxLineLength As Integer
    Private _EmptyLinesCount As Integer
    Private _LineThenSpaceRatio As Decimal
    Private _DocType As DocType
    Public Htmlfound As Boolean
    Public Property RawText() As String
        Get
            Return _RawText
        End Get
        Set(ByVal value As String)
            _RawText = value
            Dim _RawTextCRLFs As String
            _RawTextCRLFs = _RawText.Replace(vbCrLf, "|").Replace(vbLf, "|").Replace(vbCr, "|")
            _Lines = _RawTextCRLFs.Split("|")
            CalcLineStats()
            _DocType = AnalyzeType()
            'MsgBox(_DocType.ToString)
            'RemoveEmptyLines()
        End Set
    End Property
    Public ReadOnly Property CleanText() As String
        Get
            Return ParseText()
        End Get
    End Property
    Private Sub RemoveEmptyLines()
        Dim ll As New List(Of String)
        Dim nn As New List(Of String)
        ll = _Lines.ToList
        For Each l In ll
            If l.Length = 0 Or l.Replace(" ", "").Length = 0 Then
            Else
                nn.Add(l)
            End If
        Next
        _Lines = nn.ToArray
    End Sub
    Private Sub CountEmptyLines()
        Dim ll As New List(Of String)
        ll = _Lines.ToList
        Dim count As Integer
        For Each l In ll
            If l.Length = 0 Or l.Replace(" ", "").Length = 0 Then
                count += 1
            Else

            End If
        Next
        _EmptyLinesCount = count
    End Sub
    Private Function ParseText() As String
        Dim paras As New List(Of String)

        Select Case _DocType
            Case DocType.HTML
                Dim parasarray() As String = _RawText.Replace("<br />", "|").Split("|")
                paras = parasarray.ToList
            Case DocType.LeaveRaw

            Case DocType.LinesIndented
                ParseLinesShortLineEnds(paras)
            Case DocType.LinesStartSpaces
                ParseLinesIndented(paras)

            Case DocType.ParasIndented
                ParseLinesIndented(paras)
            Case DocType.ParasDefinedbyBlankLine
                ParseParasDefinedbyEmptyLine(paras)

            Case Else

        End Select
        Dim returnstring As String = ""
        For Each p In paras
            returnstring = returnstring & vbCrLf & vbCrLf & p
        Next
        Return returnstring
    End Function
    ''' <summary>
    ''' Assumes separate lines, and puts a paragraph break before a line that starts with a space
    ''' 
    ''' </summary>
    ''' <param name="paras"></param>
    Private Sub ParseLinesIndented(paras As List(Of String))
        Dim currentpara As String = ""
        For i = 0 To _Lines.Length - 1
            Dim line As String = _Lines(i)
            If line.Length <> 0 AndAlso line(0) = " " Then
                paras.Add(currentpara)
                currentpara = ""
                currentpara = " " & line
            Else
                currentpara = currentpara & " " & line
            End If
        Next
        paras.Add(currentpara)
    End Sub
    Private Sub ParseParasDefinedbyEmptyLine(paras As List(Of String))
        Dim currentpara As String = ""
        For i = 0 To _Lines.Length - 1
            Dim line As String = _Lines(i)
            If line.Length = 0 Or line.Replace(" ", "").Length = 0 Then
                paras.Add(currentpara)
                currentpara = ""
            Else
                currentpara = currentpara & " " & line
            End If
        Next
        paras.Add(currentpara)

    End Sub
    Private Sub ParseLinesShortLineEnds(paras As List(Of String))
        Dim currentpara As String = ""
        For i = 0 To _Lines.Length - 1
            Dim line As String = _Lines(i)
            If line.Length < 0.9 * _MeanLineLength Then
                currentpara = " " & line
                paras.Add(currentpara)
                currentpara = ""
            Else
                currentpara = currentpara & " " & line
            End If
        Next
        paras.Add(currentpara)
    End Sub

    Private Function AnalyzeType() As DocType
        If _RawText.Contains("</") Then
            Htmlfound = True
            Return DocType.HTML
        ElseIf _LineThenSpaceRatio > 0.7 Then
            Return DocType.AlternateLinesSpaces
        ElseIf MostlinesBeginSpaces() Then
            Return DocType.LinesStartSpaces
        ElseIf _EmptyLinesCount > 5 And _emptylinescount < _Lines.Length * 0.9 Then
            Return DocType.ParasDefinedbyBlankLine
        ElseIf _MeanLineLength > 10 * _MinLineLength Then
            Return DocType.ParasIndented
        Else
            Return DocType.LinesIndented
        End If
    End Function
    Private Sub CalcLineStats()
        Dim sum As Int32
        Dim min As Int16 = 50
        Dim max As Integer
        For i = 0 To _Lines.Length - 1
            Dim length As Int32 = _Lines(i).Length
            sum = sum + length
            If length < min Then
                min = length
            ElseIf length > max Then
                max = length
            End If
        Next
        CountEmptyLines()
        LineThenSpaceRatio()
        If _Lines.Length > 1 Then
            _MeanLineLength = sum / (_Lines.Length - 1)
        End If

        _MinLineLength = min
            _MaxLineLength = max
    End Sub
    Private Sub LineThenSpaceRatio()
        Dim LastBlank As Boolean = False
        Dim ThisBlank As Boolean = False

        Dim j = 0
        For i = 0 To _Lines.Length - 1
            Dim line As String = _Lines(i)
            ThisBlank = _Lines(i).Length = 0 OrElse _Lines(i).Replace(" ", "") = ""
            If LastBlank And Not ThisBlank Then
                j += 1
            End If
            LastBlank = ThisBlank
        Next
        _LineThenSpaceRatio = j / _Lines.Length - 1
    End Sub
    Private Function MostlinesBeginSpaces() As Boolean
        Dim j = 0
        For i = 0 To _Lines.Length - 1
            If _Lines(i).Length <> 0 AndAlso _Lines(i).Chars(0) = " " Then
                j += 1
            End If
        Next
        If j > (_Lines.Length - 1) / 2 Then
            Return True
        Else
            Return False
        End If
    End Function

End Class
