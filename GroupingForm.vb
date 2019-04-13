Imports Microsoft.VisualBasic
Public Class GroupingForm
    Private mFolderPath As String
    Private mFolder As IO.DirectoryInfo
    Private mFewest As Int16 = 3
    Private mMost As Int16 = 20
    Public Property FolderPath() As String
        Get
            Return mFolderPath
        End Get
        Set(ByVal value As String)
            mFolderPath = value
            mFolder = New IO.DirectoryInfo(mFolderPath)
            MakeList(mFolder.EnumerateFiles)
            '            Subgroups(x.EnumerateFiles, 13)
        End Set
    End Property

    Private Property Starts As New List(Of String)


    'List all the starts of length MAXLENGTH, and count the number of each
    'For each start which has a count smaller than MINCOUNT, make a new start of length one shorter and repeat
    'Go through the list of files, examining MAXLENGTH strings
    'Tally matches with the check string, 
    'When the check changes, add a row IFF the tally is between MINCOUNT and MAXCOUNT
    '
    'Reduce the length by 1 and repeat

    Private Sub MakeList(Slist As IEnumerable(Of IO.FileInfo))
        Dim check As String = ""
        Dim count As Integer = 1
        Dim total As Integer
        DataGridView1.Rows.Clear()
        Starts.Clear()

        For i = 40 To 1 Step -1
            For Each s In Slist
                If LCase(Strings.Left(s.Name, i)) <> LCase(check) Then
                    'Check changed
                    If count >= mFewest And count <= mMost Then
                        Dim val As String = Starts.Find(Function(value As String)
                                                            Return InStr(LCase(value), LCase(check)) <> 0
                                                        End Function)

                        If val <> "" Then
                            '                    MsgBox(LCase(check) & " already added")
                        Else
                            DataGridView1.Rows.Add(New String() {check, count})
                            total = total + count
                            Starts.Add(LCase(check))
                            '                   MsgBox(check & " being added")
                        End If
                    End If
                    check = Strings.Left(s.Name, i)
                    count = 1

                Else
                    count += 1
                End If
            Next
        Next
        DataGridView1.Rows.Add(New String() {"TOTAL", total})

    End Sub


    Private Sub GroupingForm_Load(sender As Object, e As EventArgs) Handles Me.Load
        FolderPath = CurrentFolder
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        For Each row In DataGridView1.Rows
            If row.Equals(DataGridView1.Rows(DataGridView1.Rows.Count - 1)) Then Exit For
            Dim m As String = row.Cells(0).Value.ToString
            For Each x In mFolder.EnumerateFiles
                If InStr(x.Name, m) = 1 Then
                    'Create folder m if doesn't exist
                    'move x to m
                    Dim subfolder As New IO.DirectoryInfo(mFolder.FullName & "\" & m & "\")
                    If subfolder.Exists Then
                    Else
                        subfolder.Create()
                    End If
                    x.MoveTo(subfolder.FullName & "\" & x.Name)
                End If
            Next
        Next
        MainForm.tvMain2.RefreshTree(Me.FolderPath)
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            If TextBox1.Text <> "" Then mFewest = Val(TextBox1.Text)
            MakeList(mFolder.EnumerateFiles)

        Catch ex As Exception

            'MakeList(mFolder.EnumerateFiles, mFewest, mMost)
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            If TextBox2.Text <> "" Then mMost = Val(TextBox2.Text)
            MakeList(mFolder.EnumerateFiles)

        Catch ex As Exception
            'MakeList(mFolder.EnumerateFiles, mFewest, mMost)
        End Try
    End Sub
End Class