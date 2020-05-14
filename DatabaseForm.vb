Public Class DatabaseForm
    '   Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    'Dim ds As New DataSet()
    'Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.15.0;Data Source=Q:\TreeSize Professional File Search Export - Movies.xlsx;" & "Extended Properties=Excel 12.0;"
    'Dim excelData As New OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString)
    'excelData.TableMappings.Add("Table", "ExcelSheet")
    'excelData.Fill(ds)
    'Me.DataGridView1.DataSource = ds.Tables(0)
    'Me.Refresh()
    Dim DontFind As Boolean = True
    Dim MyConnection As System.Data.OleDb.OleDbConnection
    Dim DtSet As System.Data.DataSet
    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
    ''' <summary>
    ''' Selects the given string in the datagridview
    ''' </summary>
    ''' <returns></returns>
    Public Property ShortFilepath As String
        Set(value As String)
            Dim i = Findrow(value)
            If i >= 0 Then
                dgv.Rows(i).Selected = True
                dgv.CurrentCell = dgv.Rows(i).Cells(0)
                '               dgv.FirstDisplayedScrollingRowIndex = i
            End If

        End Set
        Get

        End Get
    End Property

    Private Sub Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Shown

        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='Q:\TreeSize Professional File Search Export - Movies.xlsx';Extended Properties=Excel 8.0;")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet2$]", MyConnection)
        MyCommand.TableMappings.Add("Name", "Containing Path")
        DtSet = New System.Data.DataSet()
        MyCommand.Fill(DtSet)
        dgv.DataSource = DtSet.Tables(0)
        ' DataGridView1.DataMember = "[Sheet3$]"
        MyConnection.Close()
        SetDataformEntry()
        'HideRows(0, "jpg")
        Me.Text = Str(dgv.Rows.Count) & " rows"
        Me.WindowState = FormWindowState.Maximized
        'dgv.AutoResizeColumns()
        DontFind = False
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv.CellMouseClick
        ' If DontFind Then Exit Sub
        Dim j = e.RowIndex
        If j >= 0 Then
            Dim s = dgv.Rows(j).Cells(1).Value & dgv.Rows(j).Cells(0).Value
            MainForm.HighlightCurrent(s)
        End If
    End Sub

    Private Sub dgv_SelectionChanged(sender As Object, e As DataGridViewRowStateChangedEventArgs) Handles dgv.RowStateChanged
        If DontFind Then Exit Sub
        If dgv.CurrentRow IsNot Nothing Then

            Try

                '                Dim j = e.Row.Index
                Dim s = e.Row.Cells(1).Value & e.Row.Cells(0).Value
                MainForm.HighlightCurrent(s)

            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub dgv_Scroll(sender As Object, e As ScrollEventArgs) Handles dgv.Scroll
        Select Case e.Type
            Case ScrollEventType.SmallIncrement, ScrollEventType.SmallDecrement
                DontFind = False
            Case Else
                DontFind = True

        End Select
    End Sub

    Private Sub dgv_KeyDown(sender As Object, e As KeyEventArgs) Handles dgv.KeyDown
        Select Case e.KeyCode
            Case Keys.Down, Keys.Up, Keys.Left, Keys.Right, Keys.PageDown, Keys.PageUp
            Case Else
                MainForm.HandleKeys(sender, e)
        End Select
    End Sub
    Private Function HideRows(cell As Integer, search As String)
        Try


            For Each row In dgv.Rows
                If row.cells(cell).value.contains(search) Then
                Else
                    row.hide
                End If
            Next
        Catch ex As Exception
        End Try
    End Function
    Private Function Findrow(s As String) As Integer
        ' If DontFind Then Exit Function

        Dim foundrow As Integer = -2
        For i = 0 To dgv.Rows.Count - 2
            Try
                If dgv.Rows(i).Cells(0).Value = s Then
                    foundrow = i
                End If
            Catch ex As Exception
            End Try
        Next
        Return foundrow
    End Function
End Class
