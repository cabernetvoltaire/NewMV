Public Class ShowListForm
    Private mItemList As List(Of String)
    Public Property ItemList() As List(Of String)
        Get
            Return mItemList
        End Get
        Set(ByVal value As List(Of String))
            mItemList = value
            For Each m In mItemList
                ListBox1.Items.Add(m)
            Next
        End Set
    End Property
    Private Sub IndexHandler(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        With sender
            Dim lbx As ListBox = CType(sender, ListBox)
            If lbx.SelectionMode = SelectionMode.One Then
                Dim i As Long = .SelectedIndex
                If i = -1 Then
                Else
                    Debug.Print(vbCrLf & vbCrLf & "NEXT SELECTION ---------------------------------------")
                    MSFiles.Listbox = sender
                    MSFiles.ListIndex = i
                    ' HighlightCurrent(lbxFiles.SelectedItem)
                    ' MainForm.HighlightCurrent(ListBox1.SelectedItem)
                End If
            End If
        End With


    End Sub

    Public Sub ShowListForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        MainForm.frmMain_KeyDown(sender, e)
        e.SuppressKeyPress = True
        Me.Show()
    End Sub
End Class