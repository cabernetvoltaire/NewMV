Public Class ListBoxHandler
#Region "Properties"
    Private mItemList As List(Of String)
    Public Property ItemList() As List(Of String)
        Get
            Return mItemList
        End Get
        Set(ByVal value As List(Of String))
            mItemList = value
        End Set
    End Property

    Private mListbox As ListBox
    Public Property ListBox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
        End Set
    End Property
#End Region
#Region "Methods"

#End Region

#Region "Events"

#End Region

#Region "List functions"

    Public Function Duplicatelist(ByVal inList As List(Of String)) As List(Of String)
        Dim out As New List(Of String)
        For Each i In inList
            out.Add(i)
        Next
        Return out
    End Function
    Public Function ListFromLinks(list As List(Of String)) As List(Of String)
        Dim s As New List(Of String)
        For Each l In list
            Dim tgt As String = LinkTarget(l)
            If s.Contains(tgt) Then
            Else
                s.Add(LinkTarget(tgt))
            End If

        Next
        FillShowbox(MainForm.lbxShowList, 0, s)
        Return s
    End Function

    Public Sub FillShowbox(lbx As ListBox, Filter As Byte, ByVal lst As List(Of String))


        If lst.Count = 0 Then Exit Sub
        If lst.Count > 1000 Then
            ProgressBarOn(lst.Count)
        End If


        lbx.Items.Clear()
        MainForm.CurrentFilterState.FileList = lst
        lst = MainForm.CurrentFilterState.FileList
        For Each s In lst
            lbx.Items.Add(s)

            'ProgressIncrement(1)
        Next

        If lbx.Name = "lbxShowList" Then
            MainForm.CollapseShowlist(False)
            Dim m As New ShowListForm
            m.Show()

            m.ItemList = lst

        End If
        '  lbx.Refresh()
        '        lbx.TabStop = True
        ProgressBarOff()



        'MainForm.UpdateFileInfo()
    End Sub
    Private Sub CopyList(list As List(Of String), list2 As SortedList(Of Date, String))
        list.Clear()
        For Each m As KeyValuePair(Of Date, String) In list2
            list.Add(m.Value)
        Next
    End Sub
#End Region


End Class
