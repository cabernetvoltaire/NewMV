Public Class ListBoxHandler
#Region "Properties"
    Public Property SortOrder As New SortHandler
    Public Property Reverse As Boolean = False
    Public Property Filter As New FilterHandler
    Public Property Random As New RandomHandler
    Public Property FolderAdvance As Boolean = True
    Public NewItem As String = ""
    Public OldItem As String = ""
    Private Nofile As String = "(Empty folder)"
    Private CheckFilter As String = "(If nothing is shown here, check filter)"
    Private mEditing As Boolean = False
    Private WithEvents mText As New TextBox
    Private mListHandler As New MediaListHandler

    Friend mItemList As New List(Of String)
    Friend mFullItemList As New List(Of String)
    Public Event ListBoxFilled(sender As Object, e As EventArgs)
    Public Event ListboxChanged(sender As Object, e As EventArgs)
    Public Event ListIndexChanged(sender As Object, e As EventArgs)
    Public Event ListItemChanged(sender As Object, e As EventArgs)
    Public Event EndReached(sender As Object, e As EventArgs)


    Public Property ItemList() As List(Of String)
        Get
            Return mItemList
        End Get
        Set(ByVal value As List(Of String))
            mListHandler.List = value
            mItemList = value
            mFullItemList = value
        End Set
    End Property
    Private Property mSelectedItems As New List(Of String)
    Public ReadOnly Property SelectedItemsList() As List(Of String)
        Get
            mSelectedItems.Clear()
            For Each m In mListbox.SelectedItems
                mSelectedItems.Add(m.ToString)
            Next
            Return mSelectedItems

        End Get
    End Property

    Public WriteOnly Property SelectItems() As List(Of String)
        Set(value As List(Of String))
            mListbox.SelectionMode = SelectionMode.MultiExtended
            mListbox.SuspendLayout()
            For Each m In value
                SetNamed(m)
            Next
            mListbox.ResumeLayout()
        End Set
    End Property

    Public ReadOnly Property Current As String
        Get
            Return mListHandler.Current
        End Get
    End Property

    Public ReadOnly Property NextOne As String
        Get
            Return mListHandler.NextOne
        End Get
    End Property

    Public ReadOnly Property Previous As String
        Get
            Return mListHandler.Previous
        End Get
    End Property

    Friend WithEvents mListbox As New ListBox
    Private _SingleLinks As Boolean

    Public Property ListBox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value
            mItemList = AllfromListbox(mListbox)
            mFullItemList = mItemList

        End Set
    End Property
    Private Property mCurrentIndex As Integer
    Public Property CurrentIndex() As Integer
        Get
            Return mListbox.SelectedIndex
        End Get
        Set(ByVal value As Integer)

            mCurrentIndex = value
            mListHandler.Index = value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Sub FilterList()
        If Filter.State <> Filter.OldState Or Filter.SingleLinks Then

            With Filter
                mFullItemList = mItemList
                .FileList = mFullItemList
                mItemList = .FileList

            End With
        End If
    End Sub
    Public Sub OrderList()

        mItemList = SetPlayOrder(SortOrder.State, ItemList, SortOrder.ReverseOrder)
    End Sub
    Public Sub HighlightList(ls As List(Of String))
        mListbox.DataSource = Nothing
        mListbox.SelectionMode = SelectionMode.MultiExtended
        For Each f In ls
            Dim i = mListbox.FindString(f)
            If i >= 0 Then mListbox.SetSelected(i, True)
        Next

    End Sub
    Public Sub AddList(list As List(Of String))
        For Each m In list
            mItemList.Add(m)
        Next
        FillBox()
    End Sub
    Public Sub IncrementIndex(Forward As Boolean)
        If mListbox.SelectionMode <> SelectionMode.One Then mListbox.SelectionMode = SelectionMode.One
        If ListBox.DataSource Is {Nofile} Or ListBox.DataSource Is {CheckFilter} Then
            RaiseEvent EndReached(ListBox, Nothing)
            Exit Sub
        End If
        If ListBox.Items.Count > 0 Then
            If Random.NextSelect Then
                SetNamed(MSFiles.NextItem)
            Else

                If Forward Then
                    ListBox.SelectedIndex = (ListBox.SelectedIndex + 1) Mod ListBox.Items.Count
                    If FolderAdvance And ListBox.SelectedIndex = 0 Then RaiseEvent EndReached(Me, Nothing)
                Else
                    If ListBox.SelectedIndex = 0 Then
                        If FolderAdvance Then RaiseEvent EndReached(Me, Nothing)
                        If ListBox.Items.Count >= 1 Then
                            ListBox.SelectedIndex = ListBox.Items.Count - 1
                        End If
                    Else
                        If ListBox.SelectedIndex > 0 Then ListBox.SelectedIndex = ListBox.SelectedIndex - 1

                    End If
                End If
            End If
        Else
            RaiseEvent EndReached(ListBox, Nothing)
        End If
    End Sub
    Public Property SingleLinks As Boolean
        Get
            Return _SingleLinks
        End Get
        Set
            _SingleLinks = Value
        End Set
    End Property

    Public Sub FillBox(Optional List As List(Of String) = Nothing)
        mListbox.DataSource = Nothing
        If List IsNot Nothing Then mItemList = List
        Filter.SingleLinks = _SingleLinks
        FilterList()
        OrderList()
        If mItemList.Count > 200 Then
            mListbox.SuspendLayout()
        End If
        mListbox.DataSource = mItemList



        If mItemList.Count = 0 Then
            If Filter.State <> FilterHandler.FilterState.All Then

                mListbox.DataSource = {CheckFilter}

            Else
                mListbox.DataSource = {Nofile}

            End If
        End If

        'SetFirst()
    End Sub
    Private Function InvertListBoxSelections(ByRef tempListBox As ListBox) As Integer
        Dim selectedind(tempListBox.SelectedItems.Count) As Integer
        Try
            For selind = 0 To tempListBox.SelectedItems.Count - 1
                selectedind.SetValue(tempListBox.Items.IndexOf(tempListBox.SelectedItems(selind)), selind)
            Next
            tempListBox.ClearSelected()
            For listitemIndex = 0 To tempListBox.Items.Count
                If Array.IndexOf(selectedind, listitemIndex) < 0 Then
                    tempListBox.SetSelected(listitemIndex, True)
                End If
            Next
            Return 1
        Catch ex As Exception
            Return 0
        End Try
    End Function
    Public Sub InvertSelected()
        '        MsgBox("Needs fixing")
        '       Exit Sub
        ListBox.SelectionMode = SelectionMode.MultiExtended
        InvertListBoxSelections(ListBox)


    End Sub
    Public Sub RemoveItems(List As List(Of String))
        If List.Count = 0 Then Exit Sub
        Dim i = ListBox.FindString(List(0))
        ListBox.DataSource = Nothing
        For Each m In List
            mItemList.Remove(m)
            'RaiseEvent ListIndexChanged(ListBox, Nothing)
        Next
        FillBox()

        If i > ListBox.Items.Count - 1 Then
            SetIndex(ListBox.Items.Count - 1)
        Else
            SetIndex(i)
        End If
    End Sub
    Public Sub SetFirst(Optional ForceFirst As Boolean = False)
        If ListBox.Items.Count > 0 Then
            If Random.OnDirChange And Not ForceFirst Then
                ListBox.SelectedIndex = Rnd() * (ListBox.Items.Count - 1)
            Else
                ListBox.SelectedIndex = 0
            End If
        End If
    End Sub
    Public Sub SetIndex(num As Integer, Optional force As Boolean = False)
        Try
            ListBox.SelectedIndex = num
            If force Then RaiseEvent ListIndexChanged(ListBox, Nothing)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Public Sub SetNamed(Name As String)
        If ListBox.SelectedItem = Name Then Exit Sub
        If ListBox.Items.Count > 0 Then
            Dim i = ListBox.FindString(Name)
            If i > -1 Then
                ListBox.SelectedIndex = i 'ListBox.FindString(Name)
            Else
                ListBox.SelectedIndex = 0
            End If
        End If
    End Sub
    'Public Sub MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles mListbox.MouseDown
    '    Dim pt As New Point(e.X, e.Y)
    '    Dim index As Integer = mListbox.IndexFromPoint(pt)


    '    mListbox.DoDragDrop(mListbox.Items(index).ToString, DragDropEffects.Link)
    'End Sub



#End Region

#Region "Events"
#Region "Editing Functions"
    Private Sub Keydown(sender As Object, e As KeyEventArgs) Handles mListbox.KeyDown
        'Exit Sub
        If mEditing Then
        Else

            Select Case e.KeyCode

                Case Keys.F2
                    EditItem()
                Case Else
                    'MainForm.Main_KeyDown(sender, e)
            End Select
        End If

        '        mListbox.SelectionMode = SelectionMode.One
        '    End If
    End Sub
    Private Sub FinishEdit(sender As Object, e As EventArgs) Handles mText.LostFocus
        FormMain.KeyPreview = True
        mText.Hide()
        mEditing = False
        If mText.Text <> mListbox.SelectedItem Then
            OldItem = mListbox.SelectedItem
            NewItem = mText.Text
            mListbox.Items(mListbox.SelectedIndex) = NewItem 'TODO Uncouple datasource
            RaiseEvent ListItemChanged(mListbox.SelectedItem, e)
        End If
    End Sub
    Private Sub EscapeField(sender As Object, e As KeyEventArgs) Handles mText.KeyDown
        If e.KeyCode = Keys.Return Or e.KeyCode = Keys.Tab Then
            FinishEdit(sender, e)
        End If

    End Sub
    Private Sub EditItem()
        FormMain.KeyPreview = False
        mEditing = True
        Dim rect As Rectangle = mListbox.GetItemRectangle(mListbox.SelectedIndex)
        mText.Parent = mListbox
        mText.Top = rect.Top
        mText.Height = rect.Height
        mText.Width = rect.Width
        mText.Text = mListbox.SelectedItem
        mText.Show()
        mText.BringToFront()
        mText.Focus()

    End Sub
#End Region
    Private Sub IndexChanged(sender As Object, e As EventArgs) Handles mListbox.SelectedIndexChanged
        RaiseEvent ListIndexChanged(Me, Nothing)
    End Sub

#End Region

#Region "List functions"


#End Region


End Class
Public Class FileboxHandler
    Inherits ListBoxHandler
    Public Event ListBoxIncreased(sender As Object, e As EventArgs)

    Private _DirectoryPath As String

    Property DirectoryPath As String
        Get
            Return _DirectoryPath
        End Get
        Set

            If Value <> "" Then
                _DirectoryPath = Value
                GetFiles()
                mFullItemList = mItemList
                FillBox(mItemList)
                FormMain.folderwatcher.Path = Value
            End If

        End Set
    End Property

    Private Sub GetFiles()
        mItemList.Clear()
        Dim dir As New IO.DirectoryInfo(_DirectoryPath)
        If dir.Exists Then
            '            mListbox.DataSource = dir.GetFiles
            For Each f In dir.GetFiles
                mItemList.Add(f.FullName)
                RaiseEvent ListBoxIncreased(Me, Nothing)
            Next
        End If
    End Sub
    Public Sub Refresh()
        Dim ind As Integer = mListbox.SelectedIndex
        GetFiles()
        FillBox(mItemList)
        SetIndex(ind)
    End Sub


End Class
