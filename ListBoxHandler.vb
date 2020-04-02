Public Class ListBoxHandler
#Region "Properties"
    Public Property SortOrder As New SortHandler
    Public Property Filter As New FilterHandler
    Public Property Random As New RandomHandler
    Public Property FolderAdvance As Boolean = True
    Public NewItem As String = ""
    Public OldItem As String = ""

    Private mEditing As Boolean = False
    Private WithEvents mText As New TextBox

    Private mItemList As New List(Of String)
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
            mItemList = value
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
            For Each m In value
                SetNamed(m)
            Next
        End Set
    End Property


    Private WithEvents mListbox As New ListBox

    Public Property ListBox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value


        End Set
    End Property
    Private Property mCurrentIndex As Integer
    Public Property CurrentIndex() As Integer
        Get
            Return mListbox.SelectedIndex
        End Get
        Set(ByVal value As Integer)
            mCurrentIndex = value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Sub FilterList()
        If Filter.State <> Filter.OldState Then

            With Filter
                .FileList = mItemList
                mItemList = .FileList
            End With
        End If
    End Sub
    Public Sub OrderList()
        mItemList = SetPlayOrder(SortOrder.State, ItemList)

    End Sub
    Public Sub AddList(list As List(Of String))
        For Each m In list
            mItemList.Add(m)
        Next
        FillBox()
    End Sub
    Public Sub IncrementIndex(Forward As Boolean)
        If mListbox.SelectionMode <> SelectionMode.One Then mListbox.SelectionMode = SelectionMode.One
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
    Public Sub FillBox(Optional List As List(Of String) = Nothing)
        'mListbox.DataSource = Nothing
        If List IsNot Nothing Then
            mItemList = List
        Else

        End If
        FilterList()
        OrderList()

        'If mItemList.Count > 200 Then
        '    mListbox.SuspendLayout()
        'End If
        mListbox.DataSource = ItemList
        If ItemList.Count = 0 And Filter.State <> FilterHandler.FilterState.All Then
            mListbox.DataSource = {"(If nothing is shown here, check filter)"}
        End If
        RaiseEvent ListBoxFilled(mListbox, Nothing)
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
        FillBox()
        RaiseEvent ListboxChanged(ListBox, Nothing)
        If i > ListBox.Items.Count - 1 Then
            SetIndex(ListBox.Items.Count - 1)
        Else
            SetIndex(i)
        End If
    End Sub
    Public Sub SetFirst()
        If ListBox.Items.Count > 0 Then
            If Random.OnDirChange Then
                ListBox.SelectedIndex = Rnd() * (ListBox.Items.Count - 1)
            Else
                ListBox.SelectedIndex = 0
            End If
        End If
    End Sub
    Public Sub SetIndex(num As Integer)
        ListBox.SelectedIndex = num
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



    Public Sub New(Lbx As ListBox)
        mListbox = Lbx
    End Sub
#End Region

#Region "Events"
    Private Sub Keydown(sender As Object, e As KeyEventArgs) Handles mListbox.KeyDown
        'Exit Sub
        If mEditing Then
        Else

            Select Case e.KeyCode

                Case Keys.F2
                    EditItem()
                Case Else
                    MainForm.Main_KeyDown(sender, e)
            End Select
        End If

        '        mListbox.SelectionMode = SelectionMode.One
        '    End If
    End Sub
    Private Sub FinishEdit(sender As Object, e As EventArgs) Handles mText.LostFocus
        MainForm.KeyPreview = True
        mText.Hide()
        mEditing = False
        If mText.Text <> mListbox.SelectedItem Then
            OldItem = mListbox.SelectedItem
            Dim x = mListbox.SelectedIndex
            mListbox.DataSource = Nothing
            NewItem = mText.Text
            ItemList(x) = NewItem
            FillBox()
            RaiseEvent ListItemChanged(mListbox.SelectedItem, e)
        End If
    End Sub
    Private Sub EscapeField(sender As Object, e As KeyEventArgs) Handles mText.KeyDown
        If e.KeyCode = Keys.Return Or e.KeyCode = Keys.Tab Then
            FinishEdit(sender, e)
        End If

    End Sub
    Private Sub EditItem()
        MainForm.KeyPreview = False
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
    Private Sub IndexChanged(sender As Object, e As EventArgs) Handles mListbox.SelectedIndexChanged
        RaiseEvent ListIndexChanged(sender, e)
    End Sub

#End Region

#Region "List functions"


#End Region


End Class
Public Class FileboxHandler
    Inherits ListBoxHandler

    Private _DirectoryPath As String

    Property DirectoryPath As String
        Get
            Return _DirectoryPath
        End Get
        Set
            _DirectoryPath = Value
            GetFiles()
            FillBox(ItemList)
        End Set
    End Property

    Private Sub GetFiles()
        ItemList.Clear()
        Dim dir As New IO.DirectoryInfo(DirectoryPath)
        If dir.Exists Then
            For Each f In dir.GetFiles
                ItemList.Add(f.FullName)
            Next
        End If
    End Sub
    Public Sub Refresh()
        GetFiles()
        FillBox(ItemList)
    End Sub
    Public Sub New(Lbx As ListBox)
        MyBase.New(Lbx)
        ListBox = Lbx
    End Sub


End Class
