Public Class ListBoxHandler
#Region "Properties"
    Public Property SortOrder As New SortHandler
    Public Property Filter As New FilterHandler
    Public Property Random As New RandomHandler
    Private mItemList As New List(Of String)
    Public Property ItemList() As List(Of String)
        Get
            Return mItemList
        End Get
        Set(ByVal value As List(Of String))
            mItemList = value
        End Set
    End Property

    Private mListbox As New ListBox
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
    Public Sub FilterList()
        With Filter
            .FileList = ItemList
            'ItemList.Clear()
            ItemList = .FileList
        End With
    End Sub
    Public Sub OrderList()
        With SortOrder
            ItemList = SetPlayOrder(SortOrder.State, ItemList)
        End With

    End Sub
    Public Sub FillBox()
        ListBox.Items.Clear()
        FilterList()
        OrderList()

        For Each f In ItemList
            ListBox.Items.Add(f)
        Next
        'SetFirst()
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
    Public Sub SetNamed(Name As String)
        If ListBox.Items.Count > 0 Then
            Dim i = ListBox.FindString(Name)
            If i > -1 Then
                ListBox.SelectedIndex = ListBox.FindString(Name)
            Else
                ListBox.SelectedIndex = 0
            End If
        End If
    End Sub

    Public Sub New(Lbx As ListBox)
        ListBox = Lbx
    End Sub
#End Region

#Region "Events"

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
            FillBox()
        End Set
    End Property

    Public Sub GetFiles()
        ItemList.Clear()
        Dim dir As New IO.DirectoryInfo(DirectoryPath)
        For Each f In dir.GetFiles
            ItemList.Add(f.FullName)
        Next
    End Sub
    Public Sub New(Lbx As ListBox)
        MyBase.New(Lbx)
        ListBox = Lbx
    End Sub

End Class
