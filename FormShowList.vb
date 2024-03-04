Public Class FormShowList
    Dim LBH As New ListBoxHandler()

    Public Sub ShowListForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        FormMain.Main_KeyDown(sender, e)
        e.SuppressKeyPress = True
        'Me.Show()
    End Sub
    Private mListofFiles As List(Of String)
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        LBH.ListBox = ListBox1
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub ListBox1_DragDrop(sender As Object, e As DragEventArgs) Handles ListBox1.DragDrop
        ListBox1.Text = e.Data.GetData(DataFormats.Text).ToString()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        FormMain.HighlightCurrent(ListBox1.SelectedItem)
    End Sub

    Public Property ListofFiles() As List(Of String)
        Get
            Return mListofFiles
        End Get
        Set(ByVal value As List(Of String))
            mListofFiles = value
            LBH.FillBox(mListofFiles)
        End Set
    End Property



    Private Sub FormShowList_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus
        LBH.ListBox.BackColor = CurrentfilterState.Colour
    End Sub

    Private Sub LoadListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadListToolStripMenuItem.Click

    End Sub
End Class