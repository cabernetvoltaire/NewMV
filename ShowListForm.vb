Public Class ShowListForm
    Dim LBH As New ListBoxHandler(ListBox1)

    Public Sub ShowListForm_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        MainForm.Main_KeyDown(sender, e)
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
    Public Property ListofFiles() As List(Of String)
        Get
            Return mListofFiles
        End Get
        Set(ByVal value As List(Of String))
            mListofFiles = value
            LBH.FillBox(mListofFiles)
        End Set
    End Property
End Class