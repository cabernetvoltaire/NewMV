Imports AxWMPLib

Public Class MovieSwapTest
    Private WithEvents mSwapper As New MovieSwapper(Test1, Test2)

    Private Sub MovieSwapTest_Load(sender As Object, e As EventArgs) Handles Me.Load
        mSwapper.Listbox = MainForm.lbxFiles
        For Each f In MainForm.lbxFiles.Items
            ListBox1.Items.Add(f)
        Next

    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        mSwapper.ListIndex = ListBox1.SelectedIndex
        '    If ListBox1.SelectedIndex < Index Then
        '        mNext.Forwards = False
        '    Else
        '        mNext.Forwards = True
        '    End If
        '    mNext.CurrentIndex = ListBox1.SelectedIndex
        '    Current = ListBox1.SelectedItem
        '    Nxt = mNext.NextItem
        '    If Test1.URL = "" Then
        '        Test1.URL = Current
        '        Test1.BringToFront()

        '    ElseIf Test1.URL = Current Then
        '        Test2.URL = Nxt
        '        mMedia2.MediaPath = Nxt
        '        Mysettings.Media = mMedia1

        '        Test1.Visible = True
        '        Test2.Visible = False
        '        Test1.BringToFront()

        '    Else
        '        Test1.URL = Nxt
        '        mMedia1.MediaPath = Nxt
        '        Mysettings.Media = mMedia1

        '        Test2.BringToFront()
        '        Test2.Visible = True
        '        Test1.Visible = False

        '    End If
        '    Index = ListBox1.SelectedIndex
    End Sub


    Private Sub Test2_Enter(sender As Object, e As EventArgs) Handles Test2.Enter

    End Sub
End Class
