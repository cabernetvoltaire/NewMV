Imports System.IO

Public Class FormDuplicates
    Private mDB As Database
    Public DupsThread As Threading.Thread()
    Public Event FileHighlighted(sender As Object, e As EventArgs)
    '    Public WithEvents DupsPanels As List(Of DuplicatePanel)
    Public Shared Groupsize As Integer = 50
    Public Property DB() As Database
        Get
            Return mDB
        End Get
        Set(ByVal value As Database)
            mDB = value
        End Set
    End Property
    Private mStartIndex As Integer = 0
    Public Property AllDuplicates As New List(Of DuplicateSet)
    Private Property mDuplicatePanels As New List(Of DuplicatePanel)

    Private Sub OnHighlightFile(sender As Object, e As EventArgs)
        FormMain.HighlightCurrent(sender.Tag.Fullpath)
        'RaiseEvent FileHighlighted(sender, e)

    End Sub
    Public Sub FindDuplicates()
        mDB.Sort(New CompareDBByFilesize)
        Dim CurrentDuplicateSet As New DuplicateSet
        For Each entry In mDB.Entries

            If CurrentDuplicateSet.DSet.Count = 0 Then
                CurrentDuplicateSet.Size = entry.Size
            End If
            If entry.Size = CurrentDuplicateSet.Size Then


                CurrentDuplicateSet.Add(entry)
                PopulateFolderList(CurrentDuplicateSet)
            Else
                If CurrentDuplicateSet.DSet.Count > 1 Then
                    AllDuplicates.Add(CurrentDuplicateSet)
                End If
                CurrentDuplicateSet = New DuplicateSet
                CurrentDuplicateSet.Add(entry)
            End If
        Next
        If AllDuplicates.Count = 0 Then
            MsgBox("No duplicates found")
        Else
            Me.Text = Me.Text & Format(" {0} duplicates found", AllDuplicates.Count)
        End If
    End Sub

    Private Sub PopulateFolderList(CurrentDuplicateSet As DuplicateSet)
        For Each e In CurrentDuplicateSet.DSet
            If ListBox1.Items.Contains(e.Path) Then
            Else
                ListBox1.Items.Add(e.Path)
            End If
            If ListBox2.Items.Contains(e.Path) Then
            Else
                ListBox2.Items.Add(e.Path)
            End If
        Next
    End Sub

    Public Sub ShowDuplicates(Startindex As Integer)
        DuplicatesTLP.Controls.Clear()

        For j = Startindex To AllDuplicates.Count - 1

            Dim DupsPanel As New DuplicatePanel
            AddHandler DupsPanel.Highlightfile, AddressOf OnHighlightFile
            DupsPanel.Duplicates = AllDuplicates(j)
            If CheckFileDuplicates(DupsPanel) Then
                If DupsPanel.Panel.Controls.Count > 0 Then
                    DuplicatesTLP.Controls.Add(DupsPanel.Panel)
                    DupsPanel.Panel.Controls.Add(New Label With {.Text = "Duplicate No. " & Str(j)})
                    mDuplicatePanels.Add(DupsPanel)
                End If
            Else
                DuplicatesTLP.Controls.Remove(DupsPanel.Panel)
            End If
            If j / Groupsize = Int(j / Groupsize) And j > 0 Then
                'If MsgBox(j & " duplicates found so far. Continue?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                'Else
                mStartIndex = j + 1
                Exit For
                'End If
            End If
        Next

    End Sub
    Public Sub RemoveAllDuplicates()
        Dim entry As DatabaseEntry
        Dim fop As New FileOperations
        For Each d In AllDuplicates
            Dim k As New KeeperChooser
            k.DuplicateSet = d
            entry = k.Keeper
            If entry IsNot Nothing Then
                For Each e In d.DSet
                    If e IsNot entry Then
                        fop.Files.Add(e.FullPath)
                        Application.DoEvents()
                    End If
                Next
            End If
        Next
        fop.DeleteFiles()
    End Sub

    Public Sub RemoveCurrentDuplicates()

        Dim fop As New FileOperations
        AddHandler fop.Filesmoving, AddressOf FormMain.OnFilesMoving
        For Each m In mDuplicatePanels
            For Each f In m.Deletees
                'MsgBox(f.FullPath & " would be deleted.")
                fop.Files.Add(f.FullPath)
            Next
        Next
        fop.DeleteFiles()
    End Sub

    Private Function CheckFileDuplicates(flp As DuplicatePanel) As Boolean
        Dim SomeFilesSame As Boolean = False
        For i = 0 To flp.Duplicates.DSet.Count - 2
            Dim path As String = flp.Duplicates.DSet(i).FullPath
            Dim p As New IO.FileInfo(path)
            If p.Exists Then

                For j = i + 1 To flp.Duplicates.DSet.Count - 1
                    Dim qath As String = flp.Duplicates.DSet(j).FullPath
                    Dim q As New IO.FileInfo(qath)
                    If q.Exists Then
                        If AreSameFile(path, qath) Then
                            SomeFilesSame = True
                        End If
                    End If
                Next
            End If
        Next
        Return SomeFilesSame

    End Function


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ShowDuplicates(ListBox2.SelectedItem)
        Button1.Text = "Next group"
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0, 1, 2, 3
                Groupsize = Val(ComboBox1.SelectedItem)
            Case Else
                Groupsize = AllDuplicates.Count - 1
        End Select


    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ForbiddenPaths.Clear()

        For Each item In ListBox1.SelectedItems
            ForbiddenPaths.Add(item)
        Next

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        RemoveCurrentDuplicates()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        RemoveAllDuplicates()
    End Sub
    Public Function ReadBinary(path As String, size As Long, start As Long) As Byte()
        Dim bytesRead As Long
        Dim bytes(size) As Byte
        Using strm As New FileStream(path, FileMode.Open, FileAccess.Read)
            strm.Position = start
            Using rdr As New BinaryReader(strm)

                bytesRead = rdr.Read(bytes, 0, size)


            End Using
        End Using
        Return bytes
        Exit Function
    End Function

    Public Function AreSameFile(path As String, qath As String) As Boolean
        Dim same As Boolean = False
        Dim pbytes, qbytes As Byte()
        pbytes = ReadBinary(path, 1000, 5000)
        qbytes = ReadBinary(qath, 1000, 5000)
        Dim k As Integer
        While k < 1000

            If pbytes(k) = qbytes(k) Then
                same = True
            Else
                same = False
                Exit While
            End If
            k += 1
        End While

        Return same
    End Function
End Class
Public Class DuplicatePanel
    Private mDuplicates As DuplicateSet

    Public Property Panel As New Panel With {.BorderStyle = BorderStyle.Fixed3D, .AutoSize = True, .Height = 120, .BackColor = Color.AliceBlue, .Dock = DockStyle.Left}
    Private Outliner As New PictureBox With {.BackColor = Color.HotPink, .Visible = False}
    Private Exclude As Boolean = False
    Public Deletees As New List(Of DatabaseEntry)
    Public Property Tooltip As New ToolTip


    Public Event Highlightfile(sender As Object, e As EventArgs)
    Public Property Duplicates() As DuplicateSet
        Get
            Return mDuplicates
        End Get
        Set(ByVal value As DuplicateSet)
            mDuplicates = value
            AddHandler Panel.MouseClick, AddressOf PanelClicked
            PopulatePanel()
            IdentifyChoice()
        End Set
    End Property

    Private Sub PanelClicked(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Exclude = True
            Panel.BackColor = Color.Chocolate

        End If
    End Sub

    Private Sub IdentifyChoice()
        Dim kp As New KeeperChooser
        kp.DuplicateSet = mDuplicates
        For Each m In mDuplicates.DSet
            If m Is kp.Keeper Then
                m.Mark = True
                If mDuplicates.DSet.Count <= Panel.Controls.Count Then
                    OutlineControl(Panel.Controls(mDuplicates.DSet.IndexOf(m)), Outliner)

                End If
            Else
                Deletees.Add(m)
                m.Mark = False
            End If
        Next
    End Sub


    Private Sub PopulatePanel()
        For i = 0 To mDuplicates.DSet.Count - 1
            Dim path As String = mDuplicates.DSet(i).FullPath
            Dim gap As Integer = 10
            Dim pic As New PictureBox With {.Height = 100, .Width = 100,
                                 .SizeMode = PictureBoxSizeMode.StretchImage}
            Try
                ThumbLoader(mDuplicates.DSet(i), path, pic)
                pic.Top = gap
                pic.Left = (pic.Width + gap) * i + gap
                Panel.Controls.Add(pic)

            Catch ex As Exception
            End Try
        Next

    End Sub
    Private Sub _Mouseclick(sender As Object, e As MouseEventArgs)
        Dim pic As New PictureBox
        ' Dim outliner As New PictureBox With {.BackColor = Color.HotPink}
        pic = DirectCast(sender, PictureBox)
        OutlineControl(pic, outliner)
        'For Each p In pic.Parent.Controls
        '    p = DirectCast(sender, PictureBox)
        '    pic.BorderStyle = BorderStyle.None
        'Next
        'pic.BorderStyle = BorderStyle.Fixed3D
    End Sub

    Private Sub _Mouseover(sender As Object, e As EventArgs)
        Dim tt As New ToolTip
        Dim pb = DirectCast(sender, PictureBox)
        Dim ds = DirectCast(pb.Tag, DatabaseEntry)
        Dim size As String = Format(ds.Size, "###,###,###")
        tt.SetToolTip(pb, ds.Filename & "(" & size & ")" & vbCrLf & ds.Path)

        RaiseEvent Highlightfile(sender, e)
    End Sub
    Private Sub ThumbLoader(m As DatabaseEntry, path As String, pic As PictureBox)
        With pic
            If FindType(path) = Filetype.Pic Or FindType(path) = Filetype.Gif Then
                .Image = Image.FromFile(path)
                Dim x As New Bitmap(.Image, 100 * .Image.Width / .Image.Height, 100)
                .Image.Dispose()
                .Image = x
                .Width = x.Width
            Else
                .BackColor = Color.DarkCyan
            End If
            .Tag = m 'Assign the database entry to the picbox
            AddHandler .MouseEnter, AddressOf _Mouseover
            AddHandler .MouseClick, AddressOf _Mouseclick

        End With
    End Sub


End Class
Public Class DuplicateSet
    Public Property DSet As New List(Of DatabaseEntry)
    Public Property Size As Long

    Public Sub Add(Entry As DatabaseEntry)
        DSet.Add(Entry)
        If DSet.Count = 1 Then Size = Entry.Size
    End Sub



End Class

Public Class KeeperChooser
    Public Property DuplicateSet As New DuplicateSet
    Private mUniqueFolders As List(Of String)
    Public Property Preserved As DatabaseEntry

    Public ReadOnly Property UniqueFolders() As List(Of String)
        'Find all the different locations of the duplicates
        Get
            For Each x In DuplicateSet.DSet
                If mUniqueFolders.Contains(x.Path) Then
                Else
                    mUniqueFolders.Add(x.Path)
                End If
            Next
            Return mUniqueFolders
        End Get
    End Property
    Public Function Keeper() As DatabaseEntry
        Dim mChosen As New DatabaseEntry
        'All pairs need to be compared
        For i = 0 To DuplicateSet.DSet.Count - 2
            For j = i + 1 To DuplicateSet.DSet.Count - 1
                Dim x1 = DuplicateSet.DSet(i)
                Dim x2 = DuplicateSet.DSet(j)
                mChosen = mChooseKeeper(x1, x2)
            Next
        Next
        Return mChosen
    End Function

    Private Function mChooseKeeper(x1 As DatabaseEntry, x2 As DatabaseEntry) As DatabaseEntry
        If SamePath(x1, x2) Then
            'Are they in the same folder?
            If SamishNames(x1, x2) Then
                'Names similar (probably bracketed)
                'Choose on alphabet
                If x1.Filename < x2.Filename Then
                    Return x1
                Else
                    Return x2
                End If
            Else
                'Otherwise choose longest filename
                If Len(x1.Filename) > Len(x2.Filename) Then
                    Return x1
                Else
                    Return x2
                End If
            End If
        ElseIf SameTree(x1, x2) Then
            'They're on the same branch 
            'Choose deepest
            If Len(x1.Path) > Len(x2.Path) Then
                Return x1
            Else
                Return x2
            End If
        Else
            If ForbiddenPaths.Contains(x1.Path) Then
                Return x2
            ElseIf ForbiddenPaths.Contains(x2.Path) Then
                Return x1
            Else
                'They're on different branches, so choose path with greatest number of files.
                Dim f As New IO.DirectoryInfo(x1.Path)
                Dim g As New IO.DirectoryInfo(x2.Path)
                If f.EnumerateFiles.Count > g.EnumerateFiles.Count Then
                    Return x1
                Else
                    Return x2
                End If
            End If
        End If
        'Anyway if one of them is forbidden then switch


    End Function

    Private Function SameTree(p As DatabaseEntry, q As DatabaseEntry) As Boolean
        If p.Path.Contains(q.Path) Or q.Path.Contains(p.Path) Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function SamishNames(p As DatabaseEntry, q As DatabaseEntry) As Boolean
        If p.Filename.Contains(q.Filename) Or q.Filename.Contains(p.Filename) <> 0 Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function SamePath(p As DatabaseEntry, q As DatabaseEntry) As Boolean
        If p.Path = q.Path Then
            Return True
        Else
            Return False
        End If
    End Function


End Class