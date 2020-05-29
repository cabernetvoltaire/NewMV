Imports System.Object
Public Class DuplicatesForm
    Private mDB As Database
    Public Event HighlightFile(sender As Object, e As EventArgs)
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
    Private Property mPanelDuplicates As New List(Of DuplicateSet)
    Private Sub OnHighlightFile(sender As Object, e As EventArgs)
        RaiseEvent HighlightFile(sender, e)

    End Sub
    Public Sub FindDuplicates()
        mDB.Sort()
        Dim CurrentDuplicateSet As New DuplicateSet
        For Each entry In mDB.Entries

            If CurrentDuplicateSet.DSet.Count = 0 Then
                CurrentDuplicateSet.Size = entry.Size
            End If
            If entry.Size = CurrentDuplicateSet.Size Then

                CurrentDuplicateSet.Add(entry)
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

        End If
    End Sub
    Public Sub ShowDuplicatesO(StartIndex As Integer)

        For j = StartIndex To AllDuplicates.Count - 1
            Dim panel As New Panel With {.Dock = DockStyle.Bottom}
            Dim tt As New ToolTip
            tt.SetToolTip(panel, "Duplicate set " & Str(j))
            Dim m As DuplicateSet = AllDuplicates.Item(j)
            mPanelDuplicates.Add(m)
            For i = 0 To m.DSet.Count - 1
                Dim path As String = m.DSet(i).FullPath
                Dim gap As Integer = 10
                Dim pic As New PictureBox With {.Height = 100, .Width = 100,
                                 .SizeMode = PictureBoxSizeMode.StretchImage}
                Try
                    'pic.Image = m.DSet(i).GetThumbnail
                    ThumbLoader(m, i, path, pic)
                    pic.Left = (pic.Width + gap) * i + gap
                    panel.Controls.Add(pic)

                Catch ex As Exception
                End Try
            Next
            If CheckFileDuplicates(panel) Then

                TableLayoutPanel3.Controls.Add(panel)
            Else
                panel.Dispose()

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
    Public Sub ShowDuplicates(StartIndex As Integer)

        For j = StartIndex To AllDuplicates.Count - 1
            Dim x As New DuplicatePanel
            AddHandler x.Highlightfile, AddressOf OnHighlightfile
            x.Duplicates = AllDuplicates(j)
            If CheckFileDuplicates(x) Then

                TableLayoutPanel3.Controls.Add(x.Panel)
            Else
                x.Panel.Dispose()

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
    '   Private OnHighlight(sender As Object, e As EventArgs) Handles 

    'Private Function CheckThumbnailDuplicates(flp As FlowLayoutPanel) As Boolean
    '    Dim SomeThumbsSame As Boolean = False
    '    For i = 0 To flp.Controls.Count - 2

    '        Dim p As PictureBox = DirectCast(flp.Controls(i), PictureBox)
    '        For j = i + 1 To flp.Controls.Count - 1
    '            Dim q As PictureBox = DirectCast(flp.Controls(j), PictureBox)
    '            If AreSameImage(p.Image, q.Image, True) Then
    '                SomeThumbsSame = True
    '            End If
    '        Next

    '    Next
    '    Return SomeThumbsSame
    'End Function
    Private Function CheckFileDuplicates(flp As Panel) As Boolean
        Dim SomeFilesSame As Boolean = False
        For i = 0 To flp.Controls.Count - 2
            Dim path As String = flp.Controls(i).Tag.Fullpath
            Dim p As New IO.FileInfo(path)
            If p.Exists Then

                For j = i + 1 To flp.Controls.Count - 1
                    Dim qath As String = flp.Controls(j).Tag.Fullpath
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

    Private Sub ThumbLoader(m As DuplicateSet, i As Integer, path As String, pic As PictureBox)
        With pic
            If FindType(path) = Filetype.Pic Or FindType(path) = Filetype.Gif Then
                .Image = Image.FromFile(path)
                Dim x As New Bitmap(.Image, 100 * .Image.Width / .Image.Height, 100)
                .Image.Dispose()
                .Image = x
                .Width = x.Width
            Else
                .BackColor = Color.HotPink
            End If
            .Tag = m.DSet(i) 'Assign the database entry to the picbox
            AddHandler .MouseEnter, AddressOf _Mouseover
            AddHandler .MouseClick, AddressOf _Mouseclick

        End With
    End Sub

    Private Sub _Mouseclick(sender As Object, e As MouseEventArgs)
        Dim pic As New PictureBox
        Dim outliner As New PictureBox With {.BackColor = Color.HotPink}
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

        RaiseEvent HighlightFile(pb, Nothing)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ShowDuplicates(mStartIndex)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Groupsize = Val(ComboBox1.SelectedItem)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



    End Sub
End Class
Public Class DuplicatePanel
    Private mDuplicates As DuplicateSet

    Public Property Panel As New Panel With {.Dock = DockStyle.Bottom, .Height = 120}
    Private Outliner As New PictureBox With {.BackColor = Color.HotPink}
    Public Property Tooltip As New ToolTip

    Public Event Highlightfile(sender As Object, e As EventArgs)
    Public Property Duplicates() As DuplicateSet
        Get
            Return mDuplicates
        End Get
        Set(ByVal value As DuplicateSet)
            mDuplicates = value
            PopulatePanel()
            IdentifyChoice()
        End Set
    End Property
    Private Sub IdentifyChoice()
        Dim kp As New KeeperChooser
        kp.DuplicateSet = mDuplicates
        For Each m In mDuplicates.DSet
            If m Is kp.Keeper Then
                OutlineControl(Panel.Controls(mDuplicates.DSet.IndexOf(m)), Outliner)
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
                'pic.Image = m.DSet(i).GetThumbnail
                ThumbLoader(mDuplicates.DSet(i), path, pic)
                pic.Left = (pic.Width + gap) * i + gap
                Panel.Controls.Add(pic)

            Catch ex As Exception
            End Try
        Next
    End Sub
    Private Sub _Mouseclick(sender As Object, e As MouseEventArgs)
        Dim pic As New PictureBox
        Dim outliner As New PictureBox With {.BackColor = Color.HotPink}
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

        RaiseEvent Highlightfile(pb, Nothing)
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
                .BackColor = Color.HotPink
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
            'They're on different branches, so choose longest filename again.
            'Otherwise choose longest filename
            If Len(x1.Filename) > Len(x2.Filename) Then
                Return x1
            Else
                Return x2
            End If
        End If

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