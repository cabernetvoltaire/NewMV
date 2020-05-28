Imports System.Object
Public Class DuplicatesForm
    Private mDB As Database
    Public Event HighlightFile(sender As Object, e As EventArgs)
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
    Public Sub ShowDuplicates(StartIndex As Integer)

        For j = StartIndex To AllDuplicates.Count - 1
            Dim flp As New FlowLayoutPanel With {.Dock = DockStyle.Top}
            Dim m As DuplicateSet = AllDuplicates.Item(j)
            For i = 0 To m.DSet.Count - 1
                Dim path As String = m.DSet(i).FullPath
                Dim pic As New PictureBox With {.Height = 100, .Width = 100,
                                 .SizeMode = PictureBoxSizeMode.StretchImage}
                Try
                    'pic.Image = m.DSet(i).GetThumbnail
                    ThumbLoader(m, i, path, pic)
                    flp.Controls.Add(pic)

                Catch ex As Exception
                End Try
            Next
            If CheckFileDuplicates(flp) Then
                Me.Controls.Add(flp)

            End If
            If j / 200 = Int(j / 200) And j > 0 Then
                If MsgBox(j & " duplicates found so far. Continue?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                Else
                    mStartIndex = j + 1
                    Exit For
                End If
            End If
        Next

    End Sub

    Private Function CheckThumbnailDuplicates(flp As FlowLayoutPanel) As Boolean
        Dim SomeThumbsSame As Boolean = False
        For i = 0 To flp.Controls.Count - 2

            Dim p As PictureBox = DirectCast(flp.Controls(i), PictureBox)
            For j = i + 1 To flp.Controls.Count - 1
                Dim q As PictureBox = DirectCast(flp.Controls(j), PictureBox)
                If AreSameImage(p.Image, q.Image, True) Then
                    SomeThumbsSame = True
                End If
            Next

        Next
        Return SomeThumbsSame
    End Function
    Private Function CheckFileDuplicates(flp As FlowLayoutPanel) As Boolean
        Dim SomeFilesSame As Boolean = False
        For i = 0 To flp.Controls.Count - 2
            Dim path As String = flp.Controls(i).Tag.Fullpath
            For j = i + 1 To flp.Controls.Count - 1
                Dim qath As String = flp.Controls(j).Tag.Fullpath
                If AreSameFile(path, qath) Then
                    SomeFilesSame = True
                End If
            Next
        Next
        Return SomeFilesSame

    End Function

    Private Shared Function AreSameFile(path As String, qath As String) As Boolean
        Dim same As Boolean = False
        Dim pbytes, qbytes As Byte()
        pbytes = ReadBinary(path, 1000)
        qbytes = ReadBinary(qath, 1000)
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
        pic = DirectCast(sender, PictureBox)
        For Each p In pic.Parent.Controls
            p = DirectCast(sender, PictureBox)
            pic.BorderStyle = BorderStyle.None
        Next
        pic.BorderStyle = BorderStyle.Fixed3D
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
End Class
Public Class DuplicateSet
    Public Property DSet As New List(Of DatabaseEntry)
    Public Property Size As Long
    Public Sub Add(Entry As DatabaseEntry)
        DSet.Add(Entry)
        If DSet.Count = 1 Then Size = Entry.Size
    End Sub
    Public Property Preserve As Byte = 0
    Public Sub Reset()
        DSet.Clear()
    End Sub


End Class

Public Class KeeperChooser
    Public Property DuplicateSet As New DuplicateSet
    Private mUniqueFolders As List(Of String)

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
    Public Sub ChooseUnBracketed()

        'All pairs need to be compared
        For i = 0 To DuplicateSet.DSet.Count - 1
            For j = i To DuplicateSet.DSet.Count - 1
                Dim x1 = DuplicateSet.DSet(i)
                Dim x2 = DuplicateSet.DSet(j)
                If x1.Path = x2.Path Then
                End If
            Next

        Next
    End Sub
    'For any pair of members of DuplicateSet
    'Are they both in the same folder?
    'If so, is one of them unbracketed? Mark it for preservation (out of this subgroup)
    'Otherwise are the filenames essentially the same at the beginning? Mark the closest to the beginning of the alphabet for preservation
    '
    'For any pair of members of duplicate set not in the same folder
    'Mark the relative depths (count "/"?)
    'Mark the number of files in that folder. 

    'For any pair of members
    'Do they have the same thumbnail? 
    'If no pairs have the same thumbnail, remove the first and repeat. 

    '

    '

End Class