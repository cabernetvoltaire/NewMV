Imports System.Object
Public Class DuplicatesForm
    Private mDB As Database

    Public Property DB() As Database
        Get
            Return mDB
        End Get
        Set(ByVal value As Database)
            mDB = value
        End Set
    End Property
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
    End Sub
    Public Sub ShowDuplicates()

        For j = 0 To 200
            Dim flp As New FlowLayoutPanel With {.Dock = DockStyle.Top}
            Dim m As DuplicateSet = AllDuplicates.Item(j)
            For i = 0 To m.DSet.Count - 1
                Dim pic As New PictureBox With {.Height = 100, .Width = 100,
                                 .SizeMode = PictureBoxSizeMode.AutoSize,
                                 .Image = GetImage(m.DSet(i).FullPath)}
                pic.Image = pic.Image.GetThumbnailImage(150, 150, Nothing, New IntPtr)
                flp.Controls.Add(pic)

            Next
            Me.Controls.Add(flp)


        Next
    End Sub
End Class
Public Class DuplicateSet
    Public Property DSet As New List(Of DatabaseEntry)
    Public Property Size As Long
    Public Sub Add(Entry As DatabaseEntry)
        DSet.Add(Entry)
        If DSet.Count = 1 Then Size = Entry.Size
    End Sub

    Public Sub Reset()
        DSet.Clear()
    End Sub


End Class
