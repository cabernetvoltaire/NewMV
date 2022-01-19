Public Class FileOperations
#Region "Members"
    Public Event Filesmoving(sender As Object, e As EventArgs)
    Private Property mFiles As New List(Of String)
    Public Property Files() As List(Of String)
        Get
            Return mFiles
        End Get
        Set(ByVal value As List(Of String))
            mFiles = value
        End Set
    End Property
    Public Property DestinationFolder As IO.DirectoryInfo
    Public Property CurrentFolder As IO.DirectoryInfo
    Public Property Filebox As FileboxHandler
    Public Property Listbox As ListBoxHandler
    Public Property Treeview As New MasaSam.Forms.Controls.FileSystemTree

#End Region
#Region "Methods"
    Public Sub New(Current As IO.DirectoryInfo, Destination As IO.DirectoryInfo)
        CurrentFolder = Current
        DestinationFolder = DestinationFolder
    End Sub
    Public Sub New()

    End Sub
    Public Sub MoveFiles()
        'Move the files
        For Each f In mFiles
            Dim file As New IO.FileInfo(f)
            RaiseEvent Filesmoving(Me, Nothing)
            If file.Exists Then file.MoveTo(AppendExistingFilename(DestinationFolder.FullName & "\" & file.Name, file))
        Next
        'Remove them from the filebox
        'Remove them from the listbox
    End Sub
    Public Sub DeleteFiles()
        'Delete the files (to Folder)
        RaiseEvent Filesmoving(Me, Nothing)
        For Each f In mFiles
            Dim file As New IO.FileInfo(f)
            If file.Exists Then My.Computer.FileSystem.DeleteFile(f, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
        Next
        RemoveFromListbox()
        RemoveFromFilebox()
    End Sub

    Private Sub RemoveFromListbox()
        If Listbox IsNot Nothing Then
            Listbox.RemoveItems(mFiles)
        End If
    End Sub
    Private Sub RemoveFromFilebox()
        If Filebox IsNot Nothing Then
            Filebox.RemoveItems(mFiles)
        End If
    End Sub

    Public Sub MakeFolder()
        'Create a folder
        'Add a node to the Treeview
    End Sub
    Public Sub RemoveFolder(EmptyOnly As Boolean)
        'Delete the folder (Optionally if it is empty)
        'Remove the node from the treeview
    End Sub
    Public Sub MoveFolder()
        'To the destination folder
    End Sub
    Public Sub RemoveEmptySubfolders()
        Dim finished As Boolean = False
        Dim x As New List(Of IO.DirectoryInfo)

        While Not finished
            'Make a list of all the empty directories

            For Each fol In CurrentFolder.GetDirectories("*", IO.SearchOption.AllDirectories)
                If fol.GetDirectories.Count = 0 And fol.GetFiles.Count = 0 Then
                    x.Add(fol)
                End If
            Next
            'If there aren't any then done
            If x.Count = 0 Then
                finished = True
                'Otherwise, delete them all
            Else
                For Each f In x
                    My.Computer.FileSystem.DeleteDirectory(f.FullName, FileIO.UIOption.OnlyErrorDialogs, FileIO.RecycleOption.SendToRecycleBin)
                    'f.Delete(False)
                Next
            End If
            x.Clear()
        End While
    End Sub

#End Region
#Region "Events"

#End Region
End Class
