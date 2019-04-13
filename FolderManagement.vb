Module FolderManagement
    Public watchfolder As FileSystemWatcher

    Public Sub Folderwatch(path As String)
        watchfolder = New System.IO.FileSystemWatcher

        watchfolder.NotifyFilter = IO.NotifyFilters.DirectoryName
        watchfolder.NotifyFilter = watchfolder.NotifyFilter Or
                                   IO.NotifyFilters.FileName
        watchfolder.NotifyFilter = watchfolder.NotifyFilter Or
                                   IO.NotifyFilters.Attributes
        AddHandler watchfolder.Changed, AddressOf logchange
        AddHandler watchfolder.Created, AddressOf logchange
        AddHandler watchfolder.Deleted, AddressOf logchange

        ' add the rename handler as the signature is different
        AddHandler watchfolder.Renamed, AddressOf logrename

        'Set this property to true to start watching
        watchfolder.EnableRaisingEvents = True

    End Sub
    Private Function logchange(ByVal source As Object, ByVal e As _
                        System.IO.FileSystemEventArgs) As String
        If e.ChangeType = IO.WatcherChangeTypes.Changed Then
            Return "File " & e.FullPath &
                                    " has been modified" & vbCrLf
        End If
        If e.ChangeType = IO.WatcherChangeTypes.Created Then
            Return "File " & e.FullPath &
                                     " has been created" & vbCrLf
        End If
        If e.ChangeType = IO.WatcherChangeTypes.Deleted Then
            Return "File " & e.FullPath &
                                    " has been deleted" & vbCrLf
        End If
    End Function
    Public Function logrename(ByVal source As Object, ByVal e As _
                            System.IO.RenamedEventArgs) As String
        Return "File" & e.OldName &
                      " has been renamed to " & e.Name & vbCrLf
    End Function
End Module
