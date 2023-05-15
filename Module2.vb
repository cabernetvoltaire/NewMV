Imports System.IO
Imports System.Text

Module Module1

    Sub Main()
        Dim drives As DriveInfo() = DriveInfo.GetDrives()
        Dim outputFile As String = "folder_list.txt"

        Using sw As StreamWriter = New StreamWriter(outputFile, False, Encoding.UTF8)
            For Each drive As DriveInfo In drives
                If drive.IsReady Then
                    Try
                        ListFolders(drive.RootDirectory.FullName, sw)
                    Catch ex As Exception
                        Console.WriteLine($"Error accessing {drive.RootDirectory.FullName}: {ex.Message}")
                    End Try
                End If
            Next
        End Using

        Console.WriteLine("Folder list saved to: " & outputFile)
    End Sub

    Sub ListFolders(path As String, sw As StreamWriter)
        Try
            For Each folder As String In Directory.GetDirectories(path)
                sw.WriteLine(folder)
                ListFolders(folder, sw)
            Next
        Catch ex As Exception
            Console.WriteLine($"Error accessing {path}: {ex.Message}")
        End Try
    End Sub

End Module
