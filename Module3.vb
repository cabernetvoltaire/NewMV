Imports System.IO

Module Module3
    Sub Main()
        Dim allDrives() As DriveInfo = DriveInfo.GetDrives()

        For Each d In allDrives
            Console.WriteLine("Drive {0}", d.Name)
            Console.WriteLine("  Drive type: {0}", d.DriveType)

            If d.IsReady = True Then
                Console.WriteLine("  Volume label: {0}", d.VolumeLabel)
                Console.WriteLine("  File system: {0}", d.DriveFormat)
                Console.WriteLine("  Available space to current user:{0, 15} bytes", d.AvailableFreeSpace)
                Console.WriteLine("  Total available space:          {0, 15} bytes", d.TotalFreeSpace)
                Console.WriteLine("  Total size of drive:            {0, 15} bytes ", d.TotalSize)
            End If
        Next
    End Sub
End Module
