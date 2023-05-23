Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.IO


Module VideoTrim


    Property StartTime As Integer
    Property Finish As Integer
    Property InputFile1 As String
    Property Active As Boolean = False
    Public Event Finished(sender As Object, e As EventArgs)
    Private Function CommandLineToArgvW(<MarshalAs(UnmanagedType.LPWStr)> lpCmdLine As String, ByRef pNumArgs As Integer) As IntPtr
    End Function

    Private Function EscapeArg(arg As String) As String
        Dim argc As Integer
        Dim argv As IntPtr = CommandLineToArgvW("dummy.exe " & arg, argc)

        If argv = IntPtr.Zero Then
            Throw New Win32Exception()
        End If

        Dim escapedArg As String = Marshal.PtrToStringUni(Marshal.ReadIntPtr(argv, IntPtr.Size))
        Marshal.FreeHGlobal(argv)

        Return escapedArg
    End Function
    Sub Main()
        Dim duration As Integer = Finish - StartTime
        ' Replace inputFile with the path of the first input file
        Dim inputDirectory As String = Path.GetDirectoryName(InputFile1)
        Dim inputFileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(InputFile1)

        Dim OutputFile As String = Path.Combine(inputDirectory, $"{inputFileNameWithoutExtension}_start_{StartTime}_duration_{duration}.mp4")

        ExtractVideoSegment(InputFile1, StartTime, duration, outputFile)
        Dim extractVideoSegmentTask As Task = Task.Run(Sub() ExtractVideoSegment(InputFile1, StartTime, duration, outputFile))


    End Sub
    Private Sub ExtractVideoSegment(inputFile As String, startTime As Integer, duration As Integer, outputFile As String)
        ' Create an FFmpeg process
        Using ffmpegProcess As New Process()
            ffmpegProcess.StartInfo.FileName = "C:\ffmpeg.exe"
            ffmpegProcess.StartInfo.Arguments = $"-i ""{inputFile}"" -ss {startTime} -t {duration} -c copy ""{outputFile}"""

            ffmpegProcess.StartInfo.UseShellExecute = False
            ffmpegProcess.StartInfo.CreateNoWindow = True
            ffmpegProcess.StartInfo.RedirectStandardOutput = True
            ffmpegProcess.StartInfo.RedirectStandardError = True

            ' Start the FFmpeg process and wait for it to finish
            ffmpegProcess.Start()
            ffmpegProcess.WaitForExit()
            RaiseEvent Finished(outputFile, Nothing)
            ' Check if the process exited successfully
            If ffmpegProcess.ExitCode = 0 Then
                Console.WriteLine("The portion of the MP4 file was successfully extracted.")
            Else
                Console.WriteLine("An error occurred during the extraction process.")
                Console.WriteLine("Error code: " & ffmpegProcess.ExitCode)
                Console.WriteLine("Error message: " & ffmpegProcess.StandardError.ReadToEnd())
            End If
        End Using ' This will call Dispose() on the process, cleaning up resources
    End Sub

    ' Rest of your code
End Module

