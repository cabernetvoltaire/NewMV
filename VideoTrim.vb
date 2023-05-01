Imports System.Diagnostics
Imports System.Runtime.InteropServices
Imports System.ComponentModel
Imports System.IO


Module VideoTrim


    Property StartTime As Integer
    Property Finish As Integer
    Property InputFile1 As String
    Property Active As Boolean = False


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

        Dim outputFile As String = Path.Combine(inputDirectory, $"{inputFileNameWithoutExtension}_start_{StartTime}_duration_{duration}.mp4")
        ExtractVideoSegment(InputFile1, StartTime, duration, outputFile)
        Dim extractVideoSegmentTask As Task = Task.Run(Sub() ExtractVideoSegment(InputFile1, StartTime, duration, outputFile))

        Exit Sub
        ' Time range to extract (in seconds)


        ' Create an FFmpeg process
        Dim ffmpegProcess As New Process()
        ffmpegProcess.StartInfo.FileName = "C:\ffmpeg.exe"
        ffmpegProcess.StartInfo.Arguments = $"-i ""{InputFile1}"" -ss {StartTime} -t {duration} -c copy ""{outputFile}"""

        ffmpegProcess.StartInfo.UseShellExecute = False
        ffmpegProcess.StartInfo.CreateNoWindow = True
        ffmpegProcess.StartInfo.RedirectStandardOutput = True
        ffmpegProcess.StartInfo.RedirectStandardError = True
        Console.WriteLine($"Input file: {InputFile1}")
        Console.WriteLine($"Output file: {outputFile}")

        ' Start the FFmpeg process and wait for it to finish
        ffmpegProcess.Start()
        ffmpegProcess.WaitForExit()

        ' Check if the process exited successfully
        If ffmpegProcess.ExitCode = 0 Then
            Console.WriteLine("The portion of the MP4 file was successfully extracted.")
        Else
            Console.WriteLine("An error occurred during the extraction process.")
            Console.WriteLine("Error code: " & ffmpegProcess.ExitCode)
            Console.WriteLine("Error message: " & ffmpegProcess.StandardError.ReadToEnd())
        End If
    End Sub
    Sub ExtractVideoSegment(inputFile As String, startTime As String, duration As String, outputFile As String)
        Try
            Dim startInfo As New ProcessStartInfo()
            startInfo.FileName = "C:\ffmpeg.exe"
            startInfo.Arguments = $"-i ""{inputFile}"" -ss {startTime} -t {duration} -c copy ""{outputFile}"""
            startInfo.CreateNoWindow = True
            startInfo.UseShellExecute = False
            startInfo.RedirectStandardError = True

            Dim process As New Process()
            process.StartInfo = startInfo

            Console.WriteLine("Starting FFmpeg process...")
            process.Start()

            Console.WriteLine("Reading standard error...")
            Dim errorOutput As String = process.StandardError.ReadToEnd()

            Console.WriteLine("Waiting for process to exit...")
            process.WaitForExit()

            Console.WriteLine("Process exited.")
            MsgBox("Extraction Done")

            If process.ExitCode <> 0 Then
                Console.WriteLine($"Error code: {process.ExitCode}")
                Console.WriteLine($"Error message: {errorOutput}")
                Throw New Exception("An error occurred during the extraction process.")
            End If

        Catch ex As Exception
            Console.WriteLine($"Exception: {ex.Message}")
        End Try
    End Sub

    ' Rest of your code
End Module

