Imports System.IO
Imports System.Runtime.InteropServices

Module FileMetadataModule
    Private Const FILE_ATTRIBUTE_HIDDEN As Integer = 2
    Private Const FILE_ATTRIBUTE_SYSTEM As Integer = 4
    Private Const FILE_ATTRIBUTE_READONLY As Integer = 1

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
    Private Function SetFileAlternateDataStreams(ByVal fileName As String, ByVal streamName As String, ByVal data As Byte(), ByVal length As Integer) As Boolean
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)>
    Private Function GetFileAlternateDataStreams(ByVal fileName As String, ByVal streamName As String, <Out> ByRef data As Byte(), ByRef length As Integer) As Boolean
    End Function

    Public Sub WriteMetadata(ByVal filePath As String, ByVal values As List(Of Long), ByVal tags As List(Of String))
        Dim tagString As String = GenerateTagString(values, tags)
        Dim fileInfo As New FileInfo(filePath)

        ' Set the hidden, system, and read-only attributes
        Dim attributes As FileAttributes = fileInfo.Attributes
        attributes = attributes Or FileAttributes.Hidden
        attributes = attributes Or FileAttributes.System
        attributes = attributes Or FileAttributes.ReadOnly

        ' Set the attributes using File.SetAttributes
        File.SetAttributes(filePath, attributes)

        ' Convert the tag string to byte array
        Dim tagData As Byte() = System.Text.Encoding.Unicode.GetBytes(tagString)

        ' Write the tag data to the alternate data stream
        SetFileAlternateDataStreams(filePath, "metadata", tagData, tagData.Length)
    End Sub

    Public Sub ReadMetadata(ByVal filePath As String, ByRef values As List(Of Long), ByRef tags As List(Of String))
        Dim fileInfo As New FileInfo(filePath)

        ' Read the tag data from the alternate data stream
        Dim tagData As Byte()
        Dim length As Integer = 0
        Dim success As Boolean = GetFileAlternateDataStreams(filePath, "metadata", tagData, length)

        If success Then
            ' Convert the byte array to tag string
            Dim tagString As String = System.Text.Encoding.Unicode.GetString(tagData, 0, length)
            ParseTagString(tagString, values, tags)
        Else
            ' No metadata found, initialize empty lists
            values = New List(Of Long)()
            tags = New List(Of String)()
        End If
    End Sub

    Private Function GenerateTagString(ByVal values As List(Of Long), ByVal tags As List(Of String)) As String
        Dim tagString As New Text.StringBuilder()

        ' Append values to the tag string
        For Each value As Long In values
            tagString.Append(value.ToString())
            tagString.Append("|") ' Use a separator
        Next

        ' Append tags to the tag string
        For Each tag As String In tags
            tagString.Append(tag)
            tagString.Append("|") ' Use a separator
        Next

        Return tagString.ToString()
    End Function

    Private Sub ParseTagString(ByVal tagString As String, ByRef values As List(Of Long), ByRef tags As List(Of String))
        values = New List(Of Long)()
        tags = New List(Of String)()

        ' Split the tag string using the separator
        Dim parts As String() = tagString.Split("|"c)

        ' Extract values from the tag string
        For Each part As String In parts
            Dim value As Long
            If Long.TryParse(part, value) Then
                values.Add(value)
            Else
                tags.Add(part)
            End If
        Next
    End Sub
End Module
