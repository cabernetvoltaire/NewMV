Imports Microsoft.Office.Interop

Public Class Spreadsheet
    Public oXL As Excel.Application
    Public oWB As Excel.Workbook
    Public oSheet As Excel.Worksheet
    Public oRng As Excel.Range

    Public Sub New()
        oXL = CreateObject("Excel.Application")
        oXL.Visible = True

        ' Get a new workbook.
        oWB = oXL.Workbooks.Add
        oSheet = oWB.ActiveSheet

        ' Add table headers going cell by cell.
        oSheet.Cells(1, 1).Value = "Path"
        oSheet.Cells(1, 2).Value = "DateTime"
        oSheet.Cells(1, 3).Value = "Size"
        oSheet.Cells(1, 4).Value = "LinkPath"
        oSheet.Cells(1, 5).Value = "Marker"



    End Sub

    Public Sub AddLinksFromDirectory(dir As String)
        Dim dirinfo As New IO.DirectoryInfo(dir)
        Dim i = 1
        For Each f In dirinfo.EnumerateFiles("*.lnk", IO.SearchOption.AllDirectories)
            Try
                Dim file As New IO.FileInfo(LinkTarget(f.FullName))
                i += 1

                oSheet.Cells(i, 1) = file.FullName
                oSheet.Cells(i, 2) = file.CreationTime
                oSheet.Cells(i, 3) = file.Length
                oSheet.Cells(i, 4) = f.FullName
                oSheet.Cells(i, 5) = BookmarkFromLinkName(f.FullName)
            Catch ex As Exception
                Continue For
            End Try

        Next
    End Sub

    Public Sub SetFilters()
        oRng = oSheet.Range("A1:E1")
        oRng.AdvancedFilter(Excel.XlFilterAction.xlFilterInPlace)
    End Sub

    Public Sub CatalogueThisDirectory(dir As String, Optional recursive As Boolean = False)
        SetFilters()
        Dim dirinfo As New IO.DirectoryInfo(dir)
        Dim i = 1
        Dim j = 5
        Dim m As IO.SearchOption
        If recursive Then
            m = IO.SearchOption.AllDirectories
        Else
            m = IO.SearchOption.TopDirectoryOnly
        End If

        For Each f In dirinfo.EnumerateFiles("*", m)
            Try
                Dim file As New IO.FileInfo(LinkTarget(f.FullName))
                i += 1

                oSheet.Cells(i, 1) = f.FullName
                oSheet.Cells(i, 2) = f.CreationTime
                oSheet.Cells(i, 3) = f.Length
                oSheet.Cells(i, 4) = f.FullName
                Dim x As List(Of String) = AllFaveMinder.GetLinksOf(f.FullName)
                Dim xl As New List(Of Long)
                For Each fl In x
                    xl.Add(BookmarkFromLinkName(fl))
                Next
                xl.Sort()

                For Each nm In xl
                    oSheet.Cells(i, j) = nm
                    j += 1
                Next
                j = 5

            Catch ex As Exception
                Continue For
            End Try

        Next

    End Sub
End Class
