
Public Class BundleHandler
#Region "Properties"
    Private mDirectory As New IO.DirectoryInfo("C:\")
    Public Property Path As String
        Get
            Path = mPath
        End Get
        Set(value As String)
            Dim mdir As New IO.DirectoryInfo(value)
            mDirectory = mdir
            mPath = value
        End Set

    End Property
    Private Property mPath As String
    Public Property FSTree As MasaSam.Forms.Controls.FileSystemTree
    Public Property ListBox As New ListBox
    Private Property mList As New List(Of String)
    Public Property List As List(Of String)
        Get
            List = mList
        End Get
        Set(value As List(Of String))
            mList = value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Sub New(Tree As MasaSam.Forms.Controls.FileSystemTree, Box As ListBox, Path As String)
        FSTree = Tree
        ListBox = Box
        mPath = Path
        mDirectory = New IO.DirectoryInfo(mPath)
    End Sub
    Public Sub Bundle(name As String)
        'For all the files in the list
        'Create a folder (asking for name if necessary)
        'Put the files in it. 
        mList = ListfromSelectedInListbox(ListBox)

        MoveFiles(mList, Path, ListBox)
        FSTree.RefreshTree(Path)
        'Add a folder node representing it. 
    End Sub
    Public Async Function Burst(Optional Harvest As Boolean = False) As Task
        Dim Maxfiles As Integer = Val(InputBox("Min no. of files to preserve folder? (Blank means all folders burst)",, ""))
        'Get the parent folder

        Dim folders As New List(Of IO.DirectoryInfo)
        Dim Parentfolder As New IO.DirectoryInfo(Path)
        If Harvest Then
        Else
            Parentfolder = Parentfolder.Parent
        End If
        'Get all the folders with more than maxfiles
        For Each subfolder In mDirectory.GetDirectories("*", IO.SearchOption.TopDirectoryOnly)
            Try

                If subfolder.GetFiles.Count > Maxfiles Then
                    folders.Add(subfolder)
                Else
                    Dim list As New List(Of String)
                    For Each f In subfolder.GetFiles
                        list.Add(f.FullName)
                    Next
                    MoveFiles(list, Path, ListBox)
                    Await RemoveEmptyFolders(subfolder.FullName, True)
                End If
            Catch ex As Exception

            End Try
        Next
        'Promote them all to the parent folder(Retaining hierarchy)
        If Not Harvest Then
            For Each f In folders
                MoveFolderNode(f.FullName, Parentfolder.FullName)
                FSTree.RemoveNode(f.FullName)
            Next
            'Get all the remaining files
            'Promote them all to the parent folder
            Dim list As New List(Of String)
            For Each f In mDirectory.EnumerateFiles("*", IO.SearchOption.AllDirectories)
                list.Add(f.FullName)
            Next
            Report("Moving Files", 2)
            MoveFiles(list, Parentfolder.FullName, ListBox)
            Report("Deleting directory", 1)
            'mDirectory.Delete(True)
            FSTree.RemoveNode(Path)
        End If
        'Remove all the nodes of the folders removed.

    End Function
    Public Async Function HarvestBelow(Maxfiles As Integer) As Task
        Await Burst(True)

    End Function
    Public Sub MoveFolderNode(Source As String, destination As String)
        FSTree.MoveNode(Source, destination)
        'Move the folder
        'Update the nodes
        FSTree.RefreshTree(destination)

        FSTree.RemoveNode(Source)

    End Sub

    Public Async Function RemoveEmptyFolders(path As String, recurse As Boolean) As Task
        Dim folder As New IO.DirectoryInfo(path)
        Try
            Dim NoOfFiless = folder.EnumerateFiles.Count
            Dim NoOfFolders = folder.EnumerateDirectories.Count
            If NoOfFiless = 0 And NoOfFolders = 0 Then
                FSTree.RemoveNode(folder.FullName)
                folder.Delete()
                DirectoriesList.Remove(folder.FullName)
            Else
                If recurse Then

                    For Each s In folder.EnumerateDirectories("*", IO.SearchOption.AllDirectories)
                        Await RemoveEmptyFolders(s.FullName, recurse)
                        'DirectoriesList.Remove(s.FullName)
                    Next
                End If
            End If
        Catch ex As Exception


        End Try

    End Function

#End Region
End Class
