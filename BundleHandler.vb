
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

        MoveFiles(mList, Path, ListBox,, name)
        FSTree.RefreshTree(Path)
        'Add a folder node representing it. 
    End Sub
    Public Sub Burst(Optional Harvest As Boolean = False)
        Dim Maxfiles As Integer = Val(InputBox("Min no. of files to preserve folder? (Blank means all folders burst)",, "25"))
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
                    For Each f In subfolder.GetFiles
                        MoveFiles(f.FullName, Path, ListBox)

                    Next
                    RemoveEmptyFolders(subfolder.FullName)
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
            For Each f In mDirectory.EnumerateFiles("*", IO.SearchOption.AllDirectories)
                MoveFiles(f.FullName, Parentfolder.FullName, ListBox)
            Next
            mDirectory.Delete(True)
            FSTree.RemoveNode(Path)
        End If
        'Remove all the nodes of the folders removed.

    End Sub
    Public Sub HarvestBelow(Maxfiles As Integer)
        Burst(True)
        'For each folder under Path
        'Burst it. 

    End Sub
    Public Sub MoveFolderNode(Source As String, destination As String)
        FSTree.MoveNode(Source, destination)
        'Move the folder
        'Update the nodes
        FSTree.RefreshTree(destination)

        FSTree.RemoveNode(Source)

    End Sub

    Public Sub RemoveEmptyFolders(path)
        Dim folder As New IO.DirectoryInfo(path)
        Try
            Dim NoOfFiless = folder.EnumerateFiles.Count
            Dim NoOfFolders = folder.EnumerateDirectories.Count
            If NoOfFiless = 0 And NoOfFolders = 0 Then
                FSTree.RemoveNode(folder.FullName)
                folder.Delete()
                DirectoriesList.Remove(folder.FullName)
            Else
                For Each s In folder.EnumerateDirectories("*", IO.SearchOption.AllDirectories)
                    RemoveEmptyFolders(s.FullName)
                    ' DirectoriesList.Remove(s.FullName)
                Next
            End If
        Catch ex As Exception


        End Try

        'FSTree.RemoveNode(folder.FullName)
        'folder.Delete()
        'DirectoriesList.Remove(folder.FullName)
    End Sub

#End Region
End Class
