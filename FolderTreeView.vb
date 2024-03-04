Imports System.IO

Public Class FolderTreeView
    Private treeView As TreeView

    Public Sub New(ByVal tv As TreeView)
        treeView = tv
        InitializeTreeView()
    End Sub

    Private Sub InitializeTreeView()
        treeView.Nodes.Clear()
        ' Optional: Setup TreeView properties such as ShowLines, ShowPlusMinus, etc.
    End Sub

    Public Sub PopulateTreeView(ByVal rootFolderPath As String)
        treeView.Nodes.Clear()
        Dim rootNode As TreeNode = New TreeNode(rootFolderPath)
        rootNode.Tag = rootFolderPath ' Use Tag to store the full path
        treeView.Nodes.Add(rootNode)
        AddDirectories(rootNode, rootFolderPath)
        rootNode.Expand()
    End Sub

    Private Sub AddDirectories(ByVal currentNode As TreeNode, ByVal path As String)
        Dim directoryInfo As DirectoryInfo = New DirectoryInfo(path)
        For Each directory As DirectoryInfo In directoryInfo.GetDirectories()
            Dim directoryNode As TreeNode = New TreeNode(directory.Name)
            directoryNode.Tag = directory.FullName
            currentNode.Nodes.Add(directoryNode)
            AddDirectories(directoryNode, directory.FullName) ' Recursion to add subdirectories
        Next
    End Sub

    Public Sub CreateFolder(ByVal folderName As String)
        If treeView.SelectedNode IsNot Nothing Then
            Dim selectedPath As String = treeView.SelectedNode.Tag.ToString()
            Dim newPath As String = Path.Combine(selectedPath, folderName)
            Directory.CreateDirectory(newPath)
            Dim newNode As TreeNode = New TreeNode(folderName)
            newNode.Tag = newPath
            treeView.SelectedNode.Nodes.Add(newNode)
            treeView.SelectedNode.Expand()
        End If
    End Sub

    Public Sub RemoveSelectedFolder()
        If treeView.SelectedNode IsNot Nothing AndAlso treeView.SelectedNode.Parent IsNot Nothing Then
            Dim selectedPath As String = treeView.SelectedNode.Tag.ToString()
            Directory.Delete(selectedPath, True) ' True to remove directories and files within it
            treeView.Nodes.Remove(treeView.SelectedNode)
        End If
    End Sub


End Class
