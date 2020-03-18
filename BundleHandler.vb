﻿
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
        blnSuppressCreate = False

        MoveFiles(mList, Path, ListBox)
        FSTree.RefreshTree(Path)

        'Add a folder node representing it. 
    End Sub
    Public Property Maxfiles As Integer
    Public Async Function Burst(CurrentFolder As IO.DirectoryInfo, Optional Harvest As Boolean = False) As Task
        Dim DestinationFolder As New IO.DirectoryInfo(CurrentFolder.FullName)
        If Not Harvest Then DestinationFolder = CurrentFolder.Parent
        Dim folders As New List(Of IO.DirectoryInfo)
        For Each subfolder In CurrentFolder.GetDirectories("*", IO.SearchOption.AllDirectories)
            If subfolder.GetFiles.Count <= Maxfiles Or Maxfiles = 0 Then
                folders.Add(subfolder)
            End If
        Next
        For Each fol In folders
            Await Burst(fol, False)
        Next
        Dim list As New List(Of String)
        For Each f In CurrentFolder.GetFiles
            list.Add(f.FullName)
        Next
        blnSuppressCreate = True
        MoveFiles(list, DestinationFolder.FullName, ListBox)
        FSTree.RefreshTree(DestinationFolder.FullName)

        Await RemoveEmptyFolders(CurrentFolder.FullName, True)
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
            If recurse Then

                For Each s In folder.EnumerateDirectories("*", IO.SearchOption.AllDirectories)
                    Await RemoveEmptyFolders(s.FullName, recurse)
                    ' DirectoriesList.Remove(s.FullName)
                Next
            End If
            If NoOfFiless = 0 And NoOfFolders = 0 Then
                FSTree.RemoveNode(folder.FullName)
                folder.Delete()
                DirectoriesList.Remove(folder.FullName)
            End If
        Catch ex As Exception


        End Try

        'FSTree.RemoveNode(folder.FullName)
        'folder.Delete()
        'DirectoriesList.Remove(folder.FullName)
    End Function

#End Region
End Class
