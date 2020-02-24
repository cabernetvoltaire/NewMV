Public Class FilterHandler
    Public Enum FilterState As Byte
        All
        PicVid
        Vidonly
        Piconly
        LinkOnly
        NoPicVid
    End Enum
    Private mFileList As List(Of String)
    Public Property FileList As List(Of String)
        Get
            Return FilterLBList()
        End Get
        Set(value As List(Of String))
            mFileList = value
        End Set
    End Property


    Public Event StateChanged(sender As Object, e As EventArgs)
    Private mColour = {Color.MintCream, Color.LemonChiffon, Color.LightSeaGreen, Color.LightPink, Color.LightBlue, Color.PaleTurquoise}
    Private mDescription = {"All files", "Only pictures and videos", "Only videos", "Only pictures", "Only links", "No pictures or videos"}
    Public ReadOnly Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
    End Property
    Private mDescList As New List(Of String)

    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To 5
                mDescList.Add(mDescription(i))
            Next
            Descriptions = mDescList
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return mDescription(mState)
        End Get
    End Property
    Private mState As Byte

    Public Sub New()
        Me.State = FilterState.All
    End Sub
    Public Property OldState As Byte = 0
    Public Property State() As Byte
        Get
            Return mState
        End Get
        Set(ByVal value As Byte)
            If value = FilterState.LinkOnly Then OldState = mState
            mState = value

            RaiseEvent StateChanged(Me, New EventArgs)
        End Set
    End Property
    Public Sub IncrementState(Limited As Boolean)
        If Limited Then
            Me.State = (Me.State + 1) Mod (FilterState.PicVid + 1)

        Else
            Me.State = (Me.State + 1) Mod (FilterState.NoPicVid + 1)

        End If
    End Sub
    Private Function FilterLBList() As List(Of String)
        Dim lst As New List(Of String)
        Try


            For Each f In mFileList
                lst.Add(f)
            Next
            For Each m In mFileList
                Dim f As New IO.FileInfo(m)
                Select Case CurrentfilterState.State
                    Case FilterHandler.FilterState.All
                    Case FilterHandler.FilterState.NoPicVid
                        If InStr(PICEXTENSIONS & VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And f.Extension.Length <> 0 Then
                        Else
                            lst.Remove(m)
                        End If
                    Case FilterHandler.FilterState.LinkOnly
                        If LCase(f.Extension) = ".lnk" Then
                        Else
                            lst.Remove(m)
                        End If
                    Case FilterHandler.FilterState.Piconly
                        If InStr(PICEXTENSIONS, LCase(f.Extension)) = 0 And Len(f.Extension) <> 0 Then
                            lst.Remove(m)
                        Else
                        End If
                    Case FilterHandler.FilterState.PicVid
                        If InStr(PICEXTENSIONS & VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And Len(f.Extension) <> 0 Then
                            lst.Remove(m)
                        Else
                        End If
                    Case FilterHandler.FilterState.Vidonly
                        If InStr(VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And Len(f.Extension) <> 0 Then
                            lst.Remove(m)
                        Else
                        End If
                End Select
            Next
        Catch ex As Exception
            MsgBox("Error in file list" & ex.Message)
            Return Nothing
            Exit Function
        End Try

        Return lst
    End Function

End Class
