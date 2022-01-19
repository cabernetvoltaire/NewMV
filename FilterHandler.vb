Public Class FilterHandler
    Public Enum FilterState As Byte
        All
        PicVid
        Vidonly
        Piconly
        LinkOnly
        Linkless
        Linked
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
    Public Property SingleLinks As Boolean

    Public Event StateChanged(sender As Object, e As EventArgs)
    Private mColour = {Color.MintCream, Color.LemonChiffon, Color.LightSeaGreen, Color.LightPink, Color.LightBlue, Color.Azure, Color.Blue, Color.PaleTurquoise}
    Private mDescription = {"All files", "Only pictures and videos", "Only videos", "Only pictures", "Only links", "Videos without Links", "Videos with Links", "No pictures or videos"}
    Public ReadOnly Property Colour() As Color
        Get
            Return mColour(mState)
        End Get
    End Property
    Private mDescList As New List(Of String)

    Public ReadOnly Property Descriptions As List(Of String)
        Get
            For i = 0 To mDescription.length - 1
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
        Dim lst As New List(Of String) 'TODO: This removes items, we'd rather it hid them.
        'Try


        For Each f In mFileList
            lst.Add(f)
        Next
        Dim filelist As New Dictionary(Of String, String)
            For Each m In mFileList
                Dim f As New IO.FileInfo(m)
            Select Case mState
                Case FilterHandler.FilterState.All
                Case FilterHandler.FilterState.NoPicVid
                    If InStr(PICEXTENSIONS & VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And f.Extension.Length <> 0 Then
                    Else
                        lst.Remove(m)
                    End If
                Case FilterHandler.FilterState.LinkOnly
                    If LCase(f.Extension) = LinkExt Then
                        Dim x As String = LinkTarget(m)

                        If x = "" Then
                            lst.Remove(m)
                            Exit Select
                        End If
                        If x.EndsWith(".mvl") Then
                            lst.Remove(m)
                            Exit Select
                        End If
                        If x.Contains(vbNullChar) Then
                            lst.Remove(m)
                            Exit Select
                        End If
                        If SingleLinks Then
                            'For Each link In lst
                            If Not filelist.ContainsKey(x) Then filelist.Add(x, m)
                        End If


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
                    End If
                Case FilterHandler.FilterState.Linkless
                    If InStr(VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And Len(f.Extension) <> 0 Then
                        lst.Remove(m)
                    End If
                    If AllFaveMinder.GetLinksOf(m).Count = 0 Then
                    Else
                        lst.Remove(m)
                    End If
                Case FilterHandler.FilterState.Linked
                    If InStr(VIDEOEXTENSIONS, LCase(f.Extension)) = 0 And Len(f.Extension) <> 0 Then
                        lst.Remove(m)
                    End If
                    If AllFaveMinder.GetLinksOf(m).Count = 0 Then
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
            If FilterHandler.FilterState.LinkOnly And SingleLinks Then
                lst.Clear()

            For Each v In filelist.Keys
                lst.Add(v)
            Next
        End If


        Return lst
    End Function

End Class
