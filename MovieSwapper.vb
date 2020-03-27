Imports AxWMPLib
Public Class MediaSwapper

    Public WithEvents NextF As New NextFile
    Public WithEvents Media1 As New MediaHandler("mMedia1")
    Public WithEvents Media2 As New MediaHandler("mMedia2")
    Public WithEvents Media3 As New MediaHandler("mMedia3")

    Public RandomNext As Boolean = False
    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
    Public CurrentURLS As New List(Of String)
    <CodeAnalysis.SuppressMessage("Microsoft.Design", "CA1009:DeclareEventHandlersCorrectly")>
    Public Event LoadedMedia(MH As MediaHandler)
    Public Event MediaNotFound(MH As MediaHandler)
    Private Property mForceLoad As Boolean
    Public Property ForceLoad As Boolean
        Set(value As Boolean)
            Media1.Forceload = value
            Media2.Forceload = value
            Media3.Forceload = value
            mForceLoad = value
        End Set
        Get
            Return mForceload
        End Get
    End Property

    Public Property NextItem As String
    Public Property CurrentItem As String
    Public Property PreviousItem As String
    Public Event MediaShown(MH As MediaHandler)
#Region "Properties"

    ''' <summary>
    ''' Assigns the listbox which this Media Swapper controls
    ''' </summary>
    ''' <returns></returns>
    Public Property Listbox() As ListBox
        Get
            Return mListbox
        End Get
        Set(ByVal value As ListBox)
            mListbox = value

            NextF.Listbox = value

            For Each m In mListbox.Items
                mFileList.Add(m)
            Next
            GetNext()
        End Set
    End Property
    Private Sub GetNext()
        If RandomNext Then
            NextItem = NextF.RandomItem
            'Do Until NextItem <> CurrentItem
            '    NextItem = NextF.RandomItem
            'Loop

        Else
                NextItem = NextF.NextItem
        End If
    End Sub
    ''' <summary>
    ''' Sets the listindex to make the currently shown medium.
    ''' </summary>
    ''' <returns></returns>
    Public Property ListIndex() As Integer
        Get
            Return mListIndex
        End Get
        Set(ByVal value As Integer)
            If value < 0 Then Exit Property
            mListIndex = value

            NextF.CurrentIndex = value
            SetIndex(value)

        End Set
    End Property
    Public Sub New(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer, ByRef PB1 As PictureBox, ByRef PB2 As PictureBox, ByRef PB3 As PictureBox)

        AssignPlayers(MP1, MP2, MP3)
        AssignPictures(PB1, PB2, PB3)

    End Sub
    Public Sub New()

    End Sub
#End Region
#Region "Methods"
    Public Sub AssignPictures(ByRef PB1 As PictureBox, ByRef PB2 As PictureBox, ByRef PB3 As PictureBox)
        Media1.Picture = PB1
        Media2.Picture = PB2
        Media3.Picture = PB3

    End Sub
    Public Sub AssignPlayers(ByVal MP1 As AxWindowsMediaPlayer, ByVal MP2 As AxWindowsMediaPlayer, ByVal MP3 As AxWindowsMediaPlayer)

        Media1.Player = MP1
        Media2.Player = MP2
        Media3.Player = MP3

    End Sub
    ''' <summary>
    ''' Loads the media marked index, and puts the previous and next files into the other two media handlers
    ''' </summary>
    ''' <param name="index"></param>
    Private Sub SetIndex(index As Integer)
        mListcount = Listbox.Items.Count
        If mListcount >= 0 Then

            Static oldindex As Integer
            'Find the next and previous items for easy forward and backward
            NextF.Forwards = (mListIndex > oldindex) Or (mListIndex = 0)
            NextF.CurrentIndex = index

            CurrentItem = NextF.CurrentItem
            GetNext()
            PreviousItem = NextF.PreviousItem
            Report("Current:" & CurrentItem, 0)
            Report("Next:" & NextItem, 0)
            Report("Previous:" & PreviousItem, 0)
            Select Case CurrentItem

                Case Media2.MediaPath
                    RotateMedia(Media2, Media3, Media1)
                Case Media3.MediaPath
                    RotateMedia(Media3, Media1, Media2)
                Case Else
                    RotateMedia(Media1, Media2, Media3)
            End Select


            oldindex = index
        Else
            'Maybe use this later, to avoid multiple loadings which is inefficient.
            'For the moment - meh.
            'Select Case mListcount
            '    Case 1
            '        Prepare(Media1, CurrentItem)
            '        ShowPlayer(Media1)
            '    Case 2

            'End Select
        End If
    End Sub
    ''' <summary>
    ''' Loads the files into the Media Handler and sets it going, but doesn't release the pause (reset). 
    ''' </summary>
    ''' <param name="MH"></param>
    ''' <param name="path"></param>
    ''' <returns></returns>


    Private Function Prepare(ByRef MH As MediaHandler, path As String) As Boolean
        Debug.Print("PREPARE: " & MH.Player.Name & " with " & path)
        If MH.MediaPath <> path Or ForceLoad Then
            MH.MediaPath = path
            '       Else
            'MH.LoadMedia()
        End If
        Select Case MH.MediaType
            Case Filetype.Movie
                If blnFullScreen Then
                Else
                    If separate Then
                        MH.Player.uiMode = "mini"
                    Else

                        MH.Player.uiMode = "Full"
                    End If
                End If
                MH.Player.Visible = True
                MH.PlaceResetter(True)
                Return True
                RaiseEvent LoadedMedia(MH) 'Currently does nothing.
            Case Filetype.Pic
                MH.PlaceResetter(False)
                MH.MediaPath = path
                MH.Picture.Visible = True
                MH.Picture.Tag = path
                Return True

            Case Else
                Return False

        End Select
        'NB Duration is not loaded by this point. 
        CurrentURLS.Add(path)
        MH.IsCurrent = False
    End Function
    Function DisposeMedia(player As AxWindowsMediaPlayer) As Integer
        player.close()
        player.currentPlaylist.clear()
        Return 0
    End Function
    Public Sub CancelURL(filepath As String)
        If Media1.Player.URL = filepath Then DisposeMedia(Media1.Player)
        If Media2.Player.URL = filepath Then DisposeMedia(Media2.Player)
        If Media3.Player.URL = filepath Then DisposeMedia(Media3.Player)
        If Media1.Picture.Tag = filepath Then DisposePic(Media1.Picture)
        If Media2.Picture.Tag = filepath Then DisposePic(Media2.Picture)
        If Media3.Picture.Tag = filepath Then DisposePic(Media3.Picture)

    End Sub
    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)
        CurrentURLS.Clear()

        Prepare(PrevMH, PreviousItem)
        Prepare(NextMH, NextItem)
        Prepare(ThisMH, CurrentItem)
        ThisMH.IsCurrent = True

        Select Case ThisMH.MediaType
            Case Filetype.Movie
                ShowPlayer(ThisMH)
            Case Filetype.Pic
                ShowPicture(ThisMH)
        End Select

    End Sub
    Public Sub SetStartStates(ByRef SH As StartPointHandler)
        Media1.SPT.State = SH.State
        Media2.SPT.State = SH.State
        Media3.SPT.State = SH.State


    End Sub
    Public Sub SetStartpoints(ByRef SH As StartPointHandler) 'Only called when bars changed
        SetStartStates(SH)

        Media1.SPT = SH

        Media2.SPT = SH

        Media3.SPT = SH


    End Sub
    Public Sub URLSZero()
        Try

            Media1.Player.URL = ""
            Media2.Player.URL = ""
            Media3.Player.URL = ""
            DisposePic(Media1.Picture)
            DisposePic(Media2.Picture)
            DisposePic(Media3.Picture)
            ''Media1.Picture.Image = Nothing
            'Media2.Picture.Image = Nothing
            'Media3.Picture.Image = Nothing
            GC.Collect()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub ResettersOff()
        Media1.PlaceResetter(False)
        Media2.PlaceResetter(False)
        Media3.PlaceResetter(False)
    End Sub

    Public Sub MuteAll()
        Media1.Player.settings.mute = True
        Media2.Player.settings.mute = True
        Media3.Player.settings.mute = True

    End Sub
    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        MuteAll()

        MHX.PlaceResetter(False)
        With MHX.Player
            .Visible = True
            .BringToFront()
            .settings.mute = Muted
            RaiseEvent MediaShown(MHX)
        End With


    End Sub
    Private Sub OnMediaPlaying(sender As Object, e As EventArgs) Handles Media1.MediaPlaying, Media2.MediaPlaying, Media3.MediaPlaying
        'RaiseEvent MediaShown(sender)
    End Sub
    Private Sub OnRandomChanged(sender As Object, e As EventArgs) Handles NextF.RandomChanged
        NextItem = NextF.NextItem
    End Sub

    Public Sub ClickAllPics()
        PicClick(Media1.Picture)
        PicClick(Media2.Picture)
        PicClick(Media3.Picture)

    End Sub
    Private Sub HideMedias(CurrentMH As MediaHandler)
        ' Exit Sub
        If Media1 IsNot CurrentMH Then
            Media1.Picture.Visible = False
            Media1.Player.Visible = False
        End If
        If Media2 IsNot CurrentMH Then
            Media2.Picture.Visible = False
            Media2.Player.Visible = False
        End If
        If Media3 IsNot CurrentMH Then
            Media3.Picture.Visible = False
            Media3.Player.Visible = False
        End If

    End Sub
    Private Sub ShowPicture(ByRef MHX As MediaHandler)
        MuteAll()

        MHX.Picture.Visible = True
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX)

    End Sub
    Public Sub DockMedias(separated As Boolean)
        If separated Then
            Media1.Picture.Dock = DockStyle.None
            Media2.Picture.Dock = DockStyle.None
            Media3.Picture.Dock = DockStyle.None
            Media1.Player.Dock = DockStyle.None
            Media2.Player.Dock = DockStyle.None
            Media3.Player.Dock = DockStyle.None
            Media1.Picture.SetBounds(0, 0, 600, 400)
            Media2.Picture.SetBounds(650, 0, 600, 400)
            Media3.Picture.SetBounds(250, 480, 600, 400)
            Media1.Player.SetBounds(0, 0, 600, 400)
            Media2.Player.SetBounds(650, 0, 600, 400)
            Media3.Player.SetBounds(250, 480, 600, 400)

        Else
            Media1.Picture.Dock = DockStyle.Fill
            Media2.Picture.Dock = DockStyle.Fill
            Media3.Picture.Dock = DockStyle.Fill
            Media1.Player.Dock = DockStyle.Fill
            Media2.Player.Dock = DockStyle.Fill
            Media3.Player.Dock = DockStyle.Fill

        End If
    End Sub
#End Region

End Class
