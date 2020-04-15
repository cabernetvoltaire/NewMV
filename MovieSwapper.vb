Imports AxWMPLib
Public Class MediaSwapper

    Public WithEvents NextF As New NextFile
    Public WithEvents Media1 As New MediaHandler("mMedia1")
    Public WithEvents Media2 As New MediaHandler("mMedia2")
    Public WithEvents Media3 As New MediaHandler("mMedia3")
    Public MediaHandlers As New List(Of MediaHandler) From {Media1, Media2, Media3}
    'Public WithEvents Pic1 As New PictureHandler(Media1.Picture)
    'Public WithEvents Pic2 As New PictureHandler(Media2.Picture)
    'Public WithEvents Pic3 As New PictureHandler(Media3.Picture)
    Private Outliner As New PictureBox With {.BackColor = Color.HotPink}
    Private WithEvents PauseAll As New Timer With {.Interval = 3000, .Enabled = False}

    Public RandomNext As Boolean = False
    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
    Public CurrentURLS As New List(Of String)

    Public Event MediaPlaying(sender As Object, e As EventArgs)
    Public Event MediaNotFound(sender As Object, e As EventArgs)
    Private Property mForceLoad As Boolean
    Public Property ForceLoad As Boolean
        Set(value As Boolean)
            For Each m In MediaHandlers
                m.Forceload = value
            Next
            mForceLoad = value
        End Set
        Get
            Return mForceload
        End Get
    End Property

    Public Property NextItem As String
    Public Property CurrentItem As String
    Public Property PreviousItem As String
    Public Event MediaShown(sender As Object, e As EventArgs)
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
    Public Sub AssignMedias(Med1 As MediaHandler, Med2 As MediaHandler, Med3 As MediaHandler)
        Media1 = Med1
        Media2 = Med2
        Media3 = Med3
    End Sub
    Public Sub AssignPictures(ByRef PB1 As PictureBox, ByRef PB2 As PictureBox, ByRef PB3 As PictureBox)
        Media1.PicHandler.PicBox = PB1
        Media2.PicHandler.PicBox = PB2
        Media3.PicHandler.PicBox = PB3
        Media3.Picture = PB3
        Media2.Picture = PB2
        Media1.Picture = PB1


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
            'Report("Current:" & CurrentItem, 0)
            'Report("Next:" & NextItem, 0)
            'Report("Previous:" & PreviousItem, 0)
            Select Case CurrentItem 'TODO: First time round, should load all items. This is the flaw that's been bugging us.

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
        ' MH.Forceload = True
        'If MH.MediaPath <> path Or ForceLoad Then
        MH.MediaPath = path 'Still not right for pics
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
              '  RaiseEvent LoadedMedia(MH) 'Currently does nothing.
            Case Filetype.Pic
                'MH.PlaceResetter(False)
                MH.Picture.Visible = True
                MH.Picture.Tag = path 'Important for thumbnail mouseover events. 
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
        For Each m In MediaHandlers
            m.CancelMedia()
        Next

    End Sub
    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)
        'CurrentURLS.Clear()
        ' HideMedias(ThisMH)
        Prepare(PrevMH, PreviousItem)
        Prepare(NextMH, NextItem)
        Prepare(ThisMH, CurrentItem)
        ThisMH.IsCurrent = True

        Select Case ThisMH.MediaType
            Case Filetype.Movie

                OutlineControl(ThisMH.Player, Outliner)
                ShowPlayer(ThisMH)
            Case Filetype.Pic
                OutlineControl(ThisMH.Picture, Outliner)
                ShowPicture(ThisMH)
        End Select

    End Sub
    Public Sub SetStartStates(ByRef SH As StartPointHandler)
        For Each m In MediaHandlers
            m.SPT.State = SH.State
        Next
    End Sub
    Public Sub SetStartpoints(ByRef SH As StartPointHandler) 'Only called when bars changed
        SetStartStates(SH)
        For Each m In MediaHandlers
            m.SPT = SH
        Next
    End Sub
    Public Sub URLSZero()
        Try
            For Each m In MediaHandlers
                m.Player.URL = ""
                DisposePic(m.Picture)
            Next
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Public Sub ResettersOff()
        For Each m In MediaHandlers
            m.PlaceResetter(False)
        Next
    End Sub

    Public Sub MuteAll()
        For Each m In MediaHandlers
            m.Player.settings.mute = True
        Next

    End Sub
    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        MuteAll()

        MHX.PlaceResetter(False) 'Starts the video playing

        With MHX.Player
            .Visible = True
            .BringToFront()
            .settings.mute = Muted
            'MHX.Speed.Paused = False
            RaiseEvent MediaShown(MHX, Nothing)
        End With


    End Sub
    Private Sub OnMediaPlaying(sender As Object, e As EventArgs) Handles Media1.MediaPlaying, Media2.MediaPlaying, Media3.MediaPlaying
        ' RaiseEvent MediaPlaying(sender, e)
    End Sub
    Private Sub OnRandomChanged(sender As Object, e As EventArgs) Handles NextF.RandomChanged
        NextItem = NextF.NextItem
    End Sub

    Public Sub ClickAllPics(sender As Object, e As MouseEventArgs)
        For Each m In MediaHandlers
            m.PicHandler.PicClick(m, e)
        Next
    End Sub
    Private Sub HideMedias(CurrentMH As MediaHandler)
        For Each m In MediaHandlers
            If m IsNot CurrentMH Then
                m.Picture.Visible = False
                m.Player.Visible = False
            End If
        Next

    End Sub
    Private Sub ShowPicture(ByRef MHX As MediaHandler)
        MuteAll()

        MHX.Picture.Visible = True
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX, Nothing)

    End Sub
    Public Sub PauseMovies() Handles PauseAll.Tick

        Media1.Speed.Paused = True
        Media2.Speed.Paused = True
        Media3.Speed.Paused = True

    End Sub
    Public Sub SetPicState(state As Byte)
        For Each m In MediaHandlers
            If m.PicHandler.State <> state Then
                m.PicHandler.SetState(state)

            End If
        Next

    End Sub
    Public Sub ZoomPics(zoomfactor As Integer)
        For Each m In MediaHandlers
            m.PicHandler.ZoomFactor = zoomfactor
        Next
    End Sub

    Public Sub DockMedias(separated As Boolean)
        If separated Then
            For Each m In MediaHandlers
                m.Picture.Dock = DockStyle.None
                m.Player.Dock = DockStyle.None

            Next

            Media1.Picture.SetBounds(0, 0, 600, 400)
            Media2.Picture.SetBounds(650, 0, 600, 400)
            Media3.Picture.SetBounds(250, 480, 600, 400)
            Media1.Player.SetBounds(0, 0, 600, 400)
            Media2.Player.SetBounds(650, 0, 600, 400)
            Media3.Player.SetBounds(250, 480, 600, 400)

        Else
            For Each m In MediaHandlers
                m.Picture.Dock = DockStyle.Fill
                m.Player.Dock = DockStyle.Fill
            Next

        End If
    End Sub
#End Region

End Class
