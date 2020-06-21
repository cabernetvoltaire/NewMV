Imports AxWMPLib
Public Class MediaSwapper

    Public WithEvents NextF As New NextFile
    Public WithEvents Media1 As New MediaHandler("mMedia1")
    Public WithEvents Media2 As New MediaHandler("mMedia2")
    Public WithEvents Media3 As New MediaHandler("mMedia3")
    Public MediaHandlers As New List(Of MediaHandler) From {Media1, Media2, Media3}


    Private Outliner As New PictureBox With {.BackColor = Color.HotPink}
    Private WithEvents PauseAll As New Timer With {.Interval = 3000, .Enabled = False}
    Private mRandomNext As Boolean = False
    Public WriteOnly Property RandomNext As Boolean
        Set(value As Boolean)
            mRandomNext = value
            For Each m In MediaHandlers
                m.PicHandler.RandomNext = value
            Next
        End Set

    End Property

    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
    Public CurrentURLS As New List(Of String)

    Public Event MediaPlaying(sender As Object, e As EventArgs)
    Public Event MediaNotFound(sender As Object, e As EventArgs)

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
        If mRandomNext Then
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
        Report("PREPARE: " & MH.Player.Name & " with " & path)
        'BreakExecution()
        If MH.MediaPath <> path Then MH.MediaPath = path 'Still not right for pics
        Select Case MH.MediaType
            Case Filetype.Movie
                If blnFullScreen Then
                Else
                    MH.Player.uiMode = "Full"
                End If
                '                MH.Player.Visible = False
                '               MH.Player.SendToBack()
                MH.PlaceResetter(True)
                DisposePic(MH.PicHandler.PicBox)
                Return True
              '  RaiseEvent LoadedMedia(MH) 'Currently does nothing.
            Case Filetype.Pic
                'MH.PlaceResetter(False)
                MH.DisposeMovie()
                MH.Picture.Visible = True
                MH.Picture.Tag = path 'Important for thumbnail mouseover events. 
                Return True
            Case Filetype.Doc

                Return True
            Case Else
                Return False

        End Select
        'NB Duration is not loaded by this point. 
        CurrentURLS.Add(path)
        MH.IsCurrent = False
    End Function

    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)
        'CurrentURLS.Clear()
        ' HideMedias(ThisMH)

        Prepare(ThisMH, CurrentItem)
        ThisMH.IsCurrent = True

        Select Case ThisMH.MediaType
            Case Filetype.Movie

                If separate Then OutlineControl(ThisMH.Player, Outliner)
                ShowPlayer(ThisMH)
            Case Filetype.Pic
                If separate Then OutlineControl(ThisMH.Picture, Outliner)
                ShowPicture(ThisMH)
            Case Filetype.Doc
                If separate Then OutlineControl(ThisMH.Picture, Outliner)
                ShowTextFile(ThisMH)
        End Select
        Prepare(PrevMH, PreviousItem)
        Prepare(NextMH, NextItem)
    End Sub
    Private Sub ShowPicture(ByRef MHX As MediaHandler)
        MuteAll()
        ' HideMedias(MHX)
        MHX.Picture.Visible = True
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX, Nothing)

    End Sub
    Private Sub ShowTextFile(ByRef MHX As MediaHandler)
        MuteAll()
        ' HideMedias(MHX)

        MHX.Visible = False
        MHX.Textbox.BringToFront()
        MHX.Textbox.Visible = True
        RaiseEvent MediaShown(MHX, Nothing)

    End Sub
    Private Sub ShowPlayer(ByRef MHX As MediaHandler)
        MuteAll()
        '  HideMedias(MHX)

        ResetPositionsAgain()
        MHX.PlaceResetter(False) 'Starts the video playing
        MHX.Visible = False
        With MHX.Player
            .Visible = True
            .BringToFront()
            .settings.mute = Muted
            RaiseEvent MediaShown(MHX, Nothing)
        End With


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
    Function DisposeMedia(player As AxWindowsMediaPlayer) As Integer
        player.close()
        player.currentPlaylist.clear()
        Return 0
    End Function
    Private Sub ResetPositionsAgain()
        For Each m In MediaHandlers
            m.ResetPositionCanceller.Enabled = True
            m.PlaceResetter(True)
        Next
    End Sub
    Public Sub CancelURL(filepath As String)
        For Each m In MediaHandlers
            If m.MediaPath = filepath Then m.CancelMedia()
        Next

    End Sub
    Public Sub CancelURL()
        For Each m In MediaHandlers
            m.CancelMedia()
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
                m.Visible = False
            End If
        Next

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
    'Public Sub DockMedias(separated As Boolean)
    '    If separated Then
    '        For Each m In MediaHandlers
    '            m.Picture.Dock = DockStyle.None
    '            m.Player.Dock = DockStyle.None

    '        Next

    '        Media1.Picture.SetBounds(0, 0, 600, 400)
    '        Media2.Picture.SetBounds(650, 0, 600, 400)
    '        Media3.Picture.SetBounds(250, 480, 600, 400)
    '        Media1.Player.SetBounds(0, 0, 600, 400)
    '        Media2.Player.SetBounds(650, 0, 600, 400)
    '        Media3.Player.SetBounds(250, 480, 600, 400)

    '    Else
    '        For Each m In MediaHandlers
    '            m.Picture.Dock = DockStyle.Fill
    '            m.Player.Dock = DockStyle.Fill
    '        Next

    '    End If
    'End Sub
    Public Sub DockMedias(separated As Boolean)
        If separated Then
            Dim i As Int16 = 0
            Dim j As Int16 = 0
            For Each m In MediaHandlers
                Dim x As Int16
                Dim y As Int16
                Select Case i
                    Case 0
                        x = 0
                        y = 0
                    Case 1
                        x = 650
                        y = 0
                    Case 2
                        x = 250
                        y = 480
                End Select

                m.Picture.Dock = DockStyle.None
                m.Player.Dock = DockStyle.None
                m.Textbox.Dock = DockStyle.None
                m.Textbox.SetBounds(x, y, 600, 400)
                m.Picture.SetBounds(x, y, 600, 400)
                m.Player.SetBounds(x, y, 600, 400)
                i += 1
                'm.Player.SendToBack()
                'm.PicHandler.PicBox.SendToBack()

                'm.Textbox.SendToBack()

            Next
            Outliner.Visible = True

        Else
            For Each m In MediaHandlers
                m.Picture.Dock = DockStyle.Fill
                m.Player.Dock = DockStyle.Fill
                ' m.Textbox.Dock = DockStyle.Fill
            Next

        End If
    End Sub
#End Region

End Class
