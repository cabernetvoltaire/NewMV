Imports AxWMPLib
Public Class NewMediaSwapper

    Public WithEvents NextF As New NextFile
    Private Property mMediaHandlers As New List(Of MediaHandler)
    Public Property MediaHandlers As List(Of MediaHandler)
        Set(value As List(Of MediaHandler))
            mMediaHandlers = value
        End Set
        Get
            Return mMediaHandlers
        End Get
    End Property

    Private Outliner As New PictureBox With {.BackColor = Color.HotPink}
    Private mRandomNext As Boolean = False
    Public WriteOnly Property RandomNext As Boolean
        Set(value As Boolean)
            mRandomNext = value
        End Set
    End Property

    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListcount As Integer

    Public Event MediaPlaying(sender As Object, e As EventArgs)
    Public Event MediaNotFound(sender As Object, e As EventArgs)

    Public Property NextItem As String
    Public Property CurrentItem As String
    Public Property PreviousItem As String
    Public Event MediaShown(sender As Object, e As EventArgs)
#Region "Properties"
    Public Sub New(MHList As List(Of MediaHandler))
        mMediaHandlers = MHList
    End Sub
#End Region
#Region "Methods"

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
        MH.IsCurrent = False
    End Function

    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)
        'CurrentURLS.Clear()
        ' HideMedias(ThisMH)

        Prepare(ThisMH, CurrentItem)
        ThisMH.IsCurrent = True
        ThisMH.Activate()
        Prepare(PrevMH, PreviousItem)
        Prepare(NextMH, NextItem)
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
