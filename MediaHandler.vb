Imports AxWMPLib
Public Class MediaHandler
#Region "Members"

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event StartChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event SpeedChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaNotFound(ByVal sender As Object, ByVal e As EventArgs)
    Private WithEvents Sound As New AxWindowsMediaPlayer
    Private Property mSndH As New SoundController 'With {.SoundPlayer = Sound, .CurrentPlayer = Player}


    Private WithEvents ResetPosition As New Timer
    Public WithEvents PositionUpdater As New Timer
    Private DefaultFile As String = "C:\exiftools.exe"
    Public WithEvents StartPoint As New StartPointHandler
    Public WithEvents Speed As New SpeedHandler
    Public DisplayerName As String
    Private mLoop As Boolean = False
    Private mType As Filetype
    Public Name As String = Me.Name
    Public Property MediaType() As Filetype
        Get
            If mMediaPath = "" Then
            Else

                mType = FindType(mMediaPath)
            End If
            If mType = Filetype.Link Then
                mIsLink = True
                IsLink = True
                mLinkPath = LinkTarget(mMediaPath)
                mType = FindType(mLinkPath)

            Else
                mIsLink = False
            End If
            Return mType
        End Get
        Set(ByVal value As Filetype)
            mType = value
        End Set
    End Property

    Private mPicBox As New PictureBox
    Public Property Picture() As PictureBox
        Get
            Return mPicBox
        End Get
        Set(ByVal value As PictureBox)
            mPicBox = value
        End Set
    End Property
    Public Property LoopMovie As Boolean
        Get
            Return mLoop
        End Get
        Set(value As Boolean)
            mLoop = value
        End Set
    End Property

    Private WithEvents mPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Public Property Player() As AxWMPLib.AxWindowsMediaPlayer
        Get
            Return mPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            mPlayer = value

            mPlayer.settings.enableErrorDialogs = False
            mPlayer.settings.autoStart = True
        End Set
    End Property
    Private mPlayPosition As Long
    Public Property Position() As Long
        Get
            '   mPlayPosition = mPlayer.Ctlcontrols.currentPosition
            Return mPlayPosition
        End Get
        Set(ByVal value As Long)
            mPlayPosition = value
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
            Debug.Print("Position was set to " & mPlayPosition)
        End Set
    End Property

    Private mDuration As Long
    Public Property Duration() As Long
        Get
            ' mDuration = mPlayer.currentMedia.duration
            Return mDuration
        End Get
        Set(value As Long)
            mDuration = value
            StartPoint.Duration = value
        End Set
    End Property
    Private mFrameRate As Int32
    Public Property FrameRate() As Int32
        Get
            Return mFrameRate
        End Get
        Set(ByVal value As Int32)
            mFrameRate = mPlayer.network.frameRate
        End Set
    End Property
    Private mMarkers As New List(Of Long)
    Public Property Markers() As List(Of Long)
        Get
            Return mMarkers
        End Get
        Set(ByVal value As List(Of Long))
            mMarkers = value
        End Set
    End Property
    Private mMediaPath As String
    Public Property MediaPath() As String
        Get
            Return mMediaPath
        End Get
        Set(ByVal value As String)
            'If path changes, we need to check it exists, and if so, change stored directory as well, 
            'And raise a media changed event. 
            Dim b As String = mMediaPath
            If value = b Then
                'Path hasn't changed, so nothing to do. 
            ElseIf value = "" Then
                'Deals with absent file
                mMediaPath = DefaultFile
                mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                RaiseEvent MediaChanged(Me, New EventArgs)
            Else
                mMediaPath = value
                mType = FindType(value)
                Dim f As New IO.FileInfo(value)
                If f.Exists Then
                    If mType = Filetype.Link Then
                        mIsLink = True
                        mLinkPath = LinkTarget(f.FullName)
                    Else
                        mIsLink = False
                        mLinkPath = ""
                    End If
                    LoadMedia()

                    mMediaDirectory = f.Directory.FullName
                Else
                    mMediaPath = DefaultFile
                    mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                End If

                RaiseEvent MediaChanged(Me, New EventArgs)
            End If


        End Set
    End Property
    Public Sub New(value As String)
        '      Player = mPlayer
        Name = value
        PositionUpdater.Interval = 500
        PositionUpdater.Enabled = False
        ResetPosition.Interval = 1000
        'mSndH.SoundPlayer = Sound
        'mSndH.CurrentPlayer = Player
        '     StartPoint = Media.StartPoint
    End Sub
    Private mMediaDirectory As String
    Public ReadOnly Property MediaDirectory() As String
        Get
            Return mMediaDirectory
        End Get
    End Property

    Private mIsLink As Boolean = False
    Public Property IsLink() As Boolean
        Get
            Return mIsLink
        End Get
        Set(ByVal value As Boolean)
            mIsLink = value
            If mIsLink Then
                GetBookmark()
            Else
                mBookmark = -1
                mLinkPath = ""
            End If
        End Set
    End Property
    Private mBookmark As Long = -1
    Public Property Bookmark() As Long
        Get
            Return mBookmark
        End Get
        Set(ByVal value As Long)
            mBookmark = value
        End Set
    End Property

    Private mLinkPath As String

    Public ReadOnly Property LinkPath() As String

        Get
            Return mLinkPath
        End Get

    End Property
#End Region

#Region "Methods"
    Public Sub Pause(Pause As Boolean)
        Exit Sub
        If Pause Then
            mPlayer.Ctlcontrols.pause()
            Speed.Paused = True
            Speed.Fullspeed = False
        Else
            mPlayer.Ctlcontrols.play()
            Speed.Fullspeed = True
            Speed.Paused = False
        End If

    End Sub

    Private Sub GetBookmark()
        If InStr(mMediaPath, "%") <> 0 Then

            Dim s As String()
            s = mMediaPath.Split("%")
            mBookmark = Val(s(1))
        Else
            mBookmark = -1
        End If

    End Sub
    Public Function UpdateBookmark(path As String, time As String) As String
        If Right(path, 4) <> ".lnk" Then
            Return path
            Exit Function
        End If
        If InStr(path, "%") <> 0 Then
            Dim m() As String = path.Split("%")
            path = m(0) & "%" & time & "%" & m(m.Length - 1)
        Else
            path = path.Replace(".lnk", "%" & time & "%.lnk")
        End If
        mMediaPath = path
        Return path

    End Function

    Private mLinkCounter As Integer = 0
    Public Function IncrementLinkCounter(Forward As Boolean) As Integer
        If mMarkers.Count = 0 Then Exit Function
        If Forward Then
            mLinkCounter += 1
        Else
            mLinkCounter -= 1
        End If
        If mLinkCounter < 0 Then
            mLinkCounter = mLinkCounter + mMarkers.Count
        End If
        mLinkCounter = mLinkCounter Mod (mMarkers.Count)
        Return mLinkCounter
    End Function
    Public Sub MediaJumpToMarker(Optional ToEnd As Boolean = False)
        'It's a link with a bookmark
        If mBookmark > -1 And Speed.PausedPosition = 0 Then 'And mMarkers.Count = 0 Then
            If StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute Then
                mPlayPosition = mBookmark
            Else
                mPlayPosition = StartPoint.StartPoint
            End If
        Else
            'Not a link
            If ToEnd Then
                Dim m As New StartPointHandler With {
                        .Duration = mDuration,
                        .State = StartPointHandler.StartTypes.NearEnd
                                    }
                mPlayPosition = m.StartPoint
            Else
                If Speed.PausedPosition <> 0 Then
                    mPlayPosition = Speed.PausedPosition
                Else
                    If mMarkers.Count <> 0 AndAlso StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute Then
                        Try
                            mPlayPosition = mMarkers.Item(mLinkCounter)
                        Catch ex As Exception
                            mPlayPosition = StartPoint.StartPoint
                        End Try
                    Else
                        mPlayPosition = StartPoint.StartPoint
                    End If
                End If
            End If
        End If
        If mPlayPosition > mDuration Then
            Report(mPlayPosition & "Over-reach" & mDuration, 2, False)
        Else

            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
            'Sound.Ctlcontrols.currentPosition = mPlayPosition
        End If
        Debug.Print("MediaJumpMarker set" & mMediaPath & " position to " & mPlayPosition)
    End Sub
    Private Sub LoadMedia()

        Select Case mType
            Case Filetype.Doc

            Case Filetype.Link
                Select Case FindType(mLinkPath)
                    Case Filetype.Movie
                        If mBookmark > -1 Then
                            If StartPoint.State = StartPointHandler.StartTypes.ParticularAbsolute Then

                                StartPoint.Absolute = mBookmark
                            End If
                        End If
                        HandleMovie(mLinkPath)
                    Case Filetype.Pic
                        HandlePic(mLinkPath)
                End Select
            Case Filetype.Movie
                HandleMovie(mMediaPath)
            Case Filetype.Pic
                HandlePic(mMediaPath)
            Case Filetype.Unknown

                Exit Sub
        End Select
    End Sub
    Private Sub HandleMovie(URL As String)

        Static LastURL As String
        If URL <> LastURL Then
            If mPlayer Is Nothing Then
            Else
                Try
                    mPlayer.URL = URL
                    Sound.URL = URL
                    LastURL = URL
                Catch EX As Exception

                End Try

            End If
        Else
            GetBookmark()
            MediaJumpToMarker()
        End If
        DisplayerName = mPlayer.Name
    End Sub
    Public Sub HandlePic(path As String)

        Dim img As Image
        If Not Picture.Image Is Nothing Then
            DisposePic(Picture)
        End If
        img = GetImage(path)
        If img Is Nothing Then
            Exit Sub
        End If
        MainForm.OrientPic(img)
        'Resume if in middle of slideshow
        'If blnRestartSlideShowFlag Then
        '    tmrSlideShow.Enabled = True
        '    blnRestartSlideShowFlag = False
        'End If
        currentPicBox = mPicBox
        MainForm.MovietoPic(img)
        DisplayerName = mPicBox.Name
    End Sub
    Public Sub MovietoPic(img As Image)
        PreparePic(currentPicBox, img)
        'SndH.Muted = True

    End Sub




    Public Sub OrientPic(img As Image)
        Select Case ImageOrientation(img)
            Case ExifOrientations.BottomRight
                img.RotateFlip(RotateFlipType.Rotate180FlipNone)
            Case ExifOrientations.RightTop
                img.RotateFlip(RotateFlipType.Rotate90FlipNone)
            Case ExifOrientations.LeftBottom
                img.RotateFlip(RotateFlipType.Rotate270FlipNone)

        End Select
    End Sub

#End Region

#Region "Event Handlers"
    Private Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles mPlayer.PlayStateChange
        'Dim lbl As New Label
        'lbl = MainForm.lblNavigateState
        'Select Case e.newState

        '    Case 0 ' Undefined
        '        lbl.Text = lbl.Text & vbCrLf & "Undefined"

        '    Case 1 ' Stopped
        '        lbl.Text = lbl.Text & vbCrLf & "Stopped"

        '    Case 2 ' Paused
        '        lbl.Text = lbl.Text & vbCrLf & "Paused"

        '    Case 3 ' Playing
        '        lbl.Text = lbl.Text & vbCrLf & "Playing"

        '    Case 4 ' ScanForward
        '        lbl.Text = lbl.Text & vbCrLf & "ScanForward"

        '    Case 5 ' ScanReverse
        '        lbl.Text = lbl.Text & vbCrLf & "ScanReverse"

        '    Case 6 ' Buffering
        '        lbl.Text = lbl.Text & vbCrLf & "Buffering"

        '    Case 7 ' Waiting
        '        lbl.Text = lbl.Text & vbCrLf & "Waiting"

        '    Case 8 ' MediaEnded
        '        lbl.Text = lbl.Text & vbCrLf & "MediaEnded"

        '    Case 9 ' Transitioning
        '        lbl.Text = lbl.Text & vbCrLf & "Transitioning"

        '    Case 10 ' Ready
        '        lbl.Text = lbl.Text & vbCrLf & "Ready"

        '    Case 11 ' Reconnecting
        '        lbl.Text = lbl.Text & vbCrLf & "Reconnecting"

        '    Case 12 ' Last
        '        lbl.Text = lbl.Text & vbCrLf & "Last"

        '    Case Else
        '        lbl.Text = lbl.Text & vbCrLf & ("Unknown State: " + e.newState.ToString())

        'End Select
        Debug.Print(e.newState.ToString())
        Select Case e.newState

            Case WMPLib.WMPPlayState.wmppsStopped
                mPlayer.Ctlcontrols.play()

            Case WMPLib.WMPPlayState.wmppsMediaEnded
                'Debug.Print("Ended:" & StartPoint.StartPoint & " " & StartPoint.Duration)
                '  If mLoop Then
                ' mPlayer.Ctlcontrols.play()
                'End If
                Debug.Print(mPlayer.URL & ", duration " & mDuration & ", playposition" & mPlayPosition & " STOPPED")
                If MainForm.tmrAutoTrail.Enabled = False AndAlso mPlayer.Equals(Media.Player) AndAlso mType = Filetype.Movie AndAlso Not LoopMovie Then
                    MainForm.AdvanceFile(True, False)
                Else
                    MediaJumpToMarker()
                    ' mPlayer.Ctlcontrols.play()

                End If
            Case WMPLib.WMPPlayState.wmppsPlaying
                'ReportTime("Playing")
                mSndH.Slow = False
                PositionUpdater.Enabled = True
                Duration = mPlayer.currentMedia.duration
                StartPoint.Duration = mPlayer.currentMedia.duration
                Report("Duration:" & mDuration & vbCrLf & "Startpoint:" & StartPoint.StartPoint, 2)
                MediaJumpToMarker()
                Debug.Print(mPlayer.URL & " ` playstatehandler")

                If FullScreen.Changing Or Speed.Unpause Then 'Hold current position if switching to FS or back. 
                    mPlayPosition = Speed.PausedPosition
                End If
            Case WMPLib.WMPPlayState.wmppsPaused ', WMPLib.WMPPlayState.wmppsTransitioning
                If Not Speed.Fullspeed Then
                    mSndH.Slow = True
                Else

                End If
            Case Else

        End Select
    End Sub
    Private Sub OnSpeedChange(sender As Object, e As EventArgs) Handles Speed.SpeedChanged
        MainForm.OnSpeedChange(sender, e)

    End Sub

    Private Sub OnStartChange(sender As Object, e As EventArgs) Handles StartPoint.StartPointChanged, StartPoint.StateChanged
        '  MediaJumpToMarker()
        '  reportStartpoint(Me)
        RaiseEvent StartChanged(sender, e)

    End Sub

    Private Sub UpdatePosition() Handles PositionUpdater.Tick
        Try
            mPlayPosition = mPlayer.Ctlcontrols.currentPosition
            Duration = mPlayer.currentMedia.duration

        Catch ex As Exception

        End Try

    End Sub
    Private Sub ResetPos() Handles ResetPosition.Tick
        Try
            MediaJumpToMarker()
        Catch ex As Exception

        End Try

    End Sub
    Public Sub PlaceResetter(ResetOn As Boolean)
        ResetPosition.Enabled = ResetOn
        'Speed.Paused = ResetOn
        Exit Sub
        If ResetOn Then
            MediaJumpToMarker()
            mPlayer.Ctlcontrols.pause()
        Else
            mPlayer.Ctlcontrols.play()


        End If

    End Sub



#End Region

End Class