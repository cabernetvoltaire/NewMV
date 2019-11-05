Imports AxWMPLib
Public Class MediaHandler
#Region "Members"

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event StartChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event SpeedChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaNotFound(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaPlaying(ByVal sender As Object, ByVal e As EventArgs)
    Private WithEvents Sound As New AxWindowsMediaPlayer
    Private Property mSndH As New SoundController 'With {.SoundPlayer = Sound, .CurrentPlayer = Player}
    Public IsCurrent As Boolean = False

    Private WithEvents ResetPosition As New Timer With {.Interval=10}
    Public WithEvents PositionUpdater As New Timer With {.Interval = 10}
    Public WithEvents ResetPositionCanceller As New Timer With {.Interval = 150000}

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
            mPlayPosition = mPlayer.Ctlcontrols.currentPosition
            Return mPlayPosition
        End Get
        Set(ByVal value As Long)
            mPlayPosition = value
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
            If Speed.Paused Then

                Speed.PausedPosition = mPlayPosition
                '   Speed.Paused = False

            End If
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
            StartPoint.Markers = mMarkers
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
                '    'Path hasn't changed, so nothing to do. 
            ElseIf value = "" Then
                'Deals with absent file
                mMediaPath = DefaultFile
                mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                RaiseEvent MediaChanged(Me, New EventArgs)
            Else
                mMediaPath = value
                mType = FindType(value)
                Try
                    Dim f As New IO.FileInfo(value)
                    If f.Exists Then
                        If mType = Filetype.Link Then
                            mIsLink = True
                            mLinkPath = LinkTarget(f.FullName)
                        Else
                            mIsLink = False
                            mLinkPath = ""
                        End If
                        If Not MainForm.Initialising Then LoadMedia()

                        mMediaDirectory = f.Directory.FullName
                    Else
                        mMediaPath = DefaultFile
                        mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                        If Not MainForm.Initialising Then LoadMedia()
                    End If
                Catch ex As Exception

                End Try

                RaiseEvent MediaChanged(Me, New EventArgs)

            End If


        End Set
    End Property
    Public Sub New(Nomen As String)
        '      Player = mPlayer
        Name = Nomen
        ' PositionUpdater.Interval = 500
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

    Private Sub GetBookmark()
        If InStr(mMediaPath, "%") <> 0 Then

            Dim s As String()
            s = mMediaPath.Split("%")
            mBookmark = Val(s(1))
        Else
            mBookmark = -1
        End If

    End Sub
    Public Function UpdateBookmark(path As String, time As Long) As String
        If Right(path, 4) <> ".lnk" Then
            Return path
            Exit Function
        End If
        If InStr(path, "%") <> 0 Then
            Dim m() As String = path.Split("%")
            path = m(0) & "%" & time & "%" & m(m.Length - 1)
        Else
            path = path.Replace(".lnk", "%" & Str(time) & "%.lnk")
        End If
        mMediaPath = path
        Return path

    End Function
    Private Property mlinkcounter As Integer
#Region "Linkhandling"

    Public Property LinkCounter As Integer
        Get
            Return mlinkcounter
        End Get
        Set
            mlinkcounter = Value
        End Set
    End Property

    Public Function IncrementLinkCounter(Forward As Boolean) As Integer
        If mMarkers.Count = 0 Then
            Return 0
            Exit Function
        End If
        If Forward Then
            mlinkcounter += 1
        Else
            mlinkcounter -= 1
        End If
        If mlinkcounter < 0 Then
            mlinkcounter = mlinkcounter + mMarkers.Count
        End If
        mlinkcounter = mlinkcounter Mod (mMarkers.Count)

        Return mlinkcounter
    End Function
    Public Function RandomCounter() As Integer
        Static done As New List(Of Integer)

        If mMarkers.Count = 0 Then
            Return -1
            Exit Function
        End If
        Dim ret As Integer
        ret = Int(Rnd() * mMarkers.Count)
        While done.Contains(ret) And done.Count < mMarkers.Count
            ret = Int(Rnd() * mMarkers.Count)
        End While
        If done.Count = mMarkers.Count Then
            done.Clear()
            ret = Int(Rnd() * mMarkers.Count)

        End If
        done.Add(ret)
        Return ret
    End Function
    Public Function FindNearestCounter(Last As Boolean) As Integer
        If mMarkers.Count = 0 Then
            Return -1
            Exit Function
        End If
        Dim ret As Integer
        Dim i As Integer = 0
        For i = 0 To mMarkers.Count - 1
            If Last Then
                If mMarkers(i) > Position Then
                    ret = i - 1
                    Return ret
                End If
            Else
                If mMarkers(i) > Position Then
                    ret = i
                    Return ret
                End If
            End If
        Next
        Return 0

    End Function
    Public Sub SetLink(Optional num = 0)
        If num = 0 And mMarkers.Count > 0 Then
            'mLinkCounter = mMarkers.Count - 1
            mlinkcounter = 0
        Else
            mlinkcounter = num

        End If

    End Sub
#End Region
    Public Sub MediaJumpToMarker(Optional ToEnd As Boolean = False, Optional ToMarker As Boolean = False)
        'There are markers, so jump to the next one
        If ToEnd Then 'Special Case where jump to end button pressed
            Dim m As New StartPointHandler With {
                        .Duration = mDuration,
                        .State = StartPointHandler.StartTypes.NearEnd
                                    }
            mPlayPosition = m.StartPoint
            'Or it's a link with a bookmark
        ElseIf mBookmark > -1 And Speed.PausedPosition = 0 Then 'And mMarkers.Count = 0 Then
            If StartPoint.State = StartPointHandler.StartTypes.FirstMarker Then
                mPlayPosition = mBookmark
            Else
                mPlayPosition = StartPoint.StartPoint
            End If
        ElseIf Speed.PausedPosition <> 0 Then
            mPlayPosition = Speed.PausedPosition
            Speed.PausedPosition = 0
        ElseIf mMarkers.Count <> 0 AndAlso (ToMarker Or StartPoint.State = StartPointHandler.StartTypes.FirstMarker) Then
            StartPoint.Absolute = mMarkers.Item(mlinkcounter)
            mPlayPosition = StartPoint.StartPoint
        Else
            mPlayPosition = StartPoint.StartPoint
        End If
        If mPlayPosition > mDuration Then
            Report(mPlayPosition & "Over-reach" & mDuration, 0, False)
        Else
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
            'Sound.Ctlcontrols.currentPosition = mPlayPosition
        End If
        Report("Start point state is:" & StartPoint.Descriptions(StartPoint.State) & " and has been set to " & StartPoint.StartPoint & "(" & mPlayPosition & ") with a duration of " & StartPoint.Duration, 1)

    End Sub
    Private Sub LoadMedia()

        Select Case mType
            Case Filetype.Doc

            Case Filetype.Link
                Select Case FindType(mLinkPath)
                    Case Filetype.Movie

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

                    '     Sound.URL = URL
                    LastURL = URL 'Prevents reloading into a given player
                Catch EX As Exception
                    Debug.Print(EX.Message)
                End Try

            End If
            MediaJumpToMarker(ToMarker:=True) 'Jump to a new position
        Else
            mlinkcounter = 0
            GetBookmark()
            MediaJumpToMarker() 'Jump to a new position
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
    Private Sub Uhoh() Handles mPlayer.ErrorEvent
        ' MsgBox("Error in MediaPlayer")
    End Sub

    Private mResetCounter As Integer

    Private Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles mPlayer.PlayStateChange
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsStopped
                mPlayer.Ctlcontrols.play()
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If MainForm.tmrAutoTrail.Enabled = False AndAlso mPlayer.Equals(Media.Player) AndAlso mType = Filetype.Movie AndAlso Not LoopMovie Then
                    MainForm.AdvanceFile(True, False)
                Else
                    MediaJumpToMarker()
                End If

            Case WMPLib.WMPPlayState.wmppsPlaying
                mSndH.Slow = False
                PositionUpdater.Enabled = True
                ResetPositionCanceller.Enabled = True
                ' mResetCounter = 0
                mDuration = mPlayer.currentMedia.duration
                StartPoint.Duration = mDuration
                MediaJumpToMarker()
                RaiseEvent MediaPlaying(Me, Nothing)
                ' MainForm.DrawScrubberMarks()

                If FullScreen.Changing Or Speed.Paused Then 'Hold current position if switching to FS or back. 
                    mPlayPosition = Speed.PausedPosition
                    Speed.Paused = False
                    Speed.PausedPosition = 0
                End If
            Case WMPLib.WMPPlayState.wmppsPaused ', WMPLib.WMPPlayState.wmppsTransitioning
                If Not Speed.Fullspeed Then
                    mSndH.Slow = True
                End If

            Case Else

        End Select
    End Sub
    Private Sub OnSpeedChange(sender As Object, e As EventArgs) Handles Speed.SpeedChanged

        RaiseEvent SpeedChanged(sender, e)

    End Sub
    Private Sub OnStartChange(sender As Object, e As EventArgs) Handles StartPoint.StartPointChanged, StartPoint.StateChanged

        RaiseEvent StartChanged(sender, e)

    End Sub
    ''' <summary>
    ''' Ensures mPlayPosition is always up to date (to within interval)
    ''' </summary>
    Private Sub UpdatePosition() Handles PositionUpdater.Tick
        Exit Sub
        Try
            If Speed.Paused Then

            Else
                mPlayPosition = mPlayer.Ctlcontrols.currentPosition 'TODO Buggering up startpoint?
                mDuration = mPlayer.currentMedia.duration
            End If
        Catch ex As Exception
        End Try

    End Sub
    Private Sub ResetPos() Handles ResetPosition.Tick
        ' PositionUpdater.Enabled = False
        Try
            '  mPlayer.Ctlcontrols.currentPosition = StartPoint.StartPoint
            MediaJumpToMarker()
        Catch ex As Exception

        End Try

    End Sub
    Public Sub PlaceResetter(ResetOn As Boolean)
        ResetPosition.Enabled = ResetOn
    End Sub
    Private Sub PositionUpdaterCanceller_Tick(sender As Object, e As EventArgs) Handles ResetPositionCanceller.Tick
        ResetPosition.Enabled = False
    End Sub


#End Region

End Class