Imports AxWMPLib
Imports System.Text.UTF8Encoding
Imports System.IO

Public Class MediaHandler
#Region "Members"

    Public Event MediaFinished(ByVal sender As Object, ByVal e As EventArgs)
    Public Event StartChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event SpeedChanged(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaNotFound(ByVal sender As Object, ByVal e As EventArgs)
    Public Event MediaPlaying(ByVal sender As Object, ByVal e As EventArgs)
    Public Event Zoomchanged(ByVal sender As Object, ByVal e As EventArgs)

    Private WithEvents mPlayer As New AxWMPLib.AxWindowsMediaPlayer
    Private WithEvents Sound As New AxWindowsMediaPlayer
    ' Private Property mSndH As New SoundController(mPlayer, Sound) ' With {.SoundPlayer = Sound, .CurrentPlayer = mPlayer}
    Public Property Textbox As New RichTextBox With {
        .Multiline = True,
        .Dock = DockStyle.Fill,
        .BorderStyle = BorderStyle.None,
        .ScrollBars = ScrollBars.Both, .WordWrap = True,
        .Font = New Font("Calibri", 16, FontStyle.Regular),
        .BackColor = Color.Black,
        .ForeColor = Color.AntiqueWhite}
    Private MetaDirectory As IReadOnlyList(Of MetadataExtractor.Directory)
    Public ShowMeta As Boolean = False
    Public IsCurrent As Boolean = False

    Private WithEvents ResetPosition As New Timer With {.Interval = 1000} 'Changing can affect loading
    Public WithEvents PositionUpdater As New Timer With {.Interval = 100} 'Too short causes a crash on exiting.
    Public WithEvents ResetPositionCanceller As New Timer With {.Interval = 30000}
    Public WithEvents PicHandler As New PictureHandler(Picture)
    Public Metadata As String = ""

    Private ReadOnly DefaultFile As String = "C:\exiftools.exe"
    Public WithEvents SPT As New StartPointHandler
    Public WithEvents Speed As New SpeedHandler
    Public DisplayerName As String

    Public Autotrail As Boolean
    Public Playing As Boolean
    Public DontLoad As Boolean

    Private mLoop As Boolean = False
    Private mType As Filetype
    Public Name As String = Me.Name
    Public WriteOnly Property Visible As Boolean
        Set(value As Boolean)
            mPlayer.Visible = value
            Textbox.Visible = value
            PicHandler.PicBox.Visible = value

        End Set
    End Property


    Public Property MediaType() As Filetype
        Get
            If mMediaPath = "" Then
            Else

                mType = FindType(mMediaPath)
            End If
            mIsLink = False
            Select Case mType
                Case Filetype.Link
                    mIsLink = True
                    '  IsLink = True
                    mLinkPath = LinkTarget(mMediaPath)
                    mType = FindType(mLinkPath)
                    GetBookmark()
                Case Filetype.Doc


                Case Else
                    mLinkPath = ""
                    mBookmark = -1
            End Select
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
            PicHandler.PicBox = value
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

    Public Property Player() As AxWMPLib.AxWindowsMediaPlayer
        Get

            Return mPlayer
        End Get
        Set(ByVal value As AxWMPLib.AxWindowsMediaPlayer)
            mPlayer = value
            'mSndH.CurrentPlayer = value
            mPlayer.settings.enableErrorDialogs = False
            mPlayer.settings.autoStart = True
        End Set
    End Property
    Private mPlayPosition As Double
    Public Property Position() As Double
        Get
            If mType = Filetype.Movie Then

                mPlayPosition = mPlayer.Ctlcontrols.currentPosition
                Return mPlayPosition
            Else
                Return 0
            End If
        End Get
        Set(ByVal value As Double)
            mPlayPosition = value
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
            'If Speed.Paused Then
            '    Speed.PausedPosition = mPlayPosition
            'End If
            Debug.Print("Position was set to " & mPlayPosition)
        End Set
    End Property

    Private mDuration As Double
    Public Property Duration() As Double
        Get
            If mDuration = 0 Then
                Report("Duration not set yet", 0)
            End If
            '   mDuration = mPlayer.currentMedia.duration
            Return mDuration
        End Get
        Set(value As Double)
            mDuration = value
            SPT.Duration = value
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
    Private mMarkers As New List(Of Double)
    Public Property Markers() As List(Of Double)
        Get
            Return mMarkers
        End Get
        Set(ByVal value As List(Of Double))
            mMarkers = value
            SPT.Markers = mMarkers
        End Set
    End Property
    Private mMediaPath As String
    ''' <summary>
    ''' Load the medium into the appropriate control unless it's already in it.
    ''' </summary>
    ''' <returns></returns>
    Public Property MediaPath() As String
        Get
            Return mMediaPath
        End Get
        Set(ByVal value As String)
            'If path changes, we need to check it exists, and if so, change stored directory as well, 
            'And raise a media changed event. 

            'TODO Fix this
            If value = "" Then
                'Deals with absent file
                'mMediaPath = DefaultFile
                'mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                'RaiseEvent MediaChanged(Me, New EventArgs)

            Else

                mMediaPath = value
                mType = FindType(value)
                Dim f As New IO.FileInfo(value)
                If f.Exists Then
                    If mType = Filetype.Link Then
                        mIsLink = True
                        mLinkPath = LinkTarget(f.FullName)
                        ' MsgBox(mLinkPath)
                        If mLinkPath = "" Or mLinkPath.Contains(vbNullChar) Then
                            mLinkPath = mMediaPath
                            MsgBox("Destination not found")
                        End If
                    Else
                        mIsLink = False
                        mLinkPath = ""
                    End If
                    If Not DontLoad Then LoadMedia()
                    mMediaDirectory = f.Directory.FullName
                Else
                    mMediaPath = DefaultFile
                    mMediaDirectory = New IO.FileInfo(mMediaPath).Directory.FullName
                    If Not DontLoad Then LoadMedia()
                End If

                mMarkers = GetMarkersFromLinkList()
                mMarkers.Sort()
                'RaiseEvent MediaChanged(Me, New EventArgs)

            End If

        End Set
    End Property



    Public Function GetMarkersFromLinkList() As List(Of Double)
        Dim List As New List(Of String)
        Dim markerslist As New List(Of Double)

        If mIsLink Then
            List = AllFaveMinder.GetLinksOf(mLinkPath)

        Else
            List = AllFaveMinder.GetLinksOf(mMediaPath)
        End If
        Dim i = 0
        For Each m In List
            Dim n = BookmarkFromLinkName(m)
            If n > 0 Then
                If markerslist.Contains(n) Then
                Else
                    markerslist.Add(n)
                End If
                i += 1
            End If
        Next
        Return markerslist
    End Function

    Public Sub New(Nomen As String)
        Name = Nomen
        PicHandler.PicBox = mPicBox
        PositionUpdater.Enabled = False
        'mSndH.SoundPlayer = Sound
        'mSndH.CurrentPlayer = Player
        '     StartPoint = Media.StartPoint
    End Sub
    Public Sub New(Nomen As String, mplayer As AxWindowsMediaPlayer, pic As PictureBox)
        Player = mplayer
        Name = Nomen
        PositionUpdater.Enabled = False
        mPicBox = pic
        PicHandler.PicBox = mPicBox
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

        End Set
    End Property
    Private mBookmark As Double = -1
    Public Property Bookmark() As Double
        Get
            Return mBookmark
        End Get
        Set(ByVal value As Double)
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

#Region "BookmarkMethods"
    Private Sub GetBookmark()
        If InStr(mMediaPath, "%") <> 0 Then

            Dim s As String()
            s = mMediaPath.Split("%")
            mBookmark = Val(s(s.Length - 2))
        Else
            mBookmark = -1
        End If

    End Sub
    Public Function UpdateBookmark(path As String, time As Double) As String
        If Right(path, 4) <> LinkExt Then
            Return path
            Exit Function
        End If
        If InStr(path, "%") <> 0 Then
            Dim m() As String = path.Split("%")
            path = m(0) & "%" & time & "%" & m(m.Length - 1)
        Else
            path = path.Replace(LinkExt, "%" & Str(time) & "%." & LinkExt)
        End If
        mMediaPath = path
        Return path

    End Function
#End Region
    Public Function Activate()
        Select Case Me.MediaType
            Case Filetype.Movie

            Case Filetype.Link
            Case Filetype.Doc

        End Select
    End Function
    Private Property mlinkcounter As Integer
#Region "Linkhandling"

    Public Property LinkCounter As Integer
        Get
            If mlinkcounter > mMarkers.Count - 1 Or mMarkers.Count = 1 Then
                mlinkcounter = 0
            End If
            Return mlinkcounter
        End Get
        Set
            mlinkcounter = Value
        End Set
    End Property

    Public Function IncrementLinkCounter(Forward As Boolean) As Integer

        'Get current position
        Dim p = mPlayPosition
        'mMarkers.Sort()
        Dim count = mMarkers.Count
        For i = 0 To count - 1 'Find preceding marker to position

            If mMarkers(i) < p + 1 Then
                mlinkcounter = i
            End If

        Next
        If Forward Then
            mlinkcounter += 1
            If mlinkcounter = count Then
                mlinkcounter = 0
            End If
        Else
            mlinkcounter -= 1
            If mlinkcounter = -1 Then mlinkcounter = count - 1
        End If
        If mlinkcounter > count - 1 Then mlinkcounter = 0

        Return mlinkcounter


    End Function
    Public Function RandomCounter() As Integer
        Static done As New List(Of Integer) 'Choose random counter, but don't repeat it. 

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
            If mMarkers(i) > Position Then
                If Last Then
                    ret = i - 1
                Else
                    ret = i
                End If
            End If
        Next
        Return ret

    End Function
    Public Function FindNextCounter(Last As Boolean) As Integer
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
#Region "Media Handlers"
    Public Sub LargeJump(e As KeyEventArgs, Small As Boolean, Forward As Boolean)
        If Playing Then

            Dim Extra As Integer = 1
            If Small Then Extra = 10
            If Forward Then
                Position = Math.Min(Duration, Position + Duration / (Extra * Speed.FractionalJump))
                If Position = Duration Then RaiseEvent MediaFinished(Me, Nothing)
            Else
                Position = Math.Min(Duration, Position - Duration / (Extra * Speed.FractionalJump))
            End If
        Else
            RaiseEvent MediaFinished(Me, Nothing)
        End If

    End Sub
    Public Sub SmallJump(e As KeyEventArgs, Small As Boolean, Forward As Boolean)
        Dim extra As Integer = 1
        If Small Then extra = 10
        If Forward Then
            Media.Position = Media.Position + Speed.AbsoluteJump / extra
        Else
            Media.Position = Media.Position - Speed.AbsoluteJump / extra
        End If

    End Sub
    Public Sub MediaJumpToMarker()
        If SPT.State = StartPointHandler.StartTypes.FirstMarker Then
            If mBookmark > -1 Then
                mPlayPosition = mBookmark
            Else
                If mMarkers.Count > 0 Then
                    SPT.Marker = mMarkers.Item(LinkCounter)
                Else
                End If
                mPlayPosition = SPT.StartPoint
            End If
        Else
            If Speed.PausedPosition <> 0 Then
                mPlayPosition = Speed.PausedPosition
            Else
                mPlayPosition = SPT.StartPoint
            End If
        End If
        mPlayer.Ctlcontrols.currentPosition = mPlayPosition

    End Sub
    Public Sub MediaJumpToMarker(Optional ToEnd As Boolean = False, Optional ToMarker As Boolean = False)
        'This is a logical mess.
        'There are markers, so jump to the next one
        If ToEnd Then 'Special Case where jump to end button pressed
            Dim m As New StartPointHandler With {
                        .Duration = mDuration,
                        .State = StartPointHandler.StartTypes.NearEnd
                                    }
            mPlayPosition = m.StartPoint
            m = Nothing
        End If

        'Or it's a link with a bookmark
        'ElseIf mIsLink And mBookmark > -1 AndAlso Speed.PausedPosition = 0 AndAlso Not ToMarker Then 'And mMarkers.Count = 0 Then
        '    If SPT.State = StartPointHandler.StartTypes.FirstMarker Then
        '        mPlayPosition = mBookmark
        '    Else
        '        mPlayPosition = SPT.StartPoint
        '    End If
        'ElseIf Speed.PausedPosition <> 0 Then
        '    mPlayPosition = Speed.PausedPosition
        '    'Speed.PausedPosition = 0
        If ToMarker Then
            If mMarkers.Count > 0 Then
                'If Not ToMarker Then mlinkcounter = 0 'Reset when it's a new file
                SPT.Marker = mMarkers.Item(LinkCounter)
            End If
            mPlayPosition = SPT.StartPoint
        End If
        If mPlayPosition > mDuration Then
            Report(mPlayPosition & "Over-reach" & mDuration, 0, False)
        Else
            mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        End If

    End Sub
    Public Sub Mute(Optional Off As Boolean = False)
        mPlayer.settings.mute = Not Off
    End Sub
    Public Sub LoadMedia()
        ' Mysettings.LastTimeSuccessful = False
        '        PreferencesSave()

        Playing = False
        Select Case mType
            Case Filetype.Doc
                HandleDoc(mMediaPath)

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
            Case Filetype.Browsable
            Case Filetype.Unknown

                'mPlayer.URL = ""
                'mPicBox.Dispose()
                Exit Sub
        End Select
        Try


        Catch ex As Exception

        End Try
    End Sub
    Private Sub HandleDoc(url As String)
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(url)
        Try
            Textbox.LoadFile(url)
        Catch ex As Exception
            Dim s As String = ExtractParagraphs(fileReader, Me)
            If s = "No" Then
            Else
                Textbox.Text = s

            End If

        End Try
        FormMain.ctrPicAndButtons.Panel1.Controls.Add(Textbox)
    End Sub
    Private Sub HandleMovie(URL As String)
        If URL.EndsWith("exe") Then BreakExecution()

        If URL <> mPlayer.URL Then
            If mPlayer Is Nothing Then
            Else
                Try

                    mPlayer.URL = URL
                    Sound.URL = URL

                Catch EX As Exception
                    'BreakExecution()
                    Debug.WriteLine(EX.Message)
                End Try

            End If
            MediaJumpToMarker() 'Jump to a new position
        Else
            mlinkcounter = 0
            GetBookmark()
            MediaJumpToMarker() 'Jump to a new position
        End If
        DisplayerName = mPlayer.Name
    End Sub
    Public Sub PlaceResetter(ResetOn As Boolean)
        ResetPosition.Enabled = ResetOn

    End Sub
    Public Sub HandlePic(path As String)
        '    If PicHandler.PicBox.Tag = path Then
        '    Else
        PicHandler.GetImage(path)
        If PicHandler.PicBox.Image Is Nothing Then
        Else
            If Not path.EndsWith(".gif") Then

                Metadata = ImageDate(PicHandler.PicBox.Image)
            End If
        End If
        DisplayerName = mPicBox.Name
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
    Function DisposeMovie() As Integer

        Try
            mPlayer.close()

        Catch ex As Exception
            Exit Function
        End Try
        Return 0
    End Function
    Public Sub CancelMedia()
        Select Case mType
            Case Filetype.Movie
                DisposeMovie()
            Case Filetype.Pic
                DisposePic(PicHandler.PicBox)

        End Select

    End Sub
#End Region
#End Region

#Region "Event Handlers"
    Public Sub OnPicZoomed(sender As Object, e As EventArgs) Handles PicHandler.ZoomChange, PicHandler.StateChanged
        RaiseEvent Zoomchanged(sender, e)
    End Sub

    Private Sub Uhoh() Handles mPlayer.ErrorEvent

        'MsgBox("Error in MediaPlayer")
    End Sub

    Private ReadOnly mResetCounter As Integer

    Private Sub PlaystateChange(sender As Object, e As _WMPOCXEvents_PlayStateChangeEvent) Handles mPlayer.PlayStateChange
        FormMain.SP = Speed
        Select Case e.newState
            Case WMPLib.WMPPlayState.wmppsStopped
                Try
                    mPlayer.Ctlcontrols.play()
                Catch ex As Exception
                End Try
            Case WMPLib.WMPPlayState.wmppsMediaEnded
                If Autotrail = False AndAlso mPlayer.Equals(Media.Player) AndAlso mType = Filetype.Movie AndAlso Not LoopMovie Then
                    Playing = False
                    RaiseEvent MediaFinished(sender, Nothing)
                Else
                    MediaJumpToMarker()
                End If
            Case WMPLib.WMPPlayState.wmppsPlaying

                'mSndH.Slow = False
                PositionUpdater.Enabled = True
                mDuration = mPlayer.currentMedia.duration

                SPT.Duration = mDuration
                'MediaJumpToMarker()
                Playing = True
                RaiseEvent MediaPlaying(Me, Nothing)
                If FullScreen.Changing Then 'Hold current position if switching to FS or back. 
                    mPlayPosition = Speed.PausedPosition
                    ' Speed.Paused = False
                    'Speed.PausedPosition = 0
                End If
            Case WMPLib.WMPPlayState.wmppsPaused ', WMPLib.WMPPlayState.wmppsTransitioning
                If Not Speed.Fullspeed Then
                    'mSndH.Slow = True
                End If

            Case Else
                PositionUpdater.Enabled = False
        End Select
        Speed = FormMain.SP
    End Sub

    Private Sub OnStartChange(sender As Object, e As EventArgs) Handles SPT.StartPointChanged, SPT.StateChanged

        RaiseEvent StartChanged(sender, e)

    End Sub
    ''' <summary>
    ''' Ensures mPlayPosition is always up to date (to within interval)
    ''' </summary>
    Private Sub UpdatePosition() Handles PositionUpdater.Tick
        If Speed.Paused Then
            'Don't want to change if media paused.
        Else

            Try
                mPlayPosition = mPlayer.Ctlcontrols.currentPosition 'TODO Buggering up startpoint?
                '   Speed.PausedPosition = mPlayPosition
                mDuration = mPlayer.currentMedia.duration
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub ResetPosition_Tick(sender As Object, e As EventArgs) Handles ResetPosition.Tick
        'Keeps resetting the position until ready to play.
        'mPlayer.Ctlcontrols.currentPosition = mPlayPosition
        MediaJumpToMarker()
        FormMain.DrawScrubberMarks()
    End Sub
    ''' <summary>
    ''' Only there to prevent Position resetting from starting indefinitely. 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ResetPositionCanceller_Tick(sender As Object, e As EventArgs) Handles ResetPositionCanceller.Tick
        Exit Sub
        ResetPosition.Enabled = False
        ResetPositionCanceller.Enabled = False
    End Sub

#End Region
    Private Function ImageDate(img As Image) As String
        If img Is Nothing Then
            Return ""
            Exit Function
        End If
        'Return MetaDirectory(2).Tags(2).ToString
        Exit Function



        Dim r As New System.Text.RegularExpressions.Regex(":")
        Dim text As String = ""
        Dim datetaken As String
        Dim m As Imaging.PropertyItem

        Try
            m = img.GetPropertyItem(36867)
            datetaken = r.Replace(UTF8.GetString(m.Value), "-", 2)
            text = DateTime.Parse(datetaken)

        Catch ex As Exception

        End Try


        Return text
    End Function

End Class
Public Class PicGrabber
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Structure RECT
        Dim Left As Long
        Dim Top As Long
        Dim Right As Long
        Dim Bottom As Long
    End Structure
    'Private Sub Form_Load()
    '    Dim hWndWMP As Long, hDCWMP As Long, rc As RECT
    '    hWndWMP = FindWindow("WMPlayerApp", "Windows Media Player")
    '    hDCWMP = GetWindowDC(hWndWMP)
    '    GetWindowRect(hWndWMP, rc)
    '    BitBlt(hDCWMP, 0, 0, rc.Right - rc.Left, rc.Bottom - rc.Top, hDCWMP, 0, 0)
    '    ReleaseDC(hWndWMP, hDCWMP)
    'End Sub
    Public Player As New AxWindowsMediaPlayer
    Public Sub GrabScreen()
        Using x As New Bitmap(100, 100)
            '        x = Player.DrawToBitmap(x, New Rectangle With {.Height = 1000, .Width = 100, .Location = New Point(0, 0)})
        End Using
    End Sub
End Class