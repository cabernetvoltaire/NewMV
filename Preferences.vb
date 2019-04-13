Public Class Preferences
    Private mHorSplit As Int64
    Public Property HorSplit() As Int64
        Get
            Return mHorSplit
        End Get
        Set(ByVal value As Int64)
            mHorSplit = value
        End Set
    End Property
    Private mVertSplit As Int64
    Public Property VertSplit() As Int64
        Get
            Return mVertSplit
        End Get
        Set(ByVal value As Int64)
            mVertSplit = value
        End Set
    End Property
    Private mLastFilePath As String
    Public Property LastFilePath() As String
        Get
            Return mLastFilePath
        End Get
        Set(ByVal value As String)
            mLastFilePath = value
        End Set
    End Property
    Private mFavesFolder As String
    Public Property FavesFolder() As String
        Get
            Return mFavesFolder
        End Get
        Set(ByVal value As String)
            mFavesFolder = value
        End Set
    End Property
    Private mOptionZoneSize As Decimal
    Public Property OptionZoneSize() As Decimal
        Get
            Return mOptionZoneSize
        End Get
        Set(ByVal value As Decimal)
            mOptionZoneSize = value
        End Set
    End Property
    Private Sub Preferences_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LayoutPreferences()
        'VertSplit
        'HorSplit

    End Sub
    Private Sub CurrentFilesandFolders()
        'Lastopenedfile
        'FavouritesFolder
        '
    End Sub
    Private Sub CurrentState()

    End Sub
    Private Sub Options()
        'ZoneSize
        'iSSPeeds
        'IPlaybackSpeed
        'PlaybackSpeed
        'Autotrail options

    End Sub
    Public Sub SetPreferences()
        HorSplit = mHorSplit
        VertSplit = mVertSplit
        'etc

    End Sub
    Public Sub SetDefaults()
        With My.Computer.Registry.CurrentUser
            Dim s As String = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            'CurrentFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)
            Dim fol As New IO.DirectoryInfo(CurrentFolder)


            '  Media.MediaPath = New IO.DirectoryInfo(CurrentFolder).EnumerateFiles("*", IO.SearchOption.AllDirectories).First.FullName
            FavesFolderPath = s & "\Favourites\"
            ButtonFilePath = Media.MediaPath
        End With
        With MainForm
            .ctrFileBoxes.SplitterDistance = .ctrFileBoxes.Height / 4
            .ctrMainFrame.SplitterDistance = .ctrFileBoxes.Width / 2

            .CurrentFilterState.State = 0
            .PlayOrder.State = 0
            .NavigateMoveState.State = 0
            iCurrentAlpha = 0
            ButtonFilePath = ""
            '.SetValue("LastButtonFolder", strButtonfile)
        End With
        Media.StartPoint.State = 0
        PreferencesSave()
    End Sub

End Class