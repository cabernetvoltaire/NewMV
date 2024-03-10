Imports Newtonsoft.Json

Public Class AppSettings
    Public Property VertSplit As Integer
    Public Property HorSplit As Integer
    Public Property Filter As FilterHandler.FilterState
    Public Property SortOrder As SortHandler.Order
    Public Property Startpoint As StartPointHandler.StartTypes
    Public Property State As StateHandler.StateOptions
    Public Property CurrentAlpha As Integer
    Public Property FavouritesPath As String
    Public Property LastButtonFile As String
    Public Property LastGoodFilePath As String
    Public Property ThumbnailDestination As String
    Public Property RootScanPath As String
    Public Property ButtonFolder As String
    Public Property DirectoriesPath As String
    Public Property GlobalFavesPath As String
    Public Property LastTimeSuccessful As Boolean
    Public Property PreviewLinks As Boolean
    Public Property RandomNextFile As Boolean
    Public Property RandomOnDirectoryChange As Boolean
    Public Property RandomAutoTrail As Boolean
    Public Property RandomAutoLoadButtons As Boolean
    Public Property ShowAttributes As Boolean
    Public Property Encrypt As Boolean
    Public Property AutoAdvance As Boolean
    Public Property Separate As Boolean
    Public Property MovieScan As Boolean
    Public Property ScanBookmarks As Boolean
    Public Property SingleLinks As Boolean
    Public Property DontLoadVidsShowPic As Boolean
    Public Property DontPreLoadVids As Boolean
    Public Property SlowmoSoundOption As Integer
    Public Property MovieSlideShowSpeed As Integer
    Public Property FractionalJump As Integer
    Public Property AbsoluteJump As Integer
    Public Property Speed As Integer
    Public Property StateColours As Color()
    Public Property ButtonSplit As Integer


    Public Sub New()
        Me.CurrentAlpha = 0
        Me.AutoAdvance = False
        Me.AbsoluteJump = 35
        Me.SortOrder = 0
        Me.Separate = False
        Me.Startpoint = StartPointHandler.StartTypes.FirstMarker
        Me.Filter = FilterHandler.FilterState.All
        Me.DontPreLoadVids = False
        Me.DontLoadVidsShowPic = False
        Me.Encrypt = False
        Me.FractionalJump = 8
        Me.PreviewLinks = False
        Me.MovieScan = False
        Me.MovieSlideShowSpeed = 4000
        Me.Speed = -1
        Me.ScanBookmarks = False
        Me.LastTimeSuccessful = False
        Me.SlowmoSoundOption = 0
        Me.SingleLinks = False
        Me.ShowAttributes = False
        Me.RandomAutoTrail = False
        Me.RandomAutoLoadButtons = True
        Me.State = StateHandler.StateOptions.Navigate
        Me.ButtonFolder = "Q:\.msb"
        Me.LastButtonFile = "Q:\.msb\General2.msb"
        Me.GlobalFavesPath = "Q:\Favourites"
        Me.FavouritesPath = "Q:\Favourites\Bookmarks\2024"
        Me.RootScanPath = "Q:\"
        Me.HorSplit = 500
        Me.VertSplit = 500
        Me.ButtonSplit = 1500
        Me.StateColours = {Color.Aqua, Color.Orange, Color.LightPink, Color.LightGreen, Color.Tomato, Color.BurlyWood, Color.BlueViolet}


    End Sub
End Class
Public Class SettingsManager
    Private Const PreferencesFilePath As String = "C:\Users\paulc\AppData\Roaming\Metavisua\Preferences\Preferences.json"

    Public Shared Function LoadSettings() As AppSettings
        If Not IO.File.Exists(PreferencesFilePath) Then
            Return New AppSettings() ' Return defaults if file doesn't exist
        End If

        Dim json As String = IO.File.ReadAllText(PreferencesFilePath)
        Return JsonConvert.DeserializeObject(Of AppSettings)(json)
    End Function

    Public Shared Sub SaveSettings(settings As AppSettings)
        Dim json As String = JsonConvert.SerializeObject(settings)
        IO.File.WriteAllText(PreferencesFilePath, json)
    End Sub
End Class
