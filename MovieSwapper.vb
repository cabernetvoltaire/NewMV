Imports AxWMPLib
Public Class MediaSwapper

    Public NextF As New NextFile
    Public Media1 As New MediaHandler("mMedia1")
    Public Media2 As New MediaHandler("mMedia2")
    Public Media3 As New MediaHandler("mMedia3")


    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
    Public Event LoadedMedia(MH As MediaHandler)
    Public Event MediaNotFound(MH As MediaHandler)
    Public Property NextItem As String
    Public Property PreviousItem As String
    Public Event MediaShown(MH As MediaHandler)
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
        End Set
    End Property
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
            '  Listbox.SetSelected(value, True)
        End Set
    End Property
    Public Sub New(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer, ByRef PB1 As PictureBox, ByRef PB2 As PictureBox, ByRef PB3 As PictureBox)

        AssignPlayers(MP1, MP2, MP3)
        AssignPictures(PB1, PB2, PB3)

    End Sub
    Public Sub AssignPictures(ByRef PB1 As PictureBox, ByRef PB2 As PictureBox, ByRef PB3 As PictureBox)
        Media1.Picture = PB1
        Media2.Picture = PB2
        Media3.Picture = PB3

    End Sub
    Public Sub AssignPlayers(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer)
        Media1.Player = MP1
        Media2.Player = MP2
        Media3.Player = MP3

    End Sub
    ''' <summary>
    ''' Loads the media marked index, and puts the previous and next files into the other two media handlers
    ''' </summary>
    ''' <param name="index"></param>
    Private Sub SetIndex(index As Integer)
        Dim Current As String
        Dim Nxt As String
        Dim Prev As String
        mListcount = Listbox.Items.Count
        Static oldindex As Integer

        NextF.Forwards = (mListIndex > oldindex) Or (mListIndex = 0)
        NextF.CurrentIndex = index

        Current = NextF.CurrentItem
        Nxt = NextF.NextItem
        NextItem = Nxt
        Prev = NextF.PreviousItem

        Select Case Current

            Case Media2.MediaPath
                RotateMedia(Media2, Media3, Media1)
            Case Media3.MediaPath
                RotateMedia(Media3, Media1, Media2)
            Case Else
                RotateMedia(Media1, Media2, Media3)
        End Select


        oldindex = index
    End Sub

    Private Sub Prepare(ByRef MH As MediaHandler, path As String)
        Debug.Print("PREPARE: " & MH.Player.Name)
        MH.MediaPath = path
        Select Case MH.MediaType
            Case Filetype.Movie
                MH.Player.Visible = True
                MH.PlaceResetter(True)
                RaiseEvent LoadedMedia(MH) 'Currently does nothing.
            Case Filetype.Pic
                MH.PlaceResetter(False)
                MH.MediaPath = path
                MH.Picture.Visible = True
            Case Else

        End Select
        'NB Duration is not loaded by this point. 

    End Sub
    Private Sub RotateMedia(ByRef ThisMH As MediaHandler, ByRef NextMH As MediaHandler, ByRef PrevMH As MediaHandler)

        Prepare(PrevMH, NextF.PreviousItem)
        Prepare(NextMH, NextF.NextItem)
        Prepare(ThisMH, NextF.CurrentItem)

        Select Case ThisMH.MediaType
            Case Filetype.Movie
                ShowPlayer(ThisMH)
            Case Filetype.Pic
                ShowPicture(ThisMH)
        End Select

    End Sub
    Public Sub SetStartStates(ByRef SH As StartPointHandler)
        Media1.StartPoint.State = SH.State
        Media2.StartPoint.State = SH.State
        Media3.StartPoint.State = SH.State


    End Sub
    Public Sub SetStartpoints(ByRef SH As StartPointHandler)
        SetStartStates(SH)

        Media1.StartPoint = SH

        Media2.StartPoint = SH

        Media3.StartPoint = SH


    End Sub
    Public Sub URLSZero()
        Media1.Player.URL = ""
        Media2.Player.URL = ""
        Media3.Player.URL = ""
        Media1.Picture.Image = Nothing
        Media2.Picture.Image = Nothing
        Media3.Picture.Image = Nothing

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
        If Not separate Then
            HideMedias(MHX)
        End If
        MHX.PlaceResetter(False)
        With MHX.Player


            .Visible = True
            .BringToFront()
            .settings.mute = Muted
        End With
        RaiseEvent MediaShown(MHX)

    End Sub
    Public Sub ClickAllPics()
        PicClick(Media1.Picture)
        PicClick(Media2.Picture)
        PicClick(Media3.Picture)

    End Sub
    Private Sub HideMedias(CurrentMH As MediaHandler)
        Exit Sub
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
        HideMedias(MHX)
        MHX.Picture.Visible = True
        MHX.Picture.BringToFront()
        RaiseEvent MediaShown(MHX)

    End Sub

End Class
