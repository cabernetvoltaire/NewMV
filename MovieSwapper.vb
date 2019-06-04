﻿Imports AxWMPLib
Public Class MediaSwapper

    Public NextF As New NextFile
    Private mMedia1 As New MediaHandler("mMedia1")
    Private mMedia2 As New MediaHandler("mMedia2")
    Private mMedia3 As New MediaHandler("mMedia3")


    Private mFileList As New List(Of String) '
    Private mListIndex As Integer
    Private mListbox As New ListBox
    Private mListcount As Integer
    Public Event LoadedMedia(MH As MediaHandler)
    Public Event MediaNotFound(MH As MediaHandler)
    Public Property NextItem As String
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
        mMedia1.Picture = PB1
        mMedia2.Picture = PB2
        mMedia3.Picture = PB3

    End Sub
    Public Sub AssignPlayers(ByRef MP1 As AxWindowsMediaPlayer, ByRef MP2 As AxWindowsMediaPlayer, ByRef MP3 As AxWindowsMediaPlayer)
        mMedia1.Player = MP1
        mMedia2.Player = MP2
        mMedia3.Player = MP3

    End Sub

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
            Case mMedia2.MediaPath
                RotateMedia(mMedia2, mMedia3, mMedia1)
            Case mMedia3.MediaPath
                RotateMedia(mMedia3, mMedia1, mMedia2)
            Case Else
                RotateMedia(mMedia1, mMedia2, mMedia3)
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
        mMedia1.StartPoint.State = SH.State
        mMedia2.StartPoint.State = SH.State
        mMedia3.StartPoint.State = SH.State


    End Sub
    Public Sub SetStartpoints(ByRef SH As StartPointHandler)
        SetStartStates(SH)
        '  Dim dur As Long

        'dur = mMedia1.StartPoint.Duration
        'dur = mMedia3.StartPoint.Duration
        'dur = mMedia2.StartPoint.Duration
        mMedia1.StartPoint = SH

        mMedia2.StartPoint = SH

        mMedia3.StartPoint = SH
        '  SetIndex(ListIndex)
        'mMedia2.StartPoint.Duration = dur
        ' mMedia1.StartPoint.Duration = dur
        'mMedia3.StartPoint.Duration = dur

    End Sub
    Public Sub URLSZero()
        mMedia1.Player.URL = ""
        mMedia2.Player.URL = ""
        mMedia3.Player.URL = ""
        mMedia1.Picture.Image = Nothing
        mMedia2.Picture.Image = Nothing
        mMedia3.Picture.Image = Nothing

    End Sub
    Public Sub ResettersOff()
        mMedia1.PlaceResetter(False)
        mMedia2.PlaceResetter(False)
        mMedia3.PlaceResetter(False)
    End Sub

    Public Sub MuteAll()
        mMedia1.Player.settings.mute = True
        mMedia2.Player.settings.mute = True
        mMedia3.Player.settings.mute = True

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
        PicClick(mMedia1.Picture)
        PicClick(mMedia2.Picture)
        PicClick(mMedia3.Picture)

    End Sub
    Private Sub HideMedias(CurrentMH As MediaHandler)
        Exit Sub
        If mMedia1 IsNot CurrentMH Then
            mMedia1.Picture.Visible = False
            mMedia1.Player.Visible = False
        End If
        If mMedia2 IsNot CurrentMH Then
            mMedia2.Picture.Visible = False
            mMedia2.Player.Visible = False
        End If
        If mMedia3 IsNot CurrentMH Then
            mMedia3.Picture.Visible = False
            mMedia3.Player.Visible = False
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
