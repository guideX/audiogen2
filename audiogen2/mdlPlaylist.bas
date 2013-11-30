Attribute VB_Name = "mdlPlaylist"
Option Explicit
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Sub FillTreeViewWithPlaylist(lTreeView As TreeView)
On Local Error Resume Next
Dim i As Integer
lTreeView.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    If DoesFileExist(frmMain.tvwPlaylist.Nodes(i).Key) And Len(frmMain.tvwPlaylist.Nodes(i).Key) <> 0 And DoesTreeViewItemExistByText(frmMain.tvwPlaylist.Nodes(i).Text, lTreeView) = False Then
        lTreeView.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Text, 3
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub FillTreeViewWithPlaylist(lTreeView As TreeView)", Err.Description, Err.Number
End Sub

Public Sub AddToPlaylistDelay(lFilename As String)
On Local Error Resume Next
If Len(lFilename) <> 0 And DoesFileExist(lFilename) = True Then
    frmMain.lstAddToPlaylist.AddItem lFilename
    frmMain.tmrDelayAddToPlaylist.Enabled = True
End If
End Sub

Public Sub AddToPlaylist(lFilename As String, Optional lSearchCurrentDirectory As Boolean)
On Local Error Resume Next
Dim lFileTitle As String, lMP3 As MP3File, lArtist As String, lAlbum As String, lTitle As String, i As Integer, msg As String, msg2 As String, t As tSearch
With frmMain.tvwPlaylist
    DoEvents
    If DoesFileExist(lFilename) = True And DoesItemExistInPlaylist(lFilename) = False Then
        lFileTitle = GetFileTitle(lFilename)
        Select Case LCase(Right(lFileTitle, 4))
        Case ".mp3"
            lMP3.HasIDv2 = ReadID3v2(lFilename, lMP3.IDv2)
            If lMP3.HasIDv2 = True Then
                lMP3.IDv2.Artist = CleanInterpreteItems(lMP3.IDv2.Artist)
                lMP3.IDv2.Album = CleanInterpreteItems(lMP3.IDv2.Album)
                lMP3.IDv2.Title = CleanInterpreteItems(lMP3.IDv2.Title)
                lArtist = lMP3.IDv2.Artist
                lAlbum = lMP3.IDv2.Album
                lTitle = lMP3.IDv2.Title
            Else
                lMP3.HasIDv1 = ReadID3v1(lFilename, lMP3.IDv1)
                If lMP3.HasIDv1 = True Then
                    lMP3.IDv1.Artist = CleanInterpreteItems(lMP3.IDv1.Artist)
                    lMP3.IDv1.Album = CleanInterpreteItems(lMP3.IDv1.Album)
                    lMP3.IDv1.Title = CleanInterpreteItems(lMP3.IDv1.Title)
                    lArtist = lMP3.IDv1.Artist
                    lAlbum = lMP3.IDv1.Album
                    lTitle = lMP3.IDv1.Title
                End If
            End If
            lArtist = FormatCase(lArtist)
            lTitle = FormatCase(lTitle)
            lAlbum = FormatCase(lAlbum)
            If Len(Trim(lArtist)) <> 0 Then
                lArtist = Replace(lArtist, "the ", "")
                lArtist = Replace(lArtist, "The ", "")
                lArtist = Replace(lArtist, "THE ", "")
                AddToTreeView frmMain.tvwPlaylist, "MP3", tvwChild, lArtist, lArtist, 7, False, True
                If Len(Trim(lTitle)) <> 0 And Len(Trim(lAlbum)) <> 0 Then
                    AddToTreeView frmMain.tvwPlaylist, lArtist, tvwChild, lAlbum, lAlbum, 7, False: DoEvents
                    AddToTreeView frmMain.tvwPlaylist, lAlbum, tvwChild, lFilename, lTitle, 3, True, True
                ElseIf Len(Trim(lTitle)) <> 0 And Len(Trim(lAlbum)) = 0 Then
                    AddToTreeView frmMain.tvwPlaylist, lArtist, tvwChild, lFilename, lTitle, 3, True, True
                ElseIf Len(Trim(lTitle)) = 0 And Len(Trim(lAlbum)) <> 0 Then
                    AddToTreeView frmMain.tvwPlaylist, lArtist, tvwChild, lAlbum, lAlbum, 7, False, True: DoEvents
                    AddToTreeView frmMain.tvwPlaylist, lAlbum, tvwChild, lFilename, lFileTitle, 3, True, True
                ElseIf Len(Trim(lTitle)) = 0 And Len(Trim(lAlbum)) = 0 Then
                    AddToTreeView frmMain.tvwPlaylist, lArtist, tvwChild, lFilename, lFileTitle, 3, True, True
                End If
            Else
                If Len(lTitle) <> 0 And Len(lAlbum) <> 0 Then
                    AddToTreeView frmMain.tvwPlaylist, "MP3", tvwChild, lAlbum, lAlbum, 7, False, True
                    AddToTreeView frmMain.tvwPlaylist, lAlbum, tvwChild, lFilename, lFileTitle, 3, True, True
                ElseIf Len(lTitle) <> 0 And Len(lAlbum) = 0 Then
                    AddToTreeView frmMain.tvwPlaylist, "MP3", tvwChild, lFilename, lTitle, 3, True, True
                ElseIf Len(lTitle) = 0 And Len(lAlbum) <> 0 Then
                    AddToTreeView frmMain.tvwPlaylist, "MP3", tvwChild, lAlbum, lAlbum, 7, False, True
                    AddToTreeView frmMain.tvwPlaylist, lAlbum, tvwChild, lFilename, lFileTitle, 3, True, True
                ElseIf Len(lTitle) = 0 And Len(lAlbum) = 0 Then
                    AddToTreeView frmMain.tvwPlaylist, "MP3", tvwChild, lFilename, lFileTitle, 3, True, True
                End If
            End If
        Case ".wav"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "Wave", tvwChild, lFilename, lFileTitle, 3
        Case ".avi"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "Avi", tvwChild, lFilename, lFileTitle, 4
        Case "mpeg"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "Mpeg", tvwChild, lFilename, lFileTitle, 4
        Case ".mpg"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "Mpeg", tvwChild, lFilename, lFileTitle, 4
        Case ".asf"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "ASF", tvwChild, lFilename, lFileTitle, 4
        Case ".wma"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "WMA", tvwChild, lFilename, lFileTitle, 3
        Case ".wmv"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "WMV", tvwChild, lFilename, lFileTitle, 4
        Case ".wmx"
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "WMX", tvwChild, lFilename, lFileTitle, 4
        Case Else
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwPlaylist) = False Then .Nodes.Add "WMA", tvwChild, lFilename, lFileTitle, 3
        End Select
        If lSearchCurrentDirectory = True Then
            msg = Left(lFilename, Len(lFilename) - Len(lFileTitle))
            GetFiles msg, lFileFormats.fSupportedTypes, vbNormal, t
            DoEvents
            For i = 0 To t.Count
                If Len(t.Path(i)) <> 0 Then
                    AddToPlaylist t.Path(i)
                End If
            Next i
        End If
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub AddToPlaylist(lFilename As String)", Err.Description, Err.Number
End Sub

Public Function DoesItemExistInPlaylist(lFilename As String) As Boolean
On Local Error Resume Next
Dim i As Integer
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    If LCase(frmMain.tvwPlaylist.Nodes(i).Key) = LCase(lFilename) Then
        DoesItemExistInPlaylist = True
        Exit For
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError "Public Function DoesItemExistInPlaylist(lFilename As String) As Boolean", Err.Description, Err.Number
End Function

Public Sub SavePlaylist(lFilename As String)
On Local Error Resume Next
Dim msg As String, i As Integer
If DoesFileExist(lFilename) = True Then Kill lFilename
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    If InStr(frmMain.tvwPlaylist.Nodes(i).Key, ":") And Len(frmMain.tvwPlaylist.Nodes(i).Key) <> 0 And DoesFileExist(frmMain.tvwPlaylist.Nodes(i).Key) = True Then
        If Len(msg) <> 0 Then
            msg = msg & vbCrLf & frmMain.tvwPlaylist.Nodes(i).Key
        Else
            msg = frmMain.tvwPlaylist.Nodes(i).Key
        End If
    End If
Next i
If Len(msg) <> 0 Then
    SaveFile lFilename, msg
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub SavePlaylist(lFilename As String)", Err.Description, Err.Number
End Sub

Public Function SaveFile(lFilename As String, lText As String) As Boolean
On Local Error GoTo ErrHandler
Dim i As Integer
i = FreeFile
If Len(lFilename) <> 0 And Len(lText) <> 0 Then
    Open lFilename For Output As #i
    Print #i, lText
    Close #i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function SaveFile(lFilename As String, lText As String) As Boolean", Err.Description, Err.Number
End Function

Public Sub TreeViewLabelEdit(lTreeView As TreeView, lOldString As String, lNewString As String)
On Local Error Resume Next
Dim i As Integer, lTag As ID3Tag, lMP3 As MP3File
i = FindTreeViewIndex(lOldString, lTreeView)
If Right(LCase(lTreeView.Nodes(i).Key), 4) = ".mp3" Then
    lTag = RenderMp3Tag(lTreeView.Nodes(i).Key)
    lTag.Title = lNewString
    lMP3.HasIDv2 = ReadID3v2(lTreeView.Nodes(i).Key, lMP3.IDv2)
    If lMP3.HasIDv2 = True Then
        lTag.Album = lMP3.IDv2.Album
        lTag.Artist = lMP3.IDv2.Artist
        lTag.SongYear = lMP3.IDv2.SongYear
        lTag.Comment = lMP3.IDv2.Comment
        lTag.Genre = lMP3.IDv2.Genre
        WriteID3v2 lTreeView.Nodes(i).Key, lTag
    Else
        lMP3.HasIDv1 = ReadID3v1(lTreeView.Nodes(i).Key, lMP3.IDv1)
        If lMP3.HasIDv1 = True Then
            lTag.Album = lMP3.IDv1.Album
            lTag.Artist = lMP3.IDv1.Artist
            lTag.SongYear = lMP3.IDv1.SongYear
            lTag.Comment = lMP3.IDv1.Comment
            lTag.Genre = lMP3.IDv1.Genre
            WriteID3v1 lTreeView.Nodes(i).Key, lTag
        Else
            WriteID3v1 lTreeView.Nodes(i).Key, lTag
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub TreeViewLabelEdit(lTreeView As TreeView, lOldString As String, lNewString As String)", Err.Description, Err.Number
End Sub

Public Sub LoadPlaylist(lFilename As String, Optional lReloadPlaylist As Boolean)
On Local Error Resume Next
Dim msg As String, msg2 As String, i As Integer, p As clsEditPlaylistEntry, c As Integer
Dim msg3() As String
Set p = New clsEditPlaylistEntry
frmMain.tvwPlaylist.Visible = False
If lReloadPlaylist = True Then p.ResetPlaylist
If Len(lFilename) <> 0 And DoesFileExist(lFilename) = True Then
    msg = ReadFile(lFilename)
    If msg = lFilename Then msg = ""
    If Len(msg) <> 0 Then
        msg3 = Split(msg, Chr(13))
        For i = 0 To UBound(msg3)
            msg3(i) = Trim(msg3(i))
            If Len(msg3(i)) <> 0 And DoesFileExist(msg3(i)) = True Then
                MsgBox "!" & msg3(i) & "!"
            End If
        Next i
        'LockWindowUpdate frmMain.hwnd
        'Do Until Len(msg) = 0
        '    If InStr(msg, Chr(13)) Then
        '        msg2 = Trim(Left(msg, 1) & Parse(msg, Left(msg, 1), Chr(13)))
        '        msg2 = Replace(msg2, Chr(10), "")
        '        msg2 = Replace(msg2, Chr(13), "")
        '        If Len(msg2) <> 0 Then msg = Right(msg, Len(msg) - Len(msg2))
        '        If Left(msg, 1) = Chr(10) Or Left(msg, 1) = Chr(13) Then msg = Right(msg, Len(msg) - 1)
        '        If Left(msg, 1) = Chr(10) Or Left(msg, 1) = Chr(13) Then msg = Right(msg, Len(msg) - 1)
        '    Else
        '        msg2 = msg
        '        msg = ""
        '    End If
        '    If Len(msg2) <> 0 Then
        '        AddToPlaylistDelay msg2
        '        If Len(msg2) = 0 Then Exit Sub
        '    End If
        '    c = c + 1
        '    If c = 5000 Then
        '        frmMain.tvwPlaylist.Visible = True
        '        LockWindowUpdate 0
        '        Exit Sub
        '    End If
        'Loop
        'LockWindowUpdate 0
    End If
End If
frmMain.tvwPlaylist.Visible = True
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub LoadPlaylist(lFilename As String, Optional lReloadPlaylist As Boolean)", Err.Description, Err.Number
End Sub
