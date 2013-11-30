VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEditPlaylistEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lPlaylistFileName As String

Public Sub SortPlaylistByType()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
    If Err.Number <> 0 Then Err.Clear
Next i
ResetPlaylist
If Err.Number <> 0 Then Err.Clear
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case "Avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "WMA"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.Title) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                If Err.Number <> 0 Then Err.Clear
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                If Err.Number <> 0 Then Err.Clear
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.Title) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                If Err.Number <> 0 Then Err.Clear
            End If
        End If
    Case Else
        
    End Select
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
frmMain.tvwPlaylist.Visible = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistByYear()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
    If Err.Number <> 0 Then Err.Clear
Next i
ResetPlaylist
If Err.Number <> 0 Then Err.Clear
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case "Avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "WMA"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.SongYear) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.SongYear, lID3.SongYear, 6
                If Err.Number <> 0 Then Err.Clear
            End If
            If Len(lID3.Title) <> 0 Then
                If Len(lID3.SongYear) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.SongYear, tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            Else
                If Len(lID3.SongYear) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.SongYear, tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.SongYear) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.SongYear, lID3.SongYear, 6
                    If Err.Number <> 0 Then Err.Clear
                End If
                If Len(lID3.Title) <> 0 Then
                    If Len(lID3.SongYear) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.SongYear, tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                Else
                    If Len(lID3.SongYear) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.SongYear, tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                If Err.Number <> 0 Then Err.Clear
            End If
        End If
    Case Else
        
    End Select
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
frmMain.tvwPlaylist.Visible = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistByAlbum()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
    If Err.Number <> 0 Then Err.Clear
Next i
ResetPlaylist
If Err.Number <> 0 Then Err.Clear
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case "Avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "WMA"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.Album) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Album, lID3.Album, 6
                If Err.Number <> 0 Then Err.Clear
            End If
            If Len(lID3.Title) <> 0 Then
                If Len(lID3.Album) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            Else
                If Len(lID3.Album) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.Album) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Album, lID3.Album, 6
                    If Err.Number <> 0 Then Err.Clear
                End If
                If Len(lID3.Title) <> 0 Then
                    If Len(lID3.Album) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                Else
                    If Len(lID3.Album) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                If Err.Number <> 0 Then Err.Clear
            End If
        End If
    Case Else
        
    End Select
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
frmMain.tvwPlaylist.Visible = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistByComment()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
    If Err.Number <> 0 Then Err.Clear
Next i
ResetPlaylist
If Err.Number <> 0 Then Err.Clear
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case "Avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "WMA"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
        If Err.Number <> 0 Then Err.Clear
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.Comment) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Comment, lID3.Comment, 6
                If Err.Number <> 0 Then Err.Clear
            End If
            If Len(lID3.Title) <> 0 Then
                If Len(lID3.Comment) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Comment, tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            Else
                If Len(lID3.Comment) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Comment, tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                    If Err.Number <> 0 Then Err.Clear
                End If
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.Comment) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Comment, lID3.Comment, 6
                    If Err.Number <> 0 Then Err.Clear
                End If
                If Len(lID3.Title) <> 0 Then
                    If Len(lID3.Comment) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Comment, tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                Else
                    If Len(lID3.Comment) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Comment, tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                If Err.Number <> 0 Then Err.Clear
            End If
        End If
    Case Else
        
    End Select
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Visible = True
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistByArtist()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, f As Integer, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
Next i
ResetPlaylist
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case ".avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
    Case ".wma"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.Artist) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Artist, lID3.Artist, 6
                If Err.Number <> 0 Then Err.Clear
            End If
            If Len(lID3.Title) <> 0 Then
                If Len(lID3.Artist) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, lID3.Title, 3
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                End If
            Else
                If Len(lID3.Artist) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, GetFileTitle(msg), 3
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                End If
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.Artist) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Artist, lID3.Artist, 6
                    If Err.Number <> 0 Then Err.Clear
                End If
                If Len(lID3.Title) <> 0 Then
                    If Len(lID3.Artist) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, lID3.Title, 3
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                    End If
                Else
                    If Len(lID3.Artist) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, GetFileTitle(msg), 3
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                    End If
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
            End If
        End If
    Case Else
    End Select
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Visible = True
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
frmMain.tvwPlaylist.Nodes(f).Selected = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub ResetPlaylist()
On Local Error Resume Next
frmMain.tvwPlaylist.Nodes.Clear
frmMain.tvwPlaylist.Nodes.Add , , "MP3", "Mpeg Audio (*.mp1;*.mp2;*.mp3)", 6
frmMain.tvwPlaylist.Nodes.Add , , "Wave", "Wave Audio (*.wav)", 6
frmMain.tvwPlaylist.Nodes.Add , , "Mpeg", "Mpeg Video (*.mpg;*.mpeg)", 7
frmMain.tvwPlaylist.Nodes.Add , , "Avi", "Avi and DivX Video (*.avi)", 7
frmMain.tvwPlaylist.Nodes.Add , , "WMA", "Windows Media (*.wma;*.wmv;*.asf;*.wmx)", 7
frmMain.tvwPlaylist.Nodes.Add , , "M4A", "M4A Audio (*.m4a)", 6
frmMain.tvwPlaylist.Nodes.Add , , "MOV", "Movie Files (*.mov)", 7
frmMain.tvwPlaylist.Nodes.Add , , "VOB", "DVD Video Files (*.vob)", 7
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ResetPlaylist()", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistByArtistAndAlbum()
On Local Error Resume Next
Dim i As Integer, lID3 As ID3Tag, b As Boolean, msg As String, msg2 As String
frmMain.tvwPlaylist.Visible = False
msg2 = frmMain.tvwPlaylist.SelectedItem.Text
frmMain.tvwTemp.Nodes.Clear
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwTemp.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Key
Next i
For i = 1 To frmMain.tvwTemp.Nodes.Count
    Select Case LCase(Right(frmMain.tvwTemp.Nodes(i).Key, 4))
    Case ".avi"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Avi", tvwChild, msg, GetFileTitle(msg), 3
    Case ".wma"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "WMA", tvwChild, msg, GetFileTitle(msg), 3
    Case "mpeg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
    Case ".mpg"
        msg = frmMain.tvwTemp.Nodes(i).Key
        frmMain.tvwPlaylist.Nodes.Add "Mpeg", tvwChild, msg, GetFileTitle(msg), 3
    Case ".mp3"
        msg = frmMain.tvwTemp.Nodes(i).Key
        b = ReadID3v2(msg, lID3)
        If b = True Then
            If Len(lID3.Artist) <> 0 Then
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Artist, lID3.Artist, 6
                If Err.Number <> 0 Then Err.Clear
                If Len(lID3.Album) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, lID3.Album, lID3.Album, 6
                    If Err.Number <> 0 Then Err.Clear
                    If Len(lID3.Title) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                Else
                    If Len(lID3.Title) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, GetFileTitle(msg)
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
            Else
                If Len(lID3.Title) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title
                    If Err.Number <> 0 Then Err.Clear
                Else
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg)
                    If Err.Number <> 0 Then Err.Clear
                End If
            End If
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                If Len(lID3.Artist) <> 0 Then
                    frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, lID3.Artist, lID3.Artist, 6
                    If Err.Number <> 0 Then Err.Clear
                    If Len(lID3.Album) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, lID3.Album, lID3.Album, 6
                        If Err.Number <> 0 Then Err.Clear
                        If Len(lID3.Title) <> 0 Then
                            frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, lID3.Title, 3
                            If Err.Number <> 0 Then Err.Clear
                        Else
                            frmMain.tvwPlaylist.Nodes.Add lID3.Album, tvwChild, msg, GetFileTitle(msg), 3
                            If Err.Number <> 0 Then Err.Clear
                        End If
                    Else
                        If Len(lID3.Title) <> 0 Then
                            frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, lID3.Title, 3
                            If Err.Number <> 0 Then Err.Clear
                        Else
                            frmMain.tvwPlaylist.Nodes.Add lID3.Artist, tvwChild, msg, GetFileTitle(msg), 3
                            If Err.Number <> 0 Then Err.Clear
                        End If
                    End If
                Else
                    If Len(lID3.Title) <> 0 Then
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, lID3.Title, 3
                        If Err.Number <> 0 Then Err.Clear
                    Else
                        frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
                        If Err.Number <> 0 Then Err.Clear
                    End If
                End If
            Else
                frmMain.tvwPlaylist.Nodes.Add "MP3", tvwChild, msg, GetFileTitle(msg), 3
            End If
        End If
    Case Else
        
    End Select
    If Err.Number <> 0 Then Err.Clear
Next i
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(i).Expanded = True
Next i
frmMain.tvwPlaylist.Nodes(FindTreeViewIndex(msg2, frmMain.tvwPlaylist)).Selected = True
frmMain.tvwPlaylist.Visible = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub EditPlaylistEntry(lFilename As String)
On Local Error GoTo ErrHandler
Dim msg As String, arr() As String, arr1() As String, i As Integer, o As Integer, n As Integer, f As frmPlaylistEntry
If Len(lFilename) <> 0 Then
    msg = ReadFile(lIniFiles.iPlaylistTreeView)
    arr = Split(msg, vbCrLf)
    For i = 0 To UBound(arr)
        If InStr(arr(i), lFilename) Then
            arr1 = Split(arr(i), "|")
            n = i
            Exit For
        End If
    Next i
    If n <> 0 Then
        Set f = New frmPlaylistEntry
        f.Show
        For o = 0 To UBound(arr1)
            If InStr(LCase(arr1(o)), "sub") Then arr1(o) = Replace(arr1(o), "Sub:", "")
            Select Case o
            Case 0
                f.txtFoundIN.Text = arr1(o)
            Case 1
                f.txtFilename.Text = arr1(o)
            Case 2
                f.txtDisplay.Text = arr1(o)
            End Select
        Next o
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EditPlaylistEntry(lFileName As String)", Err.Description, Err.Number
End Sub

Public Sub UnSortPlaylistNode(lIndex As Integer)
On Local Error GoTo ErrHandler
Dim i As Integer
If lIndex <> 0 Then
    frmMain.tvwPlaylist.Nodes(lIndex).Sorted = False
    For i = 0 To frmMain.tvwPlaylist.Nodes(lIndex).children
        frmMain.tvwPlaylist.Nodes(i).Sorted = False
    Next i
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SortPlaylistNode(lIndex As Integer)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylistNode(lIndex As Integer)
On Local Error GoTo ErrHandler
Dim i As Integer
If lIndex <> 0 Then
    frmMain.tvwPlaylist.Nodes(lIndex).Sorted = True
    For i = 0 To frmMain.tvwPlaylist.Nodes(lIndex).children
        frmMain.tvwPlaylist.Nodes(i).Sorted = True
    Next i
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SortPlaylistNode(lIndex As Integer)", Err.Description, Err.Number
End Sub

Public Sub SortPlaylist()
On Local Error GoTo ErrHandler
Dim i As Integer, n As Integer
If frmMain.mnuAlphabetic.Checked = True Then
    frmMain.tvwPlaylist.Sorted = False
    frmMain.mnuAlphabetic.Checked = False
    For i = 1 To frmMain.tvwPlaylist.Nodes.Count
        frmMain.tvwPlaylist.Nodes(i).Sorted = False
    Next i
Else
    frmMain.tvwPlaylist.Sorted = True
    frmMain.mnuAlphabetic.Checked = True
    For i = 1 To frmMain.tvwPlaylist.Nodes.Count
        frmMain.tvwPlaylist.Nodes(i).Sorted = True
    Next i
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAlphabetic_Click()", Err.Description, Err.Number
End Sub
