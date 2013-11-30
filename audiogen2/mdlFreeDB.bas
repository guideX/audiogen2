Attribute VB_Name = "mdlFreeDB"
Option Explicit

Public Sub SendReadDiscID(lCatagory As String, lDiscID As String, lWinsock As Winsock)
On Local Error GoTo ErrHandler
If Len(lCatagory) <> 0 And Len(lDiscID) <> 0 Then
    lWinsock.SendData "cddb read " & lCatagory & " " & lDiscID & vbCrLf
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SendReadDiscID(lCatagory As String, lDiscID As String, lWinsock As Winsock)", Err.Description, Err.Number
End Sub

Public Sub ConnectToFreeDB(lWinsock As Winsock)
On Local Error GoTo ErrHandler
lWinsock.Close
lWinsock.Connect "freedb.freedb.org", 8880
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ConnectToFreeDB(lWinsock As Winsock)", Err.Description, Err.Number
End Sub

Public Sub RefreshCDTracks()
On Local Error Resume Next
Dim i As Integer, msg2 As String
For i = 1 To frmMain.lvwCD.ListItems.Count
    msg2 = ReturnRipPath & frmMain.lvwCD.ListItems(i).SubItems(1) & " - " & frmMain.lvwCD.ListItems(i).SubItems(2) & "\" & frmMain.lvwCD.ListItems(i).SubItems(1) & " - " & frmMain.lvwCD.ListItems(i).SubItems(3) & ".wav"
    If DoesFileExist(msg2) = True Then
        frmMain.lvwCD.ListItems(i).SubItems(5) = "Copied"
        frmMain.lvwCD.ListItems(i).ForeColor = vbRed
        frmMain.lvwCD.ListItems(i).ListSubItems(1).ForeColor = vbRed
        frmMain.lvwCD.ListItems(i).ListSubItems(2).ForeColor = vbRed
        frmMain.lvwCD.ListItems(i).ListSubItems(3).ForeColor = vbRed
        frmMain.lvwCD.ListItems(i).ListSubItems(4).ForeColor = vbRed
        frmMain.lvwCD.ListItems(i).ListSubItems(5).ForeColor = vbRed
        
        frmMain.lvwCD.ListItems(i).Tag = msg2
    End If
    msg2 = Left(msg2, Len(msg2) - 4) & ".mp3"
    If DoesFileExist(msg2) = True Then
        frmMain.lvwCD.ListItems(i).SubItems(5) = "Converted"
        frmMain.lvwCD.ListItems(i).ForeColor = vbBlue
        frmMain.lvwCD.ListItems(i).ListSubItems(1).ForeColor = vbBlue
        frmMain.lvwCD.ListItems(i).ListSubItems(2).ForeColor = vbBlue
        frmMain.lvwCD.ListItems(i).ListSubItems(3).ForeColor = vbBlue
        frmMain.lvwCD.ListItems(i).ListSubItems(4).ForeColor = vbBlue
        frmMain.lvwCD.ListItems(i).ListSubItems(5).ForeColor = vbBlue
        frmMain.lvwCD.Refresh
        frmMain.lvwCD.ListItems(i).Tag = msg2
    End If
Next i
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub RefreshCDTracks()", Err.Description, Err.Number
End Sub

Public Sub ProcessFreeDBString(lData As String, lWinsock As Winsock)
On Local Error Resume Next
Dim lAlbum As String, lArtist As String, lCatagory As String, lDiscID As String, i As Integer, msg As String, msg2 As String
If Len(lData) <> 1 And Len(lData) <> 2 And Len(lData) <> 3 Then
    Select Case Left(lData, 3)
    Case "200"
        If InStr(LCase(lData), "hello and welcome") Then
            SendFreeDBQueryString lWinsock, frmMain.lvwCD.Tag
        Else
            lData = Right(lData, Len(lData) - 4)
            lDiscID = ReturnDiscID(frmMain.lvwCD.Tag)
            lAlbum = Trim(Parse(lData, "/", Right(lData, 3)) & Right(lData, 3))
            lData = Left(lData, Len(lData) - Len(lAlbum))
            lCatagory = Left(lData, 1) & Trim(Parse(lData, Left(lData, 1), lDiscID))
            lArtist = Parse(lData, lDiscID, "/")
            lArtist = ReturnDirCompliant(Trim(Right(lArtist, Len(lArtist) - Len(lDiscID) + 1)))
            lAlbum = ReturnDirCompliant(Left(lAlbum, Len(lAlbum) - 2))
            With frmMain.lvwCD
                For i = 1 To .ListItems.Count
                    .ListItems(i).SubItems(1) = lArtist
                    .ListItems(i).SubItems(2) = lAlbum
                Next i
            End With
            msg = ReturnRipPath & Trim(lArtist) & " - " & Trim(lAlbum) & "\"
            MakeNewDir msg
            If Err.Number <> 0 Then Err.Clear
            frmMain.dirCopyTo.Path = msg
            frmMain.dirCopyTo.Refresh
            SendReadDiscID lCatagory, lDiscID, lWinsock
            WriteINI lIniFiles.iDiscDB, frmMain.ReturnCDDiscID, "Artist", lArtist
            WriteINI lIniFiles.iDiscDB, frmMain.ReturnCDDiscID, "Album", lAlbum
        End If
    Case "201"
    Case "210"
        WriteINI lIniFiles.iDiscDB, frmMain.ReturnCDDiscID, "Count", frmMain.lvwCD.ListItems.Count
        For i = 1 To frmMain.lvwCD.ListItems.Count
            If i = frmMain.lvwCD.ListItems.Count Then
                msg = Parse(lData, "TTITLE" & Trim(Str(i - 1)) & "=", "EXTD=")
            Else
                msg = Parse(lData, "TTITLE" & Trim(Str(i - 1)) & "=", "TTITLE" & Trim(Str(i))) & "="
            End If
            If Len(msg) > 7 Then msg = Right(msg, Len(msg) - 7)
            If Len(msg) <> 0 Then msg = Left(msg, Len(msg) - 1)
            If Left(msg, 1) = "=" Then msg = Right(msg, Len(msg) - 1)
            frmMain.lvwCD.ListItems(i).SubItems(3) = ReturnDirCompliant(msg)
            WriteINI lIniFiles.iDiscDB, frmMain.ReturnCDDiscID, Trim(Str(i)), Trim(ReturnDirCompliant(msg))
        Next i
        RefreshCDTracks
    End Select
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub ConnectToFreeDB(lWinsock As Winsock)", Err.Description, Err.Number
End Sub

Public Function DoesDiscIDExistInLibrary(lDiscID As String) As Boolean
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lDiscID) <> 0 Then
    i = Int(ReadINI(lIniFiles.iCDTracks, lDiscID, "TrackCount", 0))
    If i <> 0 Then DoesDiscIDExistInLibrary = True
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function DoesDiscIDExistInLibrary(lDiscID As String) As Boolean", Err.Description, Err.Number
End Function

Public Sub SendFreeDBQueryString(lWinsock As Winsock, lToc As String)
On Local Error GoTo ErrHandler
lWinsock.SendData "cddb query " & ReplacePlusWithSpace(ReturnFreeDBQueryString(lToc)) & vbCrLf
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SendFreeDBQueryString(lWinsock As Winsock, lToc As String)", Err.Description, Err.Number
End Sub

Public Function ReplacePlusWithSpace(lData As String) As String
On Local Error GoTo ErrHandler
ReplacePlusWithSpace = Replace(lData, "+", " ")
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReplacePlusWithSpace(lData As String) As String", Err.Description, Err.Number
End Function

Public Function ReturnFreeDBQueryString(lToc As String) As String
On Local Error GoTo ErrHandler
Dim strTocData() As String, sum As Long, tmp As Long, idx As Integer, msg As String, msg2 As String, lTrackNum As String, mediaID As String, lDiscLength As String
If Len(lToc) <> 0 Then
    msg = Trim$(lToc)
    If (msg = "" Or InStr(1, msg, " ") = 0) Then Exit Function
    strTocData = Split(msg, " ", 100, vbTextCompare)
    lTrackNum = UBound(strTocData)
    lDiscLength = (Val(strTocData(lTrackNum)) \ 75) - (Val(strTocData(0)) \ 75)
    For idx = 0 To lTrackNum - 1
        tmp = Val(strTocData(idx)) \ 75
        Do While tmp > 0
            sum = sum + (tmp Mod 10)
            tmp = tmp \ 10
        Loop
    Next idx
    mediaID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(lDiscLength), 4) & LeftZeroPad(Hex$(lTrackNum), 2))
    msg2 = mediaID & "+" & lTrackNum
    For idx = 0 To lTrackNum - 1
        msg2 = msg2 & "+" & strTocData(idx)
    Next
    ReturnFreeDBQueryString = msg2 & "+" & (Val(strTocData(lTrackNum)) \ 75)
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnFreeDBQueryString(lToc As String) As String", Err.Description, Err.Number
End Function

Public Function ReturnDiscID(lToc As String) As String
On Local Error GoTo ErrHandler
Dim strTocData() As String, sum As Long, tmp As Long, i As Integer, msg As String, lTrackNum As String, lDiscLength As String
If Len(lToc) <> 0 Then
    msg = Trim$(lToc)
    If (msg = "" Or InStr(1, msg, " ") = 0) Then Exit Function
    strTocData = Split(msg, " ", 100, vbTextCompare)
    lTrackNum = UBound(strTocData)
    lDiscLength = (Val(strTocData(lTrackNum)) \ 75) - (Val(strTocData(0)) \ 75)
    For i = 0 To lTrackNum - 1
        tmp = Val(strTocData(i)) \ 75
        Do While tmp > 0
            sum = sum + (tmp Mod 10)
            tmp = tmp \ 10
        Loop
    Next i
    ReturnDiscID = LCase$(LeftZeroPad(Hex$(sum Mod &HFF), 2) & LeftZeroPad(Hex$(lDiscLength), 4) & LeftZeroPad(Hex$(lTrackNum), 2))
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnDiscID(lToc As String) As String", Err.Description, Err.Number
End Function

Private Function LeftZeroPad(s As String, n As Integer) As String
On Local Error GoTo ErrHandler
If Len(s) < n Then
    LeftZeroPad = String$(n - Len(s), "0") & s
Else
    LeftZeroPad = s
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Private Function LeftZeroPad(s As String, n As Integer) As String", Err.Description, Err.Number
End Function
