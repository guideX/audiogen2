Attribute VB_Name = "mdlMultimedia"
Option Explicit
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Dim glo_from As Long
Dim glo_to As Long
Dim glo_AliasName As String
Dim glo_hWnd As Long

Public Function OpenMediaFile(lFilename As String, Optional lAutoPlay As Boolean) As Boolean
Dim lAliasName As String, typeDevice As String, lResult As String, lResultMsg As Integer, msg As String
If LCase(lFilename) = "mp3" Or LCase(lFilename) = "wav" Or LCase(lFilename) = "mpg" Or LCase(lFilename) = "avi" Or lFilename = "" Then Exit Function
frmMain.ctlMovie1.OpenMovie lFilename
'frmMain.ctlMovie1.PlayMovie
frmMain.ctlMovie1.Visible = True
frmMain.ctlMP3.Visible = False
frmMain.picVideo.Visible = False
frmMain.cmdPlay.Enabled = True
frmMain.cmdPausePlayback.Enabled = True
frmMain.cmdStop.Enabled = True
frmMain.cmdOpen.Enabled = False
frmMain.lblFilename.Caption = GetFileTitle(lFilename)
frmMain.lblFilename.Tag = lFilename
If FindListViewIndexByFileTitle(GetFileTitle(lFilename), frmMain.lvwProcesses) = 0 Then AddProcess GetFileTitle(lFilename), "", Right(LCase(lFilename), 3), "Play"
If lAutoPlay = True Then frmMain.ctlMovie1.PlayMovie
'Exit Function
'CloseAll
'If DoesFileExist(lFilename) = True Then
'    lAliasName = GetFileTitle(lFilename)
'    If Right(lFilename, 4) = ".avi" Then
'        typeDevice = "AviVideo"
'    ElseIf Right(lFilename, 4) = ".mp3" Then
'        typeDevice = "MP3"
'    ElseIf Right(lFilename, 4) = ".rmi" Or Right(lFilename, 4) = ".mid" Then
'        typeDevice = "sequencer"
'    ElseIf Right(lFilename, 4) = ".VOB" Or Right(lFilename, 4) = ".vob" Then
'        lResultMsg = MsgBox("Would you like to open this file as MPEG video? Click false for DVD Video.", vbYesNo + vbQuestion)
'        If lResultMsg = vbYes Then
'            typeDevice = "MPEGVideo"
'        Else
'            typeDevice = "DvDVideo"
'        End If
'    Else
'        typeDevice = "MPEGVideo"
'    End If
'    If typeDevice = "MP3" Then
'        frmMain.ctlMP3.Visible = True
'        frmMain.ctlMP3.Stop
'        frmMain.ctlMP3.OscilloType = otSpectrum
'        frmMain.ctlMP3.Tag = GetFileTitle(lFilename)
'        frmMain.lblFilename.Caption = lAliasName
'        frmMain.lblFilename.Tag = lFilename
'        frmMain.cmdPlay.Enabled = True
'        ResizeVideo
'        If lAutoPlay = True Then
'            frmMain.ctlMP3.Play lFilename
'            frmMain.cmdStop.Enabled = True
'            frmMain.cmdPausePlayback.Enabled = True
'            frmMain.picVideo.Visible = True
'        End If
'        Exit Function
'    End If
'    frmMain.ctlMP3.Stop
'    frmMain.ctlMP3.Visible = False
'    lResult = OpenMultimedia(frmMain.picVideo.hWnd, lAliasName, lFilename, typeDevice)
'    If lResult = "Success" Then
'        frmMain.lblFilename.Caption = lAliasName
'        frmMain.lblFilename.Tag = lFilename
'        frmMain.cmdPlay.Enabled = True
'        If lAutoPlay = True Then
'            frmMain.cmdStop.Enabled = True
'            frmMain.cmdPausePlayback.Enabled = True
'            frmMain.picVideo.Visible = True
'            ResizeVideo
'            PlayMediaFile lFilename
'        End If
'        OpenMediaFile = True
''        lResult = PlayMultimedia(lAliasName, 0, 100)
''        MsgBox lResult
'        'OptnChannelAllOn(Index).value = True
'        'LbActualCx(Index).Caption = GetSize(lAliasName, "cx")
'        'LbActualCy(Index).Caption = GetSize(lAliasName, "cy")
'        'LbFramesPerSecond(Index) = GetFramesPerSecond(lAliasName)
'        'LbTotalFrames(Index) = GetTotalframes(lAliasName)
'        'LbTotalTime(Index) = GetTotalTimeByMS(lAliasName) / 1000
'        'SliderMoveMultimedia(Index).Max = LbTotalFrames(Index) / (LbFramesPerSecond(Index) * 2)
'        'TimerMisc(Index).Enabled = True
'    Else
'        OpenMediaFile = False
'        'MsgBox lResult
'    End If
'End If
End Function

Public Sub PlayMediaFile(lFilename As String)
Dim lAliasName As String, lResult As String
lAliasName = GetFileTitle(lFilename)
lResult = PlayMultimedia(lAliasName, vbNullString, vbNullString)
If lResult = "Success" Then
    frmMain.tmrProgress.Enabled = True
    frmMain.cmdPausePlayback.Enabled = True
    frmMain.cmdPlay.Enabled = False
    frmMain.cmdStop.Enabled = True
    frmMain.cmdOpen.Enabled = True
End If
End Sub

Public Function OpenMultimedia(hWnd As Long, AliasName As String, FileName As String, typeDevice As String) As String
Dim cmdToDo As String * 255
Dim dwReturn As Long
Dim ret As String * 128
Dim tmp As String * 255
Dim lenShort As Long
Dim ShortPathAndFile As String
Const WS_CHILD = &H40000000
lenShort = GetShortPathName(FileName, tmp, 255)
ShortPathAndFile = Left$(tmp, lenShort)
cmdToDo = "open " & ShortPathAndFile & " type " & typeDevice & " Alias " & AliasName & " parent " & hWnd & " Style " & WS_CHILD
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    OpenMultimedia = ret: Exit Function
End If
OpenMultimedia = "Success"
End Function

Public Function PlayMultimedia(AliasName As String, from_where As String, to_where As String) As String
If from_where = vbNullString Then from_where = 0
If to_where = vbNullString Then to_where = GetTotalframes(AliasName)
If AliasName = glo_AliasName Then
    glo_from = from_where
    glo_to = to_where
End If
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
cmdToDo = "play " & AliasName & " from " & from_where & " to " & to_where
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PlayMultimedia = ret
    Exit Function
End If
PlayMultimedia = "Success"
End Function

Public Function CloseMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    CloseMultimedia = ret
    Exit Function
End If
If AliasName = glo_AliasName Then
KillTimer glo_hWnd, 500
End If
CloseMultimedia = "Success"
End Function

Public Function PauseMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Pause " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PauseMultimedia = ret
    Exit Function
End If
PauseMultimedia = "Success"
End Function

Public Function StopMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Stop " & AliasName, 0&, 0&, 0&) 'stop
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    StopMultimedia = ret
    Exit Function
End If
StopMultimedia = "Success"
End Function

Public Function ResumeMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Resume " & AliasName, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    ResumeMultimedia = ret
    Exit Function
End If
ResumeMultimedia = "Success"
End Function

Public Function GetStatusMultimedia(AliasName As String) As String
Dim dwReturn As Long
Dim status As String * 128
Dim ret As String * 128
dwReturn = mciSendString("status " & AliasName & " mode", status, 128, 0&)
If Not dwReturn = 0 Then
    GetStatusMultimedia = "ERROR"
    Exit Function
End If
Dim i As Integer
Dim CharA As String
Dim RChar As String
RChar = Right$(status, 1)
For i = 1 To Len(status)
    CharA = Mid(status, i, 1)
    If CharA = RChar Then Exit For
    GetStatusMultimedia = GetStatusMultimedia + CharA
Next i
End Function

Public Function GetTotalframes(AliasName As String) As Long
Dim dwReturn As Long
Dim Total As String * 128
dwReturn = mciSendString("set " & AliasName & " time format frames", Total, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", Total, 128, 0&)
If Not dwReturn = 0 Then
    GetTotalframes = -1
    Exit Function
End If
GetTotalframes = Val(Total)
End Function

Public Function GetTotalTimeByMS(AliasName As String) As Long
Dim dwReturn As Long
Dim TotalTime As String * 128
dwReturn = mciSendString("set " & AliasName & " time format ms", TotalTime, 128, 0&)
dwReturn = mciSendString("status " & AliasName & " length", TotalTime, 128, 0&)
mciSendString "set " & AliasName & " time format frames", 0&, 0&, 0&
If Not dwReturn = 0 Then
    GetTotalTimeByMS = -1
    Exit Function
End If
GetTotalTimeByMS = Val(TotalTime)
End Function

Public Function MoveMultimedia(AliasName As String, to_where As Long) As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("seek " & AliasName & " to " & to_where, 0&, 0&, 0&)
mciSendString "Play " & AliasName, 0&, 0&, 0&
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    MoveMultimedia = ret
    Exit Function
End If
MoveMultimedia = "Success"
End Function

Public Function GetCurrentMultimediaPos(AliasName As String) As Long
Dim dwReturn As Long
Dim pos As String * 128
dwReturn = mciSendString("status " & AliasName & " position", pos, 128, 0&)
If Not dwReturn = 0 Then
    GetCurrentMultimediaPos = -1
    Exit Function
End If
GetCurrentMultimediaPos = Val(pos)
End Function

Public Function PutMultimedia(hWnd As Long, AliasName As String, Left As Long, Top As Long, Width As Long, Height As Long) As String
Dim dwReturn As Long
Dim ret As String * 128
If Width = 0 Or Height = 0 Then
    Dim rec As RECT
    Call GetWindowRect(hWnd, rec)
    Width = rec.Right - rec.Left
    Height = rec.Bottom - rec.Top
End If
dwReturn = mciSendString("put " & AliasName & " window at " & Left & " " & Top & " " & Width & " " & Height, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    PutMultimedia = ret
    Exit Function
End If
PutMultimedia = "Success"
End Function

Public Function GetPercent(AliasName As String) As Long
'On Local Error Resume Next
Dim TotalFrames As Long
Dim currframe As Long
TotalFrames = GetTotalframes(AliasName)
currframe = GetCurrentMultimediaPos(AliasName)
If TotalFrames = -1 Or currframe = -1 Then
    GetPercent = -1
    Exit Function
End If
GetPercent = currframe * 100 / TotalFrames
End Function

Public Function GetFramesPerSecond(AliasName As String) As Long
Dim TotalFrames As Long
Dim TotalTime As Long
TotalTime = GetTotalTimeByMS(AliasName)
TotalFrames = GetTotalframes(AliasName)
If TotalFrames = -1 Or TotalTime = -1 Then
    GetFramesPerSecond = -1
    Exit Function
End If
GetFramesPerSecond = TotalFrames / (TotalTime / 1000)
End Function

Public Function GetSize(AliasName As String, CxOrCy As String) As Long
If Not CxOrCy = "cx" And Not CxOrCy = "cy" Then GetSize = -1: Exit Function
Dim dwReturn As Long
Dim Size As String * 128
Dim s1, s2, s3, Width, Height As Long
dwReturn = mciSendString("Where " & AliasName & " destination", Size, 128, 0&)
If Not dwReturn = 0 Then
    GetSize = -1
    Exit Function
End If
s1 = InStr(1, Size, " "): s2 = InStr(s1 + 1, Size, " "): s1 = InStr(s2 + 1, Size, " ")
Width = Mid(Size, s2, s1 - s2): Height = Mid(Size, s1 + 1)
If CxOrCy = "cx" Then
GetSize = Width
ElseIf CxOrCy = "cy" Then
GetSize = Height
End If
End Function

Public Function CloseAll() As String
Dim dwReturn As Long
Dim ret As String * 128
dwReturn = mciSendString("Close All", 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    CloseAll = ret
    Exit Function
End If
CloseAll = "Success"
End Function

Public Function ChannelsControl(AliasName As String, channel As String, OnOrOFF As String) As String
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
cmdToDo = "set " & AliasName & " audio " & channel & " " & OnOrOFF
dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
    ChannelsControl = ret
    Exit Function
End If
ChannelsControl = "Success"
End Function

Public Function AreMultimediaAtEnd(AliasName As String, lastFrame As Long) As Boolean
Dim currpos As Long
If lastFrame = 0 Then lastFrame = GetTotalframes(AliasName)
currpos = Val(GetCurrentMultimediaPos(AliasName))
If currpos = -1 Or lastFrame = -1 Then
    AreMultimediaAtEnd = False
    Exit Function
End If
If lastFrame = currpos Or (lastFrame - 1) < currpos Then
AreMultimediaAtEnd = True
Else
AreMultimediaAtEnd = False
End If
End Function

Public Function SetAutoRepeat(hWnd As Long, AliasName As String, first_frame As String, last_frame As String, autoTrueOrFalse As Boolean) As Boolean
Dim Result As String
If first_frame = vbNullString Then first_frame = 0
If last_frame = vbNullString Then last_frame = GetTotalframes(AliasName)
glo_from = first_frame
glo_to = last_frame
glo_hWnd = hWnd
If autoTrueOrFalse = True Then
    glo_AliasName = AliasName
    Result = SetTimer(hWnd, 500, 100, AddressOf TimerFunction)
Else
    glo_AliasName = vbNullString
    Result = KillTimer(hWnd, 500)
End If
If Result = 0 Then
    SetAutoRepeat = False
Else
    SetAutoRepeat = True
End If
End Function

Sub TimerFunction()
Dim currpos As Long
Dim Result As String
currpos = Val(GetCurrentMultimediaPos(glo_AliasName))
If currpos = -1 Then Exit Sub
If Val(glo_to) = currpos Or (Val(glo_to) - 1) < currpos Then
    Result = PlayMultimedia(glo_AliasName, Str(glo_from), Str(glo_to))
    If Not Result = "Success" Then KillTimer glo_hWnd, 500
End If
End Sub

Public Sub SetDefaultDevice(typeDevice As String, drvDefaultDevice As String)
Dim Res As String
Dim tmp As String * 255
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = WritePrivateProfileString("MCI", typeDevice, drvDefaultDevice, Windir & "\" & "system.ini")
End Sub

Public Function GetDefaultDevice(typeDevice As String) As String
Dim tmp As String * 255
Dim Res As String
Dim Windir As String
Res = GetWindowsDirectory(tmp, 255)
Windir = Left$(tmp, Res)
Res = GetPrivateProfileString("MCI", typeDevice, "None", tmp, 255, Windir & "\" & "system.ini")
GetDefaultDevice = Left$(tmp, Res)
End Function

Public Sub test()
Dim cmdToDo As String * 128
Dim dwReturn As Long
Dim ret As String * 128
Dim Value As String * 128
cmdToDo = "status movie2 channels"
dwReturn = mciSendString(cmdToDo, Value, 128, 0&)
If Not dwReturn = 0 Then
    mciGetErrorString dwReturn, ret, 128
MsgBox ret
End If
End Sub
