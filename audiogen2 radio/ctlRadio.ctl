VERSION 5.00
Begin VB.UserControl ctlRadio 
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   1890
   ScaleWidth      =   2175
End
Attribute VB_Name = "ctlRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
'"http://64.236.34.196/stream/2001"
'"http://205.188.234.129:8024"
'"http://64.236.34.97/stream/1006"
'"http://206.98.167.99:8406"
'"http://160.79.1.141:8000"
'"http://206.98.167.99:8006"
'"http://205.188.234.4:8016"
'"http://205.188.234.4:8014"
'"http://server2.somafm.com:8000"
'"http://server1.somafm.com:8082"

Public Sub StopStream()
On Local Error Resume Next
Call BASS_ChannelStop(chan)
End Sub

Public Sub PauseStream()
On Local Error Resume Next
Call BASS_ChannelPause(chan)
End Sub

Public Sub PlayStream(lURL As String)
On Local Error Resume Next
Dim threadid As Long
If (cthread) Then
    Call Beep
Else
    SetURL lURL
    cthread = CreateThread(ByVal 0&, 0, AddressOf OpenURL, 1, 0, threadid)
End If
End Sub

Public Sub RecordStream(lURL As String, lFilename As String)
On Local Error Resume Next

End Sub

Private Sub UserControl_Initialize()
On Local Error Resume Next
ChDrive App.Path
ChDir App.Path
If (Not FileExists(RPP(App.Path) & "bass.dll")) Then
'    MsgBox "BASS.DLL does not exists", vbCritical, "Audiogen 2 Radio"
    Exit Sub
End If
If (BASS_GetVersion <> MakeLong(2, 1)) Then
'    MsgBox "Incorrect version of BASS.DLL. Please download BASS.DLL 2.1", vbCritical, "Audiogen 2 Radio"
    Exit Sub
End If
If (BASS_Init(1, 44100, 0, frmMain.hWnd, 0) = 0) Then
'    Call Error_("Can't initialize device")
    Exit Sub
End If
Set WriteFile = New clsFileIo
cthread = 0
End Sub

Private Function isIDEmode() As Boolean
On Local Error Resume Next
Dim sFileName As String, lCount As Long
sFileName = String(255, 0)
lCount = GetModuleFileName(App.hInstance, sFileName, 255)
sFileName = UCase(GetFileName(Mid(sFileName, 1, lCount)))
isIDEmode = (sFileName = "VB6.EXE")
End Function

Private Sub UserControl_Terminate()
On Local Error Resume Next
Call BASS_Free
End Sub

Private Function FileExists(ByVal fp As String) As Boolean
On Local Error Resume Next
FileExists = (Dir(fp) <> "")
End Function

Private Function RPP(ByVal fp As String) As String
On Local Error Resume Next
RPP = IIf(Mid(fp, Len(fp), 1) <> "\", fp & "\", fp)
End Function

Private Function GetFileName(ByVal fp As String) As String
On Local Error Resume Next
GetFileName = Mid(fp, InStrRev(fp, "\") + 1)
End Function
