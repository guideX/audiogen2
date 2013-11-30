Attribute VB_Name = "mdlNetRadio"
Option Explicit
Public chan As Long
Private mURL As String
Public WriteFile As clsFileIo
Public FileIsOpen As Boolean, GotHeader As Boolean
Public DownloadStarted As Boolean, DoDownload As Boolean
Public DlOutput As String, SongNameUpdate As Boolean
Public cthread As Long
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Sub Main()

End Sub

Public Sub SetURL(lURL As String)
On Local Error Resume Next
mURL = lURL
End Sub

Public Sub Error_(ByVal Message As String)
On Local Error Resume Next
MsgBox Message & vbCrLf & vbCrLf & "Error Code : " & BASS_ErrorGetCode, vbInformation, "Audiogen 2"
End Sub

Sub DoMeta(ByVal meta As Long)
On Local Error Resume Next
Dim p As String, tmpMeta As String
If meta = 0 Then Exit Sub
tmpMeta = VBStrFromAnsiPtr(meta)
If ((Mid(tmpMeta, 1, 13) = "StreamTitle='")) Then
    GotHeader = False
    DownloadStarted = False
    p = Mid(tmpMeta, 14)
    DlOutput = App.Path & "\" & RemoveSpecialChar(Mid(p, 1, InStr(p, ";") - 2)) & ".mp3"
End If
End Sub

Sub MetaSync(ByVal handle As Long, ByVal channel As Long, ByVal data As Long, ByVal user As Long)
On Local Error Resume Next
Call DoMeta(data)
End Sub

Public Sub OpenURL()
On Local Error Resume Next
Dim icyPTR As Long, tmpICY As String
Call BASS_StreamFree(chan)
chan = BASS_StreamCreateURL(mURL, 0, BASS_STREAM_META Or BASS_STREAM_STATUS, AddressOf SUBDOWNLOADPROC, 0)
If chan = 0 Then
    Call Error_("Can't play the stream")
Else
    icyPTR = BASS_StreamGetTags(chan, BASS_TAG_ICY)
    If (icyPTR) Then
        Do
            tmpICY = VBStrFromAnsiPtr(icyPTR)
            icyPTR = icyPTR + Len(tmpICY) + 1
        Loop While (tmpICY <> "")
    End If
    Call DoMeta(BASS_StreamGetTags(chan, BASS_TAG_META))
    Call BASS_ChannelSetSync(chan, BASS_SYNC_META, 0, AddressOf MetaSync, 0)
    Call BASS_ChannelPlay(chan, BASSFALSE)
    
End If
Call CloseHandle(cthread)
cthread = 0
End Sub

Public Sub SUBDOWNLOADPROC(ByVal buffer As Long, ByVal length As Long, ByVal user As Long)
On Local Error Resume Next
If (buffer And length = 0) Then
'    frmNetRadio.lblBPS.Caption = VBStrFromAnsiPtr(buffer)
    Exit Sub
End If
If (Not DoDownload) Then
    DownloadStarted = False
    Call WriteFile.CloseFile
    Exit Sub
End If
If (Trim(DlOutput) = "") Then Exit Sub
If (Not DownloadStarted) Then
    DownloadStarted = True
    Call WriteFile.CloseFile
    If (WriteFile.OpenFile(DlOutput)) Then
        SongNameUpdate = False
    Else
        SongNameUpdate = True
        GotHeader = False
    End If
End If
If (Not SongNameUpdate) Then
    If (length) Then
        Call WriteFile.WriteBytes(buffer, length)
    Else
        Call WriteFile.CloseFile
        GotHeader = False
    End If
Else
    DownloadStarted = False
    Call WriteFile.CloseFile
    GotHeader = False
End If
End Sub

Public Function RemoveSpecialChar(strFileName As String)
On Local Error Resume Next
Dim i As Byte, SpecialChar As Boolean, SelChar As String, OutFileName As String
For i = 1 To Len(strFileName)
    SelChar = Mid(strFileName, i, 1)
    SpecialChar = InStr(":/\?*|<>" & Chr$(34), SelChar) > 0
    If (Not SpecialChar) Then
        OutFileName = OutFileName & SelChar
        SpecialChar = False
    Else
        OutFileName = OutFileName
        SpecialChar = False
    End If
Next i
RemoveSpecialChar = OutFileName
End Function
