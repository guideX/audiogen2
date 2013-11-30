Attribute VB_Name = "mdlMCI"
Option Explicit
Private Const MCI_OPEN = &H803
Private Const MCI_CLOSE = &H804
Private Const MCI_FORMAT_MSF = 2
Private Const MCI_OPEN_ELEMENT = &H200&
Private Const MCI_OPEN_TYPE = &H2000&
Private Const MCI_SET = &H80D
Private Const MCI_SET_TIME_FORMAT = &H400&
Private Const MCI_STATUS = &H814
Private Const MCI_STATUS_ITEM = &H100&
Private Const MCI_STATUS_LENGTH = &H1&
Private Const MCI_STATUS_NUMBER_OF_TRACKS = &H3&
Private Const MCI_STATUS_POSITION = &H2&
Private Const MCI_TRACK = &H10&
Private Type MCI_OPEN_PARMS
    dwCallback As Long
    wDeviceID As Long
    lpstrDeviceType As String
    lpstrElementName As String
    lpstrAlias As String
End Type
Private Type MCI_SET_PARMS
    dwCallback As Long
    dwTimeFormat As Long
    dwAudio As Long
End Type
Private Type MCI_STATUS_PARMS
    dwCallback As Long
    dwReturn As Long
    dwItem As Long
    dwTrack As Integer
End Type
Private Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciSendCommandA" (ByVal fdwError As Long, ByRef lpszErrorText As String, ByVal cchErrorText As Integer)
Private mciOpenParms As MCI_OPEN_PARMS
Private mciSetParms As MCI_SET_PARMS
Private mciStatusParms As MCI_STATUS_PARMS
Private m_DevID As Long

Public Function ReturnTrackCount(ByVal lDrive As String) As Integer
On Local Error Resume Next
Dim l As Long
mciOpenParms.lpstrDeviceType = "cdaudio"
mciOpenParms.lpstrElementName = lDrive
l = mciSendCommand(0, MCI_OPEN, (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT), mciOpenParms)
m_DevID = mciOpenParms.wDeviceID
mciSetParms.dwTimeFormat = MCI_FORMAT_MSF
l = mciSendCommand(m_DevID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
If l = 0 Then
    mciStatusParms.dwItem = MCI_STATUS_NUMBER_OF_TRACKS
    l = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM, mciStatusParms)
    If l = 0 Then
        ReturnTrackCount = mciStatusParms.dwReturn
    End If
End If
l = mciSendCommand(m_DevID, MCI_CLOSE, 0, 0)
End Function

Public Function ReturnMediaTOC(lDrive As String) As String
On Local Error Resume Next
Dim l As Long, i As Integer, t As Integer, mins As Long, secs As Long, frms As Long, offst As Long, s As String
mciOpenParms.lpstrDeviceType = "cdaudio"
mciOpenParms.lpstrElementName = lDrive
l = mciSendCommand(0, MCI_OPEN, (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT), mciOpenParms)
If l = 0 Then
    m_DevID = mciOpenParms.wDeviceID
    mciSetParms.dwTimeFormat = MCI_FORMAT_MSF
    l = mciSendCommand(m_DevID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
    If l = 0 Then
        mciStatusParms.dwItem = MCI_STATUS_NUMBER_OF_TRACKS
        l = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM, mciStatusParms)
        If l = 0 Then
            t = mciStatusParms.dwReturn
            For i = 1 To t
                mciStatusParms.dwItem = MCI_STATUS_POSITION
                mciStatusParms.dwTrack = i
                l = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
                If l = 0 Then
                    mins = (mciStatusParms.dwReturn) And &HFF
                    secs = (mciStatusParms.dwReturn \ 256) And &HFF
                    frms = (mciStatusParms.dwReturn \ 65536) And &HFF
                    offst = (mins * 60 * 75) + (secs * 75) + frms
                    s = s & " " & Format$(offst)
                End If
            Next i
            mciStatusParms.dwItem = MCI_STATUS_LENGTH
            mciStatusParms.dwTrack = t
            l = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
            If l = 0 Then
                mins = (mciStatusParms.dwReturn) And &HFF
                secs = (mciStatusParms.dwReturn \ 256) And &HFF
                frms = ((mciStatusParms.dwReturn \ 65536) And &HFF) + 1
                offst = offst + (mins * 60 * 75) + (secs * 75) + frms
                s = s & " " & Format$(offst)
                ReturnMediaTOC = Trim(s)
            End If
        End If
    End If
End If
l = mciSendCommand(m_DevID, MCI_CLOSE, 0, 0)
End Function

Public Function ReturnTrackLength(lDrive As String, lTrackNumber As Integer) As String
On Local Error Resume Next
Dim l As Long, i As Integer, t As Integer, mins As Long, secs As Long, frms As Long, offst As Long, s As String
mciOpenParms.lpstrDeviceType = "cdaudio"
mciOpenParms.lpstrElementName = lDrive
l = mciSendCommand(0, MCI_OPEN, (MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT), mciOpenParms)
If l = 0 Then
    m_DevID = mciOpenParms.wDeviceID
    mciSetParms.dwTimeFormat = MCI_FORMAT_MSF
    l = mciSendCommand(m_DevID, MCI_SET, MCI_SET_TIME_FORMAT, mciSetParms)
    If l = 0 Then
        mciStatusParms.dwItem = MCI_STATUS_LENGTH
        mciStatusParms.dwTrack = lTrackNumber
        l = mciSendCommand(m_DevID, MCI_STATUS, MCI_STATUS_ITEM Or MCI_TRACK, mciStatusParms)
        If l = 0 Then
            mins = (mciStatusParms.dwReturn) And &HFF
            secs = (mciStatusParms.dwReturn \ 256) And &HFF
            ReturnTrackLength = Format(mins, "00") & ":" & Format(secs, "00")
        End If
    End If
End If
l = mciSendCommand(m_DevID, MCI_CLOSE, 0, 0)
End Function

