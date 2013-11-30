VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCDTrackRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Declare Function LocalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Sub CR_SetActiveCDROM Lib "cdrip.dll" (ByVal nActiveDrive As Long)
Private Declare Function CR_OpenRipper Lib "cdrip.dll" (ByRef plBufferSize As Long, ByVal dwStartSector As Long, ByVal dwEndSector As Long) As Long
Private Declare Function CR_CloseRipper Lib "cdrip.dll" () As Long
Private Declare Function CR_GetPercentCompleted Lib "cdrip.dll" () As Long
Private Declare Function CR_GetNumberOfJitterErrors Lib "cdrip.dll" () As Long
Private Declare Function CR_GetJitterPosition Lib "cdrip.dll" () As Long
Private Declare Function CR_GetPeakValue Lib "cdrip.dll" () As Long
Private Declare Function CR_RipChunk Lib "cdrip.dll" (pbtStream As Any, ByRef pNumBytes As Long, bAbort As Long) As Long
Private Declare Sub CR_NormalizeChunk Lib "cdrip.dll" (pbsStream As Any, ByVal nNumSamples As Long, ByVal dScaleFactor As Double)
Private m_cdIndex As Long
Private m_startSector As Long
Private m_endSector As Long
Private m_iJitterErrCount As Long
Private m_bRipperOpen As Boolean
Private m_lBufferSize As Long
Private m_lReadSize As Long
Private m_hMem As Long
Private m_lPtrMem As Long
Private m_lPercentComplete As Long

Public Sub CreateForTrack(cTrack As cTocEntry)
On Local Error Resume Next
If CloseRipper() Then
    m_startSector = cTrack.StartSector
    m_endSector = cTrack.EndSector
    m_cdIndex = cTrack.fCDIndex
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub CreateForTrack(cTrack As cTocEntry)", Err.Description, Err.Number
End Sub

Public Sub CreateForSection(cD As cDrive, ByVal lStartSector As Long, ByVal lEndSector As Long)
On Local Error Resume Next
If CloseRipper() Then
    m_startSector = 0
    m_endSector = 0
    If (lStartSector >= 0) And (lStartSector < cD.TOC.TotalLengthSectors) Then
        CDRipErrHandler "CreateForSection", 3, False
    Else
        If (lEndSector > lStartSector) And (lEndSector < cD.TOC.TotalLengthSectors) Then
            CDRipErrHandler "CreateForSection", 3, False
        Else
            m_startSector = lStartSector
            m_endSector = lEndSector
            m_cdIndex = cD.fCDIndex
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub CreateForSection(cD As cDrive, ByVal lStartSector As Long, ByVal lEndSector As Long)", Err.Description, Err.Number
End Sub

Public Property Get JitterErrorCount() As Long
On Local Error Resume Next
JitterErrorCount = m_iJitterErrCount
End Property

Public Property Get PercentComplete() As Long
On Local Error Resume Next
PercentComplete = m_lPercentComplete
End Property

Public Function CloseRipper()
On Local Error Resume Next
Dim eErr As ECDRipErrorCode
If (m_bRipperOpen) Then
    m_bRipperOpen = False
    m_iJitterErrCount = CR_GetNumberOfJitterErrors()
    eErr = CR_CloseRipper()
    CDRipErrHandler "cCDTrackRipper.CloseRipper", eErr, True
    If (eErr = CDEX_OK) Then
        CloseRipper = True
    End If
Else
    CloseRipper = True
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Function CloseRipper()", Err.Description, Err.Number
End Function

Private Function CreateBuffer(ByVal lSize As Long) As Boolean
On Local Error Resume Next
If (lSize <= 1) Then
    Debug.Print lSize
    Debug.Assert "" = "Inappropriate buffer size"
    Exit Function
End If
If Not (lSize = m_lBufferSize) Then
    DestroyBuffer
End If
If (m_lPtrMem = 0) Then
    m_hMem = LocalAlloc(GPTR, lSize)
    If (m_hMem = 0) Then
    Else
        m_lPtrMem = LocalLock(m_hMem)
        If (m_lPtrMem = 0) Then
            DestroyBuffer
        Else
            m_lBufferSize = lSize
            CreateBuffer = True
        End If
    End If
Else
    CreateBuffer = True
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Function CloseRipper()", Err.Description, Err.Number
End Function

Private Sub DestroyBuffer()
On Local Error Resume Next
If Not (m_lPtrMem = 0) Then
    LocalUnlock m_hMem
    m_lPtrMem = 0
End If
If Not (m_hMem = 0) Then
    LocalFree m_hMem
    m_hMem = 0
End If
m_lReadSize = 0
m_lBufferSize = 0
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub DestroyBuffer()", Err.Description, Err.Number
End Sub

Public Function OpenRipper() As Boolean
On Local Error Resume Next
Dim eErr As ECDRipErrorCode, lStartSector As Long, lEndSector As Long, lBufferSize As Long
If CloseRipper() Then
    CR_SetActiveCDROM m_cdIndex
    m_lPercentComplete = 0
    lStartSector = m_startSector
    lEndSector = m_endSector
    eErr = CR_OpenRipper(lBufferSize, lStartSector, lEndSector)
    CDRipErrHandler "cCDTrackRipper.OpenRipper", eErr, True
    If (eErr = CDEX_OK) Then
        If CreateBuffer(lBufferSize) Then
            m_bRipperOpen = True
            m_iJitterErrCount = 0
            OpenRipper = True
        Else
            CDRipErrHandler "cCDTrackRipper.CreateBuffer", 7, False
        End If
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Function OpenRipper() As Boolean", Err.Description, Err.Number
End Function

Public Function Read() As Boolean
On Local Error Resume Next
Dim bAbort As Long, lNumBytes As Long, eErr As ECDRipErrorCode
If (m_bRipperOpen) Then
    eErr = CR_RipChunk(ByVal m_lPtrMem, lNumBytes, bAbort)
    m_lReadSize = lNumBytes
    m_lPercentComplete = CR_GetPercentCompleted()
    Select Case eErr
    Case CDEX_RIPPING_INPROGRESS, CDEX_OK
        Read = True
    Case CDEX_RIPPING_DONE
        CloseRipper
        Read = True
    Case CDEX_JITTER_ERROR
        Read = True
    Case CDEX_FILEOPEN_ERROR
        CloseRipper
        CDRipErrHandler "cCDTrackRipper.Read", eErr, True
        Read = False
    Case CDEX_ERROR
        CloseRipper
        CDRipErrHandler "cCDTrackRipper.Read", eErr, True
        Read = False
    End Select
Else
    m_lReadSize = 0
    Read = False
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Function Read() As Boolean", Err.Description, Err.Number
End Function

Public Property Get ReadBufferPtr() As Long
On Local Error Resume Next
ReadBufferPtr = m_lPtrMem
End Property

Public Property Get ReadBufferSize() As Long
On Local Error Resume Next
ReadBufferSize = m_lReadSize
End Property

Private Sub Class_Terminate()
On Local Error Resume Next
Dim bWasOpen As Boolean
bWasOpen = m_bRipperOpen
DestroyBuffer
CloseRipper
If (bWasOpen) Then
    Debug.Print "Warning: cCDTrackRipper instance terminated whilst ripper open"
    Debug.Assert Not (m_bRipperOpen)
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub Class_Terminate()", Err.Description, Err.Number
End Sub
