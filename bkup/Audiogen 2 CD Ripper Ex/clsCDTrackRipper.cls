VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDTrackRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub CR_SetActiveCDROM Lib "cdrip.dll" (ByVal nActiveDrive As Long)
Private Declare Function CR_OpenRipper Lib "cdrip.dll" (ByRef plBufferSize As Long, ByVal dwStartSector As Long, ByVal dwEndSector As Long) As Long
Private Declare Function CR_CloseRipper Lib "cdrip.dll" () As Long
Private Declare Function CR_GetPercentCompleted Lib "cdrip.dll" () As Long
Private Declare Function CR_GetNumberOfJitterErrors Lib "cdrip.dll" () As Long
Private Declare Function CR_GetJitterPosition Lib "cdrip.dll" () As Long
Private Declare Function CR_GetPeakValue Lib "cdrip.dll" () As Long
Private Declare Function CR_RipChunk Lib "cdrip.dll" (pbtStream As Any, ByRef pNumBytes As Long, bAbort As Long) As Long
Private Declare Sub CR_NormalizeChunk Lib "cdrip.dll" (pbsStream As Any, ByVal nNumSamples As Long, ByVal dScaleFactor As Double)
Private lCDIndex As Long
Private lStartSector As Long
Private lEndSector As Long
Private lJitterErrCount As Long
Private lRipperOpen As Boolean
Private lBufferSize As Long
Private lReadSize As Long
Private lMem As Long
Private lPtrMem As Long
Private lPercentComplete As Long

Public Sub CreateForTrack(lTrack As clsTocEntry)
'On Local Error Resume Next
If CloseRipper() Then
    lStartSector = lTrack.StartSector
    lEndSector = lTrack.EndSector
    lCDIndex = lTrack.fCDIndex
End If
End Sub

Public Sub CreateForSection(cD As clsDrive, ByVal lStartSector As Long, ByVal lEndSector As Long)
'On Local Error Resume Next
If CloseRipper() Then
    lStartSector = 0
    lEndSector = 0
    If (lStartSector >= 0) And (lStartSector < cD.TOC.TotalLengthSectors) Then
        CDRipErrHandler "CreateForSection", 3, False
    Else
        If (lEndSector > lStartSector) And (lEndSector < cD.TOC.TotalLengthSectors) Then
            CDRipErrHandler "CreateForSection", 3, False
        Else
            lStartSector = lStartSector
            lEndSector = lEndSector
            lCDIndex = cD.fCDIndex
        End If
    End If
End If
End Sub

Public Property Get JitterErrorCount() As Long
'On Local Error Resume Next
JitterErrorCount = lJitterErrCount
End Property

Public Property Get PercentComplete() As Long
'On Local Error Resume Next
PercentComplete = lPercentComplete
End Property

Public Function CloseRipper()
'On Local Error Resume Next
Dim eErr As eCDRipErrorCode
If (lRipperOpen) Then
    lRipperOpen = False
    lJitterErrCount = CR_GetNumberOfJitterErrors()
    eErr = CR_CloseRipper()
    CDRipErrHandler "cCDTrackRipper.CloseRipper", eErr, True
    If (eErr = CDEX_OK) Then
        CloseRipper = True
    End If
Else
    CloseRipper = True
End If
End Function

Private Function CreateBuffer(ByVal lSize As Long) As Boolean
'On Local Error Resume Next
If (lSize <= 1) Then
    Debug.Print lSize
    Debug.Assert "" = "Inappropriate buffer size"
    Exit Function
End If
If Not (lSize = lBufferSize) Then DestroyBuffer
If (lPtrMem = 0) Then
    lMem = LocalAlloc(GPTR, lSize)
    If (lMem = 0) Then
    
    Else
        lPtrMem = LocalLock(lMem)
        If (lPtrMem = 0) Then
            DestroyBuffer
        Else
            lBufferSize = lSize
            CreateBuffer = True
        End If
    End If
Else
    CreateBuffer = True
End If
End Function

Private Sub DestroyBuffer()
'On Local Error Resume Next
If Not (lPtrMem = 0) Then
    LocalUnlock lMem
    lPtrMem = 0
End If
If Not (lMem = 0) Then
    LocalFree lMem
    lMem = 0
End If
lReadSize = 0
lBufferSize = 0
End Sub

Public Function OpenRipper() As Boolean
'On Local Error Resume Next
Dim eErr As eCDRipErrorCode, lStartSector As Long, lEndSector As Long, lBufferSize As Long
If CloseRipper() Then
    CR_SetActiveCDROM lCDIndex
    lPercentComplete = 0
    lStartSector = lStartSector
    lEndSector = lEndSector
    eErr = CR_OpenRipper(lBufferSize, lStartSector, lEndSector)
    CDRipErrHandler "cCDTrackRipper.OpenRipper", eErr, True
    If (eErr = CDEX_OK) Then
        If CreateBuffer(lBufferSize) Then
            lRipperOpen = True
            lJitterErrCount = 0
            OpenRipper = True
        Else
            CDRipErrHandler "cCDTrackRipper.CreateBuffer", 7, False
        End If
    End If
End If
End Function

Public Function Read() As Boolean
'On Local Error Resume Next
Dim bAbort As Long, lNumBytes As Long, eErr As eCDRipErrorCode
If (lRipperOpen) Then
    eErr = CR_RipChunk(ByVal lPtrMem, lNumBytes, bAbort)
    lReadSize = lNumBytes
    lPercentComplete = CR_GetPercentCompleted()
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
    lReadSize = 0
    Read = False
End If
End Function

Public Property Get ReadBufferPtr() As Long
'On Local Error Resume Next
ReadBufferPtr = lPtrMem
End Property

Public Property Get ReadBufferSize() As Long
'On Local Error Resume Next
ReadBufferSize = lReadSize
End Property

Private Sub Class_Terminate()
'On Local Error Resume Next
Dim bWasOpen As Boolean
bWasOpen = lRipperOpen
DestroyBuffer
CloseRipper
If (bWasOpen) Then
    Debug.Print "Warning: cCDTrackRipper instance terminated whilst ripper open"
    Debug.Assert Not (lRipperOpen)
End If
End Sub
