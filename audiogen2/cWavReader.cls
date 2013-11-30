VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWavReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type
Private Type mmioinfo
    dwFlags As Long
    fccIOProc As Long
    pIOProc As Long
    wErrorRet As Long
    htask As Long
    cchBuffer As Long
    pchBuffer As String
    pchNext As String
    pchEndRead As String
    pchEndWrite As String
    lBufOffset As Long
    lDiskOffset As Long
    adwInfo(4) As Long
    dwReserved1 As Long
    dwReserved2 As Long
    hmmio As Long
End Type
Private Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type
Private Declare Function mmioClose Lib "winmm.dll" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioDescend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, lpckParent As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioDescendParent Lib "winmm.dll" Alias "mmioDescend" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal X As Long, ByVal uFlags As Long) As Long
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As mmioinfo, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioRead Lib "winmm.dll" (ByVal hmmio As Long, ByVal pch As Long, ByVal cch As Long) As Long
Private Declare Function mmioReadString Lib "winmm.dll" Alias "mmioRead" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Private Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Const MMIO_READ = &H0
Private Const MMIO_FINDCHUNK = &H10
Private Const MMIO_FINDRIFF = &H20
Private Const MM_WOM_DONE = &H3BD
Private Const MMSYSERR_NOERROR = 0
Private Const SEEK_CUR = 1
Private Const SEEK_END = 2
Private Const SEEK_SET = 0
Private Const TIME_BYTES = &H4
Private Const WHDR_DONE = &H1
Private Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal numBytes As Long)
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb As Long)
Private Declare Sub CopyMemoryFromString Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As String, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private m_hMMioIn As Long
Private m_lPtrFormat As Long
Private m_tFormat As WAVEFORMATEX
Private m_lStartPos As Long
Private m_lDataOffset As Long
Private m_lAudioLength As Long
Private m_hMemBuffer As Long
Private m_lPtrBuffer As Long
Private m_lBufferSize As Long
Private m_lCurrentReadSize As Long
Private m_sFile As String
Private Const cErrBase = 29670

Public Property Get Filename() As String
On Local Error Resume Next
Filename = m_sFile
End Property

Public Property Let Filename(ByVal sFile As String)
On Local Error Resume Next
If (OpenFile(sFile)) Then
    m_sFile = sFile
End If
End Property

Private Function AllocateBuffer() As Boolean
On Local Error Resume Next
FreeBuffer
m_hMemBuffer = GlobalAlloc(GPTR, m_lBufferSize * 4)
If Not (m_hMemBuffer = 0) Then
    m_lPtrBuffer = GlobalLock(m_hMemBuffer)
    AllocateBuffer = Not (m_lPtrBuffer = 0)
End If
End Function

Private Sub FreeBuffer()
On Local Error Resume Next
If Not (m_lPtrBuffer = 0) Then
    GlobalUnlock m_hMemBuffer
    m_lPtrBuffer = 0
End If
If Not (m_hMemBuffer = 0) Then
    GlobalFree m_hMemBuffer
    m_hMemBuffer = 0
End If
End Sub

Public Sub CloseFile()
On Local Error Resume Next
If m_hMMioIn Then
    mmioClose m_hMMioIn, 0
    m_hMMioIn = 0
    FreeBuffer
    m_lCurrentReadSize = 0
    m_lDataOffset = 0
    m_lAudioLength = 0
End If
End Sub

Public Function OpenFile(ByVal sSoundFile As String) As Boolean
On Local Error Resume Next
Dim lR As Long, mmckinfoParentIn As MMCKINFO, mmckinfoSubchunkIn As MMCKINFO, mmioinf As mmioinfo, sFormat As String, iBuffer As Long, bFailed As Boolean, lRem As Long
CloseFile
If (sSoundFile = "") Then
    Exit Function
End If
m_hMMioIn = mmioOpen(sSoundFile, mmioinf, MMIO_READ)
If (m_hMMioIn = 0) Then
    pInternalErrorHandler 2
    Exit Function
End If
mmckinfoParentIn.fccType = mmioStringToFOURCC("WAVE", 0)
lR = mmioDescendParent(m_hMMioIn, mmckinfoParentIn, 0, MMIO_FINDRIFF)
If Not (lR = MMSYSERR_NOERROR) Then
    CloseFile
    pInternalErrorHandler 3
    Exit Function
End If
mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("fmt", 0)
lR = mmioDescend(m_hMMioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
If (lR <> MMSYSERR_NOERROR) Then
    CloseFile
    pInternalErrorHandler 4
    Exit Function
End If
sFormat = String$(50, 0)
lR = mmioReadString(m_hMMioIn, sFormat, mmckinfoSubchunkIn.ckSize)
If (lR = -1) Then
    CloseFile
    pInternalErrorHandler 5
    Exit Function
End If
lR = mmioAscend(m_hMMioIn, mmckinfoSubchunkIn, 0)
CopyMemoryFromString m_tFormat, sFormat, Len(m_tFormat)
If Not (m_tFormat.wBitsPerSample = 16) Or Not (m_tFormat.nChannels = 2) Then
    CloseFile
    pInternalErrorHandler 1
    Exit Function
End If
mmckinfoSubchunkIn.ckid = mmioStringToFOURCC("data", 0)
lR = mmioDescend(m_hMMioIn, mmckinfoSubchunkIn, mmckinfoParentIn, MMIO_FINDCHUNK)
If Not (lR = MMSYSERR_NOERROR) Then
    CloseFile
    pInternalErrorHandler 6
    Exit Function
End If
m_lDataOffset = mmioSeek(m_hMMioIn, 0, SEEK_CUR)
m_lAudioLength = mmckinfoSubchunkIn.ckSize
If Not (AllocateBuffer()) Then
    CloseFile
    pInternalErrorHandler 7
    Exit Function
End If
OpenFile = True
End Function

Public Function Read() As Boolean
On Local Error Resume Next
Dim dataRemaining As Long, lR As Long
If (m_hMMioIn = 0) Then
    pInternalErrorHandler 6
    Exit Function
End If
dataRemaining = (m_lDataOffset + m_lAudioLength - mmioSeek(m_hMMioIn, 0, SEEK_CUR))
If (m_lBufferSize < dataRemaining) Then
    lR = mmioRead(m_hMMioIn, m_lPtrBuffer, m_lBufferSize)
    m_lCurrentReadSize = m_lBufferSize
    Read = True
ElseIf (dataRemaining > 0) Then
    lR = mmioRead(m_hMMioIn, m_lPtrBuffer, dataRemaining)
    m_lCurrentReadSize = dataRemaining
    Read = False
Else
    Read = False
End If
End Function

Public Property Get BufferStartPosition() As Long
On Local Error Resume Next
BufferStartPosition = mmioSeek(m_hMMioIn, 0, SEEK_CUR) - m_lDataOffset
End Property

Public Property Get AudioLength() As Long
On Local Error Resume Next
AudioLength = m_lAudioLength \ 4
End Property

Public Sub SeekAbsolute(ByVal lSample As Long)
On Local Error Resume Next
If (lSample * 4 > m_lAudioLength) Or (lSample < 0) Then
    pInternalErrorHandler 8
Else
    mmioSeek m_hMMioIn, (lSample * 4) + m_lDataOffset, SEEK_SET
End If
End Sub

Public Sub SeekRelative(ByVal lSampleOffset As Long)
On Local Error Resume Next
Dim lSample As Long
lSample = mmioSeek(m_hMMioIn, SEEK_CUR, 0) - m_lDataOffset
If (lSample + lSampleOffset * 4) > m_lAudioLength Then
    pInternalErrorHandler 8
ElseIf (lSample - lSampleOffset * 4) < 0 Then
    pInternalErrorHandler 8
Else
    mmioSeek m_hMMioIn, lSampleOffset * 4, SEEK_CUR
End If
End Sub

Public Property Get ReadBufferSize() As Long
On Local Error Resume Next
ReadBufferSize = m_lBufferSize \ 4
End Property

Public Property Let ReadBufferSize(ByVal lSize As Long)
On Local Error Resume Next
m_lBufferSize = lSize * 4
End Property

Public Property Get ReadSize() As Long
On Local Error Resume Next
ReadSize = m_lCurrentReadSize \ 4
End Property

Public Property Get ReadBufferPtr() As Long
On Local Error Resume Next
ReadBufferPtr = m_lPtrBuffer
End Property

Public Sub ZeroUnusedBufferBytes()
On Local Error Resume Next
Dim lPtr As Long, lBytes As Long
If (m_lCurrentReadSize < m_lBufferSize) Then
    lPtr = UnsignedAdd(m_lPtrBuffer, m_lCurrentReadSize)
    lBytes = m_lBufferSize - m_lCurrentReadSize
    ZeroMemory ByVal lPtr, lBytes
End If
End Sub

Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
On Local Error Resume Next
If Start And &H80000000 Then
    UnsignedAdd = Start + Incr
ElseIf (Start Or &H80000000) < -Incr Then
    UnsignedAdd = Start + Incr
Else
    UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
End If
End Function

Private Sub pInternalErrorHandler(ByVal lR As Long)
On Local Error Resume Next
Dim sMsg As String
Select Case lR
Case 1
    sMsg = "Only stereo 16 bit wave files supported."
Case 2
    sMsg = "Unable to open file."
Case 3
    sMsg = "Not a Wave file."
Case 4
    sMsg = "Unable to retrieve format chunk"
Case 5
    sMsg = "Error reading format"
Case 6
    sMsg = "No Wave File Open"
Case 7
    sMsg = "Insufficient memory"
Case 8
    sMsg = "Position out of range"
End Select
Err.Raise cErrBase + lR, App.EXEName & ".cWavePlayer", sMsg
End Sub

Private Sub Class_Initialize()
On Local Error Resume Next
m_lBufferSize = 131072
End Sub

Private Sub Class_Terminate()
On Local Error Resume Next
CloseFile
End Sub
