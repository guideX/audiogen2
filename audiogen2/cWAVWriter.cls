VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWAVWriter"
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
Private Declare Function mmioOpen Lib "winmm.dll" Alias "mmioOpenA" (ByVal szFileName As String, lpmmioinfo As Any, ByVal dwOpenFlags As Long) As Long
Private Declare Function mmioSeek Lib "winmm.dll" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function mmioStringToFOURCC Lib "winmm.dll" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function mmioAscend Lib "winmm.dll" (ByVal hmmio As Long, lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function mmioWrite Lib "winmm.dll" (ByVal hmmio As Long, pch As Any, ByVal cch As Long) As Long
Private Declare Function mmioWriteString Lib "winmm.dll" Alias "mmioWrite" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Private Declare Function mmioCreateChunk Lib "winmm.dll" (ByVal hmmio As Long, pmmcki As MMCKINFO, ByVal fuCreate As Long) As Long
Private Const MMIO_READ = &H0
Private Const MMIO_WRITE = &H1
Private Const MMIO_READWRITE = &H2
Private Const MMIO_FINDCHUNK = &H10
Private Const MMIO_FINDRIFF = &H20
Private Const MMIO_CREATERIFF = &H20
Private Const MMIO_ALLOCBUF = &H10000
Private Const MMIO_CREATE = &H1000&
Private Const MM_WOM_DONE = &H3BD
Private Const MMSYSERR_NOERROR = 0
Private Const SEEK_CUR = 1
Private Const SEEK_END = 2
Private Const SEEK_SET = 0
Private Const TIME_BYTES = &H4
Private Const WHDR_DONE = &H1
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal cb As Long)
Private Declare Sub CopyMemoryFromString Lib "KERNEL32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As String, ByVal cb As Long)
Private Declare Sub CopyMemoryToString Lib "KERNEL32" Alias "RtlMoveMemory" (ByVal dest As String, Source As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Const GMEM_FIXED = &H0
Private Const GMEM_ZEROINIT = &H40
Private Const GPTR = (GMEM_FIXED Or GMEM_ZEROINIT)
Private m_hMmio As Long
Private m_ckBlank As MMCKINFO
Private m_mmckInfoChild As MMCKINFO
Private m_mmckInfoParent As MMCKINFO
Private Const ERR_BASE = 29600

Public Function OpenFile(ByVal sSoundFile As String) As Boolean
On Local Error Resume Next
CloseFile
m_hMmio = mmioOpen(sSoundFile, ByVal 0&, MMIO_ALLOCBUF Or MMIO_READWRITE Or MMIO_CREATE)
If (m_hMmio = 0) Then
    ErrHandler 3, "OpenFile"
    Exit Function
End If
If (WriteWaveFormatHeader()) Then
    OpenFile = True
End If
End Function

Public Function WriteWavData(ByVal lPtrBuff As Long, ByVal lWriteSize As Long) As Long
On Local Error Resume Next
Dim lSize As Long
If (m_hMmio = 0) Then
    ErrHandler 5, "WriteWavData"
Else
    lSize = mmioWrite(m_hMmio, ByVal lPtrBuff, lWriteSize)
    If Not (lSize = lWriteSize) Then
        ErrHandler 6, "WriteWavData"
    End If
    WriteWavData = lSize
End If
End Function

Private Function WriteWaveFormatHeader() As Boolean
On Local Error Resume Next
Dim wavEx As WAVEFORMATEX, lSize As Long
wavEx.cbSize = 0
wavEx.nAvgBytesPerSec = 176400
wavEx.nBlockAlign = 4
wavEx.nChannels = 2
wavEx.nSamplesPerSec = 44100
wavEx.wBitsPerSample = 16
wavEx.wFormatTag = 1
LSet m_mmckInfoParent = m_ckBlank
m_mmckInfoParent.fccType = mmioStringToFOURCC("WAVE", 0)
If Not (mmioCreateChunk(m_hMmio, m_mmckInfoParent, MMIO_CREATERIFF) = 0) Then
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler 4, "WriteWaveFormatHeader"
    Exit Function
End If
LSet m_mmckInfoChild = m_ckBlank
m_mmckInfoChild.ckid = mmioStringToFOURCC("fmt", 0)
m_mmckInfoChild.ckSize = Len(wavEx)
If Not (mmioCreateChunk(m_hMmio, m_mmckInfoChild, 0) = 0) Then
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler 4, "WriteWaveFormatHeader"
    Exit Function
End If
lSize = mmioWrite(m_hMmio, wavEx, Len(wavEx))
If Not (lSize = Len(wavEx)) Then
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler 4, "WriteWaveFormatHeader"
    Exit Function
End If
If Not (mmioAscend(m_hMmio, m_mmckInfoChild, 0) = 0) Then
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler 4, "WriteWaveFormatHeader"
    Exit Function
End If
m_mmckInfoChild.ckid = mmioStringToFOURCC("data", 0)
If Not (mmioCreateChunk(m_hMmio, m_mmckInfoChild, 0) = 0) Then
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler 4, "WriteWaveFormatHeader"
    Exit Function
End If
WriteWaveFormatHeader = True
End Function

Public Sub CloseFile()
On Local Error Resume Next
Dim lErr As Long
If Not (m_hMmio = 0) Then
    If Not (mmioAscend(m_hMmio, m_mmckInfoChild, 0) = 0) Then
        lErr = 1
    End If
    If Not (mmioAscend(m_hMmio, m_mmckInfoParent, 0) = 0) Then
        lErr = 2
    End If
    mmioClose m_hMmio, 0
    m_hMmio = 0
    ErrHandler lErr, "CloseFile"
End If
End Sub

Private Sub ErrHandler(ByVal lErr As Long, ByVal sProc As String)
On Local Error Resume Next
Dim sMsg As String
Select Case lErr
Case 0
   Exit Sub
Case 1
   sMsg = "Unable to finalise data chunk; WAV file may not be usable."
Case 2
   sMsg = "Unable to finalise RIFF chunk; WAV file may not be usable."
Case 3
   sMsg = "Unable to open file for writing."
Case 4
   sMsg = "Unable to write the WAV file header."
Case 5
   sMsg = "WAV file not open."
Case 6
   sMsg = "Error writing data: bytes written does not match request, WAV file may not be usable."
End Select
Err.Raise lErr + ERR_BASE, App.EXEName & "." & sProc, sMsg
End Sub

Private Sub Class_Terminate()
On Local Error Resume Next
If Not (m_hMmio = 0) Then
    On Error Resume Next
    CloseFile
    Debug.Assert "" = "Warning: class terminated when file still open"
End If
End Sub

