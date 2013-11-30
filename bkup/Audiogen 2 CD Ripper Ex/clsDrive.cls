VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum ECDRipDriveType
    GENERIC = 0
    TOSHIBA
    TOSHIBANEW
    IBM
    NEC
    DEC
    IMS
    KODAK
    RICOH
    HP
    PHILIPS
    PLASMON
    GRUNDIGCDR100IPW
    MITSUMICDR
    PLEXTOR
    SONY
    YAMAHA
    NRC
    IMSCDD5
    CUSTOMDRIVE
    NUMDRIVETYPES
End Enum
Enum ECDRipReadMethod
    READMMC = 0
    READ10
    READNEC
    READSONY
    READMMC2
    READMMC3
    READC1
    READC2
    READC3
    READMMC4
    NUMREADMETHODS
End Enum
Enum ECDRipSetSpeed
    SPEEDNONE = 0
    SPEEDMMC
    SPEEDSONY
    SPEEDYAMAHA
    SPEEDTOSHIBA
    SPEEDPHILIPS
    SPEEDNEC
    NUMSPEEDMETHODS
End Enum
Enum ECDRipENDIAN
    BIGENDIAN = 0
    LITTLEENDIAN
    NUMENDIAN
End Enum
Enum ECDRipENABLEMODE
    ENABLENONE = 0
    ENABLESTD
    NUMENABLEMODES
End Enum
Enum ECDRipOUTPUTFORMAT
    STEREO44100 = 0
    MONO44100
    STEREO22050
    MONO22050
    STEREO11025
    MONO11025
    NUMOUTPUTFORMATS
End Enum
Private Type tDRIVETABLE
    DriveType As ECDRipDriveType
    readMethod As ECDRipReadMethod
    setSpeed As ECDRipSetSpeed
    endian As ECDRipENDIAN
    enableMode As ECDRipENABLEMODE
    nDensity As Long
    bAtapi As Long
End Type
Private Const TRANSPLAYER_ASPI = (0)
Private Const TRANSPLAYER_NTSCSI = (1)
Private Const CR_RIPPING_MODE_NORMAL = (0)
Private Const CR_RIPPING_MODE_PARANOIA = (1)
Private Type tSENSKEY
    SK As Byte
    ASC As Byte
    ASCQ As Byte
End Type
Private Type CDROMPARAMS
    lpszCDROMID As String * 255
    nNumReadSectors As Long
    nNumOverlapSectors As Long
    nNumCompareSectors As Long
    nOffsetStart As Long
    nOffsetEnd As Long
    nSpeed As Long
    nSpinUpTime As Long
    bJitterCorrection As Long
    bSwapLefRightChannel As Long
    driveTable As tDRIVETABLE
    btTargetID As Byte
    btAdapterID As Byte
    btLunID As Byte
    bAspiPosting As Long
    nAspiRetries As Integer
    nAspiTimeOut As Integer
    bEnableMultiRead As Long
    bMultiReadFirstOnly As Long
    nMultiReadCount As Integer
    bLockDuringRead As Long
    nRippingMode As Integer
    nParanoiaMode As Integer
    bJunk(0 To 1024) As Byte
End Type
Private Declare Function CR_SetCDROMParameters Lib "cdrip.dll" (pParam As CDROMPARAMS) As Long
Private Declare Function CR_GetCDROMParameters Lib "cdrip.dll" (pParam As CDROMPARAMS) As Long
Private Declare Sub CR_SetActiveCDROM Lib "cdrip.dll" (ByVal nActiveDrive As Long)
Private Declare Function CR_IsUnitReady Lib "cdrip.dll" () As Long
Private Declare Function CR_EjectCD Lib "cdrip.dll" (ByVal bEject As Long) As Long
Private Declare Function CR_IsAudioPlaying Lib "cdrip.dll" () As Long
Private Declare Function CR_PlayTrack Lib "cdrip.dll" (ByVal nTrack As Long) As Long
Private Declare Function CR_StopPlayTrack Lib "cdrip.dll" ()
Private Declare Function CR_PauseCD Lib "cdrip.dll" (ByVal bPause As Long)
Private m_index As Long
Private m_tP As CDROMPARAMS
Private m_bDirty As Boolean

Friend Property Get fCDIndex() As Long
'On Local Error Resume Next
fCDIndex = m_index
End Property

Public Property Get IsUnitReady() As Boolean
'On Local Error Resume Next
CR_SetActiveCDROM m_index
IsUnitReady = IIf(CR_IsUnitReady() = 0, False, True)
End Property

Public Property Get IsAudioPlaying() As Boolean
'On Local Error Resume Next
CR_SetActiveCDROM m_index
IsAudioPlaying = IIf(CR_IsAudioPlaying() = 0, False, True)
End Property

Public Sub EjectCD()
'On Local Error Resume Next
CR_SetActiveCDROM m_index
CR_EjectCD 1
End Sub

Public Sub CloseCD()
'On Local Error Resume Next
CR_SetActiveCDROM m_index
CR_EjectCD 0
End Sub

Public Sub PlayCDTrack(ByVal nIndex As Long)
'On Local Error Resume Next
CDRipErrHandler "cDrive.PlayCDTrack", CR_PlayTrack(nIndex - 1), True
End Sub

Public Sub UnpauseCD()
'On Local Error Resume Next
CDRipErrHandler "cDrive.UnPauseCD", CR_PauseCD(0), True
End Sub

Public Sub PauseCD()
'On Local Error Resume Next
CDRipErrHandler "cDrive.UnPauseCD", CR_PauseCD(1), True
End Sub

Public Sub StopCD()
'On Local Error Resume Next
CDRipErrHandler "cDrive.UnPauseCD", CR_StopPlayTrack, True
End Sub

Public Property Get Name() As String
'On Local Error Resume Next
Name = StripNull(m_tP.lpszCDROMID)
End Property

Public Property Get ReadSectors() As Long
'On Local Error Resume Next
ReadSectors = m_tP.nNumReadSectors
End Property

Public Property Let ReadSectors(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nNumReadSectors = value) Then
    m_tP.nNumReadSectors = value
    m_bDirty = True
End If
End Property

Public Property Get ReadOverlap() As Long
'On Local Error Resume Next
ReadOverlap = m_tP.nNumOverlapSectors
End Property

Public Property Let ReadOverlap(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nNumOverlapSectors = value) Then
    m_tP.nNumOverlapSectors = value
    m_bDirty = True
End If
End Property

Public Property Get BlockCompare() As Long
'On Local Error Resume Next
BlockCompare = m_tP.nNumCompareSectors
End Property

Public Property Let BlockCompare(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nNumCompareSectors = value) Then
    m_tP.nNumCompareSectors = value
    m_bDirty = True
End If
End Property

Public Property Get StartOffset() As Long
'On Local Error Resume Next
StartOffset = m_tP.nOffsetStart
End Property

Public Property Let StartOffset(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nOffsetStart = value) Then
    m_tP.nOffsetStart = value
    m_bDirty = True
End If
End Property

Public Property Get EndOffset() As Long
'On Local Error Resume Next
EndOffset = m_tP.nOffsetEnd
End Property

Public Property Let EndOffset(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nOffsetEnd = value) Then
    m_tP.nOffsetEnd = value
    m_bDirty = True
End If
End Property

Public Property Get CDSpeed() As Long
'On Local Error Resume Next
CDSpeed = m_tP.nSpeed
End Property

Public Property Let CDSpeed(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nSpeed = value) Then
    m_tP.nSpeed = value
    m_bDirty = True
End If
End Property

Public Property Get SpinUpTime() As Long
'On Local Error Resume Next
SpinUpTime = m_tP.nSpinUpTime
End Property

Public Property Let SpinUpTime(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nSpinUpTime = value) Then
    m_tP.nSpinUpTime = value
    m_bDirty = True
End If
End Property

Public Property Get Retries() As Long
'On Local Error Resume Next
Retries = m_tP.nAspiRetries
End Property

Public Property Let Retries(ByVal value As Long)
'On Local Error Resume Next
If Not (m_tP.nAspiRetries = value) Then
    m_tP.nAspiRetries = value
    m_bDirty = True
End If
End Property

Public Property Get DriveType() As ECDRipDriveType
'On Local Error Resume Next
DriveType = m_tP.driveTable.DriveType
End Property

Public Property Let DriveType(ByVal value As ECDRipDriveType)
'On Local Error Resume Next
If Not (m_tP.driveTable.DriveType = value) Then
    m_tP.driveTable.DriveType = value
    m_bDirty = True
End If
End Property

Public Property Get RippingMethod() As ECDRipDriveType
'On Local Error Resume Next
RippingMethod = m_tP.driveTable.readMethod
End Property

Public Property Let RippingMethod(ByVal value As ECDRipDriveType)
'On Local Error Resume Next
If Not (m_tP.driveTable.readMethod = value) Then
    m_tP.driveTable.readMethod = value
    m_bDirty = True
End If
End Property

Public Property Get TOC() As clsToc
'On Local Error Resume Next
Dim cToc As New clsToc
CR_SetActiveCDROM m_index
cToc.fInit m_index
Set TOC = cToc
End Property

Public Sub Refresh()
'On Local Error Resume Next
fInit m_index
End Sub

Public Sub Apply()
'On Local Error Resume Next
If (m_bDirty) Then
    CDRipErrHandler "cCDDriveInfo.Apply", CR_SetCDROMParameters(m_tP), True
    m_bDirty = False
End If
End Sub

Friend Sub fInit(ByVal index As Long)
'On Local Error Resume Next
m_bDirty = False
m_index = index
CR_SetActiveCDROM m_index
CDRipErrHandler "cCDDriveInfo.fInit", CR_GetCDROMParameters(m_tP), True
End Sub

Private Function StripNull(ByVal sz As String) As String
'On Local Error Resume Next
Dim i As Long
i = InStr(sz, vbNullChar)
If (i > 0) And (i < Len(sz)) Then
    StripNull = Left(sz, i - 1)
Else
    StripNull = sz
End If
End Function
