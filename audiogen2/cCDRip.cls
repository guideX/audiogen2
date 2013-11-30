VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCDRip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CR_Init Lib "cdrip.dll" (ByVal sFileName As String) As Long
Private Declare Function CR_GetCDRipVersion Lib "cdrip.dll" () As Long
Private Declare Function CR_GetNumCDROM Lib "cdrip.dll" () As Long
Private Declare Function CR_GetActiveCDROM Lib "cdrip.dll" () As Long
Private m_sIniFile As String
Private m_bCreated As Boolean

Public Sub Create(ByVal sIniFile As String)
On Local Error GoTo ErrHandler
Dim e As ECDRipErrorCode
If Not (m_bCreated) Then
    m_sIniFile = sIniFile
    e = CR_Init(sIniFile)
    If (e = CDEX_OK) Then
        m_bCreated = True
    End If
    CDRipErrHandler "cCDDriveRip_Create", e, True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub Create(ByVal sIniFile As String)", Err.Description, Err.Number
End Sub

Public Property Get CDDriveCount() As Long
On Local Error GoTo ErrHandler
CDRipErrHandler "cCDDriveRip.CDDriveCount", Abs(Not (m_bCreated)), False
CDDriveCount = CR_GetNumCDROM()
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Property Get CDDriveCount() As Long", Err.Description, Err.Number
End Property

Public Property Get Version() As Long
On Local Error Resume Next
Version = CR_GetCDRipVersion()
End Property

Public Property Get CDDrive(ByVal nIndex As Long) As cDrive
On Local Error Resume Next
Dim c As New cDrive
c.fInit nIndex - 1
Set CDDrive = c
End Property

Public Property Get ActiveCDDriveIndex() As Long
On Local Error Resume Next
ActiveCDDriveIndex = CR_GetActiveCDROM() + 1
End Property
