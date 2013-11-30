VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCDRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function CR_Init Lib "cdrip.dll" (ByVal sFileName As String) As Long
Private Declare Function CR_GetCDRipVersion Lib "cdrip.dll" () As Long
Private Declare Function CR_GetNumCDROM Lib "cdrip.dll" () As Long
Private Declare Function CR_GetActiveCDROM Lib "cdrip.dll" () As Long
Private lIniFile As String
Private lCreated As Boolean

Public Sub Create(ByVal sIniFile As String)
'On Local Error Resume Next
Dim e As eCDRipErrorCode
If Not (lCreated) Then
   lIniFile = sIniFile
   e = CR_Init(sIniFile)
   If (e = CDEX_OK) Then
      lCreated = True
   End If
   CDRipErrHandler "cCDDriveRip_Create", e, True
End If
If Err.Number <> 0 Then Err.Clear
End Sub

Public Property Get CDDriveCount() As Long
'On Local Error Resume Next
CDRipErrHandler "cCDDriveRip.CDDriveCount", Abs(Not (lCreated)), False
CDDriveCount = CR_GetNumCDROM()
If Err.Number <> 0 Then Err.Clear
End Property

Public Property Get Version() As Long
'On Local Error Resume Next
Version = CR_GetCDRipVersion()
If Err.Number <> 0 Then Err.Clear
End Property

Public Property Get CDDrive(ByVal nIndex As Long) As clsDrive
'On Local Error Resume Next
Dim c As New clsDrive
c.fInit nIndex - 1
Set CDDrive = c
If Err.Number <> 0 Then Err.Clear
End Property

Public Property Get ActiveCDDriveIndex() As Long
'On Local Error Resume Next
ActiveCDDriveIndex = CR_GetActiveCDROM() + 1
If Err.Number <> 0 Then Err.Clear
End Property
