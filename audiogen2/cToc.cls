VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cToc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type CDTEXTPACK_TAG
   packType As Byte
   TrackNumber As Byte
   sequenceNumber As Byte
   positionInfo As Byte
   data(0 To 11) As Byte
   crc0 As Byte
   crc1 As Byte
End Type
Private Declare Function CR_ReadCDText Lib "cdrip.dll" (pbtBuffer As Any, ByVal nBufferSize As Long, pnCDTextSize As Integer) As Long
Private Type tTOCENTRY
   dwStartSector As Long
   btFlag As Byte
   btTrackNumber As Byte
End Type
Private Declare Function CR_ReadToc Lib "cdrip.dll" () As Long
Private Declare Function CR_GetNumTocEntries Lib "cdrip.dll" () As Long
Private Declare Function CR_GetTocEntry Lib "cdrip.dll" (ByVal nTocEntry As Long) As Currency
Private Declare Sub CR_SetActiveCDROM Lib "cdrip.dll" (ByVal nActiveDrive As Long)
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private m_count As Long
Private m_cEntries As Collection
Private m_index As Long
Private m_lLeadOutOffset As Long
Private m_iCDDBID As Long
Private m_lPower2(0 To 31) As Long

Friend Sub fInit(ByVal nIndex As Long)
On Local Error Resume Next
Dim i As Long, lTocValue As Currency, tTOC() As tTOCENTRY, cToc As cTocEntry, lOffset As Long
ReDim tTOC(0 To m_count) As tTOCENTRY
Set m_cEntries = New Collection
m_index = nIndex
CDRipErrHandler "cCDToc.fInit", CR_ReadToc(), True
m_count = CR_GetNumTocEntries()
If (m_count > 0) Then
For i = 0 To m_count
   lTocValue = CR_GetTocEntry(i)
   CopyMemory tTOC(i).dwStartSector, lTocValue, 6
   CopyMemory tTOC(i).btTrackNumber, ByVal (VarPtr(lTocValue) + 5), 1
   CopyMemory tTOC(i).btFlag, ByVal (VarPtr(lTocValue) + 4), 1
Next i
For i = 0 To m_count - 1
   Set cToc = New cTocEntry
   cToc.fInit _
      m_index, _
      tTOC(i).dwStartSector, _
      tTOC(i).btFlag, tTOC(i).btTrackNumber, _
      tTOC(i + 1).dwStartSector
   m_cEntries.Add cToc, "T" & (i + 1)
Next i
m_lLeadOutOffset = tTOC(m_count).dwStartSector + 150
getCDDBID
End If
End Sub

Public Property Get TotalLengthSeconds()
On Local Error Resume Next
TotalLengthSeconds = m_lLeadOutOffset \ 75
End Property

Public Property Get TotalLengthSectors()
On Local Error Resume Next
TotalLengthSectors = m_lLeadOutOffset
End Property

Public Property Get CDDBId() As Long
On Local Error Resume Next
CDDBId = m_iCDDBID
End Property

Public Property Get CDDBQuery() As String
On Local Error Resume Next
Dim sQuery As String, i As Long, cToc As cTocEntry
sQuery = "cddb query"
sQuery = sQuery & " " & Hex$(m_iCDDBID)
sQuery = sQuery & " " & m_count
For i = 1 To m_count
    Set cToc = m_cEntries("T" & i)
    sQuery = sQuery & " " & cToc.offset
Next i
sQuery = sQuery & " " & TotalLengthSeconds
CDDBQuery = LCase(sQuery)
End Property

Public Sub Refresh()
On Local Error Resume Next
CR_SetActiveCDROM m_index
fInit m_index
End Sub

Public Property Get Count() As Long
On Local Error Resume Next
Count = m_count
End Property

Public Property Get Entry(ByVal nIndex As Long) As cTocEntry
On Local Error Resume Next
Set Entry = m_cEntries("T" & nIndex)
End Property

Private Sub getCDDBID()
On Local Error Resume Next
Dim i As Long, numTracks As Long, cToc As cTocEntry, n As Long, t As Long, Res As Long, leadOutMinutes As Long, leadOutSeconds As Long
For i = 1 To m_count
    Set cToc = m_cEntries("T" & i)
    n = UnsignedAdd(n, CDDBSum(60 * cToc.StartTimeMinutes + cToc.StartTimeSeconds))
Next i
leadOutMinutes = Int(m_lLeadOutOffset / (75 * 60))
leadOutSeconds = (Int(m_lLeadOutOffset / 75)) Mod 60
t = 60 * leadOutMinutes + leadOutSeconds
Set cToc = m_cEntries("T1")
t = t - (60 * cToc.StartTimeMinutes + cToc.StartTimeSeconds)
m_iCDDBID = RShift((n Mod &HFF), 24) Or RShift(t, 8) Or m_count
Debug.Print Hex(m_iCDDBID)
End Sub

Private Function CDDBSum(ByVal n As Long) As Long
On Local Error Resume Next
Dim retVal As Long
Do While (n > 0)
    retVal = UnsignedAdd(retVal, (n Mod 10))
    n = n \ 10
Loop
CDDBSum = retVal
End Function

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

Private Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long
On Local Error Resume Next
If (lBits <= 0) Then
    RShift = lThis
ElseIf (lBits > 63) Then
ElseIf (lBits > 31) Then
    RShift = 0
Else
    If (lThis And m_lPower2(31 - lBits)) = m_lPower2(31 - lBits) Then
        RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits) Or m_lPower2(31)
    Else
        RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits)
    End If
End If
End Function

Private Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long
On Local Error Resume Next
If (lBits <= 0) Then
    LShift = lThis
ElseIf (lBits > 63) Then
ElseIf (lBits > 31) Then
    LShift = 0
Else
    If (lThis And m_lPower2(31)) = m_lPower2(31) Then
        LShift = (lThis And &H7FFFFFFF) \ m_lPower2(lBits) Or m_lPower2(31 - lBits)
    Else
        LShift = lThis \ m_lPower2(lBits)
    End If
End If
End Function

Private Sub Init()
On Local Error Resume Next
m_lPower2(0) = &H1&
m_lPower2(1) = &H2&
m_lPower2(2) = &H4&
m_lPower2(3) = &H8&
m_lPower2(4) = &H10&
m_lPower2(5) = &H20&
m_lPower2(6) = &H40&
m_lPower2(7) = &H80&
m_lPower2(8) = &H100&
m_lPower2(9) = &H200&
m_lPower2(10) = &H400&
m_lPower2(11) = &H800&
m_lPower2(12) = &H1000&
m_lPower2(13) = &H2000&
m_lPower2(14) = &H4000&
m_lPower2(15) = &H8000&
m_lPower2(16) = &H10000
m_lPower2(17) = &H20000
m_lPower2(18) = &H40000
m_lPower2(19) = &H80000
m_lPower2(20) = &H100000
m_lPower2(21) = &H200000
m_lPower2(22) = &H400000
m_lPower2(23) = &H800000
m_lPower2(24) = &H1000000
m_lPower2(25) = &H2000000
m_lPower2(26) = &H4000000
m_lPower2(27) = &H8000000
m_lPower2(28) = &H10000000
m_lPower2(29) = &H20000000
m_lPower2(30) = &H40000000
m_lPower2(31) = &H80000000
End Sub

Private Sub Class_Initialize()
On Local Error Resume Next
Init
End Sub


