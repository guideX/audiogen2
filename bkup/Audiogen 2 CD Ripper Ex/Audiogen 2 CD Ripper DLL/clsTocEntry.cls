VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTocEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Enum ECDRipTOCEntryType
   CDROMDATAFLAG = &H4
   AUDIOTRKFLAG = &H10
End Enum
Private m_cdIndex As Long
Private m_startSector As Long
Private m_endSector As Long
Private m_trackNumber As Byte
Private m_eEntryType As ECDRipTOCEntryType
Private m_lOffset As Long

Friend Sub fInit(ByVal lCDIndex As Long, ByVal lStartSector As Long, ByVal bFlag As Byte, ByVal bTrackNumber As Byte, ByVal lEndSector As Long)
On Local Error Resume Next
m_cdIndex = lCDIndex
m_startSector = lStartSector
m_trackNumber = bTrackNumber
m_eEntryType = bFlag
m_endSector = lEndSector
End Sub

Public Property Get StartSector() As Long
On Local Error Resume Next
StartSector = m_startSector
End Property

Public Property Get EndSector() As Long
On Local Error Resume Next
EndSector = m_endSector
End Property

Public Property Get TrackNumber() As Byte
On Local Error Resume Next
TrackNumber = m_trackNumber
End Property

Public Property Get TOCEntryType() As ECDRipTOCEntryType
On Local Error Resume Next
TOCEntryType = m_eEntryType
End Property

Public Property Get Offset() As Long
On Local Error Resume Next
Offset = m_startSector + 150
End Property

Public Property Get SizeBytes() As Long
On Local Error Resume Next
Dim lFrames As Long
lFrames = m_endSector - m_startSector
SizeBytes = ((lFrames * 4&) / 75&) * 44100
End Property

Public Property Get StartTimeMinutes() As Long
On Local Error Resume Next
StartTimeMinutes = Int(Offset / (75 * 60))
End Property

Public Property Get StartTimeSeconds() As Long
On Local Error Resume Next
StartTimeSeconds = (Int(Offset / 75)) Mod 60
End Property

Public Property Get StartTimeFrames() As Long
On Local Error Resume Next
StartTimeFrames = Offset Mod 75
End Property

Public Property Get FormattedStartTime() As String
On Local Error Resume Next
FormattedStartTime = FormattedTime(StartTimeMinutes, StartTimeSeconds, StartTimeFrames)
End Property

Public Property Get FormattedLength() As String
On Local Error Resume Next
FormattedLength = FormattedTime(LengthMinutes, LengthSeconds, LengthFrames)
End Property

Private Function FormattedTime(ByVal minutes As Long, ByVal seconds As Long, ByVal frames As Long) As String
On Local Error Resume Next
FormattedTime = Format(minutes, "##0") & ":" & Format(seconds, "00") & "." & Format((frames * 60) / 75, "00")
End Function

Public Property Get LengthMinutes() As Long
On Local Error Resume Next
LengthMinutes = Int((m_endSector - m_startSector) / (75 * 60))
End Property

Public Property Get LengthSeconds() As Long
On Local Error Resume Next
LengthSeconds = (Int((m_endSector - m_startSector) / 75)) Mod 60
End Property

Public Property Get LengthFrames() As Long
On Local Error Resume Next
LengthFrames = (m_endSector - m_startSector) Mod 75
End Property

Friend Property Get fCDIndex() As Long
On Local Error Resume Next
fCDIndex = m_cdIndex
End Property
