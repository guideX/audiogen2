VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLameEncoderVersion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub beVersion Lib "lameencshim.dll" (beVersion As BE_VERSION)
Private Const BE_MAX_HOMEPAGE As Long = 128
Private Type BE_VERSION
    byDLLMajorVersion As Byte
    byDLLMinorVersion As Byte
    byMajorVersion As Byte
    byMinorVersion As Byte
    byDay As Byte
    byMonth As Byte
    btYearHomepageDetail(0 To 133) As Byte
    btReserved(0 To 124) As Byte
End Type
Private m_tBeVer As BE_VERSION
Private m_dReleaseDate As Date
Private m_sHomePage As String

Public Property Get DllMajorVersion() As Byte
On Local Error Resume Next
DllMajorVersion = m_tBeVer.byDLLMajorVersion
End Property

Public Property Get DllMinorVersion() As Byte
On Local Error Resume Next
DllMinorVersion = m_tBeVer.byDLLMinorVersion
End Property

Public Property Get MajorVersion() As Byte
On Local Error Resume Next
MajorVersion = m_tBeVer.byMajorVersion
End Property

Public Property Get MinorVersion() As Byte
On Local Error Resume Next
MinorVersion = m_tBeVer.byMinorVersion
End Property

Public Property Get ReleaseDate() As Date
On Local Error Resume Next
ReleaseDate = m_dReleaseDate
End Property

Public Property Get HomePage() As String
On Local Error Resume Next
HomePage = m_sHomePage
End Property

Public Property Get AlphaLevel() As Byte
On Local Error Resume Next
AlphaLevel = m_tBeVer.btYearHomepageDetail(131)
End Property

Public Property Get BetaLevel() As Byte
On Local Error Resume Next
BetaLevel = m_tBeVer.btYearHomepageDetail(132)
End Property

Public Property Get MMXEnabled() As Byte
On Local Error Resume Next
MMXEnabled = m_tBeVer.btYearHomepageDetail(133)
End Property

Private Sub decodeReleaseDate()
On Local Error Resume Next
Dim iYear As Long
iYear = m_tBeVer.btYearHomepageDetail(0)
iYear = iYear Or (m_tBeVer.btYearHomepageDetail(1) * &H100&)
m_dReleaseDate = DateSerial(iYear, m_tBeVer.byMonth, m_tBeVer.byDay)
End Sub

Private Sub decodeHomePage()
On Local Error Resume Next
Dim i As Long
i = 2
Do While m_tBeVer.btYearHomepageDetail(i) > 0 And (i <= BE_MAX_HOMEPAGE + 2)
    m_sHomePage = m_sHomePage & Chr(m_tBeVer.btYearHomepageDetail(i))
    i = i + 1
Loop
End Sub

Private Sub Class_Initialize()
On Local Error Resume Next
beVersion m_tBeVer
decodeReleaseDate
decodeHomePage
End Sub
