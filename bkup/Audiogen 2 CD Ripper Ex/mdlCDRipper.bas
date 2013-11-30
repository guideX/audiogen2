Attribute VB_Name = "mdlCDRipper"
Option Explicit

Enum eCDRipErrorCode
   CDEX_OK = &H0
   CDEX_ERROR = &H1
   CDEX_FILEOPEN_ERROR = &H2
   CDEX_JITTER_ERROR = &H3
   CDEX_RIPPING_DONE = &H4
   CDEX_RIPPING_INPROGRESS = &H5
End Enum
Private Const ERR_BASE = 29500
'Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Sub Main()
'On Local Error Resume Next

End Sub

Public Sub CDRipErrHandler(ByVal sProc As String, ByVal lErr As Long, ByVal bCDRipError As Boolean)
'On Local Error Resume Next
Dim msg As String
If (bCDRipError) Then
   Select Case lErr
   Case CDEX_OK
      Exit Sub
   Case CDEX_ERROR
      msg = "CDRip Error"
   Case CDEX_FILEOPEN_ERROR
      msg = "CDRip File Open Error"
   Case CDEX_JITTER_ERROR = &H3
      msg = "CDRip Jitter Error"
   Case CDEX_RIPPING_DONE
      msg = "CDRip Ripping Done"
      Exit Sub
   Case CDEX_RIPPING_INPROGRESS = &H5
      msg = "CDRip Ripping in Progresss"
      Exit Sub
   End Select
Else
   Select Case lErr
   Case 0
      Exit Sub
   Case 1
      msg = "CDRip Not Initialised."
   Case 2
      msg = "CD Buffer not open for reading."
   Case 3
      msg = "Invalid sector specified"
   Case 7
      msg = "Failed to create memory buffer to read CD to."
   End Select
End If
Err.Raise lErr + ERR_BASE, App.EXEName & "." & sProc, msg
End Sub
