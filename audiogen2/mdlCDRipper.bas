Attribute VB_Name = "mdlCDRipper"
Option Explicit
Enum ECDRipErrorCode
   CDEX_OK = &H0
   CDEX_ERROR = &H1
   CDEX_FILEOPEN_ERROR = &H2
   CDEX_JITTER_ERROR = &H3
   CDEX_RIPPING_DONE = &H4
   CDEX_RIPPING_INPROGRESS = &H5
End Enum
Private Const ERR_BASE = 29500

Public Sub CDRipErrHandler(ByVal sProc As String, ByVal lErr As Long, ByVal bCDRipError As Boolean)
On Local Error GoTo ErrHandler
Dim sMsg As String
If (bCDRipError) Then
    Select Case lErr
    Case CDEX_OK
        Exit Sub
    Case CDEX_ERROR
        sMsg = "CDRip Error"
    Case CDEX_FILEOPEN_ERROR
        sMsg = "CDRip File Open Error"
    Case CDEX_JITTER_ERROR = &H3
        sMsg = "CDRip Jitter Error"
    Case CDEX_RIPPING_DONE
        sMsg = "CDRip Ripping Done"
        Exit Sub
    Case CDEX_RIPPING_INPROGRESS = &H5
        sMsg = "CDRip Ripping in Progresss"
        Exit Sub
    End Select
Else
    Select Case lErr
    Case 0
        Exit Sub
    Case 1
        sMsg = "CDRip Not Initialised."
    Case 2
        sMsg = "CD Buffer not open for reading."
    Case 3
        sMsg = "Invalid sector specified"
    Case 7
        sMsg = "Failed to create memory buffer to read CD to."
    End Select
End If
Err.Raise lErr + ERR_BASE, App.EXEName & "." & sProc, sMsg
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub CDRipErrHandler(ByVal sProc As String, ByVal lErr As Long, ByVal bCDRipError As Boolean)", Err.Description, Err.Number
End Sub
