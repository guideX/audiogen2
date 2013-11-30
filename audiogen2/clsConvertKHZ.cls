VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvertKHZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private lConverting As Boolean
Private lConvertProgress As Integer

Public Sub ConvertKHZ(lFilename As String, lEncoder As MEnc, lDecoder As MDec)
On Local Error GoTo ErrHandler
If Len(lFilename) <> 0 Then
    If DoesFileExist(lFilename) = True Then
        Select Case LCase(Right(lFilename, 4))
        Case ".wav"
            With lEncoder
                .bitrate = 256000
                .channels = 0
                .OPENFILENAME = lFilename
                .savefilename = Left(lFilename, Len(lFilename)) - 4 & ".mp3"
                .Encode
            End With
        Case ".mp3"
            With lDecoder
                .OPENFILENAME = lFilename
                .savefilename = Left(lFilename, Len(lFilename)) - 4 & ".wav"
                .Decode
            End With
        End Select
    End If
End If
ErrHandler:
    ProcessRuntimeError "Public Sub ConvertKHZ(lFilename As String, lEncoder As MEnc, lDecoder As MDec)", Err.Description, Err.Number
End Sub
