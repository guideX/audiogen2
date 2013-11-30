Attribute VB_Name = "mdlSpectrum"
Option Explicit
Private lSpectrumIndex As Integer
Private lSpectrumCount As Integer

Public Function ReturnSpectrumCount() As Integer
On Local Error GoTo ErrHandler
lSpectrumCount = ReadINI(lIniFiles.iSpectrum, "Settings", "Count", 0)
ReturnSpectrumCount = lSpectrumCount
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnSpectrumCount() As Integer", Err.Description, Err.Number
End Function

Public Function ReturnSpectrumIndex() As Integer
On Local Error GoTo ErrHandler
ReturnSpectrumIndex = lSpectrumIndex
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnSpectrumIndex() As Integer", Err.Description, Err.Number
End Function

Public Sub LoadSpectrum(lIndex As Integer)
On Local Error GoTo ErrHandler
Dim s As String
s = lIniFiles.iSpectrum
lSpectrumIndex = lIndex
If lIndex <> 0 Then
    With frmMain.ctlMP3Player
        .Tag = ReadINI(s, Trim(Str(lIndex)), "Name", "")
        If Len(.Tag) <> 0 Then
            lSpectrumIndex = lIndex
            .BackColor = ReadINI(s, Trim(Str(lIndex)), "Backcolor", 0)
            .BottomBandsColor = ReadINI(s, Trim(Str(lIndex)), "BottomBandsColor", 0)
            .Bands = ReadINI(s, Trim(Str(lIndex)), "Bands", 0)
            .DividerColor = ReadINI(s, Trim(Str(lIndex)), "DividerColor", 0)
            .LeftChanColor = ReadINI(s, Trim(Str(lIndex)), "LeftChanColor", 0)
            .PeaksColor = ReadINI(s, Trim(Str(lIndex)), "PeaksColor", 0)
            .RightChanColor = ReadINI(s, Trim(Str(lIndex)), "RightChanColor", 0)
            .ShowPeaks = ReadINI(s, Trim(Str(lIndex)), "ShowPeaks", 0)
            .SpectrumMode = ReadINI(s, Trim(Str(lIndex)), "SpectrumMode", 0)
            .TopBandsColor = ReadINI(s, Trim(Str(lIndex)), "TopBandsColor", 0)
            .VISFrameRate = ReadINI(s, Trim(Str(lIndex)), "VISFrameRate", 0)
        End If
    End With
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub LoadSpectrum(lIndex As Integer)", Err.Description, Err.Number
End Sub
