Attribute VB_Name = "mdlRip"

Public Sub PlayCD()
''On Local Error GoTo ErrHandler

Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub PlayCD()", Err.Description, Err.Number
End Sub


Public Sub SetCDRipperSpeed(lSpeed As Integer)
'On Local Error Resume Next

End Sub

