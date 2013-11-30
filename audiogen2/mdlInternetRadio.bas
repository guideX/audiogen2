Attribute VB_Name = "mdlInternetRadio"
Option Explicit

Public Function ReturnInternetRadioSaveToDisk(lName As String) As String
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, lRadioFile As String, i As Integer, c As Integer
lRadioFile = App.Path & "\data\playlists\radio.ini"
If Len(lName) <> 0 Then
    If DoesFileExist(lRadioFile) = True Then
        c = Int(ReadINI(lRadioFile, "Settings", "Count", 0))
        If c <> 0 Then
            For i = 1 To c
                msg2 = ReadINI(lRadioFile, Trim(Str(i)), "Name", "")
                If LCase(msg2) = LCase(lName) Then
                    msg2 = ""
                    msg2 = ReadINI(lRadioFile, Trim(Str(i)), "SaveToDisk", "")
                    ReturnInternetRadioSaveToDisk = msg2
                    Exit For
                End If
            Next i
        End If
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnInternetRadioSaveToDisk(lName As String) As String", Err.Description, Err.Number
End Function

Public Function ReturnInternetRadioName(lIndex As Integer) As String
On Local Error GoTo ErrHandler
Dim lRadioFile As String
lRadioFile = App.Path & "\data\playlists\radio.ini"
ReturnInternetRadioName = ReadINI(lRadioFile, Trim(Str(lIndex)), "Name", "")
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnInternetRadioName(lIndex As Integer) As String", Err.Description, Err.Number
End Function

Public Function ReturnInternetRadioCount() As Integer
On Local Error GoTo ErrHandler
ReturnInternetRadioCount = ReadINI(App.Path & "\data\playlists\radio.ini", "Settings", "Count", 0)
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnInternetRadioCount() As Integer", Err.Description, Err.Number
End Function

Public Sub CleanUpInternetRadio()
On Local Error GoTo ErrHandler
Dim msg2 As String, msg(255) As String, msg3(255) As String, msg4(255) As String, i As Integer, c As Integer
For i = 1 To 255
    msg2 = ReadINI(App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "Name", "")
    If Len(msg2) <> 0 Then
        c = c + 1
        msg(c) = msg2
        msg3(c) = ReadINI(App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "URL", "")
        msg4(c) = ReadINI(App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "SaveToDisk", "")
    End If
Next i
For i = 1 To c
    WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "Name", msg(i)
    WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "URL", msg3(i)
    WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "SaveToDisk", msg4(i)
Next i
WriteINI App.Path & "\data\playlists\radio.ini", "Settings", "Count", c
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub CleanUpInternetRadio()", Err.Description, Err.Number
End Sub

Public Function ReturnInternetRadioAddress(lName As String) As String
On Local Error Resume Next
Dim msg As String, msg2 As String, lRadioFile As String, i As Integer, c As Integer
lRadioFile = App.Path & "\data\playlists\radio.ini"
If Len(lName) <> 0 Then
    If DoesFileExist(lRadioFile) = True Then
        c = Int(ReadINI(lRadioFile, "Settings", "Count", 0))
        If c <> 0 Then
            For i = 1 To c
                msg2 = ReadINI(lRadioFile, Trim(Str(i)), "Name", "")
                If LCase(msg2) = LCase(lName) Then
                    msg2 = ""
                    msg2 = ReadINI(lRadioFile, Trim(Str(i)), "URL", "")
                    ReturnInternetRadioAddress = msg2
                    Exit For
                End If
            Next i
        End If
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnInternetRadioAddress(lName As String) As String", Err.Description, Err.Number
End Function

Public Sub LoadInternetRadio()
On Local Error Resume Next
Dim lRadioFile As String, c As Integer, t As Integer, i As Integer, msg As String, msg2 As String, b As Boolean, msg3 As String
t = FindTreeViewIndexByFileTitle("Radio", frmMain.tvwFunctions)
If t <> 0 Then
    lRadioFile = App.Path & "\data\playlists\radio.ini"
    If DoesFileExist(lRadioFile) = True Then
        c = Int(ReadINI(lRadioFile, "Settings", "Count", 0))
        If c <> 0 Then
            For i = 1 To c
                msg = ""
                msg2 = ""
                msg = ReadINI(lRadioFile, Trim(Str(i)), "Name", "")
                msg2 = ReadINI(lRadioFile, Trim(Str(i)), "URL", "")
                If Len(msg) <> 0 And Len(msg2) <> 0 Then
                    frmMain.tvwFunctions.Nodes.Add "Radio", tvwChild, msg2, msg, 3
                    If Err.Number <> 0 Then
                        msg3 = "The internet radio entry: " & msg & " has the same url as another entry and will not be added" & vbCrLf & msg3
                        b = True
                        Err.Clear
                    End If
                End If
            Next i
        End If
    End If
End If
If b = True Then
    MsgBox "Duplicate entries were found in your internet radio list" & vbCrLf & vbCrLf & msg3, vbExclamation
End If
If Err.Number <> 0 Then
    ProcessRuntimeError "Public Sub LoadInternetRadio()", Err.Description, Err.Number
End If
End Sub

Public Sub PlayInternetRadio(lRadioControl As ctlRadio, lAddress As String)
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lAddress) <> 0 Then
    If Len(frmMain.lblFilename.Caption) <> 0 Then StopPlayback frmMain.lblFilename.Tag
    lRadioControl.PlayStream lAddress
    frmMain.lblFilename.Caption = lAddress
    frmMain.cmdPlay.Enabled = False
    frmMain.cmdPausePlayback.Enabled = True
    frmMain.cmdStop.Enabled = True
    frmMain.cmdOpen.Enabled = False
    frmMain.cmdBackward.Enabled = False
    frmMain.cmdForeward.Enabled = False
    frmMain.cmdFullScreen.Enabled = False
    For i = 0 To frmMain.fraFunction.Count - 1
        frmMain.tblTop.Buttons(i + 1).Value = tbrUnpressed
        frmMain.fraFunction(i).Visible = False
    Next i
    frmMain.tblTop.Buttons(6).Value = tbrPressed
    frmMain.fraFunction(5).Visible = True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub PlayInternetRadio(lRadioControl As ctlRadio, lAddress As String)", Err.Description, Err.Number
End Sub

Public Sub StopInternetRadio(lRadioControl As ctlRadio)
On Local Error GoTo ErrHandler
lRadioControl.StopStream
frmMain.cmdPlay.Enabled = False
frmMain.cmdOpen.Enabled = True
frmMain.cmdPausePlayback.Enabled = False
frmMain.cmdBackward.Enabled = False
frmMain.cmdForeward.Enabled = False
frmMain.cmdFullScreen.Enabled = False
frmMain.lblFilename.Caption = ""
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub StopInternetRadio(lRadioControl As ctlRadio)", Err.Description, Err.Number
End Sub

Public Function PromptAddInternetRadio() As Integer
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String
msg = InputBox("Enter Name of Internet Radio Entry", "Audiogen 2 Radio", "")
If Len(msg) <> 0 Then
    msg2 = InputBox("Enter URL of Internet Radio Entry", "Audiogen 2 Radio", "")
    If Len(msg2) <> 0 Then PromptAddInternetRadio = AddtoInternetRadio(msg, msg2, False)
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function PromptAddInternetRadio() As Integer", Err.Description, Err.Number
End Function

Public Function AddtoInternetRadio(lName As String, lAddress As String, lSave As Boolean) As Integer

On Local Error GoTo ErrHandler
Dim c As Integer, lRadioFile As String
If Len(lName) <> 0 And Len(lAddress) <> 0 Then
    lRadioFile = App.Path & "\data\playlists\radio.ini"
    If DoesFileExist(lRadioFile) = True Then
        c = Int(ReadINI(lRadioFile, "Settings", "Count", 0)) + 1
        WriteINI lRadioFile, "Settings", "Count", Trim(Str(c))
        WriteINI lRadioFile, Trim(Str(c)), "Name", lName
        WriteINI lRadioFile, Trim(Str(c)), "URL", lAddress
        WriteINI lRadioFile, Trim(Str(c)), "SaveToDisk", Trim(Str(lSave))
        frmMain.tvwFunctions.Nodes.Add "Radio", tvwChild, lAddress, lName, 3
        AddtoInternetRadio = c
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function AddtoInternetRadio(lName As String, lAddress As String, lSave As Boolean) As Integer", Err.Description, Err.Number
End Function

Public Function DeleteInternetRadio(lName As String) As Boolean
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, lRadioFile As String, i As Integer, c As Integer
lRadioFile = App.Path & "\data\playlists\radio.ini"
If Len(lName) <> 0 Then
    If DoesFileExist(lRadioFile) = True Then
        c = Int(ReadINI(lRadioFile, "Settings", "Count", 0))
        If c <> 0 Then
            For i = 1 To c
                msg2 = ReadINI(lRadioFile, Trim(Str(i)), "Name", "")
                If LCase(msg2) = LCase(lName) Then
                    WriteINI lRadioFile, Trim(Str(i)), "Name", vbNullString
                    WriteINI lRadioFile, Trim(Str(i)), "URL", vbNullString
                    WriteINI lRadioFile, Trim(Str(i)), "SaveToDisk", vbNullString
                    DeleteInternetRadio = True
                    Exit For
                End If
            Next i
        End If
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function DeleteInternetRadio(lName As String) As Boolean", Err.Description, Err.Number
End Function

Public Function DeleteTreeViewInternetRadio(lTreeView As TreeView, lName As String) As Boolean
On Local Error GoTo ErrHandler
Dim i As Integer
i = FindTreeViewIndex(lName, lTreeView)
If i <> 0 Then
    If DeleteInternetRadio(lName) = True Then
        lTreeView.Nodes.Remove i
        DeleteTreeViewInternetRadio = True
    End If
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function DeleteTreeViewInternetRadio(lTreeView As TreeView, lName As String) As Boolean", Err.Description, Err.Number
End Function
