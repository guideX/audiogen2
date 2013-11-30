Attribute VB_Name = "mdlFunctions"
Option Explicit
Private lViewSize As Integer
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRootAs As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlagsAs As Long
    lpfnCallbackAs As Long
    lParam As Long
    iImage As Long
End Type
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Enum SpecialFolderIDs
    sfidPROGRAMS = &H2
End Enum
Public Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As SpecialFolderIDs, ByRef pIdl As Long) As Long
Public Declare Function SHGetPathFromIDListA Lib "shell32" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Public Const NOERROR = 0
Private Const BIF_RETURNONLYFSDIRS = 1
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "KERNEL32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Dim lDecoderObjectCount As Integer

Public Function ReturnDirCompliant(lData As String) As String
On Local Error GoTo ErrHandler
lData = Replace(lData, "/", "")
lData = Replace(lData, "\", "")
lData = Replace(lData, ":", "")
lData = Replace(lData, "|", "")
lData = Replace(lData, "?", "")
lData = Replace(lData, "<", "")
lData = Replace(lData, ">", "")
ReturnDirCompliant = lData
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnDirCompliant(lData As String) As String", Err.Description, Err.Number
End Function

Public Function MakeNewDir(lDirectory As String) As Boolean
On Local Error GoTo ErrHandler
Dim b As Boolean
b = Dir(lDirectory, vbDirectory) <> ""
If b = False Then MkDir lDirectory
MakeNewDir = True
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function MakeNewDir(lDirectory As String) As Boolean", Err.Description, Err.Number
End Function

Public Function DoesDirectoryExist(lDirectory) As Boolean
On Local Error GoTo ErrHandler
Dim msg As String
msg = Dir(lDirectory)
If Len(msg) <> 0 Then DoesDirectoryExist = True
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function DoesDirectoryExist(lDirectory) As Boolean", Err.Description, Err.Number
End Function

Public Function GetMyDocumentsDir() As String
On Local Error GoTo ErrHandler
Dim sPath As String, IDL As Long, strPath As String, lngPos As Long
If SHGetSpecialFolderLocation(0, sfidPROGRAMS, IDL) = NOERROR Then
    sPath = String$(255, 0)
    SHGetPathFromIDListA IDL, sPath
    lngPos = InStr(sPath, Chr(0))
    If lngPos > 0 Then
        strPath = Left$(sPath, lngPos - 1)
    End If
End If
GetMyDocumentsDir = Left(strPath, Len(strPath) - 19)
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetMyDocumentsDir() As String", Err.Description, Err.Number
End Function

Public Function OpenMediaFile(lFilename As String, Optional lAutoPlay As Boolean) As Boolean
On Local Error GoTo ErrHandler
Dim lAliasName As String, typeDevice As String, lResult As String, lResultMsg As Integer, msg As String, i As Integer
If frmMain.ctlMP3Player.PlayState = 1 Then
    frmMain.ctlMP3Player.Stop
    Sleep 1.2
End If
If LCase(lFilename) = "mp3" Or LCase(lFilename) = "wav" Or LCase(lFilename) = "mpg" Or LCase(lFilename) = "avi" Or lFilename = "" Then Exit Function
If DoesFileExist(lFilename) = False Then Exit Function
If Left(LCase(frmMain.lblFilename.Caption), 7) = "http://" Then
    StopInternetRadio frmMain.ctlRadio1
    frmMain.lblFilename.Caption = ""
End If
frmMain.lblFilename.Caption = GetFileTitle(lFilename)
frmMain.lblFilename.Tag = lFilename
If FindListViewIndexByFileTitle(GetFileTitle(lFilename), frmMain.lvwBurn) = 0 Then AddProcess GetFileTitle(lFilename), "", Right(LCase(lFilename), 3), "Play"
Select Case LCase(Right(lFilename, 4))
Case ".mp3"
    If Right(LCase(lFilename), 4) = ".mp3" Then OpenID3File lFilename
    frmMain.ctlMP3Player.OscilloType = otSpectrum
    frmMain.ctlMP3Player.Play lFilename
Case Else
    frmMain.ctlMovie1.OpenMovie lFilename
    frmMain.ctlMovie1.SetVolume frmMain.sldVolume.Value * 10
    frmMain.ctlMovie1.Visible = True
    If frmMain.ctlMovie1.ReturnTotalSeconds <> 0 Then frmMain.sldProgress.Max = frmMain.ctlMovie1.ReturnTotalSeconds
    If lAutoPlay = True Then frmMain.ctlMovie1.PlayMovie
End Select
frmMain.tmrProgress.Enabled = True
frmMain.ActiveateResize
OpenMediaFile = True
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function OpenMediaFile(lFilename As String, Optional lAutoPlay As Boolean) As Boolean", Err.Description, Err.Number
End Function

Public Function AddToBurnQue(lFile As String) As Boolean
On Local Error GoTo ErrHandler
Dim vFile As Variant, lItem As ListItem, i As Integer, msg As String, l As Long, lErr As Long, br As Boolean, cWav As New cWavReader
msg = lFile
msg = GetFileTitle(msg)
For i = 1 To frmMain.lvwBurn.ListItems.Count
    If Trim(LCase(frmMain.lvwBurn.ListItems(i).SubItems(2))) = Trim(LCase(msg)) Then
        Exit Function
    End If
Next i
br = cWav.OpenFile(lFile)
lErr = Err.Number
On Error GoTo 0
If (br And (lErr = 0)) Then
    Select Case LCase(Right(lFile, 4))
    Case ".wav"
        Set lItem = frmMain.lvwBurn.ListItems.Add(, lFile, "Ready to Burn")
        lItem.SubItems(1) = "0%"
        lItem.SubItems(2) = msg
        lItem.SubItems(3) = "Track " & Str(Trim(lItem.Index))
        lItem.SubItems(4) = "WAV"
        lItem.SubItems(5) = "CDA"
        lItem.SubItems(6) = Left(lFile, Len(lFile) - Len(msg))
        lItem.SubItems(7) = frmMain.cboBurnDrives.Text
    End Select
Else
    MsgBox "'" & lFile & "' is not a 16bit stereo 44.1kHz Wave File.", vbInformation
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function AddToBurnQue(lFile As String) As Boolean", Err.Description, Err.Number
End Function

Public Function ReturnFreeFile() As Integer
On Local Error GoTo ErrHandler
ReturnFreeFile = FreeFile
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnFreeFile() As Integer", Err.Description, Err.Number
End Function

Public Function FindComboBoxIndex(lCombo As ComboBox, lText As String) As Integer
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 0 To lCombo.ListCount
        If Trim(LCase(lCombo.List(i))) = Trim(LCase(lText)) Then
            FindComboBoxIndex = i
            Exit For
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function FindComboBoxIndex(lCombo As ComboBox, lText As String) As Integer", Err.Description, Err.Number
End Function

Public Function FormatCase(lString As String) As String
On Local Error GoTo ErrHandler
Dim lSplt() As String, i As Integer, m As String
lString = LCase(lString)
lSplt = Split(lString, " ")
For i = 0 To UBound(lSplt)
    m = UCase(Left(lSplt(i), 1))
    lSplt(i) = m & Right(lSplt(i), Len(lSplt(i)) - 1)
Next i
FormatCase = Join(lSplt, " ")
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function FormatCase(lString As String) As String", Err.Description, Err.Number
End Function

Public Function ReturnTitleBarHeight(lForm As Form) As Long
On Local Error GoTo ErrHandler
Dim i As Long, r As RECT, r2 As RECT
i = GetWindowRect(lForm.hwnd, r)
i = GetClientRect(lForm.hwnd, r2)
ReturnTitleBarHeight = r.Bottom - r.Top - r2.Bottom
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnTitleBarHeight(lForm As Form) As Long", Err.Description, Err.Number
End Function

Public Function ReturnMyDocumentsDirectory() As String
On Local Error GoTo ErrHandler
Dim sPath As String, IDL As Long, strPath As String, lngPos As Long
If SHGetSpecialFolderLocation(0, sfidPROGRAMS, IDL) = NOERROR Then
    sPath = String$(255, 0)
    SHGetPathFromIDListA IDL, sPath
    lngPos = InStr(sPath, Chr(0))
    If lngPos > 0 Then
        strPath = Left$(sPath, lngPos - 1)
    End If
End If
ReturnMyDocumentsDirectory = Left(strPath, Len(strPath) - 19)
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnMyDocumentsDirectory() As String", Err.Description, Err.Number
End Function

Public Function ReturnWindowBorder(lForm As Form) As Long
On Local Error GoTo ErrHandler
Dim i As Long, r As RECT, r2 As RECT
i = GetWindowRect(lForm.hwnd, r)
i = GetClientRect(lForm.hwnd, r2)
ReturnWindowBorder = r.Right - r.Left - r2.Right - 2
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnWindowBorder(lForm As Form) As Long", Err.Description, Err.Number
End Function

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String
On Local Error GoTo ErrHandler
Dim iNull As Integer, lpIDList As Long, lResult As Long, sPath As String, udtBI As BrowseInfo
With udtBI
    .hWndOwner = hWndOwner
    .lpszTitle = lstrcat(sPrompt, "")
    .ulFlagsAs = BIF_RETURNONLYFSDIRS
End With
lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
    sPath = String$(260, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    iNull = InStr(sPath, vbNullChar)
    If iNull Then sPath = Left$(sPath, iNull - 1)
End If
BrowseForFolder = sPath
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String", Err.Description, Err.Number
End Function

Public Function ReadFile(lFilename As String) As String
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, i As Integer
Open lFilename For Input As #1
While Not EOF(1)
    Line Input #1, msg
    If Len(msg2) <> 0 Then
        msg2 = msg2 & vbCrLf & msg
    Else
        msg2 = msg
    End If
Wend
Close #1
ReadFile = msg2
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReadFile(lFilename As String) As String", Err.Description, Err.Number
End Function

Public Function EndProgram()
On Local Error Resume Next
frmMain.Visible = False
frmMain.wskFreeDB.Close
WriteINI lIniFiles.iWindowPositions, "sldVolume", "Value", frmMain.sldVolume.Value
DoEvents
frmMain.ctlMovie1.StopMovie
frmMain.ctlMovie1.CloseMovie
frmMain.ctlRadio1.StopStream
If frmMain.ctlMP3Player.PlayState = 1 Then frmMain.ctlMP3Player.Stop
WindowPosition frmMain, True
SavePlaylist lIniFiles.iPlaylist
SaveTVToFile frmMain.tvwPlaylist, lIniFiles.iPlaylistTreeView
DoEvents
Unload frmMain
If Err.Number <> 0 Then Err.Clear
End
End Function

Public Function DoesFileExist(lFilename As String) As Boolean
On Local Error GoTo ErrHandler
Dim msg As String
If LCase(Left(lFilename, 7)) = "http://" Then Exit Function
msg = Dir(lFilename)
If msg <> "" Then
    DoesFileExist = True
Else
    DoesFileExist = False
End If
Exit Function
ErrHandler:
    Err.Clear
End Function

Public Sub MakeDir(lDirectory As String)
On Local Error GoTo ErrHandler
If Len(lDirectory) <> 0 Then
    If Right(lDirectory, 1) <> "\" Then lDirectory = lDirectory & "\"
    MkDir lDirectory
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub MakeDir(lDirectory As String)", Err.Description, Err.Number
End Sub

Public Function Parse(lWhole As String, lStart As String, lEnd As String)
On Local Error GoTo ErrHandler
Dim len1 As Integer, len2 As Integer, Str1 As String, Str2 As String
If Len(Trim(lStart)) <> 0 And Len(Trim(lEnd)) <> 0 Then
    len1 = InStr(lWhole, lStart)
    len2 = InStr(lWhole, lEnd)
    Str1 = Right(lWhole, Len(lWhole) - len1)
    Str2 = Right(lWhole, Len(lWhole) - len2)
    Parse = Left(Str1, Len(Str1) - Len(Str2) - 1)
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function Parse(lWhole As String, lStart As String, lEnd As String)", Err.Description, Err.Number
End Function

Public Function GetRnd(Num As Long) As Long
On Local Error GoTo ErrHandler
Randomize Timer
GetRnd = Int((Num * Rnd) + 1)
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetRnd(Num As Long) As Long", Err.Description, Err.Number
End Function

Public Function GetFileTitle(lFilename As String) As String
On Local Error GoTo ErrHandler
Dim msg() As String
If Len(lFilename) <> 0 Then
    msg = Split(lFilename, "\", -1, vbTextCompare)
    GetFileTitle = msg(UBound(msg))
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetFileTitle(lFilename As String) As String", Err.Description, Err.Number
End Function
