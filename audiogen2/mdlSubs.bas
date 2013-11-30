Attribute VB_Name = "mdlSubs"
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOPMOST = -1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Const SWP_HIDEWINDOW = &H80
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Enum eWindowTypes
    eAnyWindow = 0
    eMainWindow = 1
End Enum
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Global lPartialPosition As RECT
Private Const SW_SHOWNORMAL = 1
Private lMP3Bitrate As String
Private lMP3SampleRate As String
Private lMP3Channels As String
Private lAutoDeleteWave As Boolean
Private lAttributes As String
Private lAutoEject As Boolean
Private lTestMode As Boolean
Private lAutoNormalize As Boolean
Private lCDSpeed As Integer
Private lRipPath As String
Private lSelectedDirectory As String
Private Type SHELLEXECUTEINFO
     cbSize As Long
     fMask As Long
     hwnd As Long
     lpVerb As String
     lpFile As String
     lpParameters As String
     lpDirectory As String
     nShow As Long
     hInstApp As Long
     lpIDList As Long
     lpClass As String
     hkeyClass As Long
     dwHotKey As Long
     hIcon As Long
     hProcess As Long
End Type
Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400
Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long


Public Sub Sleep(lMilSec)
On Local Error GoTo ErrHandler
Dim c
c = Timer
Do While Timer - c < Val(lMilSec)
    DoEvents
Loop
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub Sleep(lMilSec)", Err.Description, Err.Number
End Sub

Public Sub ShowFileProperties(FormHwnd As Long, sFileName As String)
On Local Error GoTo ErrHandler
Dim udtSEI As SHELLEXECUTEINFO
With udtSEI
       .cbSize = Len(udtSEI)
       .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
       .hwnd = FormHwnd
       .lpVerb = "properties"
       .lpFile = sFileName
       .lpParameters = vbNullChar
       .lpDirectory = vbNullChar
       .nShow = 0
       .hInstApp = 0
       .lpIDList = 0
End With
Call ShellExecuteEX(udtSEI)
If udtSEI.hInstApp <= 32 Then MsgBox sFileName & "not found, There is an error", vbCritical, "Error"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ShowFileProperties(FormHwnd As Long, sFileName As String)", Err.Description, Err.Number
End Sub

Public Function ReturnCDSpeed() As Integer
On Local Error GoTo ErrHandler
ReturnCDSpeed = lCDSpeed
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnCDSpeed() As Integer", Err.Description, Err.Number
    Err.Clear
End Function

Public Sub Surf(lUrl As String, lHwnd As Long)
On Local Error GoTo ErrHandler
Dim msg As String, c As Integer, i As Integer, l As Long
l = ShellExecute(lHwnd, vbNullString, lUrl, vbNullString, "C:\", SW_SHOWNORMAL)
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub Surf(lUrl As String, lHwnd As Long)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub SetCheckBoxValue(lValue As Boolean, lCheckBox As CheckBox)
On Local Error GoTo ErrHandler
If lValue = True Then
    lCheckBox.Value = 1
Else
    lCheckBox.Value = 0
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetCheckBoxValue(lValue As Boolean, lCheckBox As CheckBox)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub EncodeFile(lSourceFile As String, Optional lDestinationFile As String)
On Local Error GoTo ErrHandler
Select Case LCase(Right(lSourceFile, 4))
Case ".wav"
    If Len(lDestinationFile) = 0 Then
        lDestinationFile = lSourceFile
        lDestinationFile = Left(lDestinationFile, Len(lDestinationFile) - 4) & ".mp3"
    End If
    With frmMain.ctlMP3Encode
        .channels = Int(lMP3Channels)
        .bitrate = CLng(lMP3Bitrate)
        .OPENFILENAME = lSourceFile
        .savefilename = lDestinationFile
        .Tag = lDestinationFile
        .Encode
    End With
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub EncodeFile(lSourceFile As String, Optional lDestinationFile As String)", Err.Description, Err.Number
End Sub

Public Sub ResetFrames()
On Local Error GoTo ErrHandler
Dim i As Integer
For i = 0 To frmMain.fraFunction.Count - 1
    frmMain.fraFunction(i).Visible = False
Next i
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ResetFrames()", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub ToggleFullScreen(lValue As Boolean)
On Local Error GoTo ErrHandler
If lValue = False Then
    AlwaysOnTop frmBlackBack, False
    AlwaysOnTop frmMain, False
    Unload frmBlackBack
    frmMain.Left = lPartialPosition.Left
    frmMain.Top = lPartialPosition.Top
    frmMain.Width = lPartialPosition.Right
    frmMain.Height = lPartialPosition.Bottom
    frmMain.ctlMovie1.Top = 350
    frmMain.tblTop.Visible = True
    frmMain.fraFunction(5).Top = 810
    ChangeBorderStyle frmMain, vbBSNone
    ShowWindowsTaskbar True
    frmMain.SetFocus
    frmMain.Refresh
ElseIf lValue = True Then
    lPartialPosition.Left = frmMain.Left
    lPartialPosition.Top = frmMain.Top
    lPartialPosition.Right = frmMain.Width
    lPartialPosition.Bottom = frmMain.Height
    If frmMain.WindowState = vbMaximized Then frmMain.WindowState = vbNormal
    frmMain.Top = -280
    frmMain.Left = 0
    frmMain.Width = Screen.Width
    frmMain.Height = Screen.Height + 200
    frmMain.fraFunction(5).Top = 0
    frmMain.ctlMovie1.Top = 0
    frmMain.ctlMovie1.Height = 1900
    frmMain.tblTop.Visible = False
    frmMain.ActiveateResize
    ChangeBorderStyle frmMain, vbBSNone
    ShowWindowsTaskbar False
    frmBlackBack.Show
    AlwaysOnTop frmBlackBack, True
    AlwaysOnTop frmMain, True
    frmMain.SetFocus
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ToggleFullScreen(lValue As Boolean)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)
On Local Error GoTo ErrHandler
Dim lFlag As Integer
If SetOnTop Then
    lFlag = HWND_TOPMOST
Else
    lFlag = HWND_NOTOPMOST
End If
SetWindowPos myfrm.hwnd, lFlag, myfrm.Left / Screen.TwipsPerPixelX, myfrm.Top / Screen.TwipsPerPixelY, myfrm.Width / Screen.TwipsPerPixelX, myfrm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub AlwaysOnTop(myfrm As Form, SetOnTop As Boolean)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub ShowWindowsTaskbar(lShow As Boolean)
On Local Error GoTo ErrHandler
Dim rtn As Long
If lShow = False Then
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Else
    rtn = FindWindow("Shell_traywnd", "")
    Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ShowWindowsTaskbar(lShow As Boolean)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub WindowPosition(lWindow As Form, lSave As Boolean)
On Local Error GoTo ErrHandler
Dim msg As String, l As Integer, t As Integer, H As Integer, w As Integer
If lSave = False Then
    l = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "DefaultLeft", "0"))
    t = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "DefaultTop", "0"))
    w = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "DefaultWidth", "0"))
    H = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "DefaultHeight", "0"))
    If l = 0 Then
        Select Case LCase(lWindow.Name)
        Case "frmmain"
            l = 0
        End Select
    End If
    If t = 0 Then
        Select Case LCase(lWindow.Name)
        Case "frmmain"
            t = 0
        End Select
    End If
    If H = 0 Then
        Select Case LCase(lWindow.Name)
        Case "frmmain"
            H = 5900
        End Select
    End If
    If w = 0 Then
        Select Case LCase(lWindow.Name)
        Case "frmmain"
            w = 7900
        End Select
    End If
    lWindow.Width = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "Width", Trim(Str(w))))
    lWindow.Height = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "Height", Trim(Str(H))))
    lWindow.Left = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "Left", Trim(Str(l))))
    lWindow.Top = Int(ReadINI(lIniFiles.iWindowPositions, lWindow.Name, "Top", Trim(Str(t))))
Else
    WriteINI lIniFiles.iWindowPositions, lWindow.Name, "Width", lWindow.Width
    WriteINI lIniFiles.iWindowPositions, lWindow.Name, "Height", lWindow.Height
    WriteINI lIniFiles.iWindowPositions, lWindow.Name, "Top", lWindow.Top
    WriteINI lIniFiles.iWindowPositions, lWindow.Name, "Left", lWindow.Left
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ShowWindowsTaskbar(lShow As Boolean)", Err.Description, Err.Number
    Err.Clear
End Sub

Public Sub StopFlicker(lHwnd As Long)
On Local Error Resume Next
LockWindowUpdate lHwnd
End Sub

Public Sub FormDrag(lHwnd As Long)
On Local Error Resume Next
ReleaseCapture
Call SendMessage(lHwnd, &HA1, 2, 0&)
End Sub

Public Sub LoadSettings()
On Local Error GoTo ErrHandler
lIniFiles.iSpectrum = App.Path & "\data\config\spectrum.ini"
lIniFiles.iSettings = App.Path & "\data\config\settings.ini"
lIniFiles.iRegInfo = App.Path & "\data\config\reg.ini"
lIniFiles.iFileMenu = App.Path & "\data\menus\file.mnu"
lIniFiles.itvwFilesMenu = App.Path & "\data\menus\tvwFiles.mnu"
lIniFiles.itvwFunctionsMenu = App.Path & "\data\menus\tvwFunctions.mnu"
lIniFiles.iPlaylist = App.Path & "\data\playlists\default.m3u"
lIniFiles.iPlaylistTreeView = App.Path & "\data\playlists\default.tvw"
lIniFiles.iWindowPositions = App.Path & "\data\config\settings.ini"
lIniFiles.iCDTracks = App.Path & "\data\config\cdtracks.ini"
lIniFiles.iAttributes = App.Path & "\data\config\attributes.ini"
lIniFiles.iDiscDB = App.Path & "\data\config\discdb.ini"
lIniFiles.iFavorites = App.Path & "\data\config\favorites.ini"
lFileFormats.fSupportedTypes = "*.m3u;*.m4a;*.avi;*.mpg;*.mpeg;*.mpe;*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff;*.aifc;*.au;*.mv1;*.mov;*.mpa;*.qt;*.snd;*.mpm;*.mpv;*.enc;*.mid;*.rmi;*.vob;*.wma;*.wmv"
lMP3Bitrate = ReadINI(lIniFiles.iSettings, "MP3Encoder", "Bitrate", "256000")
lMP3SampleRate = ReadINI(lIniFiles.iSettings, "MP3Encoder", "SampleRate", "44100")
lMP3Channels = ReadINI(lIniFiles.iSettings, "MP3Encoder", "Channels", "2")
lAutoDeleteWave = ReadINI(lIniFiles.iSettings, "MP3Encoder", "AutoDeleteWave", False)
lAttributes = ReadINI(lIniFiles.iSettings, "MP3Decoder", "Attributes", "")
lAutoEject = ReadINI(lIniFiles.iSettings, "CDBurner", "AutoEject", True)
lTestMode = ReadINI(lIniFiles.iSettings, "CDBurner", "TestMode", False)
lAutoNormalize = ReadINI(lIniFiles.iSettings, "CDBurner", "AutoNormalize", False)
lCDSpeed = ReadINI(lIniFiles.iSettings, "CDRipper", "CDSpeed", 0)
lRipPath = ReadINI(lIniFiles.iSettings, "CDRipper", "Rip Path", App.Path & "\data\wave\")
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub LoadSettings()", Err.Description, Err.Number
End Sub

Public Sub HideAllFrames()
On Local Error Resume Next
Dim i As Integer
For i = 0 To frmMain.fraFunction.Count
    frmMain.fraFunction(i).Visible = False
Next i
DoEvents
End Sub

Public Sub AddToTreeView(lTreeView As TreeView, Optional lRelative As String, Optional lRelationship, Optional lKey As String, Optional lText As String, Optional lImage As Integer, Optional IsFile As Boolean, Optional lExpandParent As Boolean)
On Local Error Resume Next
Dim i As Integer, t As Integer
If Len(lText) <> 0 Then
    If DoesTreeViewItemExistByText(lText, lTreeView) = False Then
        If IsFile = True Then
            If DoesFileExist(lKey) = False Then Exit Sub
        End If
        lTreeView.Nodes.Add Trim(lRelative), lRelationship, lKey, lText, lImage
        If Err.Number = 35601 Then
            lTreeView.Nodes.Add , , lRelative, lRelative, lImage
            lTreeView.Nodes.Add lRelative, lRelationship, lKey, lText, lImage
            Err.Clear
        End If
        i = FindTreeViewIndexByFileTitle(lRelative, lTreeView)
        If i <> 0 And lExpandParent = True Then lTreeView.Nodes(i).Expanded = True
    End If
End If
If Err.Number <> 0 Then Err.Clear
End Sub

Public Function RemoveTreeViewMask(lData As String)
On Local Error Resume Next
Dim i As Integer
For i = 0 To 9
    lData = Replace(lData, "a<" & Trim(Str(i)) & ">", "")
Next i
RemoveTreeViewMask = lData
End Function

Public Sub SetRipPath(lData As String)
On Local Error Resume Next
lRipPath = lData
WriteINI lIniFiles.iSettings, "CDRipper", "Rip Path", lData
End Sub

Public Function ReturnRipPath() As String
On Local Error Resume Next
ReturnRipPath = lRipPath
If Err.Number <> 0 Then ProcessRuntimeError "Public Function ReturnRipPath() As String", Err.Description, Err.Number
End Function

Public Sub SetSelectedDirectory(lData As String)
On Local Error Resume Next
lSelectedDirectory = lData
End Sub

Public Function ReturnSelectedDirectory() As String
On Local Error Resume Next
ReturnSelectedDirectory = lSelectedDirectory
End Function

Public Sub FunctionsTreeView()
On Local Error Resume Next
Dim msg2 As String, msg As String, f As tSearch, d As tSearch, t As tSearch, i As Integer, lFullPath As String, lFileTitle As String, c As Integer
LockWindowUpdate frmMain.hwnd
frmMain.tvwFunctions.Visible = False
frmMain.tvwFiles.Visible = False
lFullPath = frmMain.tvwFunctions.SelectedItem.FullPath
If Err.Number <> 0 Then Err.Clear
If Left(LCase(lFullPath), 9) = "cd drives" Then
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
ElseIf InStr(LCase(lFullPath), "wma files") Then
    frmMain.tvwFiles.Nodes.Clear
    For i = 1 To frmMain.tvwPlaylist.Nodes.Count
        If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4)) = ".wma" Then
            frmMain.tvwFiles.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Text, 3
        End If
    Next i
ElseIf InStr(LCase(lFullPath), "wave files") Then
    frmMain.tvwFiles.Nodes.Clear
    For i = 1 To frmMain.tvwPlaylist.Nodes.Count
        If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4)) = ".wav" Then
            frmMain.tvwFiles.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Text, 3
        End If
    Next i
ElseIf InStr(LCase(lFullPath), "mp3 files") Then
    frmMain.tvwFiles.Nodes.Clear
    For i = 1 To frmMain.tvwPlaylist.Nodes.Count
        If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4)) = ".mp3" Then
            frmMain.tvwFiles.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, frmMain.tvwPlaylist.Nodes(i).Text, 3
        End If
    Next i
ElseIf InStr(LCase(lFullPath), "m4a files") Then
    frmMain.tvwFiles.Nodes.Clear
    For i = 0 To frmMain.tvwPlaylist.Nodes.Count
        If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4)) = ".m4a" Then
            frmMain.tvwFiles.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, GetFileTitle(frmMain.tvwPlaylist.Nodes(i).Key), 3
        End If
    Next i
ElseIf InStr(LCase(lFullPath), "m3u files") Then
    frmMain.tvwFiles.Nodes.Clear
    For i = 0 To frmMain.tvwPlaylist.Nodes.Count
        If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4)) = ".m3u" Then
            frmMain.tvwFiles.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, GetFileTitle(frmMain.tvwPlaylist.Nodes(i).Key), 3
        End If
    Next i
ElseIf InStr(LCase(lFullPath), "internet radio") And Len(lFullPath) = 14 Then
    frmMain.tvwFunctions.SelectedItem.Sorted = True
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
ElseIf Left(LCase(lFullPath), 8) = "playlist" Then
    FillTreeViewWithPlaylist frmMain.tvwFiles
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
ElseIf Left(LCase(lFullPath), 15) = "internet radio\" Then
    c = Int(ReadINI(App.Path & "\data\playlists\radio.ini", "Settings", "Count", 0))
    If c <> 0 Then
        For i = 1 To c
            msg = ReadINI(App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "Name", "")
            If Len(msg) <> 0 And msg = Right(lFullPath, Len(lFullPath) - 15) Then
                frmMain.cboPath.Text = ReadINI(App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "URL", "")
                Exit For
            End If
        Next i
    End If
    frmMain.tvwFiles.Visible = True
    frmMain.tvwFunctions.Visible = True
    If frmMain.fraFunction(0).Visible = True Then frmMain.tvwFunctions.SetFocus
    LockWindowUpdate 0
    Exit Sub
ElseIf Left(LCase(lFullPath), 8) = "settings" Then
    frmMain.tvwFiles.Nodes.Clear
    If Len(lFullPath) <> 8 Then
        Select Case LCase(frmMain.tvwFunctions.SelectedItem.Text)
        Case "rip path"
            frmMain.tvwFiles.Nodes.Add , , "Current Rip Path", ReturnRipPath(), 12
            frmMain.tvwFiles.Nodes.Add , , "Change Rip Path", "Change", 12
        Case "cd ripper"
            AddToTreeView frmMain.tvwFunctions, "CD Ripper", tvwChild, "CD Speed", "CD Speed", 12
            AddToTreeView frmMain.tvwFunctions, "CD Ripper", tvwChild, "Rip Path", "Rip Path", 12
        Case "cd burner"
            AddToTreeView frmMain.tvwFunctions, "CD Burner", tvwChild, "Auto Eject", "Auto Eject", 12
            AddToTreeView frmMain.tvwFunctions, "CD Burner", tvwChild, "Test Mode", "Test Mode", 12
            AddToTreeView frmMain.tvwFunctions, "CD Burner", tvwChild, "Auto Normalize", "Auto Normalize", 12
        Case "mp3 decoder"
            AddToTreeView frmMain.tvwFunctions, "MP3 Decoder", tvwChild, "Attributes", "Attributes", 12
        Case "mp3 encoder"
            AddToTreeView frmMain.tvwFunctions, "MP3 Encoder", tvwChild, "Bitrate", "Bitrate", 12
            AddToTreeView frmMain.tvwFunctions, "MP3 Encoder", tvwChild, "Sample Rate", "Sample Rate", 12
            AddToTreeView frmMain.tvwFunctions, "MP3 Encoder", tvwChild, "Channels", "Channels", 12
            AddToTreeView frmMain.tvwFunctions, "MP3 Encoder", tvwChild, "Auto Delete Wave", "Auto Delete Wave", 12
        Case "sample rate"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "s48000", "48000", 12
            frmMain.tvwFiles.Nodes.Add , , "s44100", "44100", 12
            frmMain.tvwFiles.Nodes.Add , , "s32000", "32000", 12
            frmMain.tvwFiles.Nodes(FindTreeViewIndex(lMP3SampleRate, frmMain.tvwFiles)).Selected = True
        Case "bitrate"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "b320000", "320000", 12
            frmMain.tvwFiles.Nodes.Add , , "b256000", "256000", 12
            frmMain.tvwFiles.Nodes.Add , , "b224000", "224000", 12
            frmMain.tvwFiles.Nodes.Add , , "b192000", "192000", 12
            frmMain.tvwFiles.Nodes.Add , , "b160000", "160000", 12
            frmMain.tvwFiles.Nodes.Add , , "b128000", "128000", 12
            frmMain.tvwFiles.Nodes.Add , , "b112000", "112000", 12
            frmMain.tvwFiles.Nodes.Add , , "b96000", "96000", 12
            frmMain.tvwFiles.Nodes.Add , , "b80000", "80000", 12
            frmMain.tvwFiles.Nodes.Add , , "b64000", "64000", 12
            frmMain.tvwFiles.Nodes.Add , , "b56000", "56000", 12
            frmMain.tvwFiles.Nodes.Add , , "b48000", "48000", 12
            frmMain.tvwFiles.Nodes.Add , , "b40000", "40000", 12
            frmMain.tvwFiles.Nodes.Add , , "b32000", "32000", 12
            frmMain.tvwFiles.Nodes(FindTreeViewIndex(lMP3Bitrate, frmMain.tvwFiles)).Selected = True
        Case "channels"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "c1", "1", 12
            frmMain.tvwFiles.Nodes.Add , , "c2", "2", 12
            frmMain.tvwFiles.Nodes(FindTreeViewIndex(lMP3Channels, frmMain.tvwFiles)).Selected = True
        Case "auto delete wave"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "adwTrue", "True", 12
            frmMain.tvwFiles.Nodes.Add , , "adwFalse", "False", 12
            If lAutoDeleteWave = True Then
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("True", frmMain.tvwFiles)).Selected = True
            Else
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("False", frmMain.tvwFiles)).Selected = True
            End If
        Case "attributes"
            frmMain.tvwFiles.Sorted = False
            For i = 1 To Int(ReadINI(lIniFiles.iAttributes, "Settings", "Count", 0))
                frmMain.tvwFiles.Nodes.Add , , "Attribute " & Trim(Str(i)), ReadINI(lIniFiles.iAttributes, Trim(Str(i)), "Data", "")
            Next i
            frmMain.tvwFiles.Nodes(FindTreeViewIndex(lAttributes, frmMain.tvwFiles)).Selected = True
        Case "auto eject"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "aeTrue", "True", 12
            frmMain.tvwFiles.Nodes.Add , , "aeFalse", "False", 12
            If lAutoEject = True Then
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("True", frmMain.tvwFiles)).Selected = True
            Else
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("False", frmMain.tvwFiles)).Selected = True
            End If
        Case "test mode"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "tmTrue", "True", 12
            frmMain.tvwFiles.Nodes.Add , , "tmFalse", "False", 12
            If lTestMode = True Then
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("True", frmMain.tvwFiles)).Selected = True
            Else
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("False", frmMain.tvwFiles)).Selected = True
            End If
        Case "auto normalize"
            frmMain.tvwFiles.Sorted = False
            frmMain.tvwFiles.Nodes.Add , , "anTrue", "True", 12
            frmMain.tvwFiles.Nodes.Add , , "anFalse", "False", 12
            If lAutoNormalize = True Then
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("True", frmMain.tvwFiles)).Selected = True
            Else
                frmMain.tvwFiles.Nodes(FindTreeViewIndex("False", frmMain.tvwFiles)).Selected = True
            End If
        Case "cd speed"
            frmMain.ShowCDSpeed
        End Select
    Else
        AddToTreeView frmMain.tvwFunctions, "Settings", tvwChild, "CD Burner", "CD Burner", 12
        AddToTreeView frmMain.tvwFunctions, "Settings", tvwChild, "CD Ripper", "CD Ripper", 12
        AddToTreeView frmMain.tvwFunctions, "Settings", tvwChild, "MP3 Encoder", "MP3 Encoder", 12
        AddToTreeView frmMain.tvwFunctions, "Settings", tvwChild, "MP3 Decoder", "MP3 Decoder", 12
    End If
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
ElseIf Left(LCase(lFullPath), 11) = "hard drives" Then
    If Len(lFullPath) <> 11 Then
        frmMain.tvwFiles.Nodes.Clear
        lFullPath = Right(lFullPath, Len(lFullPath) - 12)
        For i = 1 To frmMain.tvwFunctions.Nodes.Count
            If frmMain.tvwFunctions.Nodes(i).Image = 11 Then frmMain.tvwFunctions.Nodes(i).Image = 10
        Next i
        If frmMain.tvwFunctions.SelectedItem.Image = 10 Then frmMain.tvwFunctions.SelectedItem.Image = 11
        GetFiles lFullPath, lFileFormats.fSupportedTypes, vbNormal, t
        For i = 1 To t.Count
            lFileTitle = GetFileTitle(t.Path(i))
            If DoesTreeViewItemExistByText(lFileTitle, frmMain.tvwFiles) = False Then
                frmMain.tvwFiles.Nodes.Add , , t.Path(i), lFileTitle, 3
                AddToPlaylistDelay t.Path(i)
            End If
            If Err.Number <> 0 Then Err.Clear
        Next i
        If Trim(Str(t.Count)) <> "0" Then frmMain.Caption = "Audiogen 2 - Found " & Trim(Str(t.Count)) & " supported file(s)"
        GetDirs lFullPath, vbDirectory + vbHidden + vbVolume, d
        For i = 1 To d.Count
            lFileTitle = GetFileTitle(d.Path(i))
            If Len(lFileTitle) <> 0 Then
                AddToTreeView frmMain.tvwFunctions, frmMain.tvwFunctions.SelectedItem.Text, tvwChild, lFileTitle, lFileTitle, 10
            End If
            If Err.Number <> 0 Then Err.Clear
        Next i
    Else
        lFullPath = "Installed Hard Drives"
    End If
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
ElseIf Left(LCase(lFullPath), 7) = "folders" Then
    If Len(lFullPath) <> 7 Then
        lFullPath = Right(lFullPath, Len(lFullPath) - 8)
        If Right(lFullPath, 1) <> "\" Then lFullPath = lFullPath & "\"
        lFullPath = Replace(lFullPath, "My Documents", ReturnMyDocumentsDirectory() & "My Documents\")
        lFullPath = Replace(lFullPath, "My Music", ReturnMyDocumentsDirectory() & "My Documents\My Music\")
        lFullPath = Replace(lFullPath, "Desktop", ReturnMyDocumentsDirectory() & "Desktop\")
        lFullPath = Replace(lFullPath, "Copied CD's\", lRipPath)
        lFullPath = Replace(lFullPath, "\\", "\")
        frmMain.tvwFiles.Nodes.Clear
        For i = 1 To frmMain.tvwFunctions.Nodes.Count
            If frmMain.tvwFunctions.Nodes(i).Image = 11 Then frmMain.tvwFunctions.Nodes(i).Image = 10
        Next i
        If frmMain.tvwFunctions.SelectedItem.Image = 10 Then frmMain.tvwFunctions.SelectedItem.Image = 11
        GetFiles lFullPath, lFileFormats.fSupportedTypes, vbNormal, t
        DoEvents
        For i = 1 To t.Count
            lFileTitle = GetFileTitle(t.Path(i))
            If DoesTreeViewItemExist(t.Path(i), frmMain.tvwFiles) = False Then frmMain.tvwFiles.Nodes.Add , , t.Path(i), lFileTitle, 3
            AddToPlaylistDelay t.Path(i)
        Next i
        GetDirs lFullPath, vbDirectory, d
        DoEvents
        For i = 1 To d.Count
            lFileTitle = GetFileTitle(d.Path(i))
            AddToTreeView frmMain.tvwFunctions, frmMain.tvwFunctions.SelectedItem.Key, tvwChild, d.Path(i), lFileTitle, 10
        Next i
    Else
        lFullPath = "Media Folders"
    End If
    If frmMain.tvwFunctions.SelectedItem.Expanded = True Then
        frmMain.tvwFunctions.SelectedItem.Expanded = False
    Else
        frmMain.tvwFunctions.SelectedItem.Expanded = True
    End If
End If
frmMain.tvwFiles.Visible = True
frmMain.tvwFunctions.Visible = True
frmMain.tvwFunctions.SetFocus
LockWindowUpdate 0
lFullPath = Replace(lFullPath, "Internet Radio\", "")
If Len(lFullPath) <> 0 Then
    If FindComboBoxIndex(frmMain.cboPath, lFullPath) = 0 Then
        frmMain.cboPath.AddItem lFullPath
        frmMain.cboPath.Text = lFullPath
    End If
End If
If Err.Number <> 0 Then
    MsgBox Err.Description
    Err.Clear
End If
    
End Sub

Public Sub SetAutoNormalize(lValue As String)
On Local Error GoTo ErrHandler
If Len(lValue) <> 0 Then
    Select Case LCase(lValue)
    Case "true"
        lAutoNormalize = True
        WriteINI lIniFiles.iSettings, "CDBurner", "AutoNormalize", "True"
    Case "false"
        lAutoNormalize = False
        WriteINI lIniFiles.iSettings, "CDBurner", "AutoNormalize", "False"
    End Select
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetAutoNormalize(lValue As String)", Err.Description, Err.Number
End Sub

Public Sub SetCDSpeed(lValue As Integer)
On Local Error GoTo ErrHandler
lCDSpeed = lValue
WriteINI lIniFiles.iSettings, "CDRipper", "SetCDSpeed", Trim(Str(lValue))
ErrHandler:
    ProcessRuntimeError "Public Sub SetCDSpeed(lValue As Integer)", Err.Description, Err.Number
End Sub

Public Sub SetTestMode(lValue As String)
On Local Error GoTo ErrHandler
Select Case LCase(lValue)
Case "true"
    lTestMode = True
    WriteINI lIniFiles.iSettings, "CDBurner", "TestMode", True
Case "false"
    lTestMode = False
    WriteINI lIniFiles.iSettings, "CDBurner", "TestMode", False
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetTestMode(lValue As String)", Err.Description, Err.Number
End Sub

Public Function ReturnTestMode() As Boolean
On Local Error GoTo ErrHandler
ReturnTestMode = lTestMode
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnTestMode() As Boolean", Err.Description, Err.Number
End Function

Public Sub SetAutoEject(lValue As String)
On Local Error GoTo ErrHandler
Select Case LCase(lValue)
Case "true"
    lAutoEject = True
    WriteINI lIniFiles.iSettings, "CDBurner", "AutoEject", "True"
Case "false"
    lAutoEject = False
    WriteINI lIniFiles.iSettings, "CDBurner", "AutoEject", "False"
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetAutoEject(lValue As Boolean)", Err.Description, Err.Number
End Sub

Public Function ReturnAutoEject() As Boolean
On Local Error GoTo ErrHandler
ReturnAutoEject = lAutoEject
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnAutoEject() As Boolean", Err.Description, Err.Number
End Function

Public Sub SetAttributes(lData As String)
On Local Error GoTo ErrHandler
If Len(lData) <> 0 Then
    lData = Replace(lData, "<None>", "")
    lAttributes = lData
    WriteINI lIniFiles.iSettings, "MP3Decoder", "Attributes", lData
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetAttributes(lData As String)", Err.Description, Err.Number
End Sub

Public Function ReturnAutoDeleteWave() As Boolean
On Local Error GoTo ErrHandler
ReturnAutoDeleteWave = lAutoDeleteWave
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnAutoDeleteWave() As Boolean", Err.Description, Err.Number
End Function

Public Sub SetAutoDeleteWave(lValue As String)
On Local Error GoTo ErrHandler
If LCase(lValue) = "true" Then
    lAutoDeleteWave = True
    WriteINI lIniFiles.iSettings, "MP3Encoder", "AutoDeleteWave", "True"
ElseIf LCase(lValue) = "false" Then
    lAutoDeleteWave = False
    WriteINI lIniFiles.iSettings, "MP3Encoder", "AutoDeleteWave", "False"
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetAutoDeleteWave(lValue As String)", Err.Description, Err.Number
End Sub

Public Sub SetMP3Channels(lChannels As String)
On Local Error GoTo ErrHandler
If Len(lChannels) <> 0 Then
    lMP3Channels = lChannels
    WriteINI lIniFiles.iSettings, "MP3Encoder", "Channels", lChannels
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetMP3Channels(lChannels As String)", Err.Description, Err.Number
End Sub

Public Sub SetMP3SampleRate(lSampleRate As String)
On Local Error GoTo ErrHandler
If Len(lSampleRate) <> 0 Then
    lMP3SampleRate = lSampleRate
    WriteINI lIniFiles.iSettings, "MP3Encoder", "SampleRate", lSampleRate
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetMP3Bitrate(lBitRate As String)", Err.Description, Err.Number
End Sub

Public Sub SetMP3Bitrate(lBitRate As String)
On Local Error GoTo ErrHandler
If Len(lBitRate) <> 0 Then
    lMP3Bitrate = lBitRate
    WriteINI lIniFiles.iSettings, "MP3Encoder", "Bitrate", lBitRate
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetMP3Bitrate(lBitRate As String)", Err.Description, Err.Number
End Sub

Public Sub AddProcess(lInputFilename As String, lOutputFilename As String, lConvertFrom As String, lFunction As String)
On Local Error Resume Next
Dim i As Integer, msg As String, lItem As ListItem
Exit Sub
If Len(lInputFilename) <> 0 Then
    msg = GetFileTitle(lInputFilename)
    With frmMain.lvwBurn
        Set lItem = .ListItems.Add(, lInputFilename, "Idle")
        lItem.SubItems(1) = "0%"
        lItem.SubItems(2) = GetFileTitle(lInputFilename)
        lItem.SubItems(3) = GetFileTitle(lOutputFilename)
        lItem.SubItems(4) = lConvertFrom
        lItem.SubItems(5) = lFunction
    End With
End If
If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub StartProcess(lItem As ListItem)
On Local Error GoTo ErrHandler
Dim i As Integer, lFilename As String
lFilename = lItem.SubItems(2)
If Len(lFilename) <> 0 Then
    i = FindTreeViewIndexByFileTitle(lFilename, frmMain.tvwPlaylist)
    If i <> 0 Then lFilename = frmMain.tvwPlaylist.Nodes(i).Key
    If LCase(lItem.SubItems(5)) = "none" Or LCase(lItem.SubItems(5)) = "play" Then
        lItem.Text = "Playing"
        If OpenMediaFile(lFilename, True) = False Then
            i = FindListViewIndexByFileTitle(GetFileTitle(lFilename), frmMain.lvwBurn)
            frmMain.lvwBurn.ListItems(i).Text = "Error"
        End If
        Exit Sub
    End If
    Select Case LCase(lItem.SubItems(4))
    Case "mp3"
        If LCase(LCase(lItem.SubItems(5))) = "wave" Then
            If Len(frmMain.ctlMP3Decode.Tag) = 0 Then
                If ProcessEntry(lFilename, "MP3 to Wave", True) = False Then
                    Exit Sub
                End If
            Else
                MsgBox "Unable to start the decoder now, please wait for other decode to finish", vbExclamation
                Exit Sub
            End If
        End If
    Case "wave"
        If LCase(lItem.SubItems(5)) = "mp3" Then
            If Len(frmMain.ctlMP3Encode.Tag) = 0 Then
                If ProcessEntry(lFilename, "Wave to MP3", True) = False Then
                    Exit Sub
                End If
            Else
                MsgBox "Unable to start the encoder now, please wait for other encode to finish", vbExclamation
                Exit Sub
            End If
        ElseIf LCase(lItem.SubItems(5)) = "wma" Then
            If ProcessEntry(lFilename, "Wave to WMA", True) = False Then
                Exit Sub
            End If
        End If
    Case "wma"
        If LCase(lItem.SubItems(5)) = "wave" Then
            If ProcessEntry(lFilename, "WMA to Wave", True) = False Then
                Exit Sub
            End If
        End If
    End Select
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub StartProcess(lItem As ListItem)", Err.Description, Err.Number
End Sub

Public Function ReturnProcessType(lFilename As String) As String
On Local Error GoTo ErrHandler
Select Case Right(LCase(lFilename), 4)
Case ".mp3"
    ReturnProcessType = "MP3 to Wave"
Case ".wav"
    ReturnProcessType = "Wave to MP3"
Case ".wma"
    ReturnProcessType = "WMA to Wave"
End Select
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnProcessType(lFilename As String) As String", Err.Description, Err.Number
End Function

Public Function ProcessEntry(lFilename As String, lFunction As String, lProcessNow As Boolean) As Boolean
On Local Error Resume Next
Dim mbox As VbMsgBoxResult, i As Integer
If DoesFileExist(lFilename) = True Then
    Select Case LCase(Trim(lFunction))
    Case "play"
        If lProcessNow = True Then
            OpenMediaFile lFilename, True
        ElseIf lProcessNow = False Then
            AddProcess lFilename, "None", "MP3", "None"
        End If
        ProcessEntry = True
    Case "wma to wave"
        Select Case LCase(Right(Trim(lFilename), 4))
        Case ".wma"
            If lProcessNow = True Then
                frmMain.SetBusy True
                frmMain.Caption = "Audiogen 2 - Decode " & GetFileTitle(lFilename)
                frmMain.ctlWMADecode.OPENFILENAME = lFilename
                frmMain.ctlWMADecode.savefilename = Left(lFilename, Len(lFilename) - 4) & ".wav"
                DoEvents
                frmMain.ctlWMADecode.Decode
                frmMain.ctlWMADecode.Tag = lFilename
                i = FindListViewIndexByKey(frmMain.lvwPending, frmMain.ctlWMADecode.Tag)
            ElseIf lProcessNow = False Then
                AddProcess lFilename, Left(lFilename, Len(lFilename) - 4) & ".wav", "WMA", "Wave"
            End If
        Case Else
            MsgBox "Could not decode " & lFilename, vbInformation
        End Select
    Case "mp3 to wave"
        Select Case LCase(Right(Trim(lFilename), 4))
        Case ".mp3"
            If lProcessNow = True Then
                If DoesFileExist(Left(lFilename, Len(lFilename) - 4) & ".wav") = True Then
                    i = FindListViewIndexByKey(frmMain.lvwPending, lFilename)
                    frmMain.lvwPending.ListItems.Remove i
                    Exit Function
                End If
                frmMain.SetBusy True
                frmMain.Caption = "Audiogen 2 - Decode " & GetFileTitle(lFilename)
                frmMain.ctlMP3Decode.OPENFILENAME = lFilename
                frmMain.ctlMP3Decode.savefilename = Left(lFilename, Len(lFilename) - 4) & ".wav"
                DoEvents
                frmMain.ctlMP3Decode.Decode
                frmMain.ctlMP3Decode.Tag = lFilename
            ElseIf lProcessNow = False Then
                AddProcess lFilename, Left(lFilename, Len(lFilename) - 4) & ".wav", "MP3", "Wave"
            End If
        Case Else
            MsgBox "Could not decode " & lFilename, vbInformation
        End Select
        ProcessEntry = True
    Case "wave to mp3"
        If lProcessNow = True Then
            If Right(LCase(lFilename), 4) = ".wav" Then
                If DoesFileExist(Left(lFilename, Len(lFilename) - 4) & ".mp3") = True Then
                    mbox = MsgBox("The file '" & Left(lFilename, Len(lFilename) - 4) & ".mp3" & "' already exists, would you like to overwrite?", vbYesNo + vbQuestion)
                    If mbox = vbNo Then
                        Exit Function
                    Else
                        Kill Left(lFilename, Len(lFilename) - 4) & ".mp3"
                    End If
                End If
                frmMain.Caption = "Audiogen 2 - Encode " & GetFileTitle(lFilename)
                frmMain.SetBusy True
                With frmMain.ctlMP3Encode
                    .channels = Int(lMP3Channels)
                    .bitrate = Int(lMP3Bitrate)
                    .OPENFILENAME = lFilename
                    .savefilename = Left(lFilename, Len(lFilename) - 4) & ".mp3"
                    .Encode
                    .Tag = lFilename
                End With
            End If
        ElseIf lProcessNow = False Then
            If Right(LCase(lFilename), 4) = ".wav" Then AddProcess lFilename, Left(lFilename, Len(lFilename) - 4) & ".mp3", "Wave", "MP3"
        End If
        ProcessEntry = True
    Case "wave to wma"
        Select Case Right(LCase(lFilename), 4)
        Case ".wav"
            If lProcessNow = True Then
                frmMain.Caption = "Audiogen 2 - Encode " & GetFileTitle(lFilename)
                frmMain.ctlWMAEncode.Encode lFilename, Left(lFilename, Len(lFilename) - 4) & ".wma"
                frmMain.ctlWMAEncode.Tag = lFilename
                frmMain.SetBusy True
            ElseIf lProcessNow = False Then
                AddProcess lFilename, Left(lFilename, Len(lFilename) - 4) & ".wma", "Wave", "WMA"
            End If
            ProcessEntry = True
        Case Else
            ProcessEntry = False
        End Select
    Case "wave to cda"
        If lProcessNow = True Then
            AddToBurnQue lFilename
        ElseIf lProcessNow = False Then
            AddToBurnQue lFilename
        End If
    End Select
Else
    ProcessEntry = False
End If
If Err.Number <> 0 Then ProcessRuntimeError "Public Function ProcessEntry(lFilename As String, lFunction As String, lProcessNow As Boolean) As Boolean", Err.Description, Err.Number
End Function

Public Sub StopPlayback(lFilename As String)
On Local Error GoTo ErrHandler
Dim i As Integer, msg As String
If DoesFileExist(lFilename) = True Then
    msg = GetFileTitle(lFilename)
    With frmMain
        For i = 1 To frmMain.lvwBurn.ListItems.Count
            If LCase(frmMain.lvwBurn.ListItems(i).SubItems(2)) = LCase(msg) Then
                frmMain.lvwBurn.ListItems(i).Text = "Stopped"
                frmMain.lvwBurn.ListItems(i).SubItems(1) = "0%"
                Exit For
            End If
        Next i
        If LCase(Right(lFilename, 4)) = ".mp3" Then
            .ctlMP3Player.Stop
        ElseIf Left(LCase(frmMain.lblFilename.Caption), 7) = "http://" Then
            StopInternetRadio frmMain.ctlRadio1
        Else
            .ctlMovie1.StopMovie
            .ctlMovie1.Visible = False
        End If
        .lblFilename.Caption = ""
        .lblFilename.Tag = ""
        .sldProgress.Value = 0
        .sldProgress.Max = 100
    End With
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub StopPlayback(lFilename As String)", Err.Description, Err.Number
End Sub

Public Sub AddToPending(lFile As String, Optional lBurnEvent As Boolean, Optional lPlayEvent As Boolean)
On Local Error Resume Next
Dim lItem As ListItem, lFileTitle As String
lFileTitle = GetFileTitle(lFile)
With frmMain.lvwPending
    If DoesListViewItemExist(lFile, frmMain.lvwPending) = False Then
        Set lItem = .ListItems.Add(, lFile, lFileTitle)
        lItem.SubItems(1) = Left(lFile, Len(lFile) - Len(lFileTitle))
        Select Case Right(LCase(lFileTitle), 4)
        Case ".wav"
            If lBurnEvent = True Then
                lItem.SubItems(2) = "Wave to CDA"
            ElseIf lPlayEvent = True Then
                lItem.SubItems(2) = "Play"
            Else
                lItem.SubItems(2) = "Wave to MP3"
            End If
            lItem.SubItems(3) = "Wave Audio"
        Case ".mp3"
            If lPlayEvent = True Then
                lItem.SubItems(2) = "Play"
            Else
                lItem.SubItems(2) = "MP3 to Wave"
            End If
            lItem.SubItems(3) = "Mpeg Layer 3 Audio"
        Case ".wma"
            If lPlayEvent = True Then
                lItem.SubItems(2) = "Play"
            Else
                lItem.SubItems(2) = "WMA to Wave"
            End If
            lItem.SubItems(3) = "Windows Media Audio"
        Case ".mp1"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Mpeg Layer 1 Audio"
        Case ".mp2"
            lItem.SubItems(3) = "Mpeg Layer 2 Audio"
        Case ".mp4"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Mpeg Layer 4 Audio"
        Case ".avi"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Avi Video"
        Case "mpeg"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Mpeg Video"
        Case ".mpg"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Mpeg Video"
        Case ".asf"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "ASF Video"
        Case ".ogg"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "OGG Audio"
        Case ".mov"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Movie Video"
        Case ".snd"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Sound Audio"
        Case ".mid"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Midi Audio"
        Case ".vob"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Raw DVD Video"
        Case ".wmv"
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = "Windows Media Video"
        Case Else
            lItem.SubItems(2) = "Play"
            lItem.SubItems(3) = UCase(Replace(Right(lFile, 4), ".", ""))
        End Select
        lItem.SubItems(4) = "Qued"
    End If
End With
If Err.Number <> 0 Then ProcessRuntimeError "Public Sub AddToPending(lFile As String)", Err.Description, Err.Number
End Sub
