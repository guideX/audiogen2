Attribute VB_Name = "mdlCDDrives"
Option Explicit
Dim FS
Enum gDriveType
    dFloppyDrive = 1
    dHardDrive = 2
    dCDDrive = 4
End Enum
Private Type gDrive
    dDriveType As gDriveType
    dDriveNumber As Integer
    dDriveLetter As String
End Type
Private Type gDrives
    dCount As Integer
    dDrive(32) As gDrive
    dCDCount As Integer
End Type
Private lDrives As gDrives

Public Function ReturnDriveCount() As Integer
On Local Error GoTo ErrHandler
ReturnDriveCount = lDrives.dCount
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnDriveCount() As Integer", Err.Description, Err.Number
End Function

Public Function ReturnHardDriveLetter(lDriveIndex As Integer) As String
On Local Error GoTo ErrHandler
If lDrives.dDrive(lDriveIndex).dDriveType = dHardDrive Then ReturnHardDriveLetter = lDrives.dDrive(lDriveIndex).dDriveLetter
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function ReturnHardDriveLetter(lDriveIndex As Integer) As String", Err.Description, Err.Number
End Function

Public Sub FillTreeViewWithFolders(lTreeView As TreeView)
On Local Error GoTo ErrHandler
lTreeView.Nodes.Add , , "Folders", "Folders", 6
lTreeView.Nodes.Add , , "Playlist", "Playlist", 6
lTreeView.Nodes.Add "Playlist", tvwChild, "M3U", "M3U Files (*.m3u)", 6
lTreeView.Nodes.Add "Playlist", tvwChild, "MP3", "MP3 Files (*.mp3)", 6
lTreeView.Nodes.Add "Playlist", tvwChild, "WAV_", "Wave Files (*.wav)", 6
lTreeView.Nodes.Add "Playlist", tvwChild, "WMA", "WMA Files (*.wma)", 6
lTreeView.Nodes.Add "Playlist", tvwChild, "M4A", "M4A Files (*.m4a)", 6
lTreeView.Nodes.Add "Folders", tvwChild, "My Music", "My Music", 6
lTreeView.Nodes.Add "Folders", tvwChild, "My Documents", "My Documents", 6
lTreeView.Nodes.Add "Folders", tvwChild, "Copied CD's", "Copied CD's", 6
lTreeView.Nodes.Add "Folders", tvwChild, "Desktop", "Desktop", 1
lTreeView.Nodes.Add , , "Radio", "Internet Radio", 9
lTreeView.Nodes.Add , , "Settings", "Settings", 12
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub FillTreeViewWithFolders(lTreeView As TreeView)", Err.Description, Err.Number
End Sub

Public Sub FillTreeViewWithDrives(lTreeView As TreeView)
On Local Error GoTo ErrHandler
Dim lHardDrives As Node, lCDDrives As Node, i As Integer, lFloppyDisks As Node, t As Integer, f As tSearch
Set lHardDrives = lTreeView.Nodes.Add(, , "HardDrives", "Hard Drives", 8)
Set lCDDrives = lTreeView.Nodes.Add(, , "CDDrives", "CD Drives", 5)
For i = 0 To lDrives.dCount
    Select Case lDrives.dDrive(i).dDriveType
    Case 1
        If Len(lDrives.dDrive(i).dDriveLetter) <> 0 Then
            lTreeView.Nodes.Add "FloppyDrives", tvwChild, lDrives.dDrive(i).dDriveLetter, lDrives.dDrive(i).dDriveLetter, 8
        End If
    Case 2
        If Len(lDrives.dDrive(i).dDriveLetter) <> 0 Then
            lTreeView.Nodes.Add "HardDrives", tvwChild, lDrives.dDrive(i).dDriveLetter, lDrives.dDrive(i).dDriveLetter, 8
        End If
    Case 4
        If Len(lDrives.dDrive(i).dDriveLetter) <> 0 Then
            lTreeView.Nodes.Add "CDDrives", tvwChild, lDrives.dDrive(i).dDriveLetter, lDrives.dDrive(i).dDriveLetter, 5
        End If
    End Select
Next i
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub FillTreeViewWithDrives(lTreeView As TreeView)", Err.Description, Err.Number
End Sub

Public Function FindTreeViewIndex(lText As String, lTreeView As TreeView) As Integer
On Local Error GoTo ErrHandler
Dim i As Integer
For i = 1 To lTreeView.Nodes.Count
    If Trim(LCase(lTreeView.Nodes(i).Text)) = Trim(LCase(lText)) Then
        FindTreeViewIndex = i
        Exit For
    End If
Next i
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function FindTreeViewIndex(lText As String, lTreeView As TreeView) As Integer", Err.Description, Err.Number
End Function

Public Sub FillComboWithCDDrives(lComboBox As ComboBox)
On Local Error GoTo ErrHandler
Dim i As Integer
lComboBox.Clear
For i = 0 To lDrives.dCount
    If lDrives.dDrive(i).dDriveType = dCDDrive Then
        lComboBox.AddItem lDrives.dDrive(i).dDriveLetter
    End If
Next i
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub FillComboWithCDDrives(lComboBox As ComboBox)", Err.Description, Err.Number
End Sub

Public Sub LoadDrives()
On Local Error GoTo ErrHandler
Dim d, lToc As String, lCDDriveLoaded As Boolean
lDrives.dCount = 0
Set FS = CreateObject("scripting.filesystemobject")
For Each d In FS.Drives
    Select Case d.DriveType
    Case 2
        lDrives.dCount = lDrives.dCount + 1
        lDrives.dDrive(lDrives.dCount).dDriveLetter = d
        lDrives.dDrive(lDrives.dCount).dDriveType = dHardDrive
    Case 4
        lDrives.dCount = lDrives.dCount + 1
        lDrives.dCDCount = lDrives.dCDCount + 1
        lDrives.dDrive(lDrives.dCount).dDriveLetter = d
        lDrives.dDrive(lDrives.dCount).dDriveType = dCDDrive
    End Select
Next
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub LoadDrives()", Err.Description, Err.Number
End Sub

Public Sub ProcessCDDriveError(lNumber As Integer)
On Local Error GoTo ErrHandler
Dim msg As String
Select Case lNumber
Case 687
    msg = "Error: Cannot complete request, reached parameter out of range error"
Case 690
    msg = "Error: Please insert an audio cd and try again"
Case 684
    MsgBox "ASPI Drivers are not installed or are out of date. Please install ASPI drivers before using this feature.", vbCritical
Case 690
    MsgBox "Error: " & App.Title & " can not find a cd in the drive, can not continue!", vbExclamation, App.Title
Case 691
    MsgBox "Error: CD has changed!", vbExclamation, "CD Drive"
Case 692
    MsgBox "Error: Not ready!", vbExclamation, "CD Drive"
Case 693
    MsgBox "Error: Seek error!", vbExclamation, "CD Drive"
Case 694
    MsgBox "Error: Read error!", vbExclamation, "CD Drive"
Case 695
    MsgBox "Error: No CD Detected!", vbExclamation, "CD Drive"
Case 696
    MsgBox "Error: General error!", vbExclamation, "CD Drive"
Case 697
    MsgBox "Error: Illegal CD Change!", vbExclamation, "CD Drive"
Case 698
    MsgBox "Error: Drive not found!", vbExclamation, "CD Drive"
Case 699
    MsgBox "Error: DAC Unable!", vbExclamation, "CD Drive"
Case 700
    MsgBox "Error: ASPI error!", vbExclamation, "CD Drive"
Case 701
    MsgBox "Error: User break!", vbExclamation, "CD Drive"
Case 702
    MsgBox "Error: CD time out!", vbExclamation, "CD Drive"
Case 703
    MsgBox "Error: Out of memory!", vbExclamation, "CD Drive"
Case 704
    MsgBox "Error: Sector not found!", vbExclamation, "CD Drive"
Case 712
    MsgBox "Error: Free hard disk space error!", vbExclamation, "CD Drive"
Case 713
    MsgBox "Error: Device not found!", vbExclamation, "CD Drive"
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ProcessCDDriveError(lNumber As Integer)", Err.Description, Err.Number
End Sub

Public Function GetDriveType(lDrive As String) As gDriveType
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
            GetDriveType = lDrives.dDrive(i).dDriveType
            Exit For
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetDriveType(lDrive As String) As gDriveType", Err.Description, Err.Number
End Function

Public Function GetDriveIndex(lDrive As String) As Integer
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
            GetDriveIndex = i
            Exit For
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetDriveIndex(lDrive As String) As Integer", Err.Description, Err.Number
End Function

Public Function GetCDDriveIndex(lDrive As String) As Integer
On Local Error Resume Next
Dim i As Integer, m As Integer
If Len(lDrive) <> 0 Then
    For i = 0 To lDrives.dCount
        If lDrives.dDrive(i).dDriveType = dCDDrive Then
            m = m + 1
            If LCase(lDrives.dDrive(i).dDriveLetter) = LCase(lDrive) Then
                GetCDDriveIndex = m
                Exit For
            End If
        End If
    Next i
End If
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetCDDriveIndex(lDrive As String) As Integer", Err.Description, Err.Number
End Function
