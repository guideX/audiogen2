VERSION 5.00
Begin VB.UserControl ctlCDRipper 
   ClientHeight    =   1170
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1170
   ScaleWidth      =   2490
   Begin VB.CommandButton Command1 
      Caption         =   "Rip"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "ctlCDRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Event RipStarted()
Public Event RipProgress(lValue As Integer)
Public Event RipComplete()
Public Event RipCanceled()
Private lCanceled As Boolean
Private lRipping As Boolean
Private lCDRipperObject As clsCDRipper
Private lTocObject As clsToc

'Public Function ReturnTocEntry(lIndex As Integer) As clsTocEntry
''On Local Error GoTo ErrHandler
'ReturnTocEntry = lTocObject.Entry(lIndex)
'Exit Function
'ErrHandler:
'    MsgBox Err.Description
'    Err.Clear
'End Function

'Public Function ReturnDriveName(lIndex As Integer) As String
''On Local Error GoTo ErrHandler
'Dim i As Integer
'ReturnDriveName = lCDRipperObject.CDDrive(i).Name
'Exit Function
'ErrHandler:
'    MsgBox Err.Description
'    Err.Clear
'End Function

'Public Function ReturnDrive(lIndex As Integer) As clsDrive
''On Local Error GoTo ErrHandler
'ReturnDrive = lCDRipperObject.CDDrive(lIndex)
'Exit Function
'ErrHandler:
'    MsgBox Err.Description
'    Err.Clear
'End Function

'Public Sub ApplyCDDrive(lIndex As Integer)
''On Local Error GoTo ErrHandler
'lCDRipperObject.CDDrive(CLng(lIndex)).Apply
'Exit Sub
'ErrHandler:
'    MsgBox "ApplyCDDrive: " & Err.Description
'    Err.Clear
'End Sub

'Public Sub InitializeObjects()
''On Local Error Resume Next
'Set lCDRipperObject = New clsCDRipper
'Set lTocObject = New clsToc
'End Sub

Public Sub RipTrack(lTrackIndex As Integer, lFile As String)
'On Local Error GoTo ErrHandler
'Dim lWritter As clsWaveWritter, lTocEntry As clsTocEntry, lTrackRip As New clsCDTrackRipper
Dim lWritter As clsWaveWritter, lTrackRip As New clsCDTrackRipper
Set lWritter = New clsWaveWritter
If (lWritter.OpenFile(lFile)) Then
    'Set lTocEntry = lTocObject.Entry(lTrackIndex)
    
'    MsgBox "Got Past Toc Entry"
    RaiseEvent RipStarted
    lTrackRip.CreateForTrack lTocObject.Entry(lTrackIndex)
    If (lTrackRip.OpenRipper()) Then
        Do While lTrackRip.Read
            lWritter.WriteWavData lTrackRip.ReadBufferPtr, lTrackRip.ReadBufferSize
            RaiseEvent RipProgress(lTrackRip.PercentComplete)
            DoEvents
            If (lCanceled) Then Exit Do
        Loop
        lTrackRip.CloseRipper
        lWritter.CloseFile
        If lCanceled = True Then
            RaiseEvent RipCanceled
            Kill lFile
            lRipping = False
        Else
            RaiseEvent RipComplete
            lRipping = False
        End If
    End If
End If
Exit Sub
ErrHandler:
    MsgBox "RipTrack: " & Err.Description
    Err.Clear
End Sub

Public Function DriveName(lIndex As Integer) As String
'On Local Error GoTo ErrHandler
DriveName = lCDRipperObject.CDDrive(lIndex).Name
Exit Function
ErrHandler:
    MsgBox "DriveName: " & Err.Description
    Err.Clear
End Function

Public Sub ShowOptions()
'On Local Error GoTo ErrHandler
frmOptions.Show 1
Exit Sub
ErrHandler:
    MsgBox "ShowOptions: " & Err.Description
    Err.Clear
End Sub

Public Sub SetDrive(lIndex As Integer)
'On Local Error GoTo ErrHandler
lCDRipperObject.CDDrive(CLng(lIndex)).Apply
Exit Sub
ErrHandler:
    MsgBox "SetDrive: " & Err.Description
    Err.Clear
End Sub

Public Function DriveCount() As Integer
'On Local Error GoTo ErrHandler
DriveCount = lCDRipperObject.CDDriveCount
Exit Function
ErrHandler:
    MsgBox "DriveCount: " & Err.Description
    Err.Clear
End Function

Public Sub InitializeObjects()
'On Local Error GoTo ErrHandler
Set lCDRipperObject = New clsCDRipper
Set lTocObject = New clsToc
Exit Sub
ErrHandler:
    MsgBox "InitializeObjects(): " & Err.Description
    Err.Clear
End Sub

Private Sub Command1_Click()
InitializeObjects
RipTrack 1, "d:\test.wav"
End Sub

Private Sub UserControl_Initialize()
'On Local Error GoTo ErrHandler
Exit Sub
ErrHandler:
    MsgBox "UserControl_Initialize(): " & Err.Description
    Err.Clear
End Sub
