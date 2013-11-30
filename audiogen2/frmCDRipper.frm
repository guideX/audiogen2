VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCDRipper 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - CD Ripper"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCDRipper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlg 
      Left            =   960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Audiogen2.XPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Cancel"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmCDRipper.frx":1708A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdGrab 
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   2400
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmCDRipper.frx":170A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.DriveListBox cboDrv 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   4650
   End
   Begin VB.ListBox lstTracks 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00404040&
      Height          =   1800
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   2925
   End
   Begin VB.OptionButton optWAV 
      Caption         =   "Wave Audio"
      Height          =   240
      Left            =   3120
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1605
   End
   Begin VB.OptionButton optMP3 
      Caption         =   "MP3 Audio"
      Height          =   240
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1605
   End
   Begin VB.ListBox lstBitrate 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   1245
      IntegralHeight  =   0   'False
      Left            =   3120
      TabIndex        =   1
      Top             =   1035
      Width           =   1605
   End
   Begin MSComctlLib.ProgressBar prg 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
   End
End
Attribute VB_Name = "frmCDRipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cGrabber As FL_TrackGrabber
Attribute cGrabber.VB_VarHelpID = -1
Private cCDInfo As New FL_CDInfo
Private cTrackInfo As New FL_TrackInfo
Private cManager As New FL_Manager
Private strDrvID As String
Private blnCancel As Boolean

Private Sub cboDrv_Change()
On Local Error Resume Next
strDrvID = vbNullString
lstTracks.Clear
If cManager.IsCDVDDrive(cboDrv.Drive) Then
    strDrvID = cManager.DrvChr2DrvID(cboDrv.Drive)
    ShowAudioTracks
End If
End Sub

Sub ShowAudioTracks()
On Local Error Resume Next
Dim i As Integer
If Not cCDInfo.GetInfo(strDrvID) Then
    MsgBox "Could not read CD information.", vbExclamation
    Exit Sub
End If
For i = 1 To cCDInfo.Tracks
    If Not cTrackInfo.GetInfo(strDrvID, i) Then
        MsgBox "Could not get info about track " & i, vbAbortRetryIgnore
    Else
        If cTrackInfo.Mode = MODE_AUDIO Then
            lstTracks.AddItem "Track " & Format(i, "00")
            lstTracks.ItemData(lstTracks.ListCount - 1) = i
        End If
    End If
Next
If lstTracks.ListCount = 0 Then
    MsgBox "No audio tracks found!", vbExclamation
End If
End Sub

Private Sub cGrabber_Progress(ByVal Percent As Integer, ByVal Track As Integer, ByVal startLBA As Long, ByVal endLBA As Long, Cancel As Boolean)
On Local Error Resume Next
prg.Value = Percent
Cancel = blnCancel
DoEvents
End Sub

Private Sub cmdCancel_Click()
On Local Error Resume Next
blnCancel = True
End Sub

Private Sub cmdGrab_Click()
On Local Error Resume Next
Dim ret As FL_SAVETRACK
blnCancel = False
If lstTracks.ListIndex < 0 Then
    MsgBox "No track selected."
    Exit Sub
End If
If optWAV Then
    dlg.Filter = "PCM WAV (*.wav)|*.wav"
Else
    If lstBitrate.ListIndex < 0 Then
        MsgBox "No bitrate selected.", vbExclamation
        Exit Sub
    End If
    dlg.Filter = "MPEG-3 audio (*.mp3)|*.mp3"
End If
On Error GoTo ErrorHandler
dlg.ShowSave
On Error GoTo 0
cmdGrab.Enabled = Not cmdGrab.Enabled
cmdCancel.Enabled = Not cmdCancel.Enabled
If optWAV Then
    ret = cGrabber.AudioTrackToWAV(strDrvID, lstTracks.ItemData(lstTracks.ListIndex), dlg.Filename)
Else
    ret = cGrabber.AudioTrackToMP3(strDrvID, lstTracks.ItemData(lstTracks.ListIndex), dlg.Filename, lstBitrate.ItemData(lstBitrate.ListIndex))
End If
cmdGrab.Enabled = Not cmdGrab.Enabled
cmdCancel.Enabled = Not cmdCancel.Enabled
Select Case ret
    Case ST_CANCELED: MsgBox "Canceled.", vbInformation
    Case ST_ENCODER_INIT: MsgBox "Failed to initialize encoder.", vbExclamation
    Case ST_FINISHED: MsgBox "Finished.", vbInformation
    Case ST_INVALID_SESSION: MsgBox "Invalid session."
    Case ST_INVALID_TRACKMODE: MsgBox "Invalid track mode."
    Case ST_INVALID_TRACKNO: MsgBox "Invalid track number."
    Case ST_NOT_READY: MsgBox "Unit not ready.", vbExclamation
    Case ST_READ_ERR: MsgBox "Read error.", vbExclamation
    Case ST_UNKNOWN_ERR: MsgBox "Unknown error.", vbExclamation
    Case ST_WRITE_ERR: MsgBox "Write error.", vbExclamation
End Select
ErrorHandler:
End Sub

Private Sub cmdOK_Click()

End Sub

Private Sub Form_Load()
On Local Error Resume Next
If Not cManager.Init() Then
    MsgBox "No interfaces found.", vbExclamation
    Unload Me
End If
Set cGrabber = New FL_TrackGrabber
AddBitrates
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
cManager.Goodbye
End Sub

Private Sub optMP3_Click()
On Local Error Resume Next
lstBitrate.Enabled = optMP3.Value
End Sub

Private Sub optWAV_Click()
lstBitrate.Enabled = optMP3.Value
End Sub

Private Sub AddBitrates()
On Local Error Resume Next
With lstBitrate
    .AddItem "128 KBit"
    .ItemData(0) = [128 kBits]
    .AddItem "160 KBit"
    .ItemData(1) = [160 kBits]
    .AddItem "192 KBit"
    .ItemData(2) = [192 kBits]
    .AddItem "224 KBit"
    .ItemData(3) = [224 kBits]
    .AddItem "256 KBit"
    .ItemData(4) = [256 kBits]
    .AddItem "320 KBit"
    .ItemData(5) = [320 kBits]
End With
End Sub
