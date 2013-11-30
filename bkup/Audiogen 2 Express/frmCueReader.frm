VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmCueReader 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - Cue Sheet Reader"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin OsenXPCntrl.OsenXPButton cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Back"
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
      MICON           =   "frmCueReader.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdExtract 
      Height          =   375
      Left            =   5040
      TabIndex        =   4
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Extract"
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
      MICON           =   "frmCueReader.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2_Express.XP_ProgressBar prg 
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BrushStyle      =   0
      Color           =   12937777
      Scrolling       =   3
   End
   Begin OsenXPCntrl.OsenXPButton cmdBrowse 
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "..."
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
      MICON           =   "frmCueReader.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   5325
   End
   Begin MSComDlg.CommonDialog dlgBIN 
      Left            =   1440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BIN images (*.bin)|*.bin"
      Flags           =   2
   End
   Begin MSComctlLib.TreeView tvwTracks 
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   3043
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   471
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog dlgCUE 
      Left            =   2040
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Cue sheets (*.cue)|*.cue"
      Flags           =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   416
      Y1              =   176
      Y2              =   176
   End
End
Attribute VB_Name = "frmCueReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cCue  As FL_CueReader
Attribute cCue.VB_VarHelpID = -1

Private blnCancel        As Boolean

Private Function TrackMode2Str(m As FL_TrackModes) As String
    Select Case m
        Case MODE_AUDIO: TrackMode2Str = "audio"
        Case MODE_MODE1: TrackMode2Str = "mode 1"
        Case MODE_MODE2: TrackMode2Str = "mode 2"
        Case MODE_MODE2_FORM1: TrackMode2Str = "mode 2 form 1"
        Case MODE_MODE2_FORM2: TrackMode2Str = "mode 2 form 2"
    End Select
End Function

Private Sub cCue_ExtractProgress(ByVal Percent As Integer, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmImgTools.Show
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler

    dlgCUE.ShowOpen
    txtFile = dlgCUE.FileName

    Select Case cCue.OpenCue(txtFile)

        Case CUE_BINARYEXPECTED, CUE_BINFILEEXPECTED, _
             CUE_CUEFILEEXPECTED, CUE_INDEXEXPECTED, _
             CUE_INDEXMSFEXPECTED, CUE_INDEXNUMEXPECTED, _
             CUE_TRACKEXPECTED, CUE_TRACKNUMEXPECTED:
            MsgBox "Invalid fields in cue sheet.", vbExclamation

        Case CUE_BINMISSING:
            MsgBox "BIN image missing.", vbExclamation

        Case CUE_OK:
            ShowTracks

    End Select

ErrorHandler:
End Sub

Private Sub ShowTracks()

    Dim i   As Integer, j   As Integer

    tvwTracks.Nodes.Clear

    For i = 1 To cCue.TrackCount

        tvwTracks.Nodes.Add(, , "trk" & i, "Track " & Format(i, "00") & " - " & TrackMode2Str(cCue.TrackMode(i))).Tag = i

        For j = 0 To cCue.TrackIndexCount(i) - 1
            tvwTracks.Nodes.Add("trk" & i, tvwChild, , "Index " & Format(j + cCue.TrackIndexFirst(i), "00") & " (" & cCue.TrackIndexLBA(i, j) & " LBA)").Tag = i
        Next

    Next

End Sub

Private Sub cmdExtract_Click()
    On Error GoTo ErrorHandler

    If cmdExtract.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    dlgBIN.ShowSave

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdExtract.Caption = "Cancel"

    If cCue.ExtractTrack(tvwTracks.SelectedItem.Tag, dlgBIN.FileName) Then
        MsgBox "Finished", vbInformation
    Else
        MsgBox "Failed (HDD full?)", vbExclamation
    End If

    cmdBack.Enabled = Not cmdBack.Enabled
    cmdExtract.Caption = "Extract track"

ErrorHandler:
End Sub

Private Sub Form_Load()
    Set cCue = New FL_CueReader
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub
