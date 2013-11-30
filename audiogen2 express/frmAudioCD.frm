VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmAudioCD 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - Audio CD"
   ClientHeight    =   3075
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
   Icon            =   "frmAudioCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Audiogen2_Express.XP_ProgressBar prgUsed 
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2160
      Width           =   5415
      _ExtentX        =   9551
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
   Begin MSComctlLib.ListView lstTracks 
      Height          =   1545
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   2725
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "img"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#"
         Object.Width           =   1588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Length"
         Object.Width           =   2381
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Filename"
         Object.Width           =   5292
      EndProperty
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton5 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "Clear"
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
      MICON           =   "frmAudioCD.frx":1708A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton4 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "Save"
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
      MICON           =   "frmAudioCD.frx":170A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton OsenXPButton3 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "Load"
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
      MICON           =   "frmAudioCD.frx":170C2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdBack 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Back"
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
      MICON           =   "frmAudioCD.frx":170DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdWrite 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Write"
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
      MICON           =   "frmAudioCD.frx":170FA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComDlg.CommonDialog dlgFLA 
      Left            =   4425
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "FLA projects (*.fla)|*.fla"
      Flags           =   2
   End
   Begin VB.ComboBox cboDrv 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   5385
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   4950
      Top             =   750
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Supported (*.wav;*.mp3)|*.wav;*.mp3|PCM WAV (*.wav)|*.wav|MPEG-3 audio (*.mp3)|*.mp3"
   End
   Begin MSComctlLib.ImageList img 
      Left            =   5475
      Top             =   750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAudioCD.frx":17116
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAudioCD.frx":198C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   416
      Y1              =   168
      Y2              =   168
   End
   Begin VB.Label lblDrive 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   135
      Width           =   435
   End
   Begin VB.Label lblUsed 
      AutoSize        =   -1  'True
      Caption         =   "Used:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   420
   End
   Begin VB.Menu mnuRClick 
      Caption         =   "RClick"
      Visible         =   0   'False
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move up"
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move down"
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemSel 
         Caption         =   "Remove selected"
      End
   End
End
Attribute VB_Name = "frmAudioCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cAudioCD As FL_CDAudioWriter
Attribute cAudioCD.VB_VarHelpID = -1

Private cDrvNfo             As New FL_DriveInfo
Private cCDNfo              As New FL_CDInfo

Public Sub Burn()

    Dim strMsg  As String

    Me.Hide
    frmAudioCDPrg.Show

    frmAudioCDPrg.prgTotal.Max = cAudioCD.FileCount

    Select Case cAudioCD.WriteAudioToCD(strDrvID)
        Case BURNRET_CLOSE_SESSION: strMsg = "Could not close session."
        Case BURNRET_CLOSE_TRACK: strMsg = "Could not cloe track."
        Case BURNRET_FILE_ACCESS: strMsg = "Failed to access a file."
        Case BURNRET_INVALID_MEDIA: strMsg = "Invalid medium in drive."
        Case BURNRET_ISOCREATION: strMsg = "ISO creation failed."
        Case BURNRET_NO_NEXT_WRITABLE_LBA: strMsg = "Could not get next writable LBA."
        Case BURNRET_NOT_EMPTY: strMsg = "Disk is finalized."
        Case BURNRET_OK: strMsg = "Finished."
        Case BURNRET_SYNC_CACHE: strMsg = "Could not synchronize cache."
        Case BURNRET_WPMP: strMsg = "Write Parameters Page invalid"
        Case BURNRET_WRITE: strMsg = "Write error (Buffer Underrun?)"
    End Select

    MsgBox strMsg, vbInformation

    Me.Show
    'Unload frmAudioCDSettings
    Unload frmAudioCDPrg

End Sub

Public Property Let EjectDisk(aval As Boolean)
    cAudioCD.EjectAfterWrite = aval
End Property

Private Sub cAudioCD_CacheProgress(ByVal Percent As Integer, ByVal Track As Integer)
    frmAudioCDPrg.prgTrack.Value = Percent
End Sub

' ClosingSession Event will be fired
' both for FinalizeDisk and Not FinalizeDisk
Private Sub cAudioCD_ClosingSession()
    With frmAudioCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=2)
            .SubItems(1) = "Closing session..."
        End With
    End With
End Sub

Private Sub cAudioCD_ClosingTrack(ByVal Track As Integer)
    With frmAudioCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=2)
            .SubItems(1) = "Closing track..."
        End With
    End With
End Sub

Private Sub cAudioCD_Finished()
    With frmAudioCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=2)
            .SubItems(1) = "Finished"
        End With
    End With
End Sub

Private Sub cAudioCD_StartCaching()
    With frmAudioCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=2)
            .SubItems(1) = "Caching track..."
        End With
    End With
End Sub

Private Sub cAudioCD_StartWriting()
    With frmAudioCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=2)
            .SubItems(1) = "Writing track..."
        End With
    End With
End Sub

Private Sub cAudioCD_WriteProgress(ByVal Percent As Integer, ByVal Track As Integer)
    frmAudioCDPrg.prgTrack.Value = Percent
    frmAudioCDPrg.prgTotal.Value = Track
End Sub

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    UpdateUsedSpace
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdDrvNfo_Click()
'    frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdOptions_Click()
    ''PopupMenu mnuMenu, , cmdOptions.Left, cmdOptions.Top + cmdOptions.Height
End Sub

Private Sub cmdWrite_Click()

    If UpdateUsedSpace Then
        If MsgBox("Data size exceeds disk capacity." & vbCrLf & _
                  "Continue?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    frmAudioCDSettings.Show vbModal, Me

End Sub

Private Sub Form_Load()
    Set cAudioCD = New FL_CDAudioWriter
    ' directory to decode files to
    cAudioCD.TempDir = GetSetting("Audiogen", "AudioCD", "temp", cAudioCD.TempDir)
    ShowDrives
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmSelectProject.Show
    Unload Me
End Sub

Private Sub lstTracks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then mnuRemSel_Click
End Sub

Private Sub lstTracks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuRClick
End Sub

Private Sub lstTracks_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i   As Integer

    ' add dropped files to queue
    For i = 1 To Data.Files.count
        With Data.Files
            If Not cAudioCD.AddFile(.Item(i)) Then
                MsgBox "Could not add " & FileFromPath(.Item(i))
            End If
        End With
    Next

    ShowTracks

    UpdateUsedSpace

End Sub

Private Sub ShowTracks()

    Dim i   As Integer

    lstTracks.ListItems.Clear

    For i = 0 To cAudioCD.FileCount - 1
        With lstTracks.ListItems.Add(Text:=i + 1, SmallIcon:=2)
            .SubItems(1) = FormatTime(cAudioCD.TrackLength(i)) & " min"
            .SubItems(2) = FileFromPath(cAudioCD.file(i))
        End With
    Next

End Sub

Private Function FormatTime(ByVal sec As Long) As String
    FormatTime = Format(sec \ 60, "00") & ":" & _
                 Format(sec - (sec \ 60) * 60, "00")
End Function

Private Sub ShowDrives()

    Dim strDrives() As String
    Dim i           As Long

    ' show only CD writers
    strDrives = GetDriveList(OPT_CDWRITERS)

    For i = LBound(strDrives) To UBound(strDrives) - 1

        cDrvNfo.GetInfo cManager.DrvChr2DrvID(strDrives(i))

        With cDrvNfo
            cboDrv.AddItem strDrives(i) & ": " & _
                           .Vendor & " " & _
                           .Product & " " & _
                           .Revision & " [" & _
                           .HostAdapter & ":" & _
                           .Target & "]"
        End With

    Next

    cboDrv.ListIndex = 0

End Sub

Private Sub mnuClear_Click()
End Sub

Private Sub mnuLoadPrj_Click()
End Sub

Private Sub mnuMoveDown_Click()

    Dim intLastNdx  As Integer

    intLastNdx = lstTracks.SelectedItem.index

    cAudioCD.MoveIndexUp intLastNdx - 1
    ShowTracks

    lstTracks.ListItems(intLastNdx).Selected = True

End Sub

Private Sub mnuMoveUp_Click()

    Dim intLastNdx  As Integer

    intLastNdx = lstTracks.SelectedItem.index

    cAudioCD.MoveIndexDown intLastNdx - 1
    ShowTracks

    lstTracks.ListItems(intLastNdx).Selected = True

End Sub

Private Function UpdateUsedSpace() As Boolean

    Dim cMSFCD      As New FL_MSF
    Dim cMSFTrks    As New FL_MSF

    Dim m           As Integer, s   As Integer
    Dim i           As Integer
    Dim lngLength   As Long

    ' get total length of queue in seconds
    For i = 0 To cAudioCD.FileCount - 1
        lngLength = lngLength + cAudioCD.TrackLength(i)
    Next

    m = lngLength \ 60
    s = lngLength - m * 60
    cMSFTrks.m = m
    cMSFTrks.s = s

    cCDNfo.GetInfo strDrvID
    cMSFCD.LBA = cCDNfo.Capacity \ 2352

    prgUsed.Max = cMSFCD.LBA
    prgUsed.Value = cMSFTrks.LBA

    ' files fit to disk?
    If (cMSFCD.m < m) Or (cMSFCD.m = m And cMSFCD.s < s) Then
        UpdateUsedSpace = True
    End If

End Function

Private Sub mnuRemSel_Click()

    Dim i   As Integer

    With lstTracks.ListItems
        For i = .count To 1 Step -1
            If .Item(i).Selected Then
                .Remove i
                cAudioCD.RemFile i - 1
            End If
        Next
    End With

    UpdateUsedSpace

End Sub

Private Sub mnuSavePrj_Click()

End Sub

Private Sub OsenXPButton3_Click()
    On Error GoTo ErrorHandler

    dlgFLA.ShowOpen
    If cAudioCD.LoadProject(dlgFLA.FileName) Then
        ShowTracks
        UpdateUsedSpace
    Else
        MsgBox "Could not load project.", vbExclamation
    End If

ErrorHandler:

End Sub

Private Sub OsenXPButton4_Click()
    On Error GoTo ErrorHandler

    dlgFLA.ShowSave
    If Not cAudioCD.SaveProject(dlgFLA.FileName) Then
        MsgBox "Could not save project.", vbExclamation
    End If

ErrorHandler:
End Sub

Private Sub OsenXPButton5_Click()
    cAudioCD.Clear
    ShowTracks
    UpdateUsedSpace

End Sub
