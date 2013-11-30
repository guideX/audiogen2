VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmBurnISO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - Burn ISO image"
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
      TabIndex        =   11
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
      MICON           =   "frmBurnISO.frx":0000
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
      TabIndex        =   10
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Burn"
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
      MICON           =   "frmBurnISO.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkFinalize 
      Caption         =   "Finalize disk"
      Height          =   195
      Left            =   720
      TabIndex        =   7
      Top             =   1485
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkEjectDisk 
      Caption         =   "Eject disk after write"
      Height          =   195
      Left            =   720
      TabIndex        =   6
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   855
      Width           =   1650
   End
   Begin VB.ComboBox cboDrv 
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   4770
   End
   Begin VB.CheckBox chkTestmode 
      Caption         =   "Test mode"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   930
      Width           =   1740
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4725
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   285
      Left            =   5520
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   5400
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   416
      Y1              =   176
      Y2              =   176
   End
   Begin VB.Label lblWriteSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Write speed:"
      Height          =   195
      Left            =   2865
      TabIndex        =   9
      Top             =   900
      Width           =   930
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "File:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmBurnISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cISOCD   As FL_CDISOWriter
Attribute cISOCD.VB_VarHelpID = -1

Private cDrvNfo             As New FL_DriveInfo
Private cCDNfo              As New FL_CDInfo

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    intSpeeds = cDrvNfo.GetWriteSpeeds(strDrvID)

    cboSpeed.Clear
    For i = LBound(intSpeeds) To UBound(intSpeeds)
        cboSpeed.AddItem intSpeeds(i) & " KB/s (" & (intSpeeds(i) \ 176) & "x)"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = intSpeeds(i)
    Next
    cboSpeed.AddItem "Max."
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cboDrv_Click()
    strDrvID = cManager.DrvChr2DrvID(Left$(cboDrv.List(cboDrv.ListIndex), 1))
    ShowSpeeds
End Sub

Private Sub cISOCD_ClosingSession()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Closing session..."
        End With
    End With
End Sub

Private Sub cISOCD_Finished()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Finished"
        End With
    End With
End Sub

Private Sub cISOCD_Progress(Percent As Integer)
    On Error Resume Next
    frmDataCDPrg.prg.Value = Percent
End Sub

Private Sub cISOCD_StartWriting()
    With frmDataCDPrg.lstStatus.ListItems
        With .Add(SmallIcon:=1)
            .SubItems(1) = "Writing track..."
        End With
    End With
End Sub

Private Sub cmdBack_Click()
    Unload Me
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo ErrorHandler
    dlgISO.ShowOpen
    txtFile = dlgISO.FileName
ErrorHandler:
End Sub

Private Sub cmdDrvNfo_Click()
    'frmDriveInfo.Show vbModal, Me
End Sub

Private Sub cmdWrite_Click()

    Dim strMsg  As String

    If txtFile = vbNullString Then
        MsgBox "No ISO image selected.", vbExclamation
        Exit Sub
    End If

    cCDNfo.GetInfo strDrvID
    If FileLen(txtFile) > cCDNfo.Capacity Then
        If MsgBox("Image size exceeds disk capacity." & vbCrLf & _
                  "Continue?", vbYesNo Or vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If

    cISOCD.ISOFile = txtFile
    cISOCD.EjectAfterWrite = chkEjectDisk
    cISOCD.NextSessionAllowed = Not CBool(chkFinalize)
    cISOCD.TestMode = chkTestmode

    Me.Hide
    frmDataCDPrg.Show

    Select Case cISOCD.WriteISOtoCD(strDrvID)
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
    Unload frmDataCDPrg

End Sub

Private Sub Form_Load()
    Set cISOCD = New FL_CDISOWriter
    ShowDrives
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub

Private Sub ShowDrives()

    Dim strDrives() As String
    Dim i           As Long

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
