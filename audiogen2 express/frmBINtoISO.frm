VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmBINtoISO 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - BIN to ISO converter"
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
      TabIndex        =   8
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
      MICON           =   "frmBINtoISO.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdConvert 
      Height          =   375
      Left            =   5160
      TabIndex        =   7
      Top             =   2760
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Convert"
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
      MICON           =   "frmBINtoISO.frx":001C
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
      TabIndex        =   6
      Top             =   2400
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
   Begin VB.TextBox txtISO 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   4605
   End
   Begin OsenXPCntrl.OsenXPButton cmdBrowseBIN 
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
      MICON           =   "frmBINtoISO.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtBIN 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4605
   End
   Begin MSComDlg.CommonDialog dlgISO 
      Left            =   960
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "ISO images (*.iso)|*.iso"
      Flags           =   2
   End
   Begin MSComDlg.CommonDialog dlgBIN 
      Left            =   1440
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "BIN images (*.bin)|*.bin"
   End
   Begin OsenXPCntrl.OsenXPButton cmdBrowseISO 
      Height          =   255
      Left            =   5640
      TabIndex        =   4
      Top             =   480
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
      MICON           =   "frmBINtoISO.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   416
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Label Label2 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBINtoISO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents cConv    As FL_ImageConverter
Attribute cConv.VB_VarHelpID = -1

Private blnCancel   As Boolean

Private Sub cConv_Progress(ByVal Percent As Integer, Cancel As Boolean)
    prg.Value = Percent
    Cancel = blnCancel
    DoEvents
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    frmImgTools.Show
End Sub

Private Sub cmdBrowseBIN_Click()
    On Error GoTo ErrorHandler
    dlgBIN.ShowOpen
    txtBIN = dlgBIN.FileName
ErrorHandler:
End Sub

Private Sub cmdBrowseISO_Click()
    On Error GoTo ErrorHandler
    dlgISO.ShowSave
    txtISO = dlgISO.FileName
ErrorHandler:
End Sub

Private Sub cmdConvert_Click()

    Dim strMsg  As String

    If txtBIN = vbNullString Then
        MsgBox "No BIN file selected.", vbExclamation
        Exit Sub
    End If

    If txtISO = vbNullString Then
        MsgBox "No ISO file selected.", vbExclamation
        Exit Sub
    End If

    If cmdConvert.Caption = "Cancel" Then
        blnCancel = True
        Exit Sub
    End If

    cmdConvert.Caption = "Cancel"
    cmdBack.Enabled = Not cmdBack.Enabled
    
    blnCancel = False
    Select Case cConv.ConvertBINtoISO(txtBIN, txtISO)
        Case BIN2ISO_CANCELED: strMsg = "Canceled"
        Case BIN2ISO_INVALID_MODE: strMsg = "Only Mode-1 BIN supported"
        Case BIN2ISO_NOT_RAW: strMsg = "BIN is not raw or no image"
        Case BIN2ISO_OK: strMsg = "Finished"
        Case BIN2ISO_UNKNOWN: strMsg = "Unknown error"
    End Select

    MsgBox strMsg, vbInformation

    cmdConvert.Caption = "Convert"
    cmdBack.Enabled = Not cmdBack.Enabled

End Sub

Private Sub Form_Load()
    Set cConv = New FL_ImageConverter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Me.Hide
    frmImgTools.Show
End Sub
