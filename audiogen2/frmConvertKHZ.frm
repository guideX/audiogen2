VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDFCF4A3-AD96-11D4-9959-0050BACD4F4C}#1.0#0"; "MDec.ocx"
Object = "{34B82A63-9874-11D4-9E66-0020780170C6}#1.0#0"; "MEnc.ocx"
Begin VB.Form frmConvertKHZ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Convert KHZ"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConvertKHZ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin MDECLib.MDec ctlMP3Decoder 
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin MENCLib.MEnc ctlMP3Encoder 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   255
      _Version        =   65536
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   0
   End
   Begin Audiogen2.XPButton cmdSkip 
      Height          =   255
      Left            =   3240
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "Skip File"
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
      MICON           =   "frmConvertKHZ.frx":1708A
      PICN            =   "frmConvertKHZ.frx":170A6
      PICH            =   "frmConvertKHZ.frx":1712F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   4695
      TabIndex        =   1
      Top             =   840
      Width           =   4695
      Begin MSComctlLib.ProgressBar prg 
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Would you like to convert this file to the correct format now? (You will not be able to burn this file otherwise)"
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   120
         Width           =   4335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4680
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4680
         Y1              =   0
         Y2              =   0
      End
   End
   Begin Audiogen2.XPButton cmdConvertNow 
      Height          =   255
      Left            =   1800
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BTYPE           =   9
      TX              =   "Convert Now"
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
      MPTR            =   99
      MICON           =   "frmConvertKHZ.frx":171B8
      PICN            =   "frmConvertKHZ.frx":1731A
      PICH            =   "frmConvertKHZ.frx":1772F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   645
      Left            =   120
      Picture         =   "frmConvertKHZ.frx":17B46
      Top             =   120
      Width           =   645
   End
   Begin VB.Label lblTitle 
      Caption         =   "The file you selected is not a valid 16 bit 44,000 khz file."
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   675
      Left            =   120
      Picture         =   "frmConvertKHZ.frx":18424
      Top             =   2880
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "frmConvertKHZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lConvertKHZ As clsConvertKHZ

Private Sub cmdConvertNow_Click()
On Local Error GoTo ErrHandler
lConvertKHZ.ConvertKHZ Me.Tag, ctlMP3Encoder, ctlMP3Decoder
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdSkip_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSkip_Click()", Err
End Sub

Private Sub ctlMP3Encoder_PercentDone(ByVal nPercent As Long)
On Local Error GoTo ErrHandler
Select Case nPercent
Case 100
    prg.Value = 0
Case Else
    prg.Value = Int(nPercent)
End Select
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Set lConvertKHZ = New clsConvertKHZ
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub
