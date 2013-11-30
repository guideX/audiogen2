VERSION 5.00
Begin VB.Form frmRegister 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Register"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
   Begin Audiogen2.XPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   1320
      TabIndex        =   8
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "OK"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frmRegister.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Top             =   6120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Cancel"
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
      MICON           =   "frmRegister.frx":69B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdPayPal 
      Height          =   375
      Left            =   960
      TabIndex        =   13
      Top             =   3360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Launch Paypal"
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
      MICON           =   "frmRegister.frx":6B16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraGray 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3495
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   4560
         Width           =   2295
      End
      Begin VB.TextBox txtNickname 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   4200
         Width           =   2295
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nickname:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRegister.frx":6C78
         ForeColor       =   &H00404040&
         Height          =   1095
         Left            =   960
         TabIndex        =   6
         Top             =   3120
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "The $20 Registration is accepted via paypal. Click the button below to register Audiogen 2."
         ForeColor       =   &H00404040&
         Height          =   855
         Left            =   960
         TabIndex        =   5
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "How to register"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmRegister.frx":6D0D
         ForeColor       =   &H00404040&
         Height          =   975
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   2415
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3480
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3480
         Y1              =   5160
         Y2              =   5160
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Register"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Audiogen 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmRegister.frx":6D9A
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCancel_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdOK_Click()
On Local Error GoTo ErrHandler
SetRegNickname txtNickname.Text
SetRegPassword txtPassword.Text
If IsRegistered = True Then
    MsgBox "Thank you very much for registering. All of the money made from Audiogen 2 is spent on the development of Audiogen 2", vbInformation
    frmSplash.cmdRegister.Enabled = False
    frmSplash.Caption = "Audiogen 2"
    Unload Me
Else
    MsgBox "The registration code you have entered is invalid, e-mail guidex@tnexgen.com for instructions", vbExclamation
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdOK_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdPayPal_Click()
On Local Error GoTo ErrHandler
Surf "https://www.paypal.com/xclick/business=guidex%40team-nexgen.com&item_name=Audiogen+Registration&amount=20.00&no_note=1&tax=0&currency_code=USD&lc=US", Me.hwnd
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPayPal_Click()", Err.Description, Err.Number
End Sub

