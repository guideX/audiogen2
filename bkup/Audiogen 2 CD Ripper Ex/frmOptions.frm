VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 - CD Ripper Options"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6705
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
   ScaleHeight     =   3150
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOptions 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   880
      Width           =   6735
      Begin VB.TextBox txtReadSectors 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtReadOverlap 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2280
         TabIndex        =   12
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox txtStartOffset 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   120
         Width           =   1395
      End
      Begin VB.TextBox txtEndOffset 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   2280
         TabIndex        =   10
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtRetries 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5220
         TabIndex        =   9
         Top             =   120
         Width           =   1395
      End
      Begin VB.TextBox txtBlockCompare 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5220
         TabIndex        =   8
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtCDSpeed 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5220
         TabIndex        =   7
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtSpinUpTime 
         Enabled         =   0   'False
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   5220
         TabIndex        =   6
         Top             =   1200
         Width           =   1395
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   6720
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Read Sectors:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Read Overlap:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Offset:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "End Offset:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Retries:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3900
         TabIndex        =   17
         Top             =   180
         Width           =   1305
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Block Compare:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3900
         TabIndex        =   16
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "CD Speed:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3900
         TabIndex        =   15
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Spin Up Time (s):"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   3900
         TabIndex        =   14
         Top             =   1260
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      TabIndex        =   4
      Top             =   2670
      Width           =   855
   End
   Begin VB.ComboBox cboDrives 
      ForeColor       =   &H00404040&
      Height          =   315
      ItemData        =   "frmOptions.frx":0000
      Left            =   2280
      List            =   "frmOptions.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.ComboBox cboDriveType 
      Enabled         =   0   'False
      ForeColor       =   &H00404040&
      Height          =   315
      ItemData        =   "frmOptions.frx":0004
      Left            =   2280
      List            =   "frmOptions.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmOptions.frx":0008
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblDrives 
      Caption         =   "Drive:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Type:"
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
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lCDRipper As clsCDRipper
Private lLastIndex As Long

Public Property Let CRRip(lRip As clsCDRipper)
'On Local Error Resume Next
Set lCDRipper = lRip
End Property

Private Sub ShowDrives()
'On Local Error Resume Next
Dim ctl As Control, txt As TextBox, cbo As ComboBox, i As Long
For i = 1 To lCDRipper.CDDriveCount
    cboDrives.AddItem lCDRipper.CDDrive(i).Name
Next i
If (cboDrives.ListCount > 0) Then
    cboDrives.ListIndex = 0
Else
    For Each ctl In Me.Controls
        If (TypeOf ctl Is TextBox) Then
            Set txt = ctl
            txt.Enabled = False
            txt.Text = ""
            txt.BackColor = vbButtonFace
        ElseIf (TypeOf ctl Is ComboBox) Then
            Set cbo = ctl
            cbo.Enabled = False
            cbo.ListIndex = -1
            cbo.BackColor = vbButtonFace
        End If
    Next
    cmdClose.Enabled = False
End If
End Sub

Private Sub cboDrives_Click()
'On Local Error Resume Next
'Dim lDrive As clsDrive
'If (lLastIndex > 0) And (lLastIndex <> cboDrives.ListIndex + 1) Then ApplyCDDrive CLng(lLastIndex)
'Set lDrive = ReturnDrive(cboDrives.ListIndex + 1)
'txtReadSectors.Text = lDrive.ReadSectors
'txtReadOverlap.Text = lDrive.ReadOverlap
'txtStartOffset.Text = lDrive.StartOffset
'txtEndOffset.Text = lDrive.EndOffset
'txtRetries.Text = lDrive.Retries
'txtBlockCompare.Text = lDrive.BlockCompare
'txtCDSpeed.Text = lDrive.CDSpeed
'txtSpinUpTime.Text = lDrive.SpinUpTime
'cboDriveType.ListIndex = lDrive.DriveType
'lLastIndex = cboDrives.ListIndex + 1
End Sub

Private Sub cmdCancel_Click()
'On Local Error Resume Next
Unload Me
End Sub

Private Sub cmdOK_Click()
'On Local Error Resume Next
lCDRipper.CDDrive(cboDrives.ListIndex + 1).Apply
Unload Me
End Sub

Private Sub Form_Load()
'On Local Error Resume Next
cboDriveType.AddItem "GENERIC"
cboDriveType.AddItem "TOSHIBA"
cboDriveType.AddItem "TOSHIBANEW"
cboDriveType.AddItem "IBM"
cboDriveType.AddItem "NEC"
cboDriveType.AddItem "DEC"
cboDriveType.AddItem "IMS"
cboDriveType.AddItem "KODAK"
cboDriveType.AddItem "RICOH"
cboDriveType.AddItem "HP"
cboDriveType.AddItem "PHILIPS"
cboDriveType.AddItem "PLASMON"
cboDriveType.AddItem "GRUNDIGCDR100IPW"
cboDriveType.AddItem "MITSUMICDR"
cboDriveType.AddItem "PLEXTOR"
cboDriveType.AddItem "SONY"
cboDriveType.AddItem "YAMAHA"
cboDriveType.AddItem "NRC"
cboDriveType.AddItem "IMSCDD5"
cboDriveType.AddItem "CUSTOMDRIVE"
ShowDrives
End Sub
