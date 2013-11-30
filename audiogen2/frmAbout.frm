VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - About"
   ClientHeight    =   5280
   ClientLeft      =   8895
   ClientTop       =   4815
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   3465
   StartUpPosition =   1  'CenterOwner
   Begin Audiogen2.XPButton cmdClose 
      Cancel          =   -1  'True
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   12
      Top             =   4800
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Close"
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
      MICON           =   "frmAbout.frx":57E2
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
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   3495
      Begin VB.Label lblVoltNincs 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tamas Henning"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         MouseIcon       =   "frmAbout.frx":5944
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3480
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Label lblAudiogen2Message 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":5A96
         ForeColor       =   &H00404040&
         Height          =   1335
         Left            =   960
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label lblColinFoss 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Colin Foss"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         MouseIcon       =   "frmAbout.frx":5B43
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   1920
         Width           =   2415
      End
      Begin VB.Label lblLeonAiossa 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Leon Aiossa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         MouseIcon       =   "frmAbout.frx":5C95
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lblDevelopmentTeam 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Development Team"
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
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label lblTeamNexgenURL 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Team Nexgen"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Left            =   960
         MouseIcon       =   "frmAbout.frx":5DE7
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label lblPublisher 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Publisher"
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
         Height          =   195
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label lblCopyrightMessage 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "This program is protected by copyright and international treaties."
         ForeColor       =   &H00404040&
         Height          =   735
         Left            =   960
         TabIndex        =   5
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
      Begin VB.Label lblDescription 
         BackStyle       =   0  'Transparent
         Height          =   1455
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   120
      Picture         =   "frmAbout.frx":5F39
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "© 2005-2006 Team Nexgen"
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
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "<Design Mode>"
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
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label lblApplicationName 
      BackStyle       =   0  'Transparent
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
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdClose_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
lblApplicationName.Caption = App.Title
lblVersion.Caption = "Version: " & App.Major & "." & App.Revision
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub lblColinFoss_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblColinFoss.ForeColor = vbGreen
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblColinFoss_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblColinFoss_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblColinFoss.ForeColor = vbBlue
    Surf "mailto:fosscolin@hotmail.com", Me.hwnd
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblColinFoss_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblLeonAiossa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblLeonAiossa.ForeColor = vbGreen
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblLeonAiossa_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblLeonAiossa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblLeonAiossa.ForeColor = vbBlue
    Surf "http://www.tnexgen.com", Me.hwnd
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblLeonAiossa_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblTeamNexgenURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblTeamNexgenURL.ForeColor = vbGreen
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblTeamNexgenURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblTeamNexgenURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblTeamNexgenURL.ForeColor = vbBlue
    Surf "http://www.tnexgen.com", Me.hwnd
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblTeamNexgenURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblVoltNincs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lblVoltNincs.ForeColor = vbBlue
    Surf "ironrainomega@gmail.com", Me.hwnd
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblVoltNincs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub
