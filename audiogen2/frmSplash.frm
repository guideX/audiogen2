VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Audiogen 2 - Splash"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer tmrStartupDelay 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   6240
      Top             =   960
   End
   Begin VB.Frame fraWeb 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   0
      TabIndex        =   12
      Top             =   390
      Visible         =   0   'False
      Width           =   6135
      Begin SHDocVwCtl.WebBrowser ctlWeb 
         Height          =   3615
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   3495
         ExtentX         =   6165
         ExtentY         =   6376
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6240
      Top             =   480
   End
   Begin Audiogen2.XPButton cmdForum 
      Height          =   375
      Left            =   2700
      TabIndex        =   10
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Forum"
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
      MICON           =   "frmSplash.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdWebsite 
      Height          =   375
      Left            =   1350
      TabIndex        =   3
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Website"
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
      MCOL            =   1300
      MPTR            =   99
      MICON           =   "frmSplash.frx":5944
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdRegister 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Register"
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
      MICON           =   "frmSplash.frx":5AA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdAbout 
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "About"
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
      MICON           =   "frmSplash.frx":5C08
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
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   6855
      TabIndex        =   8
      Top             =   7920
      Width           =   6855
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   180
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CheckBox chkShowOnStartup 
         Appearance      =   0  'Flat
         Caption         =   "Show on Startup"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   200
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1815
      End
      Begin Audiogen2.XPButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   5520
         TabIndex        =   1
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
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
         MPTR            =   99
         MICON           =   "frmSplash.frx":5D6A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdStart 
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         TabIndex        =   0
         Top             =   90
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "&Start"
         ENAB            =   0   'False
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
         MICON           =   "frmSplash.frx":5ECC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   6840
         Y1              =   0
         Y2              =   0
      End
   End
   Begin Audiogen2.XPButton cmdSupport 
      Height          =   375
      Left            =   4050
      TabIndex        =   5
      Top             =   0
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Support"
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
      MICON           =   "frmSplash.frx":602E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblErrCount 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "frmSplash.frx":6190
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Error Count:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Download Skins"
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
      Left            =   120
      MouseIcon       =   "frmSplash.frx":62E2
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   840
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Official Audiogen 2 Forum"
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
      Left            =   120
      MouseIcon       =   "frmSplash.frx":6434
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   0
      X2              =   6840
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Image imgSplash 
      Height          =   7290
      Left            =   240
      Picture         =   "frmSplash.frx":6586
      Top             =   600
      Width           =   5925
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lOpen As Boolean

Public Function ReturnOpen() As Boolean
On Local Error GoTo ErrHandler
ReturnOpen = lOpen
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Sub ReturnOpen()", Err.Description, Err.Number
End Function

Private Sub cmdAbout_Click()
On Local Error GoTo ErrHandler
frmAbout.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdAbout_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
On Local Error GoTo ErrHandler
End
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCancel_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdForum_Click()
On Local Error GoTo ErrHandler
If cmdForum.Value = True Then
    cmdWebsite.Value = False
    ctlWeb.Navigate "http://www.tnexgen.com/forum/viewforum.php?f=7&sid=2dddc900b3e67dd9fca3c6f96394f532"
    fraWeb.Visible = True
    imgSplash.Visible = False
Else
    cmdWebsite.Value = False
    fraWeb.Visible = False
    imgSplash.Visible = True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdForum_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdRegister_Click()
On Local Error GoTo ErrHandler
frmRegister.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdRegister_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdStart_Click()
On Local Error GoTo ErrHandler
'Unload Me
frmMain.Show
Exit Sub
ErrHandler:
    MsgBox "There was a problem starting audiogen. Reinstalling will fix this problem.", vbExclamation
    ProcessRuntimeError "Private Sub cmdWebsite_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSupport_Click()
On Local Error GoTo ErrHandler
Surf "mailto:guidex@tnexgen.com", Me.hwnd
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSupport_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdWebsite_Click()
On Local Error GoTo ErrHandler
If cmdWebsite.Value = True Then
    cmdForum.Value = False
    ctlWeb.Navigate "http://www.tnexgen.com"
    imgSplash.Visible = False
    fraWeb.Visible = True
Else
    imgSplash.Visible = True
    fraWeb.Visible = False
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdWebsite_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Dim i As Integer
Me.Visible = False
LoadSettings
lblErrCount.Caption = ReadINI(App.Path & "\data\config\err.ini", "Settings", "Count", "0")
LoadRegInfo
If IsRegistered = True Then
    cmdRegister.Enabled = False
    cmdStart.Enabled = True
    chkShowOnStartup.Visible = True
    chkShowOnStartup.Enabled = True
    i = Int(ReadINI(lIniFiles.iWindowPositions, "frmSplash", "ShowOnStartup", "1"))
    If i = 0 Then
        frmMain.Show
        Unload Me
        Exit Sub
    End If
Else
    Me.Caption = "Audiogen 2 - Unregistered"
    tmrDelay.Enabled = True
    chkShowOnStartup.Visible = False
End If
tmrStartupDelay.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
imgSplash.Left = (Me.Width - imgSplash.Width) / 2
imgSplash.Top = ((Me.Height - imgSplash.Height) / 2) - 250
Line1.X2 = Me.Width
Picture1.Width = Me.Width
cmdRegister.Width = Me.Width / 5
cmdWebsite.Width = Me.Width / 5
cmdWebsite.Left = cmdRegister.Width
cmdForum.Width = Me.Width / 5
cmdForum.Left = cmdWebsite.Width + cmdWebsite.Left
cmdSupport.Width = Me.Width / 5
cmdSupport.Left = cmdForum.Width + cmdWebsite.Width + cmdWebsite.Left
cmdAbout.Width = Me.Width / 5
cmdAbout.Left = cmdSupport.Width + cmdForum.Width + cmdWebsite.Width + cmdWebsite.Left
Picture1.Top = Me.Height - Picture1.Height - 450

cmdCancel.Left = Me.Width - cmdCancel.Width - 200
cmdStart.Left = Me.Width - cmdCancel.Width - cmdStart.Width - 300
fraWeb.Width = Me.Width - 120
fraWeb.Height = Me.Height - 650 - Picture1.Height - 200
ctlWeb.Width = Me.Width - 120
ctlWeb.Height = Me.Height - 1450
If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error GoTo ErrHandler
If chkShowOnStartup.Value = 1 Then
    WriteINI lIniFiles.iWindowPositions, "frmSplash", "ShowOnStartup", "1"
Else
    WriteINI lIniFiles.iWindowPositions, "frmSplash", "ShowOnStartup", "0"
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Unload(Cancel As Integer)", Err.Description, Err.Number
End Sub

Private Sub imgSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error GoTo ErrHandler
FormDrag Me.hwnd
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub imgSplash_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lblErrCount_Click()
On Local Error GoTo ErrHandler
frmErrorReview.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lblErrCount_Click()", Err.Description, Err.Number
End Sub

Private Sub tmrDelay_Timer()
On Local Error GoTo ErrHandler
If prgProgress.Value = 100 Then
    cmdStart.Enabled = True
    tmrDelay.Enabled = False
    prgProgress.Visible = False
    cmdStart.SetFocus
    Exit Sub
End If
prgProgress.Value = prgProgress.Value + 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tmrDelay_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrStartupDelay_Timer()
On Local Error GoTo ErrHandler
tmrStartupDelay.Enabled = False
Me.Visible = True
lOpen = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tmrStartupDelay_Timer()", Err.Description, Err.Number
End Sub
