VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmPromptWaveFiles 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Add to Burn Que"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPromptWaveFiles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin Audiogen2.XPButton cmdOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&OK"
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
      MICON           =   "frmPromptWaveFiles.frx":28BA2
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
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
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
      MICON           =   "frmPromptWaveFiles.frx":28D04
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   4575
      TabIndex        =   0
      Top             =   840
      Width           =   4575
      Begin MSComctlLib.TreeView tvwWaves 
         Height          =   2955
         Left            =   0
         TabIndex        =   1
         Top             =   15
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   5212
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4560
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4560
         Y1              =   0
         Y2              =   0
      End
   End
   Begin Audiogen2.XPButton imgOther 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Other"
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
      MICON           =   "frmPromptWaveFiles.frx":28E66
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Select your burnable cd content below"
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Add Wave File(s) to Burn Que"
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
      TabIndex        =   4
      Top             =   120
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   30
      Picture         =   "frmPromptWaveFiles.frx":28FC8
      Top             =   60
      Width           =   720
   End
End
Attribute VB_Name = "frmPromptWaveFiles"
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
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub cmdOK_Click()
On Local Error GoTo ErrHandler
Dim i As Integer
For i = 1 To tvwWaves.Nodes.Count
    If tvwWaves.Nodes(i).Checked = True Then
        AddToBurnQue tvwWaves.Nodes(i).Key
    End If
Next i
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdOK_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Dim i As Integer, l As Integer
For i = 1 To frmMain.tvwPlaylist.Nodes.Count
    If LCase(Right(frmMain.tvwPlaylist.Nodes(i).Key, 4) = ".wav") Then
        tvwWaves.Nodes.Add , , frmMain.tvwPlaylist.Nodes(i).Key, GetFileTitle(frmMain.tvwPlaylist.Nodes(i).Key)
    End If
Next i
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub imgOther_Click()
'On Local Error GoTo ErrHandler
Dim msg As String, i As Integer, lNode As Node
msg = OpenDialog(Me, "Wave Files (*.wav)|*.wav|", "Add Wave File", CurDir)
If Len(msg) <> 0 Then
    i = FindTreeViewIndex(msg, frmMain.tvwPlaylist)
    If i <> 0 Then
        MsgBox frmMain.tvwPlaylist.Nodes(i).Text
    Else
        AddToPlaylist msg, False
        i = FindTreeViewIndexByFileTitle(GetFileTitle(msg), frmMain.tvwPlaylist)
        If i <> 0 Then
            Set lNode = tvwWaves.Nodes.Add(, , , GetFileTitle(msg))
            lNode.Checked = True
            lNode.Selected = True
        Else
            MsgBox "An error occured", vbExclamation
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub imgOther_Click()", Err.Description, Err.Number
End Sub
