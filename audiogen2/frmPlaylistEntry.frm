VERSION 5.00
Begin VB.Form frmPlaylistEntry 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Edit Playlist Entry"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlaylistEntry.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin Audiogen2.XPButton cmdOK 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "OK"
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
      MICON           =   "frmPlaylistEntry.frx":6852
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
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1215
      ScaleWidth      =   4815
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      Begin VB.TextBox txtFoundIN 
         Height          =   285
         Left            =   960
         TabIndex        =   8
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtDisplay 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox txtFilename 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   120
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Found In:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Display:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Filename:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1695
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4800
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4800
         Y1              =   0
         Y2              =   0
      End
   End
   Begin Audiogen2.XPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   2040
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
      MICON           =   "frmPlaylistEntry.frx":69B4
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
      Caption         =   "To change aspects of playlist entries"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "Edit Playlist Entry"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmPlaylistEntry.frx":6B16
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmPlaylistEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCancel_Click()", Err.Description, Err.Number
End Sub

