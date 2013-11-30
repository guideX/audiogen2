VERSION 5.00
Begin VB.Form frmInsertCD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Insert CD"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInsertCD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   3000
   StartUpPosition =   1  'CenterOwner
   Begin Audiogen2.XPButton cmdClose 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1920
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Close"
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
      MICON           =   "frmInsertCD.frx":28BA2
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
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   840
      Width           =   3015
      Begin VB.Label lblBody 
         BackStyle       =   0  'Transparent
         Caption         =   "You do not have a blank CD-R or blank CD-RW in the drive below. Please insert blank CD media into your cd drive and try again."
         Height          =   1455
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5040
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   5040
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Label lblStatus 
      Caption         =   "Please Insert a blank CD-R"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmInsertCD.frx":28D04
      Top             =   0
      Width           =   720
   End
End
Attribute VB_Name = "frmInsertCD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
On Local Error Resume Next
Unload Me
End Sub
