VERSION 5.00
Object = "{0843A632-51E5-4C05-9CC6-DF653E994327}#3.0#0"; "Audiogen2CDRipper.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Audiogen2CDRipper.ctlCDRipper ctlCDRipper1 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2566
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   1920
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
ctlCDRipper1.InitializeObjects: DoEvents
ctlCDRipper1.RipTrack 1, "d:\test.wav"
End Sub

Private Sub ctlCDRipper1_RipCanceled()
MsgBox "Rip Canceled"
End Sub

Private Sub ctlCDRipper1_RipProgress(lValue As Integer)
Label1.Caption = Str(lValue)
End Sub

Private Sub ctlCDRipper1_RipStarted()
MsgBox "Rip Started"
End Sub
