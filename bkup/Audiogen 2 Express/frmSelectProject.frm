VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmSelectProject 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 Express"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
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
   Icon            =   "frmSelectProject.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   205
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdExit 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Exit"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
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
      MICON           =   "frmSelectProject.frx":1708A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin OsenXPCntrl.OsenXPButton cmdOK 
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   2640
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
      MICON           =   "frmSelectProject.frx":170A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2_Express.ucCoolList lstPrjs 
      Height          =   2460
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   4339
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackNormal      =   16777215
      BackSelected    =   4210752
      BackSelectedG1  =   4210752
      BoxRadius       =   5
      Focus           =   0   'False
      ItemHeight      =   40
      ItemHeightAuto  =   0   'False
      ItemOffset      =   6
      SelectModeStyle =   4
   End
End
Attribute VB_Name = "frmSelectProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Select Case lstPrjs.ListIndex
        Case 0: frmDataCD.Show: Me.Hide
        Case 1: frmAudioCD.Show: Me.Hide
        'Case 2: frmAudioGrabber.Show: Me.Hide
        Case 2: frmImgTools.Show vbModal, Me
        Case 3: frmOptions.Show vbModal, Me
    End Select

End Sub

Private Sub Form_Load()
    With lstPrjs
        .AddItem "Data CD project"
        .AddItem "Audio CD project"
        '.AddItem "Audio CD Grabber"
        .AddItem "Image tools"
        .AddItem "Options"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'If MsgBox("Quit Audiogen 2 Express?", vbYesNo Or vbQuestion, "Quit?") = vbNo Then
    '    Cancel = 1
    '    Exit Sub
    'End If

    Dim frm As Form
    For Each frm In Forms
        Unload frm
    Next frm

End Sub

Private Sub lstPrjs_DblClick()
    cmdOK_Click
End Sub
