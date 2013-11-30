VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Begin VB.Form frmDataCDSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - Data CD Writer Settings"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5025
   ControlBox      =   0   'False
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
   ScaleHeight     =   143
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OsenXPCntrl.OsenXPButton cmdCancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   1680
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
      MPTR            =   0
      MICON           =   "frmDataCDSettings.frx":0000
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
      Left            =   3960
      TabIndex        =   6
      Top             =   1680
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
      MICON           =   "frmDataCDSettings.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboSpeed 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.CheckBox chkTestmode 
      Caption         =   "Test mode"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1740
   End
   Begin VB.CheckBox chkEjectDisk 
      Caption         =   "Eject disk after write"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkFinalize 
      Caption         =   "Finalize disk"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox chkOnTheFly 
      Caption         =   "On The Fly"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      X1              =   0
      X2              =   416
      Y1              =   104
      Y2              =   104
   End
   Begin VB.Label lblWriteSpeed 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Write speed:"
      Height          =   195
      Left            =   2145
      TabIndex        =   5
      Top             =   165
      Width           =   930
   End
End
Attribute VB_Name = "frmDataCDSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cDrvNfo As New FL_DriveInfo

Private Sub ShowSpeeds()

    Dim i           As Integer
    Dim intSpeeds() As Integer

    intSpeeds = cDrvNfo.GetWriteSpeeds(strDrvID)

    cboSpeed.Clear
    For i = LBound(intSpeeds) To UBound(intSpeeds)
        cboSpeed.AddItem intSpeeds(i) & " KB/s (" & (intSpeeds(i) \ 176) & "x)"
        cboSpeed.ItemData(cboSpeed.ListCount - 1) = intSpeeds(i)
    Next
    cboSpeed.AddItem "Max."
    cboSpeed.ItemData(cboSpeed.ListCount - 1) = &HFFFF&

    cboSpeed.ListIndex = cboSpeed.ListCount - 1

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Me.Hide

    cManager.SetCDRomSpeed strDrvID, &HFFFF&, cboSpeed.ItemData(cboSpeed.ListIndex)

    frmDataCD.OnTheFly = chkOnTheFly
    frmDataCD.Finalize = chkFinalize
    frmDataCD.TestMode = chkTestmode
    frmDataCD.EjectDisk = chkEjectDisk

    frmDataCD.Burn

End Sub

Private Sub Form_Load()
    ShowSpeeds
End Sub
