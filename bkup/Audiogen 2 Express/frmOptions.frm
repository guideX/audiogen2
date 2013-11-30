VERSION 5.00
Object = "{B3FF7FA6-B059-4900-8BEC-5C65E3D9C033}#1.0#0"; "xplook.ocx"
Object = "{6930E6FE-7D84-4DFB-BF5F-3D5D5E1A0302}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express - Options"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   418
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin prjXTab.XTab XTab1 
      Height          =   3180
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   5609
      TabCount        =   2
      TabCaption(0)   =   "Data CD Writter"
      TabContCtrlCnt(0)=   1
      Tab(0)ContCtrlCap(1)=   "picDataCD"
      TabCaption(1)   =   "Audio CD Writter"
      TabContCtrlCnt(1)=   1
      Tab(1)ContCtrlCap(1)=   "picAudioCD"
      TabTheme        =   2
      InActiveTabBackStartColor=   -2147483626
      InActiveTabBackEndColor=   -2147483626
      InActiveTabForeColor=   -2147483631
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   -2147483628
      TabStripBackColor=   -2147483626
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   -2147483627
      Begin VB.PictureBox picAudioCD 
         BorderStyle     =   0  'None
         Height          =   2265
         Left            =   -74880
         ScaleHeight     =   2265
         ScaleWidth      =   6105
         TabIndex        =   1
         Top             =   360
         Width           =   6105
         Begin VB.TextBox txtAudioCDTemp 
            Height          =   285
            Left            =   750
            TabIndex        =   2
            Top             =   120
            Width           =   4725
         End
         Begin OsenXPCntrl.OsenXPButton cmdBrowseAudioCDTemp 
            Height          =   255
            Left            =   5520
            TabIndex        =   8
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   9
            TX              =   "..."
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
            MICON           =   "frmOptions.frx":0000
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
            AutoSize        =   -1  'True
            Caption         =   "Temp:"
            Height          =   195
            Left            =   0
            TabIndex        =   3
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.PictureBox picDataCD 
         BorderStyle     =   0  'None
         Height          =   2265
         Left            =   120
         ScaleHeight     =   2265
         ScaleWidth      =   6105
         TabIndex        =   4
         Top             =   360
         Width           =   6105
         Begin OsenXPCntrl.OsenXPButton cmdBrowseDataCDTemp 
            Height          =   255
            Left            =   5520
            TabIndex        =   7
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   450
            BTYPE           =   9
            TX              =   "..."
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
            MICON           =   "frmOptions.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox txtDataCDTemp 
            Height          =   285
            Left            =   750
            TabIndex        =   5
            Top             =   120
            Width           =   4725
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Temp:"
            Height          =   195
            Left            =   0
            TabIndex        =   6
            Top             =   120
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cDataCD     As New FL_CDDataWriter
Private cAudioCD    As New FL_CDAudioWriter

Private Sub chkDigitalPlayback_Click()
    'SaveSetting "Audiogen", "Grabber", "playmode", CBool(chkDigitalPlayback)
End Sub

Private Sub cmdBrowseAudioCDTemp_Click()

    Dim strText As String

    strText = BrowseForFolder("Please select a new temp dir", txtAudioCDTemp, hWnd, True, , True)

    If strText <> vbNullString Then
        txtAudioCDTemp = AddSlash(strText)
        SaveSetting "Audiogen", "AudioCD", "temp", txtAudioCDTemp
    End If

End Sub

Private Sub cmdBrowseDataCDTemp_Click()

    Dim strText As String

    strText = BrowseForFolder("Please select a new temp dir", txtDataCDTemp, hWnd, True, , True)

    If strText <> vbNullString Then
        txtDataCDTemp = AddSlash(strText)
        SaveSetting "Audiogen", "DataCD", "temp", txtDataCDTemp
    End If

End Sub

Private Sub cmdBrowseGrabPath_Click()

    Dim strText As String

    'strText = BrowseForFolder("Please select a new default dir", txtGrabPath, hWnd, True, , True)

    If strText <> vbNullString Then
        'txtGrabPath = AddSlash(strText)
        'SaveSetting "Audiogen", "Grabber", "path", txtGrabPath
    End If

End Sub

Private Sub Form_Load()

    txtDataCDTemp = GetSetting("Audiogen", "DataCD", "temp", cDataCD.TempDir)
    txtAudioCDTemp = GetSetting("Audiogen", "AudioCD", "temp", cAudioCD.TempDir)
'    txtGrabPath = GetSetting("Audiogen", "Grabber", "path", AddSlash(App.Path))
'    chkDigitalPlayback = Abs(CBool(GetSetting("Audiogen", "Grabber", "playmode", 0)))
'    txtTimeout = GetSetting("Audiogen", "Grabber", "timeout", 8)

'    tabstrip.TabIndex = 1
    tabstrip_Click

End Sub

Private Sub tabstrip_Click()
    'Select Case tabstrip.SelectedItem.index
'        Case 1
'            picDataCD.Visible = True
'            picAudioCD.Visible = False
'            picCDDAGrab.Visible = False
'        Case 2
'            picDataCD.Visible = False
'            picAudioCD.Visible = True
'            picCDDAGrab.Visible = False
'        Case 3
'            picDataCD.Visible = False
'            picAudioCD.Visible = False
'            picCDDAGrab.Visible = True
'    End Select
End Sub

Private Sub txtTimeout_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        'SaveSetting "Audiogen", "Grabber", "timeout", txtTimeout
        KeyAscii = 0
    End If
End Sub

Private Sub txtTimeout_LostFocus()
    'SaveSetting "Audiogen", "Grabber", "timeout", txtTimeout
End Sub

