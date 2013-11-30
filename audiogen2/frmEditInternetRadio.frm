VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmEditInternetRadio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Edit Internet Radio"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditInternetRadio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   4320
   StartUpPosition =   2  'CenterScreen
   Begin Audiogen2.XPButton cmdSave 
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Save"
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
      MICON           =   "frmEditInternetRadio.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdRemove 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Remove"
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
      MICON           =   "frmEditInternetRadio.frx":69B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdCancel 
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
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
      MICON           =   "frmEditInternetRadio.frx":6B16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdAdd 
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   2160
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "&Add"
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
      MICON           =   "frmEditInternetRadio.frx":6C78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fraBlackdrop 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   4335
      Begin VB.CheckBox chkSaveToDisk 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Save to disk"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2040
         TabIndex        =   0
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2040
         TabIndex        =   6
         Top             =   120
         Width           =   2175
      End
      Begin VB.TextBox txtURL 
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   2040
         TabIndex        =   5
         Top             =   525
         Width           =   2175
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4320
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   4320
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "URL:"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   525
         Width           =   1575
      End
   End
   Begin MSComctlLib.Slider sldRadioStations 
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   90
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   2
      SelStart        =   1
      Value           =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select:"
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
   Begin VB.Image Image1 
      Height          =   675
      Left            =   120
      Picture         =   "frmEditInternetRadio.frx":6DDA
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmEditInternetRadio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String
msg = InputBox("Enter the Name:", App.Title)
msg2 = InputBox("Enter the URL of " & msg & ":", App.Title)
AddtoInternetRadio msg, msg2, False
sldRadioStations.Max = ReturnInternetRadioCount
sldRadioStations.Value = 1
sldRadioStations_Change
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdAdd_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdCancel_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCancel_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdRemove_Click()
On Local Error GoTo ErrHandler
Dim msg As String, i As Integer, msg2 As String
msg = App.Path & "\data\playlists\radio.ini"
For i = 1 To ReturnInternetRadioCount + 1
    msg2 = ReturnInternetRadioName(i)
    If Trim(LCase(msg2)) = Trim(LCase(txtName.Text)) Then
        WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "Name", ""
        WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(i)), "URL", ""
        CleanUpInternetRadio
        DoEvents
        sldRadioStations.Max = ReturnInternetRadioCount
        sldRadioStations.Value = 1
        sldRadioStations_Change
        Exit Sub
    End If
Next i
MsgBox "The entry could not be found!"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSave_Click()
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer
WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(sldRadioStations.Value)), "Name", txtName.Text
WriteINI App.Path & "\data\playlists\radio.ini", Trim(Str(sldRadioStations.Value)), "URL", txtURL.Text
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSave_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
If sldRadioStations.LargeChange <> 2 Then sldRadioStations.LargeChange = 2
If sldRadioStations.SmallChange <> 1 Then sldRadioStations.SmallChange = 1
sldRadioStations.Max = ReturnInternetRadioCount
sldRadioStations.Value = 1
sldRadioStations_Change
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub sldRadioStations_Change()
On Local Error GoTo ErrHandler
txtName.Text = ReturnInternetRadioName(sldRadioStations.Value)
If Len(txtName.Text) <> 0 Then
    txtURL.Text = ReturnInternetRadioAddress(txtName.Text)
    SetCheckBoxValue ReturnInternetRadioSaveToDisk(txtName.Text), chkSaveToDisk
Else
    txtName.Text = ""
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub sldRadioStations_Change()", Err.Description, Err.Number
End Sub
