VERSION 5.00
Begin VB.Form frmEditFavorites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Edit Favorites"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5595
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditFavorites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin Audiogen2.XPButton cmdClose 
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   2640
      Width           =   1095
      _ExtentX        =   1931
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmEditFavorites.frx":6852
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdPlay 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Play"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmEditFavorites.frx":69B4
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Audiogen2.XPButton cmdRemove 
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      BTYPE           =   9
      TX              =   "Remove"
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
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmEditFavorites.frx":6B16
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstFavorites 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmEditFavorites"
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

Private Sub cmdPlay_Click()
On Local Error GoTo ErrHandler
Dim msg As String, i As Integer
msg = ReturnInternetRadioAddress(lstFavorites.Text)
If Len(msg) = 0 Then
    i = FindTreeViewIndex(lstFavorites.Text, frmMain.tvwPlaylist)
    If i <> 0 Then
        ProcessEntry frmMain.tvwPlaylist.Nodes(i).Key, "Play", True
        Exit Sub
    End If
Else
    PlayInternetRadio frmMain.ctlRadio1, msg
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPlay_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdRemove_Click()
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer
c = Int(ReadINI(lIniFiles.iFavorites, "Settings", "Count", 0))
If c <> 0 Then
    For i = 1 To c
        If LCase(ReadINI(lIniFiles.iFavorites, Trim(Str(i)), "Data", "")) = LCase(lstFavorites.Text) Then
            WriteINI lIniFiles.iFavorites, Trim(Str(i)), vbNullString, vbNullString
        End If
    Next i
    For i = 1 To frmMain.mnuFavorite.Count - 1
        If LCase(frmMain.mnuFavorite(i).Caption) = LCase(lstFavorites.Text) Then
            frmMain.mnuFavorite(i).Caption = ""
            frmMain.mnuFavorite(i).Visible = False
        End If
    Next i
    lstFavorites.RemoveItem lstFavorites.ListIndex
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer, msg As String
c = Int(ReadINI(lIniFiles.iFavorites, "Settings", "Count", 0))
If c <> 0 Then
    For i = 1 To c
        msg = ReadINI(lIniFiles.iFavorites, Trim(Str(i)), "Data", "")
        If Len(msg) <> 0 Then lstFavorites.AddItem msg
    Next i
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub
