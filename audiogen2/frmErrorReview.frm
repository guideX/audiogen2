VERSION 5.00
Begin VB.Form frmErrorReview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiogen 2 - Error Review"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmErrorReview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumber 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Top             =   3360
      Width           =   4095
   End
   Begin VB.TextBox txtDescription 
      Height          =   615
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   4095
   End
   Begin Audiogen2.XPButton cmdClose 
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   3720
      Width           =   975
      _ExtentX        =   1720
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
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   0
      FCOLO           =   16711680
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmErrorReview.frx":57E2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox lstErrors 
      Height          =   2400
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label4 
      Caption         =   "Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "time(s)"
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label lblHowManyErrors 
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   3840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "This Error Occured:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
End
Attribute VB_Name = "frmErrorReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSort_Click()

End Sub

Private Sub cmdClose_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdClose_Click()", Err.Description, Err.Number
End Sub

Private Sub Form_Load()
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer, n As Integer, msg As String, b As Boolean
If DoesFileExist(App.Path & "\data\config\err.ini") = True Then c = Int(ReadINI(App.Path & "\data\config\err.ini", "Settings", "Count", ""))
If c <> 0 Then
    For i = 1 To c
        b = False
        msg = ReadINI(App.Path & "\data\config\err.ini", Trim(Str(i)), "Name", "")
        If Len(msg) <> 0 Then
            For n = 0 To lstErrors.ListCount
                If lstErrors.List(n) = msg Then
                    b = True
                    Exit For
                End If
            Next n
            If b = False Then lstErrors.AddItem msg
        End If
    Next i
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub lstErrors_Click()
On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer, n As Integer, msg As String
c = Int(ReadINI(App.Path & "\data\config\err.ini", "Settings", "Count", 0))
If c <> 0 Then
    For i = 1 To c
        If LCase(ReadINI(App.Path & "\data\config\err.ini", Trim(Str(i)), "Name", "")) = LCase(lstErrors.Text) Then
            txtDescription.Text = ReadINI(App.Path & "\data\config\err.ini", Trim(Str(i)), "Description", "")
            txtNumber.Text = ReadINI(App.Path & "\data\config\err.ini", Trim(Str(i)), "Number", "")
            n = n + 1
        End If
    Next i
    lblHowManyErrors.Caption = Trim(Str(n))
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lstErrors_Click()", Err.Description, Err.Number
End Sub
