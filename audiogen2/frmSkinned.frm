VERSION 5.00
Begin VB.Form frmSkinned 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   1725
   ClipControls    =   0   'False
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
   ScaleHeight     =   855
   ScaleWidth      =   1725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgArray4 
      Height          =   255
      Index           =   0
      Left            =   1200
      Top             =   480
      Width           =   255
   End
   Begin VB.Image imgArray 
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgArray3 
      Height          =   255
      Index           =   0
      Left            =   840
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgArray2 
      Height          =   255
      Index           =   0
      Left            =   480
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgArray1 
      Height          =   255
      Index           =   0
      Left            =   120
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "frmSkinned"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Enum eImageTypes
    iNone = 0
    iBurn = 5
    iCDRip = 6
End Enum
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Private Sub ImageRefresh(Optional lExclude As Integer)
'On Local Error GoTo ErrHandler
Dim i As Integer
For i = 0 To imgArray.Count
    If i <> 1 Then
        If imgArray(i).Picture <> imgArray1(i).Picture Then
            imgArray(i).Picture = imgArray1(i).Picture
        End If
    End If
Next i
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Function AddImage(lLeft As Integer, lTop As Integer, lFile1 As String, lFile2 As String, lFile3 As String, lFile4 As String, lType As eImageTypes, lScriptFile As String, lIcon As String) As Image
'On Local Error GoTo ErrHandler
Dim i As Integer
If DoesFileExist(App.Path & "\data\skins\" & lFile1) = True Then
    i = frmSkinned.imgArray1.Count + 1
    Load frmSkinned.imgArray(i)
    Load frmSkinned.imgArray1(i)
    Load frmSkinned.imgArray2(i)
    Load frmSkinned.imgArray3(i)
    Load frmSkinned.imgArray4(i)
    If Len(lFile1) <> 0 And DoesFileExist(App.Path & "\data\skins\" & lFile1) = True Then frmSkinned.imgArray1(i).Picture = LoadPicture(App.Path & "\data\skins\" & lFile1)
    If Len(lFile2) <> 0 And DoesFileExist(App.Path & "\data\skins\" & lFile2) = True Then frmSkinned.imgArray2(i).Picture = LoadPicture(App.Path & "\data\skins\" & lFile2)
    If Len(lFile3) <> 0 And DoesFileExist(App.Path & "\data\skins\" & lFile3) = True Then frmSkinned.imgArray3(i).Picture = LoadPicture(App.Path & "\data\skins\" & lFile3)
    If Len(lFile4) <> 0 And DoesFileExist(App.Path & "\data\skins\" & lFile4) = True Then frmSkinned.imgArray4(i).Picture = LoadPicture(App.Path & "\data\skins\" & lFile4)
    If Right(lScriptFile, 1) <> "\" Then
        If DoesFileExist(lScriptFile) = True Then
            frmSkinned.imgArray1(i).Tag = ReadINI(lScriptFile, "Script", "Resize", "")
            If Left(LCase(frmSkinned.imgArray1(i).Tag), 9) = "fromfile(" Then
                frmSkinned.imgArray1(i).Tag = Right(frmSkinned.imgArray1(i).Tag, Len(frmSkinned.imgArray1(i).Tag) - 9)
                frmSkinned.imgArray1(i).Tag = Left(frmSkinned.imgArray1(i).Tag, Len(frmSkinned.imgArray1(i).Tag) - 1)
            End If
            frmSkinned.imgArray2(i).Tag = ReadINI(lScriptFile, "Script", "Click", "")
        End If
    End If
    frmSkinned.imgArray(i).Picture = frmSkinned.imgArray1(i).Picture
    frmSkinned.imgArray(i).Left = lLeft
    frmSkinned.imgArray(i).Top = lTop
    frmSkinned.imgArray(i).Tag = lType
    frmSkinned.imgArray(i).Visible = True
    Select Case LCase(lIcon)
    Case "hand"
        frmSkinned.imgArray(i).MousePointer = 99
        frmSkinned.imgArray(i).MouseIcon = LoadPicture(App.Path & "\data\gfx\hand.cur")
    Case "arrow"
        frmSkinned.imgArray(i).MousePointer = 1
    Case "question"
        frmSkinned.imgArray(i).MousePointer = 14
    Case "hourglass"
        frmSkinned.imgArray(i).MousePointer = 11
    End Select
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub DisableImage(lDisable As Boolean, lIndex As Integer)
'On Local Error GoTo ErrHandler
Select Case lDisable
Case True
    imgArray(lIndex).Picture = imgArray4(lIndex).Picture
    imgArray(lIndex).Enabled = False
Case False
    imgArray(lIndex).Picture = imgArray1(lIndex).Picture
    imgArray(lIndex).Enabled = True
End Select
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function ReturnSkinCount() As Integer
'On Local Error GoTo ErrHandler
ReturnSkinCount = Int(ReadINI(App.Path & "\data\skins\skins.ini", "Settings", "Count", 0))
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Function LoadSkin(lFile As String) As Boolean
'On Local Error GoTo ErrHandler
Dim i As Integer, c As Integer
If Len(lFile) <> 0 And DoesFileExist(lFile) = True Then
    frmSkinned.Width = Int(ReadINI(lFile, "Settings", "Width", 0))
    frmSkinned.Height = Int(ReadINI(lFile, "Settings", "Height", 0))
    c = Int(ReadINI(lFile, "Settings", "ImageCount", 0))
    If c <> 0 Then
        i = c
        Do Until i = 0
            AddImage Int(ReadINI(lFile, Str(Trim(i)), "Left", 0)), Int(ReadINI(lFile, Str(Trim(i)), "Top", 0)), ReadINI(lFile, Str(Trim(i)), "File1", ""), ReadINI(lFile, Str(Trim(i)), "File2", ""), ReadINI(lFile, Str(Trim(i)), "File3", ""), ReadINI(lFile, Str(Trim(i)), "File4", ""), ReadINI(lFile, Str(Trim(i)), "Type", 0), App.Path & "\" & ReadINI(lFile, Str(Trim(i)), "Script", ""), ReadINI(lFile, Str(Trim(i)), "Icon", "arrow")
            i = i - 1
        Loop
    End If
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Private Sub Form_Load()
On Local Error Resume Next
LoadSkin App.Path & "\data\skins\audiogen\audiogen.ini"
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error GoTo ErrHandler
'FormDrag Me.hWnd
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error GoTo ErrHandler
ImageRefresh
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub imgArray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error GoTo ErrHandler
'Dim msg As String
If Button = 1 Then
    If imgArray(Index).Picture <> imgArray2(Index).Picture Then imgArray(Index).Picture = imgArray2(Index).Picture
    Select Case LCase(imgArray2(Index).Tag)
    Case "formdrag"
        FormDrag Me.hWnd
    End Select
    'MsgBox imgArray2(Index).Tag
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub imgArray_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error GoTo ErrHandler
ImageRefresh
If Button = 0 Then
    If imgArray(Index).Picture <> imgArray3(Index).Picture Then
        imgArray(Index).Picture = imgArray3(Index).Picture
    End If
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub imgArray_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Local Error GoTo ErrHandler
If Button = 1 Then
    If imgArray(Index).Picture <> imgArray1(Index).Picture Then imgArray(Index).Picture = imgArray1(Index).Picture
    Select Case imgArray(Index).Tag
    Case 5
        MsgBox "burn"
    Case 6
        MsgBox "rip"
    Case 7
        MsgBox "Stop"
    Case 8
        MsgBox "Play"
    Case 9
        MsgBox "back"
    Case 10
        MsgBox "Forward"
    Case 11
        Unload Me
    Case 12
        Me.WindowState = vbMinimized
    End Select
End If
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub
