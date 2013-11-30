Attribute VB_Name = "mdlChangeBorderStyle"
Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const GWL_STYLE = (-16)
Private Const SWP_NOMOVE = &H2
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_CHILD = &H40000000
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
Private Const WS_EX_WINDOWEDGE = &H100
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_POPUP = &H80000000
Private Const WS_SYSMENU = &H80000
Private Const WS_TABSTOP = &H10000
Private Const WS_THICKFRAME = &H40000
Private Const WS_TILED = WS_OVERLAPPED
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_ICONIC = WS_MINIMIZE
Private Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME Or WS_MINIMIZEBOX Or WS_MAXIMIZEBOX)
Private Const WS_POPUPWINDOW = (WS_POPUP Or WS_BORDER Or WS_SYSMENU)
Private Const WS_SIZEBOX = WS_THICKFRAME
Private Const WS_TILEDWINDOW = WS_OVERLAPPEDWINDOW
Private Const DS_ABSALIGN = &H1&
Private Const DS_LOCALEDIT = &H20
Private Const DS_MODALFRAME = &H80
Private Const DS_NOIDLEMSG = &H100
Private Const DS_SETFONT = &H40
Private Const DS_SETFOREGROUND = &H200
Private Const DS_SYSMODAL = &H2&
Private mvarBorderStyle As FormBorderStyleConstants

Public Sub ChangeBorderStyle(lForm As Form, vData As FormBorderStyleConstants)
On Local Error GoTo ErrHandler
Dim r As RECT
Const BLess14 = WS_CAPTION
Const BLess3 = &HC80080
Const BLess3_EX = &H101
Const BLessE = WS_THICKFRAME Or WS_CAPTION
If lForm.BorderStyle = 1 Or lForm.BorderStyle = 4 Then
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLess14
ElseIf lForm.BorderStyle = 3 Then
    SetWindowLong lForm.hwnd, GWL_EXSTYLE, GetWindowLong(lForm.hwnd, GWL_EXSTYLE) Xor BLess3_EX
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLess3
ElseIf lForm.BorderStyle <> 0 Then
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLessE
End If
mvarBorderStyle = vData
If vData = 1 Or vData = 4 Then
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLess14
ElseIf vData = 3 Then
    SetWindowLong lForm.hwnd, GWL_EXSTYLE, GetWindowLong(lForm.hwnd, GWL_EXSTYLE) Xor BLess3_EX
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLess3
ElseIf vData <> 0 Then
    SetWindowLong lForm.hwnd, GWL_STYLE, GetWindowLong(lForm.hwnd, GWL_STYLE) Xor BLessE
End If
GetWindowRect lForm.hwnd, r
SetWindowPos lForm.hwnd, 0, r.Left, r.Top, r.Right - r.Left + 1, r.Bottom - r.Top, SWP_NOMOVE
SetWindowPos lForm.hwnd, 0, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, SWP_NOMOVE
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ChangeBorderStyle(lForm As Form, vData As FormBorderStyleConstants)", Err.Description, Err.Number
End Sub
