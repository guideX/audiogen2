Attribute VB_Name = "mdlRegInfo"
Option Explicit
Private lNickname As String
Private lPassword As String
Private lRegistered As Boolean

Public Sub SetRegNickname(lNick As String)
On Local Error GoTo ErrHandler
If Len(lNick) <> 0 Then
    lNickname = lNick
    WriteINI lIniFiles.iRegInfo, "Settings", "Nickname", lNick
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetRegNickname(lNick As String)", Err.Description, Err.Number
End Sub

Public Sub SetRegPassword(lPass As String)
On Local Error GoTo ErrHandler
If Len(lPass) <> 0 Then
    lPassword = lPass
    WriteINI lIniFiles.iRegInfo, "Settings", "Password", lPass
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub SetRegPassword(lPass As String)", Err.Description, Err.Number
End Sub

Public Function GetRegNickname() As String
On Local Error GoTo ErrHandler
GetRegNickname = lNickname
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetRegNickname() As String", Err.Description, Err.Number
End Function

Public Function GetRegPassword() As String
On Local Error GoTo ErrHandler
GetRegPassword = lPassword
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetRegNickname() As String", Err.Description, Err.Number
End Function

Public Sub LoadRegInfo()
On Local Error GoTo ErrHandler
lNickname = ReadINI(lIniFiles.iRegInfo, "Settings", "Nickname", "")
lPassword = ReadINI(lIniFiles.iRegInfo, "Settings", "Password", "")
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Function GetRegNickname() As String", Err.Description, Err.Number
End Sub

Public Function IsRegistered() As Boolean
On Local Error GoTo ErrHandler
Dim msg As String
msg = KeyGen(lNickname, "pickles", 1)
If LCase(msg) = LCase(lPassword) Then
    lRegistered = True
    IsRegistered = True
Else
    lRegistered = False
    IsRegistered = False
End If
IsRegistered = lRegistered
Exit Function
ErrHandler:
    ProcessRuntimeError "Public Function GetRegNickname() As String", Err.Description, Err.Number
End Function
