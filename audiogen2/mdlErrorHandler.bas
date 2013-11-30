Attribute VB_Name = "mdlErrorHandler"
Option Explicit
Private lIniFile As String

Public Sub ProcessRuntimeError(lName As String, lDescription As String, lNumber As Integer)
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lIniFile) = 0 Then lIniFile = App.Path & "\data\config\err.ini"
If Len(lName) <> 0 Then
    i = Int(ReadINI(lIniFile, "Settings", "Count", 0)) + 1
    WriteINI lIniFile, "Settings", "Count", Trim(Str(i))
    If i <> 0 Then
        WriteINI lIniFile, Trim(Str(i)), "Name", lName
        WriteINI lIniFile, Trim(Str(i)), "Description", lDescription
        WriteINI lIniFile, Trim(Str(i)), "Number", lNumber
    End If
End If
Exit Sub
ErrHandler:
    Err.Clear
End Sub
