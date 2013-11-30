Attribute VB_Name = "mdlConvertKHZ"
Option Explicit
Private lConvertKHZOutputFile As String

Public Function ReturnConvertKHZOutputFile() As String
On Local Error GoTo ErrHandler
ReturnConvertKHZOutputFile = lConvertKHZOutputFile
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function

Public Sub SetConvertKHZOutputFile(lFilename As String)
On Local Error GoTo ErrHandler
If Len(lFilename) <> 0 And DoesFileExist(lFilename) = True Then lConvertKHZOutputFile = lFilename
Exit Sub
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Sub

Public Function NewConvertKHZDialog(lFilename As String) As String
On Local Error GoTo ErrHandler
Dim c As New frmConvertKHZ
Exit Function
If Len(lFilename) <> 0 And DoesFileExist(lFilename) = True Then
    Set c = New frmConvertKHZ
    c.Show 1
'    MsgBox ReturnConvertKHZOutputFile()
'    NewConvertKHZDialog = ReturnConvertKHZOutputFile()
End If
Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
End Function
