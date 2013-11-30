Attribute VB_Name = "mdlFavorites"
Option Explicit

Public Sub RefreshFavoritesMenu()
On Local Error Resume Next
Dim i As Integer, c As Integer
c = Int(ReadINI(lIniFiles.iFavorites, "Settings", "Count", 0))
Select Case c
Case 0
    For i = 1 To frmMain.mnuFavorite.Count
        frmMain.mnuFavorite(i).Caption = ""
        frmMain.mnuFavorite(i).Enabled = False
        frmMain.mnuFavorite(i).Visible = False
    Next i
    frmMain.mnuFavoritesSep.Visible = False
Case Else
    For i = 1 To c
        Load frmMain.mnuFavorite(i)
        frmMain.mnuFavorite(i).Caption = ReadINI(lIniFiles.iFavorites, Trim(Str(i)), "Data", "")
        If Len(frmMain.mnuFavorite(i).Caption) <> 0 Then
            frmMain.mnuFavorite(i).Visible = True
            frmMain.mnuFavorite(i).Enabled = True
        Else
            frmMain.mnuFavorite(i).Visible = False
            frmMain.mnuFavorite(i).Enabled = False
        End If
    Next i
    frmMain.mnuFavoritesSep.Visible = True
End Select
If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub AddFavorite(lData As String)
On Local Error GoTo ErrHandler
Dim i As Integer
If Len(lData) <> 0 Then
    i = Int(ReadINI(lIniFiles.iFavorites, "Settings", "Count", 0))
    i = i + 1
    WriteINI lIniFiles.iFavorites, "Settings", "Count", Trim(Str(i))
    WriteINI lIniFiles.iFavorites, Trim(Str(i)), "Data", lData
    RefreshFavoritesMenu
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub AddFavorite(lData As String)", Err.Description, Err.Number
End Sub
