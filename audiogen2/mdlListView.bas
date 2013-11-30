Attribute VB_Name = "mdlListView"
Option Explicit

Public Function FillListViewWithTracks(lDiscID As String, lListView As ListView, lINI As String) As Boolean
On Local Error Resume Next
Dim i As Integer, c As Integer, lArtist As String, lAlbum As String
If Len(lDiscID) <> 0 Then
    c = Int(ReadINI(lINI, lDiscID, "TrackCount", 0))
    If i <> 0 Then
        lArtist = ReadINI(lINI, lDiscID, "Artist", "")
        lAlbum = ReadINI(lINI, lDiscID, "Album", "")
        For i = 1 To c
            lListView.ListItems(i).SubItems(1) = lArtist
            lListView.ListItems(i).SubItems(2) = lAlbum
            lListView.ListItems(i).SubItems(3) = ReadINI(lINI, lDiscID, Trim(Str(i)), "")
        Next i
    Else
        FillListViewWithTracks = False
    End If
End If
End Function

Public Sub SetListViewTagWithToc(lListView As ListView, lDrive As String)
On Local Error Resume Next
lListView.Tag = ReturnMediaTOC(lDrive)
End Sub

Public Sub FillListViewWithTrackCount(lListView As ListView, lDrive As String)
On Local Error Resume Next
Dim t As Integer, i As Integer
t = ReturnTrackCount(lDrive)
If t <> 0 Then
    lListView.ListItems.Clear
    For i = 1 To t
        lListView.ListItems.Add , , Trim(Str(i))
        lListView.ListItems(i).SubItems(1) = "Unknown Artist"
        lListView.ListItems(i).SubItems(2) = "Unknown Album"
        lListView.ListItems(i).SubItems(3) = "Untitled"
        lListView.ListItems(i).SubItems(4) = ReturnTrackLength(lDrive, i)
        lListView.ListItems(i).SubItems(5) = "Idle"
    Next i
End If
End Sub

Public Function FindListViewIndexByFileTitle(lFileTitle As String, lListView As ListView) As Integer
On Local Error Resume Next
Dim i As Integer
For i = 1 To lListView.ListItems.Count
    If LCase(GetFileTitle(lListView.ListItems(i).SubItems(2))) = LCase(lFileTitle) Then
        FindListViewIndexByFileTitle = i
        Exit For
    End If
Next i
End Function

Public Function FindListViewIndexByKey(lListView As ListView, lKey As String) As Integer
On Local Error Resume Next
Dim i As Integer
For i = 1 To lListView.ListItems.Count
    If LCase(lListView.ListItems(i).Key) = LCase(lKey) Then
        FindListViewIndexByKey = i
        Exit For
    End If
Next i
End Function

Public Function FindListViewIndex(lListView As ListView, lText As String) As Integer
On Local Error Resume Next
Dim i As Integer
For i = 1 To lListView.ListItems.Count
    If LCase(lListView.ListItems(i).Text) = LCase(lText) Then
        FindListViewIndex = i
        Exit For
    End If
Next i
End Function

Public Function DoesListViewItemExist(lKey As String, lListView As ListView) As Boolean
On Local Error Resume Next
Dim i As Integer
If Len(lKey) <> 0 Then
    For i = 1 To lListView.ListItems.Count
        If LCase(lListView.ListItems(i).Key) = LCase(lKey) Then
            DoesListViewItemExist = True
            Exit For
        End If
    Next i
End If
End Function

Public Function SaveListViewToFile(lListView As ListView, lFilename As String) As Boolean
On Local Error Resume Next
Dim i As Integer, msg As String, b As Boolean
If Len(lFilename) <> 0 Then
    For i = 0 To lListView.ListItems.Count
        If Len(msg) <> 0 Then
            msg = msg & vbCrLf & lListView.ListItems(i).Text & "\\" & lListView.ListItems(i).SubItems(1) & "\\" & lListView.ListItems(i).SubItems(2) & "\\" & lListView.ListItems(i).SubItems(3) & "\\" & lListView.ListItems(i).SubItems(4)
        Else
            msg = lListView.ListItems(i).Text & "\\" & lListView.ListItems(i).SubItems(1) & "\\" & lListView.ListItems(i).SubItems(2) & "\\" & lListView.ListItems(i).SubItems(3) & "\\" & lListView.ListItems(i).SubItems(4)
        End If
    Next i
    If Len(msg) <> 0 Then b = SaveFile(lFilename, msg)
End If
End Function
