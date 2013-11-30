Attribute VB_Name = "mdlTreeView"
Option Explicit

Public Function FindTreeViewIndexByFileTitle(lFileTitle As String, lTreeView As TreeView) As Integer
On Local Error Resume Next
Dim i As Integer
For i = 1 To lTreeView.Nodes.Count
    If InStr(lTreeView.Nodes(i).Key, lFileTitle) Then
        FindTreeViewIndexByFileTitle = i
        Exit For
    End If
Next i
End Function

Public Sub SaveTVToFile(lTreeView As TreeView, lFilename As String)
On Local Error Resume Next
Dim n As Integer, i As Integer, nodX As Node
n = ReturnFreeFile()
Open lFilename For Output As #n
    Print #n, "Root:" & Trim(lTreeView.Nodes(1).Key) & "|" & Trim(lTreeView.Nodes(1).Text)
    For i% = 2 To lTreeView.Nodes.Count
        Set nodX = lTreeView.Nodes(i%)
        Print #n, "Sub:" & Trim(nodX.Parent.Key) & "|" & Trim(nodX.Key) & "|" & Trim(nodX.Text)
        If Err.Number <> 0 Then Err.Clear
    Next i%
Close #n
If Err.Number <> 0 Then Err.Clear
End Sub

Public Sub LoadTVFromFile(lTreeView As TreeView, lFilename As String)
On Local Error Resume Next
Dim Dummy As String, RootText As String, RootNode As String, RootKey As String, nodX As Node, SubNode As String, SubRelation As String, TempSubKey As String, SubKey As String, SubText As String, i As Integer
Open lFilename For Input As #1
While Not EOF(1)
Line Input #1, Dummy
If Left$(Dummy, 5) = "Root:" Then
    RootNode$ = Mid$(Dummy, 6, Len(Dummy) - 5)
    RootKey$ = GetBefore(RootNode$)
    RootText$ = GetAfter(RootNode$)
    If DoesTreeViewItemExist(RootKey$, lTreeView) = False Then Set nodX = lTreeView.Nodes.Add(, , RootKey$, RootText$)
End If
If Left$(Dummy, 4) = "Sub:" Then
    SubNode$ = Mid$(Dummy, 5, Len(Dummy) - 4)
    SubRelation$ = GetBefore(SubNode$)
    TempSubKey$ = GetAfter(SubNode$)
    SubKey$ = GetBefore(TempSubKey$)
    SubText$ = GetAfter(TempSubKey$)
    If Len(SubKey$) <> 0 And Len(SubText$) <> 0 Then
        If Len(SubKey$) <> Len(SubText$) Then
            If DoesFileExist(SubKey$) = True Then
                If DoesTreeViewItemExist(SubKey$, lTreeView) = False Then Set nodX = lTreeView.Nodes.Add(SubRelation$, tvwChild, SubKey$, SubText$, 3)
            End If
        Else
            If DoesTreeViewItemExist(SubKey$, lTreeView) = False Then Set nodX = lTreeView.Nodes.Add(SubRelation$, tvwChild, SubKey$, SubText$, 6)
        End If
    End If
End If
Wend
Close #1
If Err.Number <> 0 Then Err.Clear
End Sub

Private Function GetBefore(Sentence As String) As String
On Local Error Resume Next
Dim Counter As Integer, Before As String
Const Sign = "|"
Counter = 1
For Counter = 1 To Len(Sentence)
    If Mid(Sentence, Counter, 1) = Sign Then
        Exit For
    End If
Next Counter
If Counter <> Len(Sentence) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If
GetBefore = Before
End Function

Private Function GetAfter(Sentence As String) As String
On Local Error Resume Next
Dim Counter As Integer, Rest As String
Const Sign = "|"
Counter = 1
For Counter = 1 To Len(Sentence)
    If Mid(Sentence, Counter, 1) = Sign Then
        Exit For
    End If
Next Counter
If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If
GetAfter = Rest
End Function

Public Function DoesTreeViewItemExistByText(lText As String, lTreeView As TreeView) As Boolean
On Local Error Resume Next
Dim i As Integer
If Len(lText) <> 0 Then
    For i = 1 To lTreeView.Nodes.Count
        If Trim(LCase(lTreeView.Nodes(i).Text)) = Trim(LCase(lText)) Then
            DoesTreeViewItemExistByText = True
            Exit For
        End If
    Next i
End If
End Function

Public Function DoesTreeViewItemExist(lKey As String, lTreeView As TreeView) As Boolean
On Local Error Resume Next
Dim i As Integer
If Len(lKey) <> 0 Then
    For i = 1 To lTreeView.Nodes.Count
        If LCase(lTreeView.Nodes(i).Key) = LCase(lKey) Then
            DoesTreeViewItemExist = True
            Exit For
        End If
    Next i
End If
End Function

Public Function DoesTreeViewItemExistInParent(lKey As String, lParent As String, lTreeView As TreeView) As Boolean
On Local Error Resume Next
Dim i As Integer
If Len(lKey) <> 0 Then
    For i = 1 To lTreeView.Nodes.Count
        If LCase(lTreeView.Nodes(i).Key) = LCase(lKey) Then
            If LCase(lTreeView.Nodes(i).Key) = LCase(lParent) Then
                DoesTreeViewItemExistInParent = True
                Exit For
            End If
        End If
    Next i
End If
End Function

