Attribute VB_Name = "mdlID3"
Option Explicit
Public Type ID3Tag
    Artist As String
    Title As String
    Album As String
    SongYear As String
    Comment As String
    Genre As String
End Type
Private Type ID3v1Tag
    Identifier(2) As Byte
    Title(29) As Byte
    Artist(29) As Byte
    Album(29) As Byte
    SongYear(3) As Byte
    Comment(29) As Byte
    Genre As Byte
End Type
Private Type ID3v2Header
    Identifier(2) As Byte
    Version(1) As Byte
    flags As Byte
    Size(3) As Byte
End Type
Private Type ID3v2ExtendedHeader
    Size(3) As Byte
End Type
Private Type ID3v2FrameHeader
    FrameID(3) As Byte
    Size(3) As Byte
    flags(1) As Byte
End Type
Public Enum MP3SourceEnum
    SOURCE_IDV1
    SOURCE_IDV2
    SOURCE_FILENAME
    SOURCE_USERENTRY
End Enum
Public Type MP3File
    SourceFile As String
    SourceType As MP3SourceEnum
    FileInterpretItems() As String
    FileInterpretItemCnt As Long
    FileInterpretArtist As Long
    FileInterpretTitle As Long
    HasIDv1 As Boolean
    HasIDv2 As Boolean
    IDv1 As ID3Tag
    IDv2 As ID3Tag
    FileTag As ID3Tag
    UserTag As ID3Tag
End Type

Public Function ReadID3v1(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
On Local Error GoTo Failed
Dim FileNo As Integer, Fp As Long, RdTag As ID3v1Tag
FileNo = FreeFile
Open strFile For Binary As #FileNo
    Fp = LOF(FileNo) - 127
    If Fp > 0 Then
        Get #FileNo, Fp, RdTag
        If GetStringValue(RdTag.Identifier, 3, 0) = "TAG" Then
            outTag.Artist = Trim$(GetStringValue(RdTag.Artist, 30, 0))
            outTag.Title = Trim$(GetStringValue(RdTag.Title, 30, 0))
            outTag.Album = Trim$(GetStringValue(RdTag.Album, 30, 0))
            outTag.Comment = Trim$(GetStringValue(RdTag.Comment, 30, 0))
            outTag.SongYear = Trim$(GetStringValue(RdTag.SongYear, 30, 0))
            ReadID3v1 = True
        End If
    End If
Close #FileNo
Exit Function
Failed:
ReadID3v1 = False
End Function

Public Function RenderMp3Tag(lFile As String) As ID3Tag
On Local Error Resume Next
Dim lMP3 As MP3File, lTag As ID3Tag
If Right(LCase(lFile), 4) = ".mp3" Then
    If DoesFileExist(lFile) = True Then
        lMP3.HasIDv2 = ReadID3v2(lFile, lMP3.IDv2)
        If lMP3.HasIDv2 = True Then
            lMP3.IDv2.Artist = CleanInterpreteItems(lMP3.IDv2.Artist)
            lMP3.IDv2.Album = CleanInterpreteItems(lMP3.IDv2.Album)
            lMP3.IDv2.Title = CleanInterpreteItems(lMP3.IDv2.Title)
            lMP3.IDv2.Comment = CleanInterpreteItems(lMP3.IDv2.Comment)
            lMP3.IDv2.SongYear = CleanInterpreteItems(lMP3.IDv2.SongYear)
            lMP3.IDv2.Genre = CleanInterpreteItems(lMP3.IDv2.Genre)
        Else
            lMP3.HasIDv1 = ReadID3v1(lFile, lMP3.IDv1)
            If lMP3.HasIDv1 = True Then
                lMP3.IDv1.Artist = CleanInterpreteItems(lMP3.IDv1.Artist)
                lMP3.IDv1.Album = CleanInterpreteItems(lMP3.IDv1.Album)
                lMP3.IDv1.Title = CleanInterpreteItems(lMP3.IDv1.Title)
                lMP3.IDv1.Comment = CleanInterpreteItems(lMP3.IDv1.Comment)
                lMP3.IDv1.SongYear = CleanInterpreteItems(lMP3.IDv1.SongYear)
                lMP3.IDv1.Genre = CleanInterpreteItems(lMP3.IDv1.Genre)
            End If
        End If
    End If
End If
End Function

Public Function WriteID3v1(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
On Local Error GoTo Failed
Dim FileNo As Integer, Fp As Long, RdTag As ID3v1Tag, WrTag As ID3v1Tag, LocalTag As ID3Tag
If outTag.Artist = "" And outTag.Title = "" And outTag.Album = "" Then Exit Function
FileNo = FreeFile
Open strFile For Binary As #FileNo
    Fp = LOF(FileNo) - 127
    If Fp > 0 Then
        Get #FileNo, Fp, RdTag
        If GetStringValue(RdTag.Identifier, 3, 0) = "TAG" Then
            Fp = LOF(FileNo) - 127
        Else
            Fp = LOF(FileNo) + 1
        End If
    Else
        Fp = LOF(FileNo) + 1
    End If
    LocalTag.Artist = outTag.Artist
    LocalTag.Title = outTag.Title
    LocalTag.Album = outTag.Album
    LocalTag.Comment = outTag.Comment
    LocalTag.Genre = outTag.Genre
    LocalTag.SongYear = outTag.SongYear
    If Len(LocalTag.Artist) > 30 Then LocalTag.Artist = Left$(LocalTag.Artist, 30)
    If Len(LocalTag.Title) > 30 Then LocalTag.Title = Left$(LocalTag.Title, 30)
    If Len(LocalTag.Album) > 30 Then LocalTag.Album = Left$(LocalTag.Album, 30)
    If Len(LocalTag.Comment) > 30 Then LocalTag.Comment = Left$(LocalTag.Comment, 30)
    If Len(LocalTag.SongYear) > 30 Then LocalTag.SongYear = Left$(LocalTag.SongYear, 30)
    SetStringValue WrTag.Identifier, "TAG", 3
    SetStringValue WrTag.Artist, LocalTag.Artist, Len(LocalTag.Artist)
    SetStringValue WrTag.Title, LocalTag.Title, Len(LocalTag.Title)
    SetStringValue WrTag.Album, LocalTag.Album, Len(LocalTag.Album)
    WrTag.Genre = 255
    Put #FileNo, Fp, WrTag
Close #FileNo
WriteID3v1 = True
Exit Function
Failed:
WriteID3v1 = False
End Function

Public Function ReadID3v2(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
On Local Error GoTo Failed
Dim i As Integer, FileNo As Integer, Fp As Long, RdHeader As ID3v2Header, RdExtHeader As ID3v2ExtendedHeader, RdFrameHeader As ID3v2FrameHeader, FrameID As String, FrameSize As Long, TextEncoding As Byte, RdData() As Byte, RdString As String, bGotArtist As Boolean, bGotTitle As Boolean
FileNo = FreeFile
Fp = 1
Open strFile For Binary As #FileNo
    Get #FileNo, Fp, RdHeader
    If GetStringValue(RdHeader.Identifier, 3, 0) = "ID3" Then
        Fp = Loc(FileNo) + 1
        If GetBit(6, RdHeader.flags) Then
            Get #FileNo, , RdExtHeader
            Fp = Fp + GetLongValue(RdExtHeader.Size)
        End If
        Do
            Get #FileNo, Fp, RdFrameHeader
            FrameID = GetStringValue(RdFrameHeader.FrameID, 4, 0)
            FrameSize = GetLongValue(RdFrameHeader.Size)
            If Not FrameSize < 2 Then
                If FrameID = "TPE1" Or FrameID = "TIT2" Or FrameID = "TALB" Then
                    Get #FileNo, , TextEncoding
                    ReDim RdData(FrameSize - 2)
                    Get #FileNo, , RdData
                    RdString = GetStringValue(RdData, UBound(RdData) + 1, TextEncoding)
                    If FrameID = "TPE1" Then
                        outTag.Artist = RdString
                        bGotArtist = True
                    ElseIf FrameID = "TIT2" Then
                        outTag.Title = RdString
                        bGotTitle = True
                    Else
                        outTag.Album = RdString
                    End If
                End If
            End If
            Fp = Fp + 10 + FrameSize
        Loop While Not FrameSize = 0 And Not Fp > 10 + GetLongValue(RdHeader.Size)
        If bGotArtist And bGotTitle Then ReadID3v2 = True
    End If
Close #FileNo
Exit Function
Failed:
ReadID3v2 = False
End Function

Public Function WriteID3v2(ByVal strFile As String, ByRef outTag As ID3Tag) As Boolean
On Local Error GoTo Failed
Dim i As Integer, FileNo As Integer, Fp As Long, AudioData() As Byte, AudioSize As Long, TagSize As Long, Header As ID3v2Header, WrHeader As ID3v2Header
TagSize = Len(outTag.Artist) + Len(outTag.Title) + Len(outTag.Album)
If Not Len(outTag.Artist) = 0 Then TagSize = TagSize + 11
If Not Len(outTag.Title) = 0 Then TagSize = TagSize + 11
If Not Len(outTag.Album) = 0 Then TagSize = TagSize + 11
FileNo = FreeFile
Fp = 1
Open strFile For Binary As #FileNo
    AudioSize = LOF(FileNo)
    Get #FileNo, Fp, Header
    If GetStringValue(Header.Identifier, 3, 0) = "ID3" Then
        AudioSize = AudioSize - GetLongValue(Header.Size)
    End If
    ReDim AudioData(AudioSize - 1)
    Get #FileNo, LOF(FileNo) - AudioSize + 1, AudioData
Close #FileNo
Kill strFile
Open strFile For Binary As #FileNo
    SetStringValue WrHeader.Identifier, "ID3", 3
    WrHeader.Version(0) = 3
    SetLongValue WrHeader.Size, TagSize
    Put #FileNo, , WrHeader
    WriteFrame FileNo, "TPE1", outTag.Artist
    WriteFrame FileNo, "TIT2", outTag.Title
    WriteFrame FileNo, "TALB", outTag.Album
    Put #FileNo, , AudioData
Close #FileNo
WriteID3v2 = True
Exit Function
Failed:
WriteID3v2 = False
End Function

Private Sub WriteFrame(ByVal FileNo As Integer, ByVal strFrameHeader As String, ByVal strFrameData As String)
On Local Error Resume Next
Dim FrameHeader As ID3v2FrameHeader, EncData As Byte, FrameData() As Byte
If Not Len(strFrameData) = 0 Then
    SetStringValue FrameHeader.FrameID, strFrameHeader, 4
    SetLongValue FrameHeader.Size, Len(strFrameData) + 1
    Put #FileNo, , FrameHeader
    ReDim FrameData(Len(strFrameData) - 1)
    SetStringValue FrameData, strFrameData, Len(strFrameData)
    Put #FileNo, , EncData
    Put #FileNo, , FrameData
End If
End Sub

Private Function GetLongValue(ByRef SyncsafeInt() As Byte) As Long
On Local Error Resume Next
Dim i As Integer, j As Integer, BitNr As Integer
For i = 3 To 0 Step -1
    For j = 0 To 6
        If GetBit(j, SyncsafeInt(i)) Then
            GetLongValue = GetLongValue + 2 ^ BitNr
        End If
        BitNr = BitNr + 1
    Next j
Next i
End Function

Private Sub SetLongValue(ByRef SyncsafeInt() As Byte, ByVal Value As Long)
On Local Error Resume Next
Dim i As Integer, ByteNr As Integer, BitNr As Integer
ByteNr = 3
For i = 0 To 27
    If Value And 2 ^ i Then
        SetBit BitNr, SyncsafeInt(ByteNr), True
    End If
    BitNr = BitNr + 1
    If BitNr Mod 7 = 0 Then
        ByteNr = ByteNr - 1
        BitNr = 0
    End If
Next i
End Sub

Public Sub OpenID3File(lFilename As String)
On Local Error GoTo ErrHandler
Dim lID3 As ID3Tag, b As Boolean
With frmMain
If Len(lFilename) <> 0 Then
    If DoesFileExist(lFilename) = True Then
        .cboID3Type.Clear
        .txtID3File.Text = lFilename
        b = ReadID3v2(lFilename, lID3)
        If b = True Then
            .cboID3Type.AddItem "ID3 Version 2"
            .cboID3Type.ListIndex = 0
            .txtArtist.Enabled = True
            .txtAlbum.Enabled = True
            .txtComments.Enabled = True
            .txtTitle.Enabled = True
            .txtArtist.Text = lID3.Artist
            .txtAlbum.Text = lID3.Album
            .txtTitle.Text = lID3.Title
            .txtComments.Text = lID3.Comment
        Else
            b = ReadID3v1(lFilename, lID3)
            If b = True Then
                .cboID3Type.AddItem "ID3 Version 1"
                .cboID3Type.ListIndex = 0
                .txtArtist.Enabled = True
                .txtAlbum.Enabled = True
                .txtComments.Enabled = True
                .txtTitle.Enabled = True
                .txtArtist.Text = lID3.Artist
                .txtAlbum.Text = lID3.Album
                .txtTitle.Text = lID3.Title
                .txtComments.Text = lID3.Comment
            Else
                .cboID3Type.AddItem "No Tag Detected"
                .cboID3Type.ListIndex = 0
                .txtArtist.Enabled = False
                .txtAlbum.Enabled = False
                .txtComments.Enabled = False
                .txtTitle.Enabled = False
                .txtArtist.Text = ""
                .txtAlbum.Text = ""
                .txtComments.Text = ""
                .txtTitle.Text = ""
                .txtComments.Text = ""
            End If
        End If
    End If
End If
End With
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub OpenID3File(lFilename As String)", Err.Description, Err.Number
End Sub

Private Function GetStringValue(ByRef StringData() As Byte, ByVal StringLength As Integer, ByVal EncodingFormat As Byte) As String
On Local Error Resume Next
Dim i As Integer
For i = 0 To StringLength - 1
    If EncodingFormat = 0 Or EncodingFormat = 3 Then
        If StringData(i) = 0 Then Exit Function
        GetStringValue = GetStringValue & Chr$(StringData(i))
    ElseIf EncodingFormat = 1 Then
        If i >= 2 And i Mod 2 = 0 Then
            If StringData(i) = 0 Then Exit Function
            GetStringValue = GetStringValue & Chr$(StringData(i))
        End If
    ElseIf EncodingFormat = 2 Then
        If i Mod 2 = 0 Then
            If StringData(i) = 0 Then Exit Function
            GetStringValue = GetStringValue & Chr$(StringData(i))
        End If
    End If
Next i
End Function

Private Sub SetStringValue(ByRef StringData() As Byte, ByVal Value As String, ByVal StringLength As Integer)
On Local Error Resume Next
Dim i As Integer
For i = 0 To StringLength - 1
    StringData(i) = ASC(Mid$(Value, i + 1, 1))
Next i
End Sub

Private Sub SetBit(ByVal BitNr As Integer, ByRef SrcData As Byte, ByVal BitState As Boolean)
On Local Error Resume Next
Dim Pattern As Byte
If BitState Then
    Pattern = 2 ^ BitNr
    SrcData = SrcData Or Pattern
Else
    Pattern = 255 - 2 ^ BitNr
    SrcData = SrcData And Pattern
End If
End Sub

Private Function GetBit(ByVal BitNr As Byte, ByVal SrcData As Byte) As Boolean
On Local Error Resume Next
Dim Pattern As Byte
Pattern = 2 ^ BitNr
If SrcData And Pattern Then GetBit = True
End Function

Public Function CleanInterpreteItems(ByVal strData As String) As String
On Local Error Resume Next
Dim i As Long, WorkStr As String, StrField() As String
WorkStr = Replace$(strData, "___", " - ", , , vbTextCompare)
StrField = Split(WorkStr, "-", , vbTextCompare)
WorkStr = ""
For i = 0 To UBound(StrField)
    StrField(i) = Trim$(StrField(i))
    StrField(i) = CleanStr(StrField(i), False, False, False)
    If Not StrField(i) = "" Then
        If Not IsNumeric(StrField(i)) Then
            CleanInterpreteItems = CleanInterpreteItems & StrField(i) & "-"
        End If
    End If
Next i
If Not Len(CleanInterpreteItems) = 0 Then CleanInterpreteItems = Left$(CleanInterpreteItems, Len(CleanInterpreteItems) - 1)
End Function

Private Function ReplaceStr(ByVal strData As String) As String
On Local Error Resume Next
strData = Replace$(strData, "_", " ", , , vbTextCompare)
strData = Replace$(strData, "´", "'", , , vbTextCompare)
strData = Replace$(strData, "`", "'", , , vbTextCompare)
strData = Replace$(strData, "{", "(", , , vbTextCompare)
strData = Replace$(strData, "[", "(", , , vbTextCompare)
strData = Replace$(strData, "]", ")", , , vbTextCompare)
strData = Replace$(strData, "}", ")", , , vbTextCompare)
strData = Replace$(strData, "/", "", , , vbTextCompare)
strData = Replace$(strData, "\", "", , , vbTextCompare)
strData = Replace$(strData, ":", "", , , vbTextCompare)
strData = Replace$(strData, "*", "", , , vbTextCompare)
strData = Replace$(strData, "?", "", , , vbTextCompare)
strData = Replace$(strData, """", "", , , vbTextCompare)
strData = Replace$(strData, "<", "", , , vbTextCompare)
strData = Replace$(strData, ">", "", , , vbTextCompare)
strData = Replace$(strData, "|", "", , , vbTextCompare)
ReplaceStr = strData
End Function

Public Function CleanStr(ByVal strData As String, ByVal UpperCase As Boolean, ByVal LowerCase As Boolean, ByVal CutLeadingNumber As Boolean) As String
On Local Error Resume Next
Dim i As Long, SplitField() As String, NewStr As String
strData = ReplaceStr(strData)
CleanStr = Trim$(strData)
If CleanStr = "" Then Exit Function
Do While Not InStr(1, CleanStr, "  ", vbTextCompare) = 0
    CleanStr = Replace$(CleanStr, "  ", " ", , , vbTextCompare)
Loop
SplitField = Split(CleanStr, " ", , vbTextCompare)
CleanStr = ""
For i = 0 To UBound(SplitField)
    If Not i = 0 Or Not CutLeadingNumber Or Not IsNumeric(SplitField(i)) Then
        If UpperCase Then
            NewStr = UCase$(Left$(SplitField(i), 1))
        Else
            NewStr = Left$(SplitField(i), 1)
        End If
        If LowerCase Then
            NewStr = NewStr & LCase$(Right$(SplitField(i), Len(SplitField(i)) - 1))
        Else
            NewStr = NewStr & Right$(SplitField(i), Len(SplitField(i)) - 1)
        End If
        CleanStr = CleanStr & NewStr & " "
    End If
Next i
CleanStr = Trim$(CleanStr)
End Function
