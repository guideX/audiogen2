VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaveDataWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get FileExtension() As String
Attribute FileExtension.VB_Description = "Gets the extension to use for files written by the implementation."

End Property

Public Property Let Artist(ByVal value As String)
Attribute Artist.VB_Description = "Sets the artist name for the track being ripped."

End Property

Public Property Let Title(ByVal value As String)
Attribute Title.VB_Description = "Sets the title of the track being ripped."

End Property

Public Property Let TrackNumber(ByVal value As String)
Attribute TrackNumber.VB_Description = "Sets the track number of the track being ripped."

End Property

Public Property Let Year(ByVal value As String)
Attribute Year.VB_Description = "Sets the year for the track being ripped."

End Property

Public Property Let Album(ByVal value As String)
Attribute Album.VB_Description = "Sets the album name for the track being ripped."

End Property

Public Property Let Genre(ByVal value As String)
Attribute Genre.VB_Description = "Sets the genre of the track being ripped."

End Property

Public Property Let Comment(ByVal value As String)
Attribute Comment.VB_Description = "Sets the comment for the track being ripped."

End Property

Public Function OpenFile(ByVal sSoundFile As String) As Boolean
Attribute OpenFile.VB_Description = "Called to set the file name before any data is encoded."

End Function

Public Function WriteWavData(ByVal lPtrBuff As Long, ByVal lWriteSize As Long) As Long
Attribute WriteWavData.VB_Description = "Called to write the stereo 16bit 44.1kHz wave data in the buffer to the file."

End Function

Public Sub CloseFile()
Attribute CloseFile.VB_Description = "Called when ripping has been completed and there is no more data to encode."

End Sub
