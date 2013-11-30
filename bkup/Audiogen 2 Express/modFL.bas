Attribute VB_Name = "modFL"
Option Explicit
Public Type t_AudioTrack
    Album       As String
    Artist      As String
    Title       As String
    no          As Integer
    grab        As Boolean
    startLBA    As Long
    endLBA      As Long
    lenLBA      As Long
End Type
Public Type t_AudioTracks
    Track(98)   As t_AudioTrack
    count       As Integer
End Type
Public cManager     As New FL_Manager
Public strDrvID     As String
