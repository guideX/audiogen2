Attribute VB_Name = "mdlTypes"
Option Explicit
Private Type gFileFormats
    fSupportedTypes As String
End Type
Private Type gCD
    cTrackCount As Integer
End Type
Private Type gIniFiles
    iSpectrum As String
    iDiscDB As String
    iAttributes As String
    iSettings As String
    iFileMenu As String
    itvwFilesMenu As String
    itvwFunctionsMenu As String
    iWindowPositions As String
    iPlaylist As String
    iPlaylistTreeView As String
    iCDTracks As String
    iRegInfo As String
    iFavorites As String
End Type
Private Type gDirectories
    dRootUserDir As String
    dMyDocumentsDir As String
    dMyMusicDir As String
    dDesktop As String
    dSharedFolder As String
End Type
Global lIniFiles As gIniFiles, lFileFormats As gFileFormats, lDirectories As gDirectories
Global lCD As gCD
