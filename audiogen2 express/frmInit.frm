VERSION 5.00
Begin VB.Form frmInit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Audiogen 2 Express"
   ClientHeight    =   405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmInit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   27
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   253
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Searching for interfaces..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3555
   End
End
Attribute VB_Name = "frmInit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    Dim cDrvNfo         As New FL_DriveInfo
    Dim blnForceASPI    As Boolean
    Dim strDrives()     As String
    Dim i               As Long

    Me.Show

    ShowMsg "Searching for usable interface...", 500

    blnForceASPI = Command = "-aspi"
    If Not cManager.Init(blnForceASPI) Then
        MsgBox "No interfaces found!" & vbCrLf & _
               "Please install an ASPI driver!" & vbCrLf & _
               "App will exit.", vbExclamation, "Error"
        Unload Me
    End If

    ShowMsg "Used interface: " & Choose(cManager.CurrentInterface, "ASPI", "SPTI"), 500
    ShowMsg "Searching for drives...", 700

    strDrives = cManager.GetCDVDROMs()
    For i = LBound(strDrives) To UBound(strDrives) - 1
        With cDrvNfo
            .GetInfo cManager.DrvChr2DrvID(strDrives(i))
            If (.ReadCapabilities And RC_DVDROM) Then
                ShowMsg strDrives(i) & ": " & .Vendor & " " & .Product & " " & .Revision & " (DVD)", 400
            Else
                ShowMsg strDrives(i) & ": " & .Vendor & " " & .Product & " " & .Revision & " (CD)", 400
            End If
        End With
    Next

    ShowMsg "Ready.", 1000
    frmSelectProject.Show
    Unload Me

End Sub

Private Sub ShowMsg(msg As String, ms As Long)
    lblStatus = msg
    DoEvents
    Sleep ms
End Sub
