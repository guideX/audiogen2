VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{FDFCF4A3-AD96-11D4-9959-0050BACD4F4C}#1.0#0"; "MDec.ocx"
Object = "{34B82A63-9874-11D4-9E66-0020780170C6}#1.0#0"; "MEnc.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{350143A3-863D-11D5-A844-0080AE000001}#2.0#0"; "WAEnc.ocx"
Object = "{A6FC7BFB-24EE-11D7-BEB4-444553540000}#2.0#0"; "WaDec.ocx"
Object = "{53350C0B-A412-4457-95B4-A6DB803AB756}#1.0#0"; "Audiogen2Movie.ocx"
Object = "{9285FE9E-00F2-457F-A1A2-A0A5D78F3A9F}#3.0#0"; "Audiogen2Radio.ocx"
Object = "{EEB96F74-14D2-11D3-A1BB-B6FC7F000000}#1.0#0"; "Mp3OCX.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "Audiogen 2"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   870
   ClientWidth     =   11715
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   11715
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   9240
      ScaleHeight     =   3975
      ScaleWidth      =   2415
      TabIndex        =   13
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
      Begin WADECLib.WaDec ctlWMADecode 
         Height          =   375
         Left            =   720
         TabIndex        =   86
         Top             =   1440
         Width           =   375
         _Version        =   131072
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin WAENCLib.WAEnc ctlWMAEncode 
         Height          =   375
         Left            =   720
         TabIndex        =   85
         Top             =   960
         Width           =   375
         _Version        =   131072
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin VB.Timer tmrDelayExpandPlaylist 
         Enabled         =   0   'False
         Interval        =   200
         Left            =   600
         Top             =   2040
      End
      Begin VB.Timer tmrDelayAddToPlaylist 
         Interval        =   20
         Left            =   600
         Top             =   2520
      End
      Begin VB.ListBox lstAddToPlaylist 
         Appearance      =   0  'Flat
         Height          =   375
         IntegralHeight  =   0   'False
         Left            =   1080
         TabIndex        =   84
         Top             =   3480
         Width           =   375
      End
      Begin VB.Timer tmrCopyAll 
         Enabled         =   0   'False
         Interval        =   1200
         Left            =   600
         Top             =   3000
      End
      Begin Audiogen2Radio.ctlRadio ctlRadio1 
         Height          =   375
         Left            =   1080
         TabIndex        =   82
         Top             =   3000
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
      End
      Begin VB.Timer tmrProcessPending 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   600
         Top             =   3480
      End
      Begin MDECLib.MDec ctlMP3Decode 
         Height          =   375
         Left            =   720
         TabIndex        =   78
         Top             =   480
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin VB.Timer tmrDelayLoadTV 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   120
         Top             =   2520
      End
      Begin MSComctlLib.TreeView tvwTemp 
         Height          =   735
         Left            =   1560
         TabIndex        =   53
         Top             =   3120
         Visible         =   0   'False
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   1296
         _Version        =   393217
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Timer tmrRewind 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   120
         Top             =   3480
      End
      Begin VB.Timer tmrFastForward 
         Enabled         =   0   'False
         Interval        =   30
         Left            =   120
         Top             =   3000
      End
      Begin MSWinsockLib.Winsock wskFreeDB 
         Left            =   120
         Top             =   2040
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComctlLib.ImageList imlTop 
         Left            =   120
         Top             =   720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   48
         ImageHeight     =   48
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1708A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":18BDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A732
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C286
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1DDDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F92E
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21482
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":21AA9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MENCLib.MEnc ctlMP3Encode 
         Height          =   375
         Left            =   720
         TabIndex        =   23
         Top             =   120
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   0
      End
      Begin VB.Timer tmrProgress 
         Enabled         =   0   'False
         Interval        =   700
         Left            =   1080
         Top             =   2520
      End
      Begin VB.Timer tmrDragBlackLine 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1080
         Top             =   2040
      End
      Begin MSComctlLib.ImageList imlToolbar 
         Left            =   120
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   15
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22329
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2275F
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22B7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":22F8B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":233AF
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":237C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23BC9
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":23FD6
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":243DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":247F1
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24C0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2500A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSScriptControlCtl.ScriptControl ctlVBScript 
         Left            =   120
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
      End
      Begin VB.Image imgCancelBurn2 
         Height          =   450
         Left            =   1800
         Picture         =   "frmMain.frx":25425
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgCancelBurn1 
         Height          =   450
         Left            =   1320
         Picture         =   "frmMain.frx":259D8
         Top             =   1080
         Width           =   450
      End
      Begin VB.Image imgBurn2 
         Height          =   450
         Left            =   1800
         Picture         =   "frmMain.frx":25F9B
         Top             =   360
         Width           =   450
      End
      Begin VB.Image imgBurn1 
         Height          =   450
         Left            =   1320
         Picture         =   "frmMain.frx":26568
         Top             =   360
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar tblTop 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   1429
      ButtonWidth     =   1455
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imlTop"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Files"
            ImageIndex      =   1
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Playlist"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "CD Ripper"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "CD Writter"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Playback"
            ImageIndex      =   6
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Tag Editor"
            ImageIndex      =   7
            Style           =   2
         EndProperty
      EndProperty
      Begin MP3OCXLib.Mp3OCX ctlMP3Player 
         Height          =   675
         Left            =   5880
         TabIndex        =   19
         Top             =   120
         Width           =   4935
         _Version        =   65536
         _ExtentX        =   8705
         _ExtentY        =   1191
         _StockProps     =   161
         BackColor       =   -2147483633
         TopBandsColor   =   12632319
         BottomBandsColor=   128
         LeftChanColor   =   33023
         RightChanColor  =   33023
         PeaksColor      =   16777215
         OscilloType     =   1
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin Audiogen2.XPButton cmdEjectCD 
         Height          =   510
         Left            =   4320
         TabIndex        =   37
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "Eject CD"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":26B35
         PICN            =   "frmMain.frx":26C97
         PICH            =   "frmMain.frx":2723F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdAdd 
         Height          =   510
         Left            =   1440
         TabIndex        =   30
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "Add"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":277DA
         PICN            =   "frmMain.frx":2793C
         PICH            =   "frmMain.frx":27EE9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdBurn 
         Height          =   510
         Left            =   2760
         TabIndex        =   29
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "Burn CD"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":28496
         PICN            =   "frmMain.frx":285F8
         PICH            =   "frmMain.frx":28BD5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvwBurn 
         Height          =   2760
         Left            =   0
         TabIndex        =   16
         Top             =   1200
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4868
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlToolbar"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Percent"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Input File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Output File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Convert From"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Convert To"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Input Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Output Path"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prgBurn 
         Height          =   300
         Left            =   0
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.ComboBox cboBurnDrives 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   840
         Width           =   9015
      End
      Begin Audiogen2.XPButton cmdSelectFilesToBurn 
         Height          =   510
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   900
         BTYPE           =   9
         TX              =   "Select File(s)"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   4210752
         FCOLO           =   4210752
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":291B2
         PICN            =   "frmMain.frx":29314
         PICH            =   "frmMain.frx":298C1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00404040&
         X1              =   4200
         X2              =   4200
         Y1              =   0
         Y2              =   480
      End
      Begin VB.Label lblBurnStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Idle"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   8775
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin MSComctlLib.TreeView tvwPlaylist 
         Height          =   3495
         Left            =   0
         TabIndex        =   14
         Top             =   360
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6165
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imlToolbar"
         Appearance      =   0
         OLEDragMode     =   1
         OLEDropMode     =   1
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Width           =   9015
         Begin Audiogen2.XPButton cmdLoadPlaylist 
            Height          =   285
            Left            =   0
            TabIndex        =   81
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Load Playlist"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":29E6E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdSortByArtist 
            Height          =   285
            Left            =   2565
            TabIndex        =   70
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Sort by Artist"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":29FD0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdSort 
            Height          =   285
            Left            =   1320
            TabIndex        =   69
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Sort"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":2A132
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdSortByAlbum 
            Height          =   285
            Left            =   3825
            TabIndex        =   71
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Sort by Album"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":2A294
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdSortByType 
            Height          =   285
            Left            =   5085
            TabIndex        =   72
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Sort by Type"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":2A3F6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdDefaultPlaylist 
            Height          =   285
            Left            =   6330
            TabIndex        =   73
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Default Playlist"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":2A558
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdSavePlaylist 
            Height          =   285
            Left            =   7590
            TabIndex        =   74
            Top             =   60
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   "Save Playlist"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":2A6BA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   9135
      Begin VB.ComboBox cboPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   0
         TabIndex        =   1
         Text            =   "cboPath"
         Top             =   0
         Width           =   6495
      End
      Begin VB.PictureBox picResizeHorrizontal 
         BorderStyle     =   0  'None
         Height          =   80
         Left            =   60
         MousePointer    =   7  'Size N S
         ScaleHeight     =   75
         ScaleWidth      =   6120
         TabIndex        =   12
         Top             =   1920
         Width           =   6120
      End
      Begin VB.PictureBox picDrag 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   6360
         ScaleHeight     =   1575
         ScaleWidth      =   75
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   80
      End
      Begin VB.PictureBox picWhite 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1575
         Left            =   2400
         ScaleHeight     =   1575
         ScaleWidth      =   855
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox picResizeVerticle 
         BorderStyle     =   0  'None
         Height          =   1575
         Left            =   3240
         MousePointer    =   9  'Size W E
         ScaleHeight     =   105
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   5
         TabIndex        =   9
         Top             =   360
         Width           =   80
      End
      Begin MSComctlLib.TreeView tvwFiles 
         Height          =   1575
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2778
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   53
         LabelEdit       =   1
         Sorted          =   -1  'True
         Style           =   1
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         ImageList       =   "imlToolbar"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin MSComctlLib.TreeView tvwFunctions 
         Height          =   1575
         Left            =   0
         TabIndex        =   4
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   2778
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   53
         LabelEdit       =   1
         Style           =   5
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imlToolbar"
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvwPending 
         Height          =   1815
         Left            =   0
         TabIndex        =   2
         Top             =   2040
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   3201
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         Icons           =   "imlToolbar"
         ForeColor       =   4210752
         BackColor       =   16777215
         Appearance      =   0
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Function"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Format"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ProgressBar prgProgress 
         Height          =   300
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
      End
   End
   Begin VB.Frame fraFunction 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Currently Playing"
      Height          =   3975
      Index           =   7
      Left            =   0
      TabIndex        =   87
      Top             =   810
      Width           =   9135
      Begin VB.Image imgFade 
         Height          =   300
         Index           =   0
         Left            =   0
         Picture         =   "frmMain.frx":2A81C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   6000
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   6
      Left            =   0
      TabIndex        =   28
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin Audiogen2.XPButton cmdWipe 
         Height          =   375
         Left            =   2160
         TabIndex        =   52
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Wipe"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":3061E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdSave 
         Height          =   375
         Left            =   1080
         TabIndex        =   51
         Top             =   2520
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BTYPE           =   9
         TX              =   "Save"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":30780
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtComments 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   50
         Top             =   2160
         Width           =   6855
      End
      Begin VB.TextBox txtTitle 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   48
         Top             =   1800
         Width           =   6855
      End
      Begin VB.TextBox txtAlbum 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   45
         Top             =   1440
         Width           =   6855
      End
      Begin VB.TextBox txtArtist 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1080
         TabIndex        =   43
         Top             =   1080
         Width           =   6855
      End
      Begin VB.ComboBox cboID3Type 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMain.frx":308E2
         Left            =   1080
         List            =   "frmMain.frx":308EC
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   6855
      End
      Begin Audiogen2.XPButton cmdID3Select 
         Height          =   285
         Left            =   8040
         TabIndex        =   40
         Top             =   120
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         BTYPE           =   9
         TX              =   "&Select"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":3090E
         PICN            =   "frmMain.frx":3092A
         PICH            =   "frmMain.frx":30D3F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox txtID3File 
         Height          =   315
         Left            =   1080
         TabIndex        =   39
         Top             =   120
         Width           =   6855
      End
      Begin VB.Label lblComments 
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   49
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Title:"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblAlbum 
         BackStyle       =   0  'Transparent
         Caption         =   "Album:"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblArtist 
         BackStyle       =   0  'Transparent
         Caption         =   "Artist:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Type:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "File:"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   4
      Left            =   0
      TabIndex        =   6
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin Audiogen2.XPButton cmdSearch 
         Default         =   -1  'True
         Height          =   315
         Left            =   5520
         TabIndex        =   21
         Top             =   0
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   9
         TX              =   "Search"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   0
         MICON           =   "frmMain.frx":31156
         PICN            =   "frmMain.frx":31172
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboSearchType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtSearch 
         Height          =   285
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   3015
      End
      Begin MSComctlLib.ListView lvwSearch 
         Height          =   3495
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   6165
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Title"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Artist"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Album"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "File"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraFunction 
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   2
      Left            =   0
      MouseIcon       =   "frmMain.frx":313AE
      TabIndex        =   31
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin VB.DriveListBox drvRip 
         Height          =   315
         Left            =   3840
         TabIndex        =   83
         Top             =   720
         Width           =   5175
      End
      Begin Audiogen2.XPButton cmdRefreshCD 
         Height          =   300
         Left            =   4800
         TabIndex        =   79
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":31500
         PICN            =   "frmMain.frx":31662
         PICH            =   "frmMain.frx":31B58
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdStopCD 
         Height          =   300
         Left            =   3600
         TabIndex        =   77
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Stop"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":3204E
         PICN            =   "frmMain.frx":321B0
         PICH            =   "frmMain.frx":32219
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboDrives 
         Height          =   315
         ItemData        =   "frmMain.frx":32282
         Left            =   0
         List            =   "frmMain.frx":32284
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   0
         Width           =   9015
      End
      Begin VB.DirListBox dirCopyTo 
         BackColor       =   &H00FFFFFF&
         Height          =   2565
         Left            =   3840
         TabIndex        =   33
         Top             =   1080
         Width           =   5175
      End
      Begin Audiogen2.XPButton cmdCopy 
         Height          =   300
         Left            =   0
         TabIndex        =   32
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Copy"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":32286
         PICN            =   "frmMain.frx":323E8
         PICH            =   "frmMain.frx":327F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ProgressBar prgRipProgress 
         Height          =   315
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lvwCD 
         Height          =   3015
         Left            =   0
         TabIndex        =   36
         Top             =   720
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   5318
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Track"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Artist"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Album"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Title"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Length"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin Audiogen2.XPButton cmdPlayCD 
         Height          =   300
         Left            =   1200
         TabIndex        =   75
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Play"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":32BFA
         PICN            =   "frmMain.frx":32D5C
         PICH            =   "frmMain.frx":32DD6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Audiogen2.XPButton cmdPauseCD 
         Height          =   300
         Left            =   2400
         TabIndex        =   76
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         BTYPE           =   9
         TX              =   "Pause"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   16711680
         MCOL            =   12632256
         MPTR            =   99
         MICON           =   "frmMain.frx":32E50
         PICN            =   "frmMain.frx":32FB2
         PICH            =   "frmMain.frx":33030
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Frame fraFunction 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   3975
      Index           =   5
      Left            =   0
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   9135
      Begin Audiogen2MoviePlayer.ctlMovie ctlMovie1 
         Height          =   1935
         Left            =   0
         TabIndex        =   80
         Top             =   345
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3413
      End
      Begin VB.Frame fraPlaybackControls 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   350
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   9735
         Begin Audiogen2.XPButton cmdMute 
            Height          =   285
            Left            =   480
            TabIndex        =   55
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":330AE
            PICN            =   "frmMain.frx":33210
            PICH            =   "frmMain.frx":33469
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdForeward 
            Height          =   285
            Left            =   6960
            TabIndex        =   56
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":336B9
            PICN            =   "frmMain.frx":3381B
            PICH            =   "frmMain.frx":3389E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdPlay 
            Height          =   285
            Left            =   6000
            TabIndex        =   57
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":33921
            PICN            =   "frmMain.frx":33A83
            PICH            =   "frmMain.frx":33AFD
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdOpen 
            Height          =   285
            Left            =   4560
            TabIndex        =   58
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":33B77
            PICN            =   "frmMain.frx":33CD9
            PICH            =   "frmMain.frx":33D55
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdBackward 
            Height          =   285
            Left            =   5040
            TabIndex        =   59
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":33DD1
            PICN            =   "frmMain.frx":33F33
            PICH            =   "frmMain.frx":33FB4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdFullScreen 
            Height          =   285
            Left            =   0
            TabIndex        =   60
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":34035
            PICN            =   "frmMain.frx":34197
            PICH            =   "frmMain.frx":3421B
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin Audiogen2.XPButton cmdPausePlayback 
            Height          =   285
            Left            =   5520
            TabIndex        =   61
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":3429F
            PICN            =   "frmMain.frx":34401
            PICH            =   "frmMain.frx":3447F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSComctlLib.Slider sldProgress 
            Height          =   255
            Left            =   1800
            TabIndex        =   62
            ToolTipText     =   "Position"
            Top             =   75
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            TickStyle       =   3
            TickFrequency   =   0
         End
         Begin MSComctlLib.Slider sldVolume 
            Height          =   255
            Left            =   3480
            TabIndex        =   63
            ToolTipText     =   "Position"
            Top             =   75
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   450
            _Version        =   393216
            Max             =   100
            SelStart        =   100
            TickStyle       =   3
            TickFrequency   =   0
            Value           =   100
         End
         Begin Audiogen2.XPButton cmdStop 
            Height          =   285
            Left            =   6480
            TabIndex        =   64
            Top             =   30
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   503
            BTYPE           =   9
            TX              =   ""
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   16711680
            MCOL            =   12632256
            MPTR            =   99
            MICON           =   "frmMain.frx":344FD
            PICN            =   "frmMain.frx":3465F
            PICH            =   "frmMain.frx":346C8
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblFilename 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   7560
            TabIndex        =   67
            Top             =   75
            Width           =   7335
         End
         Begin VB.Label Label1 
            Caption         =   "Progress:"
            Height          =   255
            Left            =   1080
            TabIndex        =   66
            Top             =   75
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Volume:"
            Height          =   255
            Left            =   2880
            TabIndex        =   65
            Top             =   75
            Width           =   795
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuCDRipper 
         Caption         =   "Single Track Ripper"
      End
      Begin VB.Menu mnuProcessQUe 
         Caption         =   "Process Que"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuRadio 
      Caption         =   "&Radio"
      Begin VB.Menu mnuPlayStation 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuStopStation 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuSep34789263789 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddStation 
         Caption         =   "Add"
      End
      Begin VB.Menu mnuRemoveStation 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuEditInternetRadio 
         Caption         =   "Edit"
      End
      Begin VB.Menu mnuAddStationToFavorites 
         Caption         =   "Add Favorite"
      End
   End
   Begin VB.Menu mnuFavorates 
      Caption         =   "F&avorites"
      Begin VB.Menu mnuAddFavorite 
         Caption         =   "Add New Favorite"
      End
      Begin VB.Menu mnuEditFavorites 
         Caption         =   "Edit My Favorites .."
      End
      Begin VB.Menu mnuFavoritesSep 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   "<Blank Favorite>"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuPB 
      Caption         =   "&Player"
      Begin VB.Menu mnuOpenMovie 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSep37896923 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayMovie 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuPauseMovieM 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuResumeMovie 
         Caption         =   "Resume"
      End
      Begin VB.Menu mnuStopMovieM 
         Caption         =   "Stop"
      End
      Begin VB.Menu mnuSep30278903897 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFullScreenM 
         Caption         =   "Full Screen"
      End
      Begin VB.Menu mnuChangePlayRate 
         Caption         =   "Change Play Rate"
      End
      Begin VB.Menu mnuChangeMoviePosition 
         Caption         =   "Change Movie Position"
      End
      Begin VB.Menu mnuRewindM 
         Caption         =   "Rewind"
      End
      Begin VB.Menu mnuForward 
         Caption         =   "Forward"
      End
      Begin VB.Menu mnuOpenCDDoor 
         Caption         =   "Open CD Door"
      End
      Begin VB.Menu mnuCloseDoor 
         Caption         =   "Close Door"
      End
      Begin VB.Menu mnuSep3892836296 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMuteM 
         Caption         =   "Mute"
      End
      Begin VB.Menu mnuUnMuteM 
         Caption         =   "Un Mute"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuViewErrors 
         Caption         =   "View Error Log"
      End
      Begin VB.Menu mnuSep327896293 
         Caption         =   "-"
      End
      Begin VB.Menu mnuContactLeonAiossa 
         Caption         =   "Contact Leon Aiossa"
      End
      Begin VB.Menu mnuEMailColinFoss 
         Caption         =   "Contact Colin Foss"
      End
      Begin VB.Menu mnuHomeWeb 
         Caption         =   "Team Nexgen Homepage"
      End
      Begin VB.Menu mnuForumWeb 
         Caption         =   "Forum"
      End
      Begin VB.Menu mnuSep38926392693 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutAudiogen2 
         Caption         =   "About Audiogen 2"
      End
   End
   Begin VB.Menu mnuFunctionsPopup 
      Caption         =   "<tvwFunctions>"
      Visible         =   0   'False
      Begin VB.Menu mnuToggleExpanded 
         Caption         =   "Expand"
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProporties2 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnutvwFile 
      Caption         =   "<tvwFile Menu>"
      Visible         =   0   'False
      Begin VB.Menu mnuShowContainingFolder 
         Caption         =   "Show Containing Folder"
      End
      Begin VB.Menu mnuSep3786926537 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddToBatch 
         Caption         =   "Add to Batch"
      End
      Begin VB.Menu mnuAddAlltoBatch 
         Caption         =   "Add all to Batch"
      End
      Begin VB.Menu mnuAddPlayToBatch 
         Caption         =   "Add Play to Batch"
      End
      Begin VB.Menu mnuAddBurnToBatch 
         Caption         =   "Add Burn to Batch"
      End
      Begin VB.Menu mnuAddToBurnQue2 
         Caption         =   "Add to Burn Que"
      End
      Begin VB.Menu mnuAddFileToFavorites 
         Caption         =   "Add to Favorites"
      End
      Begin VB.Menu mnuSaveFilesAsPlaylist 
         Caption         =   "Save as Playlist"
      End
      Begin VB.Menu mnuSep389279037890 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayFile 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuDecodeFile1 
         Caption         =   "Convert"
      End
      Begin VB.Menu mnuDecodeFile 
         Caption         =   "Decode"
      End
      Begin VB.Menu mnuSep3892736 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEncodeToMp3 
         Caption         =   "Encode to MP3"
      End
      Begin VB.Menu mnuEncodeToWMA 
         Caption         =   "Encode to WMA"
      End
      Begin VB.Menu mnuSpe38979263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearFiles 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu mnuSep3789263 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProporties 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnutvwPlaylistMenu 
      Caption         =   "<tvwPlaylist Menu>"
      Visible         =   0   'False
      Begin VB.Menu mnuArrange 
         Caption         =   "Sort By"
         Begin VB.Menu mnuArrangeByArtistAndAlbum 
            Caption         =   "Artist and Album"
         End
         Begin VB.Menu mnuArrangeByArtist 
            Caption         =   "Artist"
         End
         Begin VB.Menu mnuArrangeByAlbum 
            Caption         =   "Album"
         End
         Begin VB.Menu mnuArrangeByFormat 
            Caption         =   "Type"
         End
         Begin VB.Menu mnuSortByComments 
            Caption         =   "Comments"
         End
         Begin VB.Menu mnuSortByYear 
            Caption         =   "Year"
         End
         Begin VB.Menu mnuSep3789263954972634 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSortParentNode 
            Caption         =   "Alphabetize"
         End
         Begin VB.Menu mnuAlphabetic 
            Caption         =   "Alphabetize All"
         End
      End
      Begin VB.Menu mnuSep378926397632 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayThisFile 
         Caption         =   "Play"
      End
      Begin VB.Menu mnuDecode 
         Caption         =   "Convert"
      End
      Begin VB.Menu mnuSepw3897289356 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFile2 
         Caption         =   "Add File"
      End
      Begin VB.Menu mnuAddToBurnQue 
         Caption         =   "Add to Burn Que"
      End
      Begin VB.Menu mnuAddPlaylistItemToFavorites 
         Caption         =   "Add to Favorites"
      End
      Begin VB.Menu mnuRemoveFile 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuClearPlaylist 
         Caption         =   "Clear Playlist"
      End
      Begin VB.Menu mnuSep397290369264 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditTag 
         Caption         =   "Edit Tag"
      End
      Begin VB.Menu mnuChaneTitle 
         Caption         =   "Change Title"
      End
      Begin VB.Menu mnuChangeArtist 
         Caption         =   "Change Artist"
      End
      Begin VB.Menu mnuChangeAlbum 
         Caption         =   "Change Album"
      End
      Begin VB.Menu mnuSep38902790837 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProporties3 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu mnuWaveFileMaint 
      Caption         =   "<Wave File Maint>"
      Visible         =   0   'False
      Begin VB.Menu mnuAddToBurnQue12345 
         Caption         =   "Add to Burn Que"
      End
      Begin VB.Menu mnuConvertWave 
         Caption         =   "Convert"
      End
   End
   Begin VB.Menu mnutvwPlaylistMenu2 
      Caption         =   "<Album/Artist Maint>"
      Visible         =   0   'False
      Begin VB.Menu mnuExpand 
         Caption         =   "Expand"
      End
      Begin VB.Menu mnuContract 
         Caption         =   "Contract"
      End
      Begin VB.Menu mnuSep837296392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFileToArtist 
         Caption         =   "Add File"
      End
      Begin VB.Menu mnuRemoveArtist 
         Caption         =   "Remove"
      End
   End
   Begin VB.Menu mnuPending 
      Caption         =   "<Pending>"
      Visible         =   0   'False
      Begin VB.Menu mnuProcessNow 
         Caption         =   "Process Que Now"
      End
      Begin VB.Menu mnuProcessThisItemOnly 
         Caption         =   "Process this item only"
      End
      Begin VB.Menu mnuSep378296392 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddToBurnQue1 
         Caption         =   "Add Entry to Burn Que"
      End
      Begin VB.Menu mnuAddAllToBurnQue1 
         Caption         =   "Add All Entries to Burn Que"
      End
      Begin VB.Menu mnuChangeFunction 
         Caption         =   "Change Function"
         Begin VB.Menu mnuToPlay 
            Caption         =   "to Play"
         End
         Begin VB.Menu mnuWaveToCDA123 
            Caption         =   "to Wave to CDA"
         End
         Begin VB.Menu mnuWaveToMp3123 
            Caption         =   "to Wave to MP3"
         End
         Begin VB.Menu mnuWavetoWMA123 
            Caption         =   "to Wave to WMA"
         End
         Begin VB.Menu mnuMP3ToWave123 
            Caption         =   "to MP3 to Wave"
         End
         Begin VB.Menu mnuWMAToWave123 
            Caption         =   "to WMA to Wave"
         End
      End
      Begin VB.Menu mnuSep387238962 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnulvwProcesses 
      Caption         =   "<lvwProcesses>"
      Visible         =   0   'False
      Begin VB.Menu mnuRemoveEntry 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuClearProcesses 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Private lLeftCorrection As Integer
Private lResizeVert As Boolean
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Enum EApplicationMode
   eIdle
   eBurning
End Enum
Private lLabelEdit As String
Private lLoadPlaylistEnabled As Boolean
Private lMP3Frames As Long
Private lInt As Integer
Private m_cRip As cCDRip
Private m_cToc As cToc
Private m_bRipping As Boolean
Private m_bCancel As Boolean
Private lRipIndex As Integer
Private WithEvents m_cDiscMaster As cDiscMaster
Attribute m_cDiscMaster.VB_VarHelpID = -1
Private m_cRedbook As cRedbookDiscMaster
Private m_eMode As EApplicationMode
Private m_CurrentTrack As Integer
Private m_var As Integer
Private lFullscreen As Boolean
Private lMute As Boolean
Private DraggedKnot As Node
Private SelectedKnot As Node
Private lBusy As Boolean
Private lRipping As Boolean
Private lDiscID As String
Private lBurnTrack As Integer
Private lSpectrumToggle As Integer
Private lFrameCount As Long
Private lProgressClicked As Boolean

Public Function ReturnCDDiscID() As String
On Local Error Resume Next
ReturnCDDiscID = lDiscID
End Function

Public Sub SetBusy(lValue As Boolean)
On Local Error Resume Next
lBusy = lValue
End Sub

Private Sub createCD()
On Local Error Resume Next
Dim lFreeBlocks As Long, lTotalBlocks As Long, lBlockSize As Long, iFile As Long, iRecorder As Long, cRecorder As cDiscRecorder, bCloseExclusive As Boolean, cInfo As cMediaInfo, cProps As cDiscRecorderProperties, cProp As cProperty, iProp As Long
Set frmMain.cmdBurn.PictureNormal = imgCancelBurn1.Picture
Set frmMain.cmdBurn.PictureOver = imgCancelBurn2.Picture
setApplicationMode eBurning
prgBurn.Visible = True
cboBurnDrives.Visible = False
showStatus "Opening Recorder for burning.."
Set m_cRedbook = m_cDiscMaster.RedbookDiscMaster
iRecorder = cboBurnDrives.ItemData(cboBurnDrives.ListIndex)
m_cDiscMaster.Recorders(iRecorder).SetAsActive
Set cRecorder = m_cDiscMaster.Recorders.ActiveRecorder
showStatus "Checking media type.."
cRecorder.OpenExclusive
bCloseExclusive = True
Set cInfo = cRecorder.MediaInfo
If (cInfo.MediaPresent) And ((cInfo.MediaFlags And MEDIA_WRITABLE) = MEDIA_WRITABLE) Then
    cRecorder.CloseExclusive
    bCloseExclusive = False
    If Not (m_bCancel) Then
        If (cInfo Is Nothing) Then
            frmInsertCD.Show 1
            Set cmdBurn.PictureNormal = imgBurn1.Picture
            Set cmdBurn.PictureOver = imgBurn2.Picture
        Else
            lFreeBlocks = cInfo.FreeBlocks
            lBlockSize = m_cRedbook.AudioBlockSize
            For iFile = 1 To lvwBurn.ListItems.Count
                If LCase(lvwBurn.ListItems(iFile).Text) = "ready to burn" Then
                    lTotalBlocks = lTotalBlocks + FileLen(lvwBurn.ListItems(iFile).SubItems(6) & lvwBurn.ListItems(iFile).SubItems(2)) \ lBlockSize
                End If
            Next iFile
            prgBurn.Value = 0
            If lTotalBlocks <> 0 Then prgBurn.Max = lTotalBlocks
            If (lTotalBlocks <= lFreeBlocks) Then
                showStatus "Creating CD Image"
                createCDImage
                If Not (m_bCancel) Then
                    Dim lErr As Long
                    m_cDiscMaster.RecordDisc ReturnTestMode(), ReturnAutoEject()
                    lErr = Err.Number
                    On Error GoTo 0
                    If lErr <> 0 Then
                        MsgBox "Error: " & DecodeIMAPIError(lErr), vbExclamation
                    End If
                End If
            Else
                frmInsertCD.Show 1
                Set cmdBurn.PictureNormal = imgBurn1.Picture
                Set cmdBurn.PictureOver = imgBurn2.Picture
            End If
        End If
    End If
Else
    frmInsertCD.Show 1
    Set cmdBurn.PictureNormal = imgBurn1.Picture
    Set cmdBurn.PictureOver = imgBurn2.Picture
End If
If (m_bCancel) Then
    On Error Resume Next
    m_cDiscMaster.ClearFormatContent
    On Error GoTo 0
End If
If (bCloseExclusive) Then
    cRecorder.CloseExclusive
End If
setApplicationMode eIdle
prgBurn.Visible = False
cboBurnDrives.Visible = True
If Err.Number <> 0 Then
    ProcessRuntimeError "Private Sub createCD()", Err.Description, Err.Number
End If
End Sub

Public Function DoesBlankCDMediaExist() As Boolean
On Local Error Resume Next
Dim lFreeBlocks As Long, lTotalBlocks As Long, lBlockSize As Long, iFile As Long, iRecorder As Long, cRecorder As cDiscRecorder, bCloseExclusive As Boolean, cInfo As cMediaInfo, cProps As cDiscRecorderProperties, cProp As cProperty, iProp As Long
Set m_cRedbook = m_cDiscMaster.RedbookDiscMaster
iRecorder = cboBurnDrives.ItemData(cboBurnDrives.ListIndex)
m_cDiscMaster.Recorders(iRecorder).SetAsActive
Set cRecorder = m_cDiscMaster.Recorders.ActiveRecorder
cRecorder.OpenExclusive
bCloseExclusive = True
Set cInfo = cRecorder.MediaInfo
If (cInfo.MediaPresent) And ((cInfo.MediaFlags And MEDIA_WRITABLE) = MEDIA_WRITABLE) Then
    cRecorder.CloseExclusive
    bCloseExclusive = False
    If Not (m_bCancel) Then
        If (cInfo Is Nothing) Then
            DoesBlankCDMediaExist = False
        Else
            DoesBlankCDMediaExist = True
        End If
    End If
Else
    DoesBlankCDMediaExist = False
End If
If (m_bCancel) Then
    On Error Resume Next
    m_cDiscMaster.ClearFormatContent
    On Error GoTo 0
End If
If (bCloseExclusive) Then cRecorder.CloseExclusive
If Err.Number <> 0 Then ProcessRuntimeError "Public Function DoesBlankCDMediaExist() As Boolean", Err.Description, Err.Number
End Function

Private Sub createCDImage()
On Local Error GoTo ErrHandler
Dim iFile As Long, sFile As String
For iFile = 1 To lvwBurn.ListItems.Count
    If LCase(lvwBurn.ListItems(iFile).Text) = "ready to burn" Then
        sFile = lvwBurn.ListItems(iFile).SubItems(6) & lvwBurn.ListItems(iFile).SubItems(2)
        showStatus "Adding track " & sFile & "..."
        addTrack sFile
        If m_bCancel Then Exit For
    End If
Next iFile
Exit Sub
ErrHandler:
    If Err.Number <> 0 Then ProcessRuntimeError "Private Sub createCDImage()", Err.Description, Err.Number
End Sub

Private Function addTrack(ByVal sFile As String) As Long
On Local Error GoTo ErrHandler
Dim lBlockSize As Long, cWav As cWavReader, lTrackSize As Long, lTrackBlocks As Long, bMore As Boolean, lReadSize As Long, lWrittenSize As Long, lWrittenBlocks As Long
lBlockSize = m_cRedbook.AudioBlockSize
Set cWav = New cWavReader
cWav.ReadBufferSize = lBlockSize \ 4
cWav.OpenFile sFile
lTrackSize = cWav.AudioLength * 4
lTrackBlocks = lTrackSize \ lBlockSize
If (lTrackSize Mod lBlockSize) > 0 Then
    lTrackBlocks = lTrackBlocks + 1
End If
m_cRedbook.CreateAudioTrack lTrackBlocks
Do
    bMore = cWav.Read
    If (bMore) Then
        lReadSize = cWav.ReadSize * 4
        cWav.ZeroUnusedBufferBytes
        m_cRedbook.AddAudioTrackBlocks cWav.ReadBufferPtr, lBlockSize
        lWrittenSize = lWrittenSize + lReadSize
        lWrittenBlocks = lWrittenBlocks + 1
        prgBurn.Value = prgBurn.Value + 1
        DoEvents
    End If
Loop While (bMore) And Not (m_bCancel)
If Not (m_bCancel) Then
    m_cRedbook.CloseAudioTrack
End If
cWav.CloseFile
addTrack = lWrittenBlocks
Exit Function
ErrHandler:
    If Err.Number <> 0 Then ProcessRuntimeError "Private Function addTrack(ByVal sFile As String) As Long", Err.Description, Err.Number
End Function

Private Sub setApplicationMode(ByVal eMode As EApplicationMode)
On Local Error GoTo ErrHandler
m_eMode = eMode
If (eMode = eBurning) Then
    m_bCancel = False
    cmdBurn.Caption = "Cancel"
Else
    cmdBurn.Caption = "&Burn"
    showStatus "Ready"
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub setApplicationMode(ByVal eMode As EApplicationMode)", Err.Description, Err.Number
End Sub

Private Sub enableControl(ctl As Control, ByVal bState As Boolean)
On Local Error GoTo ErrHandler
Dim oBackColor As OLE_COLOR
oBackColor = IIf(bState, vbWindowBackground, vbButtonFace)
ctl.Enabled = bState
If TypeOf ctl Is ListBox Then
   ctl.BackColor = oBackColor
ElseIf TypeOf ctl Is ComboBox Then
   ctl.BackColor = oBackColor
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub enableControl(ctl As Control, ByVal bState As Boolean)", Err.Description, Err.Number
End Sub

Private Sub showRecorders()
On Local Error GoTo ErrHandler
Dim l As Long
Set m_cDiscMaster = New cDiscMaster
m_cDiscMaster.Initialise
With m_cDiscMaster.Recorders
    For l = 1 To .Count
        With .Recorder(l)
            If (.SupportsRedbook) Then
                cboBurnDrives.AddItem .VendorId & " " & .ProductId & " " & .RevisionId
                cboBurnDrives.ItemData(cboBurnDrives.NewIndex) = l
            End If
        End With
    Next l
End With
If (cboBurnDrives.ListCount > 0) Then
    enableControl cboBurnDrives, True
    cboBurnDrives.ListIndex = 0
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub showRecorders()", Err.Description, Err.Number
End Sub

Private Sub showStatus(ByVal sStatus As String)
On Local Error GoTo ErrHandler
lblBurnStatus.Caption = sStatus
lblBurnStatus.Refresh
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub showStatus(ByVal sStatus As String)", Err.Description, Err.Number
End Sub

Public Function RipTrack(ByVal lTrack As Long, lFile As String) As Boolean
On Local Error GoTo ErrHandler
Dim cWriter As cWAVWriter, cTrack As cTocEntry, cTrackRip As New cCDTrackRipper
cboDrives.Visible = False
cboPath.Visible = False
prgProgress.Visible = True
prgRipProgress.Visible = True
Set cWriter = New cWAVWriter
If (cWriter.OpenFile(lFile)) Then
    Set cTrack = m_cToc.Entry(lTrack)
    cTrackRip.CreateForTrack cTrack
    lvwCD.Refresh
    If (cTrackRip.OpenRipper()) Then
        Do While cTrackRip.Read
            cWriter.WriteWavData cTrackRip.ReadBufferPtr, cTrackRip.ReadBufferSize
            prgRipProgress.Value = cTrackRip.PercentComplete
            prgProgress.Value = cTrackRip.PercentComplete
            If lvwCD.ListItems(lRipIndex).SubItems(5) <> "Copy " & Trim(Str(cTrackRip.PercentComplete)) & "%" Then
                LockWindowUpdate lvwCD.hwnd
                lvwCD.ListItems(lRipIndex).SubItems(5) = "Copy " & Trim(Str(cTrackRip.PercentComplete)) & "%"
                If lvwCD.ListItems(lTrack).ForeColor <> vbRed Then lvwCD.ListItems(lRipIndex).ForeColor = vbRed
                If lvwCD.ListItems(lTrack).ListSubItems(1).ForeColor <> vbRed Then lvwCD.ListItems(lTrack).ListSubItems(1).ForeColor = vbRed
                If lvwCD.ListItems(lTrack).ListSubItems(2).ForeColor <> vbRed Then lvwCD.ListItems(lTrack).ListSubItems(2).ForeColor = vbRed
                If lvwCD.ListItems(lTrack).ListSubItems(3).ForeColor <> vbRed Then lvwCD.ListItems(lTrack).ListSubItems(3).ForeColor = vbRed
                If lvwCD.ListItems(lTrack).ListSubItems(4).ForeColor <> vbRed Then lvwCD.ListItems(lTrack).ListSubItems(4).ForeColor = vbRed
                If lvwCD.ListItems(lTrack).ListSubItems(5).ForeColor <> vbRed Then
                    lvwCD.ListItems(lTrack).ListSubItems(5).ForeColor = vbRed
                    lvwCD.Refresh
                End If
                
                LockWindowUpdate 0
            End If
            DoEvents
            If (m_bCancel) Then
                Exit Do
            End If
        Loop
        If lvwCD.ListItems(lRipIndex).ForeColor <> vbBlack Then lvwCD.ListItems(lRipIndex).ForeColor = vbBlack
        tmrCopyAll.Enabled = True
        cTrackRip.CloseRipper
        cWriter.CloseFile
        If (m_bCancel) Then
            Kill lFile
            m_bRipping = False
        Else
            lvwCD.ListItems(lRipIndex).SubItems(5) = "Copied"
            lvwCD.ListItems(lRipIndex).Tag = Left(lFile, Len(lFile) - 4) & ".mp3"
        End If
    End If
End If
m_bRipping = False
m_bCancel = False
lRipIndex = 0
RipTrack = True
cboDrives.Visible = True
cboPath.Visible = True
prgProgress.Visible = False
prgRipProgress.Visible = False
Exit Function
ErrHandler:
    cboDrives.Visible = True
    cboPath.Visible = True
    prgProgress.Visible = False
    prgRipProgress.Visible = False
    m_bRipping = False
    m_bCancel = False
    lRipIndex = 0
    ProcessRuntimeError "Public Function RipTrack(ByVal lTrack As Long, lFile As String) As Boolean", Err.Description, Err.Number
End Function

Public Sub ShowDrives()
On Local Error GoTo ErrHandler
Dim I As Long
Set m_cRip = New cCDRip
cboDrives.Clear
m_cRip.Create App.Path & "\data\config\rip.ini"
For I = 1 To m_cRip.CDDriveCount
    cboDrives.AddItem m_cRip.CDDrive(I).Name
Next I
If (cboDrives.ListCount > 0) Then cboDrives.ListIndex = 0
ShowTracks
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ShowDrives()", Err.Description, Err.Number
End Sub

Private Sub ShowTracks()
On Local Error GoTo ErrHandler
Dim lIndex As Long, lItem As ListItem, I As Long, cD As cDrive
Set m_cToc = Nothing
lIndex = cboDrives.ListIndex + 1
lvwCD.ListItems.Clear
If (lIndex > 0) Then
    Set cD = m_cRip.CDDrive(lIndex)
    If (cD.IsUnitReady) Then
        Set m_cToc = cD.TOC
        For I = 1 To m_cToc.Count
            With m_cToc.Entry(I)
                Set lItem = lvwCD.ListItems.Add(, "T" & I, Trim(Str(I)))
                lItem.SubItems(1) = "Unknown"
                lItem.SubItems(2) = "Unknown"
                lItem.SubItems(3) = "Unknown"
                lItem.SubItems(4) = .FormattedLength
                lItem.SubItems(5) = "Idle"
                lItem.Selected = True
            End With
        Next I
    Else
        lvwCD.ListItems.Clear
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ShowTracks()", Err.Description, Err.Number
End Sub

Private Sub cboDrives_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, msg3 As String, I As Integer, c As Integer
ShowTracks
SetListViewTagWithToc lvwCD, cboDrives.Text
If Len(lvwCD.Tag) <> 0 Then
    lDiscID = ReturnDiscID(lvwCD.Tag)
    If Len(lvwCD.Tag) <> 0 Then
        c = Int(ReadINI(lIniFiles.iDiscDB, lDiscID, "Count", 0))
        If c = 0 Then
            ConnectToFreeDB wskFreeDB
        Else
            msg = ReadINI(lIniFiles.iDiscDB, lDiscID, "Artist", "")
            msg2 = ReadINI(lIniFiles.iDiscDB, lDiscID, "Album", "")
            If Len(msg) <> 0 And Len(msg2) <> 0 Then
                For I = 1 To c
                    lvwCD.ListItems(I).SubItems(1) = msg
                    lvwCD.ListItems(I).SubItems(2) = msg2
                    lvwCD.ListItems(I).SubItems(3) = ReturnDirCompliant(ReadINI(lIniFiles.iDiscDB, lDiscID, Trim(Str(I)), ""))
                Next I
            End If
            msg3 = Trim(ReturnRipPath())
            If Right(msg3, 1) <> "\" Then msg3 = msg3 & "\"
            msg3 = msg3 & Trim(msg) & " - " & Trim(msg2) & "\"
            MakeNewDir msg3
            dirCopyTo.Path = msg3
            RefreshCDTracks
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cboDrives_Click()", Err.Description, Err.Number
End Sub

Private Sub cboPath_Click()
On Local Error GoTo ErrHandler
Dim I As Integer, msg As String, msg2 As String
For I = 1 To frmMain.tvwFunctions.Nodes.Count
    msg = tvwFunctions.Nodes(I).FullPath
    msg = Replace(msg, "Hard Drives\", "") & "\"
    msg = Replace(msg, "Folders\", "")
    msg = Replace(msg, "My Documents\", lDirectories.dMyDocumentsDir)
    msg = Replace(msg, "My Music\", GetMyDocumentsDir & "\My Documents\My Music\")
    msg = Replace(msg, "Playlist\", "")
    msg = Replace(msg, "\\", "\")
    msg2 = cboPath.Text & "\"
    msg2 = Replace(msg2, "\\", "\")
    msg2 = Replace(msg2, "Installed Hard Drives", "Hard Drives")
    msg2 = Replace(msg2, "Media Folders", "Folders")
    If LCase(msg) = LCase(msg2) Then
        If InStr(msg, "Internet Radio") Then
        Else
            If I <> 0 And I <> 1 Then frmMain.tvwFunctions.Nodes(I).Parent.Expanded = True
            frmMain.tvwFunctions.Nodes(I).Selected = True
            FunctionsTreeView
            DoEvents
            cboPath.SetFocus
            If frmMain.tvwFunctions.Nodes(I).Expanded = False Then frmMain.tvwFunctions.Nodes(I).Expanded = True
            Exit For
        End If
    End If
Next I
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cboPath_Click()", Err.Description, Err.Number
End Sub

Private Sub cboPath_KeyPress(KeyAscii As Integer)
On Local Error GoTo ErrHandler
Dim I As Integer, msg As String, msg2 As String, msg3 As String, lMP3 As MP3File, lArtist As String, lAlbum As String, lTitle As String, lItem As Node
If KeyAscii = 13 Then
    KeyAscii = 0
    If Left(LCase(cboPath.Text), 7) = "http://" Then
        ctlRadio1.StopStream
        ctlRadio1.PlayStream cboPath.Text
    Else
        tvwFiles.Visible = False
        tvwFiles.Nodes.Clear
        tvwFiles.Visible = True
        For I = 1 To tvwPlaylist.Nodes.Count
            If Len(tvwPlaylist.Nodes(I).Text) <> 0 Then
                If InStr(LCase(tvwPlaylist.Nodes(I).Key), LCase(cboPath.Text)) Then
                    msg = tvwPlaylist.Nodes(I).Text
                    msg3 = GetFileTitle(tvwPlaylist.Nodes(I).Key)
                    msg2 = Left(tvwPlaylist.Nodes(I).Key, Len(tvwPlaylist.Nodes(I).Key) - Len(msg3))
                    If Len(msg2) <> 0 Then
                        Set lItem = tvwFiles.Nodes.Add(, , tvwPlaylist.Nodes(I).Key, msg3, 3)
                        lMP3.HasIDv2 = ReadID3v2(tvwPlaylist.Nodes(I).Key, lMP3.IDv2)
                        If lMP3.HasIDv2 = True Then
                            lMP3.IDv2.Artist = CleanInterpreteItems(lMP3.IDv2.Artist)
                            lMP3.IDv2.Album = CleanInterpreteItems(lMP3.IDv2.Album)
                            lMP3.IDv2.Title = CleanInterpreteItems(lMP3.IDv2.Title)
                            lArtist = lMP3.IDv2.Artist
                            lAlbum = lMP3.IDv2.Album
                            lTitle = lMP3.IDv2.Title
                        Else
                            lMP3.HasIDv1 = ReadID3v1(tvwPlaylist.Nodes(I).Key, lMP3.IDv1)
                            If lMP3.HasIDv1 = True Then
                                lMP3.IDv1.Artist = CleanInterpreteItems(lMP3.IDv1.Artist)
                                lMP3.IDv1.Album = CleanInterpreteItems(lMP3.IDv1.Album)
                                lMP3.IDv1.Title = CleanInterpreteItems(lMP3.IDv1.Title)
                                lArtist = lMP3.IDv1.Artist
                                lAlbum = lMP3.IDv1.Album
                                lTitle = lMP3.IDv1.Title
                            End If
                        End If
                    End If
                End If
            End If
        Next I
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cboPath_KeyPress(KeyAscii As Integer)", Err.Description, Err.Number
End Sub

Private Sub cmdAdd_Click()
On Local Error GoTo ErrHandler
Dim lCommonDialog As New pcCommonDialog, lFile As String, vFile As Variant, lItem As ListItem, I As Integer, msg As String, l As Long, lErr As Long, br As Boolean, cWav As New cWavReader
If (lCommonDialog.VBGetOpenFileName(lFile, MultiSelect:=True, Filter:="Wave Files (*.WAV)|*.WAV|", DefaultExt:="WAV", Owner:=Me.hwnd, flags:=OFN_EXPLORER)) Then
    For Each vFile In lCommonDialog.GetMultiSelectFileNames(lFile)
        msg = lFile
        msg = GetFileTitle(msg)
        For I = 1 To lvwBurn.ListItems.Count
            If Trim(LCase(lvwBurn.ListItems(I).SubItems(2))) = Trim(LCase(msg)) Then
                Exit Sub
            End If
        Next I
        br = cWav.OpenFile(lFile)
        lErr = Err.Number
        On Error GoTo 0
        If (br And (lErr = 0)) Then
            Select Case LCase(Right(lFile, 4))
            Case ".wav"
                Set lItem = lvwBurn.ListItems.Add(, lFile, "Ready to Burn")
                lItem.SubItems(1) = "0%"
                lItem.SubItems(2) = msg
                lItem.SubItems(3) = Left(msg, Len(msg) - 4) & ".mp3"
                lItem.SubItems(4) = "WAV"
                lItem.SubItems(5) = "CDA"
                lItem.SubItems(6) = Left(lFile, Len(lFile) - Len(msg))
                lItem.SubItems(7) = cboBurnDrives.Text
            End Select
        Else
            MsgBox "'" & lFile & "' is not a 16bit stereo 44.1kHz Wave File.", vbInformation
        End If
    Next
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdAdd_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdBackward_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
tmrRewind.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdBackward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub cmdBackward_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
tmrRewind.Enabled = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdBackward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub cmdBurn_Click()
On Local Error GoTo ErrHandler
Select Case Trim(LCase(cmdBurn.Caption))
Case "cancel"
    Set cmdBurn.PictureNormal = imgBurn1.Picture
    Set cmdBurn.PictureOver = imgBurn2.Picture
End Select
If lvwBurn.ListItems.Count <> 0 Then
    If (m_eMode = eIdle) Then
        createCD
    Else
        m_bCancel = True
    End If
Else
    Set cmdBurn.PictureNormal = imgCancelBurn1.Picture
    Set cmdBurn.PictureOver = imgCancelBurn2.Picture
    cmdBurn.Caption = "Cancel"
    MsgBox "There are no file(s) to burn!", vbExclamation, App.Title
    Set cmdBurn.PictureNormal = imgBurn1.Picture
    Set cmdBurn.PictureOver = imgBurn2.Picture
    cmdBurn.Caption = "Burn CD"
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdBurn_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdConvert_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, mbox As VbMsgBoxResult
msg = dirCopyTo.Path & "\" & lvwCD.ListItems(lvwCD.SelectedItem.index).SubItems(1) & " - " & Left(lvwCD.ListItems(lvwCD.SelectedItem.index).SubItems(3), Len(lvwCD.ListItems(lvwCD.SelectedItem.index).SubItems(3)) - 2) & ".wav"
If Err.Number <> 0 Then
    Err.Clear
    Exit Sub
End If
If DoesFileExist(msg) = True Then
    If ProcessEntry(lvwCD.SelectedItem.Key, ReturnProcessType(msg), False) = True Then
        ProcessEntry msg, ReturnProcessType(msg), True
    End If
Else
    mbox = MsgBox("This file has not yet been copied from the CD, would you like to copy it now?", vbYesNo + vbQuestion)
    If mbox = vbYes Then
        cmdCopy_Click
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdConvert_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdCopy_Click()
On Local Error GoTo ErrHandler
tmrCopyAll.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCopy_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdCopyAll_Click()
On Local Error GoTo ErrHandler
tmrCopyAll.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdCopyAll_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdDefaultPlaylist_Click()
On Local Error GoTo ErrHandler
LoadPlaylist lIniFiles.iPlaylist, True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdDefaultPlaylist_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdEjectCD_Click()
On Local Error GoTo ErrHandler
Select Case LCase(cmdEjectCD.Caption)
Case "eject cd"
    ctlMovie1.OpenCDDoor
    cmdEjectCD.Caption = "Close CD"
Case "close cd"
    ctlMovie1.CloseDoor
    cmdEjectCD.Caption = "Eject CD"
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdEjectCD_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdForeward_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
tmrFastForward.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdForeward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub cmdForeward_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
tmrFastForward.Enabled = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdForeward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub cmdFullScreen_Click()
On Local Error GoTo ErrHandler
If lFullscreen = False Then
    lFullscreen = True
    ToggleFullScreen True
Else
    lFullscreen = False
    ToggleFullScreen False
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdFullScreen_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdID3Select_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lID3 As ID3Tag, b As Boolean
msg = OpenDialog(Me, "Mpeg Layer 3 Audio (*.mp3)|*.mp3|All Files (*.*)|*.*|", "Select MP3 File", CurDir)
If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        cboID3Type.Clear
        txtID3File.Text = msg
        b = ReadID3v2(msg, lID3)
        If b = True Then
            cboID3Type.AddItem "ID3 Version 2"
            cboID3Type.ListIndex = 0
            txtArtist.Enabled = True
            txtAlbum.Enabled = True
            txtComments.Enabled = True
            txtTitle.Enabled = True
            txtArtist.Text = lID3.Artist
            txtAlbum.Text = lID3.Album
            txtTitle.Text = lID3.Title
            txtComments.Text = lID3.Comment
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                cboID3Type.AddItem "ID3 Version 1"
                cboID3Type.ListIndex = 0
                txtArtist.Enabled = True
                txtAlbum.Enabled = True
                txtComments.Enabled = True
                txtTitle.Enabled = True
                txtArtist.Text = lID3.Artist
                txtAlbum.Text = lID3.Album
                txtTitle.Text = lID3.Title
                txtComments.Text = lID3.Comment
            Else
                cboID3Type.AddItem "No Tag Detected"
                cboID3Type.ListIndex = 0
                txtArtist.Enabled = False
                txtAlbum.Enabled = False
                txtComments.Enabled = False
                txtTitle.Enabled = False
                txtArtist.Text = ""
                txtAlbum.Text = ""
                txtComments.Text = ""
                txtTitle.Text = ""
                txtComments.Text = ""
            End If
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdID3Select_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdLoadPlaylist_Click()
On Local Error GoTo ErrHandler
Dim msg As String
msg = OpenDialog(Me, "Playlist Files (*.m3u)|*.m3u|All Files (*.*)|*.*|", App.Title, CurDir)
If Len(msg) <> 0 Then
    LoadPlaylist msg, True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdMute_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdMute_Click()
On Local Error GoTo ErrHandler
If lMute = True Then
    ctlMovie1.Mute False
    lMute = False
Else
    ctlMovie1.Mute True
    lMute = True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdMute_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdOpen_Click()
On Local Error GoTo ErrHandler
Dim msg As String
msg = Trim(OpenDialog(frmMain, "Supported Types |" & lFileFormats.fSupportedTypes & "|", "Open Media", CurDir))
If Len(msg) <> 0 And DoesFileExist(msg) = True Then OpenMediaFile msg, False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdOpen_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdPauseCD_Click()
On Local Error GoTo ErrHandler
m_cRip.CDDrive(cboDrives.ListIndex + 1).PauseCD
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPauseCD_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdPausePlayback_Click()
On Local Error GoTo ErrHandler
If cmdPausePlayback.Caption = "Pause" Then
    frmMain.sldProgress.Enabled = False
    cmdPausePlayback.Caption = "Resume"
    Select Case LCase(Right(lblFilename.Caption, 4))
    Case ".mp3"
        ctlMP3Player.Pause
    Case Else
        ctlMovie1.PauseMovie
    End Select
ElseIf cmdPausePlayback.Caption = "Resume" Then
    frmMain.sldProgress.Enabled = True
    cmdPausePlayback.Caption = "Pause"
    Select Case LCase(Right(lblFilename.Caption, 4))
    Case ".mp3"
        ctlMP3Player.Pause
    Case Else
        ctlMovie1.ResumeMovie
    End Select
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPausePlayback_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdPlay_Click()
On Local Error GoTo ErrHandler
If Len(lblFilename.Tag) <> 0 Then
    ctlMovie1.PlayMovie
Else
    cmdOpen_Click
    DoEvents
    ctlMovie1.PlayMovie
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPlay_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdPlayCD_Click()
On Local Error GoTo ErrHandler
m_cRip.CDDrive(cboDrives.ListIndex + 1).PlayCDTrack CLng(lvwCD.SelectedItem.Text + 1)

Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdPlayCD_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdRefreshCD_Click()
On Local Error GoTo ErrHandler
cboDrives_Click
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdRefreshCD_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSave_Click()
On Local Error GoTo ErrHandler
Dim t As ID3Tag
If cboID3Type.Text = "ID3 Version 1" Then
    ReadID3v1 txtID3File.Text, t
Else
    ReadID3v2 txtID3File.Text, t
End If
t.Artist = txtArtist.Text
t.Album = txtAlbum.Text
t.Comment = txtComments.Text
t.Title = txtTitle.Text
If cboID3Type.Text = "ID3 Version 1" Then
    WriteID3v1 txtID3File.Text, t
Else
    WriteID3v2 txtID3File.Text, t
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSave_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSavePlaylist_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, I As Integer
For I = 1 To frmMain.tvwPlaylist.Nodes.Count
    If DoesFileExist(frmMain.tvwPlaylist.Nodes(I).Key) = True Then
        If Len(msg) <> 0 Then
            msg = msg & vbCrLf & frmMain.tvwPlaylist.Nodes(I).Key
        Else
            msg = frmMain.tvwPlaylist.Nodes(I).Key
        End If
    End If
Next I
msg2 = SaveDialog(Me, "M3U Playlists (*.m3u)|*.m3u|All Files (*.*)|*.*|", "Save Playlist As ...", CurDir)
If Len(msg2) <> 0 Then
    If Left(msg2, 4) <> ".m3u" Then msg2 = Left(msg2, Len(msg2) - 1) & ".m3u"
    SaveFile msg2, msg
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSavePlaylist_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSearch_Click()
On Local Error GoTo ErrHandler
Dim I As Integer, msg As String, msg2 As String, mItem As ListItem, msg3 As String, lMP3 As MP3File, lArtist As String, lAlbum As String, lTitle As String
If cboSearchType.ListIndex = 0 Then
    lvwSearch.ListItems.Clear
    For I = 1 To tvwPlaylist.Nodes.Count
        If Len(tvwPlaylist.Nodes(I).Text) <> 0 Then
            If InStr(LCase(tvwPlaylist.Nodes(I).Key), txtSearch.Text) Then
                msg = tvwPlaylist.Nodes(I).Text
                msg3 = GetFileTitle(tvwPlaylist.Nodes(I).Key)
                msg2 = Left(tvwPlaylist.Nodes(I).Key, Len(tvwPlaylist.Nodes(I).Key) - Len(msg3))
                If Len(msg2) <> 0 Then
                    Set mItem = lvwSearch.ListItems.Add(, , msg)
                    lMP3.HasIDv2 = ReadID3v2(tvwPlaylist.Nodes(I).Key, lMP3.IDv2)
                    If lMP3.HasIDv2 = True Then
                        lMP3.IDv2.Artist = CleanInterpreteItems(lMP3.IDv2.Artist)
                        lMP3.IDv2.Album = CleanInterpreteItems(lMP3.IDv2.Album)
                        lMP3.IDv2.Title = CleanInterpreteItems(lMP3.IDv2.Title)
                        lArtist = lMP3.IDv2.Artist
                        lAlbum = lMP3.IDv2.Album
                        lTitle = lMP3.IDv2.Title
                    Else
                        lMP3.HasIDv1 = ReadID3v1(tvwPlaylist.Nodes(I).Key, lMP3.IDv1)
                        If lMP3.HasIDv1 = True Then
                            lMP3.IDv1.Artist = CleanInterpreteItems(lMP3.IDv1.Artist)
                            lMP3.IDv1.Album = CleanInterpreteItems(lMP3.IDv1.Album)
                            lMP3.IDv1.Title = CleanInterpreteItems(lMP3.IDv1.Title)
                            lArtist = lMP3.IDv1.Artist
                            lAlbum = lMP3.IDv1.Album
                            lTitle = lMP3.IDv1.Title
                        End If
                    End If
                    mItem.SubItems(1) = lArtist
                    mItem.SubItems(2) = lAlbum
                    mItem.SubItems(3) = msg2
                    mItem.SubItems(4) = msg3
                    Select Case LCase(Right(msg3, 4))
                    Case ".mp3"
                        mItem.SubItems(5) = "Audio Clip"
                    Case ".wav"
                        mItem.SubItems(5) = "Audio Clip"
                    Case "mpeg"
                        mItem.SubItems(5) = "Video Clip"
                    Case ".avi"
                        mItem.SubItems(5) = "Video Clip"
                    End Select
                End If
            End If
        End If
    Next I
ElseIf cboSearchType.ListIndex = 1 Then
    lvwSearch.ListItems.Clear
    For I = 0 To ReturnDriveCount
        SearchHardDrive ReturnHardDriveLetter(I), lvwSearch
    Next I
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSearch_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSelectFilesToBurn_Click()
On Local Error GoTo ErrHandler
frmPromptWaveFiles.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSelectFilesToBurn_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSort_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByArtistAndAlbum
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSort_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSortByAlbum_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByAlbum
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSortByAlbum_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSortByArtist_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByArtist
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSortByArtist_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSortByComments_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByComment
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSortByComments_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSortByType_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByType
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSortByAlbum_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdSortByYear_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByYear
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdSortByYear_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdStop_Click()
On Local Error GoTo ErrHandler
StopPlayback lblFilename.Tag
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdStop_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdStopCD_Click()
On Local Error GoTo ErrHandler
m_cRip.CDDrive(cboDrives.ListIndex + 1).StopCD
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdStopCD_Click()", Err.Description, Err.Number
End Sub

Private Sub cmdWipe_Click()
On Local Error GoTo ErrHandler
Dim t As ID3Tag, I As Integer, msg As String
msg = txtID3File.Text
If Len(msg) <> 0 Then
    t.Album = ""
    t.Artist = ""
    t.Comment = ""
    t.Genre = ""
    t.SongYear = 0
    t.Title = ""
    WriteID3v1 msg, t
    WriteID3v2 msg, t
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub cmdWipe_Click()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGDoubleClick()
On Local Error GoTo ErrHandler
If lFullscreen = True Then
    lFullscreen = False
    ToggleFullScreen False
Else
    lFullscreen = True
    ToggleFullScreen True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGDoubleClick()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMovieClosed()
On Local Error GoTo ErrHandler
cmdPlay.Enabled = False
cmdPausePlayback.Enabled = False
cmdStop.Enabled = False
cmdOpen.Enabled = True
cmdMute.Enabled = False
cmdFullScreen.Enabled = True
cmdBackward.Enabled = False
cmdForeward.Enabled = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMovieClosed()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMovieOpened(lFilename As String)
On Local Error GoTo ErrHandler
cmdPlay.Enabled = True
cmdPausePlayback.Enabled = False
cmdStop.Enabled = False
cmdOpen.Enabled = False
cmdMute.Enabled = False
cmdFullScreen.Enabled = True
cmdBackward.Enabled = True
cmdForeward.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMovieOpened(lFilename As String)", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMoviePaused()
On Local Error GoTo ErrHandler
cmdPlay.Enabled = True
cmdPausePlayback.Enabled = True
cmdStop.Enabled = True
cmdOpen.Enabled = False
cmdMute.Enabled = True
cmdFullScreen.Enabled = True
cmdBackward.Enabled = False
cmdForeward.Enabled = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMoviePaused()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMoviePlay()
On Local Error GoTo ErrHandler
cmdPlay.Enabled = True
cmdPausePlayback.Enabled = True
cmdStop.Enabled = True
cmdOpen.Enabled = False
cmdMute.Enabled = True
cmdFullScreen.Enabled = True
cmdBackward.Enabled = True
cmdForeward.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMoviePlay()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMovieResumed()
On Local Error GoTo ErrHandler
cmdPlay.Enabled = True
cmdPausePlayback.Enabled = True
cmdStop.Enabled = True
cmdOpen.Enabled = False
cmdMute.Enabled = True
cmdFullScreen.Enabled = True
cmdBackward.Enabled = True
cmdForeward.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMovieResumed()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGMovieStopped()
On Local Error GoTo ErrHandler
cmdPlay.Enabled = False
cmdPausePlayback.Enabled = False
cmdStop.Enabled = False
cmdOpen.Enabled = True
cmdMute.Enabled = False
cmdFullScreen.Enabled = True
cmdBackward.Enabled = False
cmdForeward.Enabled = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGMovieStopped()", Err.Description, Err.Number
End Sub

Private Sub ctlMovie1_AGRightClick()
On Local Error GoTo ErrHandler
PopupMenu mnuPB
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMovie1_AGRightClick()", Err.Description, Err.Number
End Sub

Private Sub ctlMP3Decode_PercentDone(ByVal nPercent As Long)
On Local Error GoTo ErrHandler
Dim I As Integer
Select Case Int(nPercent)
Case 100
    I = FindListViewIndexByKey(lvwPending, ctlMP3Decode.Tag)
    If I <> 0 Then lvwPending.ListItems.Remove I
    AddToPlaylistDelay ctlMP3Decode.Tag
    ctlMP3Decode.Tag = ""
    If prgProgress.Visible = True Then prgProgress.Visible = False
    If cboPath.Visible = False Then cboPath.Visible = True
    If prgProgress.Value <> 0 Then prgProgress.Value = 0
    Caption = "Audiogen 2 - Decode Complete"
    FunctionsTreeView
    lBusy = False
    tmrProcessPending.Enabled = True
Case Else
    If prgProgress.Visible = False Then prgProgress.Visible = True
    If cboPath.Visible = True Then cboPath.Visible = False
    If prgProgress.Value <> Int(nPercent) Then prgProgress.Value = Int(nPercent)
End Select
If Len(ctlMP3Decode.Tag) <> 0 Then
    I = FindListViewIndexByKey(lvwPending, ctlMP3Decode.Tag)
    If I <> 0 Then
        Select Case Int(nPercent)
        Case 100
            lvwPending.ListItems(I).SubItems(4) = "Finished"
        Case 0
            lvwPending.ListItems(I).SubItems(4) = "Starting"
        Case Else
            If lvwPending.ListItems(I).SubItems(4) <> "Decode " & nPercent & "%" Then lvwPending.ListItems(I).SubItems(4) = "Decode " & nPercent & "%"
        End Select
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMP3Decode_PercentDone(ByVal nPercent As Long)", Err.Description, Err.Number
End Sub

Private Sub ctlMP3Encode_PercentDone(ByVal nPercent As Long)
On Local Error GoTo ErrHandler
Dim I As Integer
If Int(nPercent) = 100 Then
    RefreshCDTracks
    AddToPlaylistDelay ctlMP3Encode.Tag
    I = FindListViewIndexByKey(lvwPending, ctlMP3Encode.Tag)
    If I <> 0 Then lvwPending.ListItems.Remove I
    If DoesFileExist(ctlMP3Encode.Tag) = True And ReturnAutoDeleteWave = True Then Kill Left(ctlMP3Encode.Tag, Len(ctlMP3Encode.Tag) - 4) & ".wav"
    ctlMP3Encode.Tag = ""
    If prgProgress.Visible = True Then prgProgress.Visible = False
    If cboPath.Visible = False Then cboPath.Visible = True
    If prgProgress.Value <> 0 Then prgProgress.Value = 0
    Caption = "Audiogen 2 - Encode Complete"
    FunctionsTreeView
    lBusy = False
    tmrCopyAll.Enabled = True
    tmrProcessPending.Enabled = True
Else
    If nPercent <> 0 Then
        If prgProgress.Visible = False Then prgProgress.Visible = True
        If cboPath.Visible = True Then cboPath.Visible = False
    End If
    prgProgress.Value = nPercent
    If Len(ctlMP3Encode.Tag) <> 0 Then
        I = FindListViewIndexByKey(lvwBurn, Left(ctlMP3Encode.Tag, Len(ctlMP3Encode.Tag)))
        If I <> 0 Then
            If lvwBurn.ListItems(I).SubItems(1) <> "Encode " & Trim(Str(nPercent)) Then lvwBurn.ListItems(I).SubItems(1) = "Encode " & Trim(Str(nPercent))
        End If
        I = FindListViewIndexByKey(lvwPending, ctlMP3Encode.Tag)
        If I <> 0 Then lvwPending.ListItems(I).SubItems(4) = "Encode " & Trim(Str(nPercent)) & "%"
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMP3Encode_PercentDone(ByVal nPercent As Long)", Err.Description, Err.Number
End Sub

Private Sub ctlMP3Player_FrameNotify(ByVal Frame As Long)
On Local Error Resume Next
If lProgressClicked = False Then
    If sldProgress.Value <> Frame * 100 / lFrameCount Then sldProgress.Value = Frame * 100 / lFrameCount
End If
End Sub

Private Sub ctlMP3Player_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
Dim I As Integer, c As Integer
If Button = 1 Then
    c = ReturnSpectrumCount
    I = ReturnSpectrumIndex
    Select Case I
    Case 0
        LoadSpectrum 1
    Case 1
        If c > 1 Then LoadSpectrum (I + 1)
    Case c
        LoadSpectrum 1
    Case Else
        LoadSpectrum (I + 1)
    End Select
ElseIf Button = 2 Then
    If ctlMP3Player.OscilloType = otSpectrum Then
        ctlMP3Player.OscilloType = otWave
    Else
        ctlMP3Player.OscilloType = otSpectrum
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMP3Player_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)", Err.Description, Err.Number
End Sub

Private Sub ctlMP3Player_PeakFound()
'MsgBox "Hey"
frmMain.fraFunction(5).BackColor = GetRnd(2000)
'If frmMain.fraFunction(5).Visible = True Then
'    frmMain.fraFunction(5).Visible = False
'Else
'    frmMain.fraFunction(5).Visible = True
'End If
End Sub

Private Sub ctlMP3Player_Started(ByVal Frames As Long)
On Local Error GoTo ErrHandler
cmdPausePlayback.Enabled = True
cmdPlay.Enabled = False
cmdStop.Enabled = True
lFrameCount = Frames
frmMain.sldProgress.Max = 100
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMP3Player_Started(ByVal Frames As Long)", Err.Description, Err.Number
End Sub

Private Sub ctlMP3Player_ThreadEnded(ByVal ExitCode As MP3OCXLib.ThreadErrors)
On Local Error GoTo ErrHandler
cmdPausePlayback.Enabled = False
cmdPlay.Enabled = False
cmdStop.Enabled = False
cmdOpen.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlMP3Player_ThreadEnded(ByVal ExitCode As MP3OCXLib.ThreadErrors)", Err.Description, Err.Number
End Sub

Private Sub ctlWMADecode_PercentDone(ByVal nPercent As Double)
On Local Error GoTo ErrHandler
Dim I As Integer
Select Case Int(nPercent)
Case 100
    I = FindListViewIndexByKey(lvwPending, ctlWMADecode.Tag)
    If I <> 0 Then lvwPending.ListItems.Remove I
    AddToPlaylistDelay ctlWMADecode.Tag
    ctlWMADecode.Tag = ""
    If prgProgress.Visible = True Then prgProgress.Visible = False
    If cboPath.Visible = False Then cboPath.Visible = True
    If prgProgress.Value <> 0 Then prgProgress.Value = 0
    Caption = "Audiogen 2 - Decode Complete"
    FunctionsTreeView
    lBusy = False
    tmrProcessPending.Enabled = True
Case Else
    If prgProgress.Visible = False Then prgProgress.Visible = True
    If cboPath.Visible = True Then cboPath.Visible = False
    If prgProgress.Value <> Int(nPercent) Then prgProgress.Value = Int(nPercent)
End Select
If Len(ctlWMADecode.Tag) <> 0 Then
    I = FindListViewIndexByFileTitle(GetFileTitle(ctlWMADecode.Tag), lvwBurn)
    If I <> 0 Then
        If lvwBurn.ListItems(I).SubItems(1) <> Int(nPercent) & "%" Then
            If Int(nPercent) = 100 Then
                ctlWMADecode.Tag = ""
                lvwBurn.ListItems(I).Text = "Finished"
                lvwBurn.ListItems(I).SubItems(1) = Int(nPercent) & "%"
            Else
                If lvwBurn.ListItems(I).Text <> "Decoding" Then lvwBurn.ListItems(I).Text = "Decoding"
                lvwBurn.ListItems(I).SubItems(1) = Int(nPercent) & "%"
            End If
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlWMADecode_PercentDone(ByVal nPercent As Double)", Err.Description, Err.Number
End Sub

Private Sub ctlWMAEncode_PercentDone(ByVal nPercent As Double)
On Local Error GoTo ErrHandler
Dim I As Integer
Select Case Int(nPercent)
Case 100
    I = FindListViewIndexByKey(lvwPending, ctlWMAEncode.Tag)
    If I <> 0 Then lvwPending.ListItems.Remove I
    AddToPlaylistDelay ctlWMAEncode.Tag
    ctlWMAEncode.Tag = ""
    If prgProgress.Visible = True Then prgProgress.Visible = False
    If cboPath.Visible = False Then cboPath.Visible = True
    If prgProgress.Value <> 0 Then prgProgress.Value = 0
    Caption = "Audiogen 2 - Encode Complete"
    FunctionsTreeView
    lBusy = False
    tmrProcessPending.Enabled = True
Case Else
    If prgProgress.Visible = False Then prgProgress.Visible = True
    If cboPath.Visible = True Then cboPath.Visible = False
    If prgProgress.Value <> Int(nPercent) Then prgProgress.Value = Int(nPercent)
End Select
If Len(ctlWMAEncode.Tag) <> 0 Then
    I = FindListViewIndexByFileTitle(GetFileTitle(ctlWMAEncode.Tag), lvwBurn)
    If I <> 0 Then
        If lvwBurn.ListItems(I).SubItems(1) <> Int(nPercent) & "%" Then
            If Int(nPercent) = 100 Then
                ctlWMAEncode.Tag = ""
                lvwBurn.ListItems(I).Text = "Finished"
                lvwBurn.ListItems(I).SubItems(1) = Int(nPercent) & "%"
            Else
                If lvwBurn.ListItems(I).Text <> "Decoding" Then lvwBurn.ListItems(I).Text = "Decoding"
                lvwBurn.ListItems(I).SubItems(1) = Int(nPercent) & "%"
            End If
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub ctlWMAEncode_PercentDone(ByVal nPercent As Double)", Err.Description, Err.Number
End Sub

Private Sub drvRip_Change()
On Local Error Resume Next
dirCopyTo.Path = drvRip.drive
End Sub

Private Sub Form_Load()
On Local Error Resume Next
Dim msg As String, p As clsEditPlaylistEntry
LoadSpectrum 1
Set p = New clsEditPlaylistEntry
p.ResetPlaylist
ctlMovie1.SetBlackDrop 20, 20, 20, 20
If IsRegistered = False Then
    Caption = "Audiogen 2 - Unregistered"
Else
    Caption = "Audiogen 2"
End If
ChDir App.Path
cmdOpen.Enabled = True
ctlWMADecode.MyKey = "¥DI-SUD+FU_YTFÍ-iguDFJD-SDdefNMRR-SsrfEDS¿"
ctlWMAEncode.MyKey = "§RDEH-YRD_WODJT_PEUFNGO-JSIDFJD-SDOENMRR-SDFSEDS_WDFSF¿"
lvwBurn.ListItems.Clear
InitCommonControls
dirCopyTo.Path = ReturnRipPath
msg = GetMyDocumentsDir
ShowDrives
lDirectories.dDesktop = msg & "Desktop\"
lDirectories.dMyDocumentsDir = msg & "My Documents\"
lDirectories.dMyMusicDir = msg & "My Documents\My Music\"
lDirectories.dRootUserDir = msg
lDirectories.dSharedFolder = msg & "Shared\"
sldVolume.Value = Int(ReadINI(lIniFiles.iWindowPositions, "sldVolume", "Value", 100))
cboSearchType.AddItem "Search Playlist"
cboSearchType.AddItem "Search Hard Drives"
cboSearchType.ListIndex = 0
LoadDrives
FillComboWithCDDrives cboDrives
FillTreeViewWithDrives tvwFunctions
FillTreeViewWithFolders tvwFunctions
WindowPosition Me, False
InitResizers
If DoesFileExist(lIniFiles.iPlaylistTreeView) = True Then
    tmrDelayLoadTV.Enabled = True
Else
    Me.Visible = True
    LoadPlaylist lIniFiles.iPlaylist
    tvwPlaylist.Nodes(1).Selected = True
End If
If frmSplash.ReturnOpen = True Then
    Unload frmSplash
End If
cboPath.Text = ""
LoadInternetRadio
RefreshFavoritesMenu
frmMain.tmrDelayExpandPlaylist = True
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub Form_Load()", Err.Description, Err.Number
End Sub

Private Sub InitResizers()
On Local Error Resume Next
picDrag.Height = 80
picDrag.Width = picResizeHorrizontal.Width
picDrag.Top = picResizeHorrizontal.Top
picDrag.Left = picResizeHorrizontal.Left
picDrag.Visible = True
picResizeHorrizontal.Visible = False
picDrag.Visible = False
picWhite.Visible = False
picResizeHorrizontal.Visible = True
picResizeHorrizontal.Left = 40
lvwPending.Top = picResizeHorrizontal.Top
lvwPending.Height = Me.Height - lvwPending.Top - 1650
tvwFiles.Height = Me.Height - lvwPending.Height - 2000
tvwFunctions.Height = Me.Height - lvwPending.Height - 2000
picResizeVerticle.Height = tvwFiles.Height
picResizeVerticle.Visible = False
picDrag.Height = tvwFunctions.Height - 40
picDrag.Width = picResizeVerticle.Width
picDrag.Top = picResizeVerticle.Top + 80
picDrag.Visible = True
picWhite.Visible = True
picWhite.Width = picResizeVerticle.Width + 80
picWhite.Height = picResizeVerticle.Height + 180
picWhite.Left = picResizeVerticle.Left - 50
picWhite.Top = picResizeVerticle.Top + 50
tvwFunctions.Width = picResizeVerticle.Left
tvwFiles.Left = picResizeVerticle.Left + picResizeVerticle.Width
tvwFiles.Width = Me.ScaleWidth - (tvwFunctions.Width + 100)
picResizeVerticle.Top = cboPath.Top + cboPath.Height
picResizeVerticle.Visible = True
picDrag.Visible = False
picWhite.Visible = False
picResizeVerticle.Visible = True
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub InitResizers()", Err.Description, Err.Number
End Sub

Public Sub ActiveateResize()
On Local Error GoTo ErrHandler
Form_Resize
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ActiveateResize()", Err.Description, Err.Number
End Sub

Private Sub Form_Resize()
On Local Error Resume Next
Dim I As Integer
LockWindowUpdate Me.hwnd
If Me.ScaleWidth > 0 And Me.ScaleHeight > 0 Then
    For I = 0 To imgFade.Count
        imgFade(I).Width = Me.ScaleWidth
    Next I
    
    If Me.Width > 9000 Then ctlMP3Player.Width = Me.ScaleWidth - 6000
    LockWindowUpdate frmMain.hwnd
    fraPlaybackControls.Width = Me.ScaleWidth
    txtID3File.Width = Me.ScaleWidth - 2200
    cmdID3Select.Left = (Me.ScaleWidth - 1100)
    cboID3Type.Width = Me.ScaleWidth - 1150
    txtArtist.Width = Me.ScaleWidth - 1150
    txtTitle.Width = Me.ScaleWidth - 1150
    txtAlbum.Width = Me.ScaleWidth - 1150
    txtComments.Width = Me.ScaleWidth - 1150
    prgProgress.Width = Me.ScaleWidth
    prgRipProgress.Width = Me.ScaleWidth
    prgBurn.Width = Me.ScaleWidth
    If Me.ScaleWidth <> 0 Then tvwFunctions.Width = (Me.ScaleWidth / 2)
    If Me.ScaleWidth <> 0 Then tvwFiles.Width = (Me.ScaleWidth / 2) - 80
    tvwFiles.Left = tvwFunctions.Width + picResizeVerticle.Width - 30
    tvwPlaylist.Width = Me.Width - 100
    If Me.Height > 1660 Then tvwPlaylist.Height = Me.Height - 2000
    lvwBurn.Width = Me.Width - 120
    If Me.Height > 2850 Then lvwBurn.Height = Me.Height - 2850
    lvwCD.Width = ((Me.Width - 80) / 4) * 3
    If (Me.Height - frmMain.cboDrives.Height) > 2010 Then lvwCD.Height = (Me.Height - frmMain.cboDrives.Height) - 2060
    dirCopyTo.Width = ((Me.Width - 80) / 3) - 80
    drvRip.Width = ((Me.Width - 80) / 3) - 1100
    dirCopyTo.Left = lvwCD.Width
    drvRip.Left = lvwCD.Width
    If (Me.Height - frmMain.cboDrives.Height) > 2010 Then
        dirCopyTo.Height = (Me.Height - frmMain.cboDrives.Height) - 2000
    End If
    cboDrives.Width = Me.ScaleWidth
    If Me.ScaleHeight <> 0 Then lvwPending.Top = (Me.ScaleHeight / 2) - 30
    If Me.ScaleHeight <> 0 And Me.ScaleHeight > (lvwPending.Height + tblTop.Height) - 750 Then
        tvwFunctions.Height = Me.ScaleHeight - (lvwPending.Height + tblTop.Height) - 750
        tvwFiles.Height = Me.ScaleHeight - (lvwPending.Height + tblTop.Height) - 750
    End If
    If Me.ScaleWidth <> 0 Then lvwPending.Width = Me.ScaleWidth
    picResizeHorrizontal.Top = lvwPending.Top + 30
    picResizeHorrizontal.Width = Me.ScaleWidth
    picResizeVerticle.Left = tvwFunctions.Width
    If tvwFunctions.Height > 200 Then picResizeVerticle.Height = tvwFunctions.Height - 200
    cboPath.Width = picResizeVerticle.Width
    If (Me.ScaleHeight > lvwPending.Top) Then lvwPending.Height = (Me.ScaleHeight - lvwPending.Top)
    cboPath.Width = Me.ScaleWidth
    For I = 0 To fraFunction.UBound
        fraFunction(I).Width = Me.ScaleWidth
        fraFunction(I).Height = Me.ScaleHeight
    Next I
    lblFilename.Width = Me.ScaleWidth - 250
    txtSearch.Width = Me.ScaleWidth / 1.09
    If cboSearchType.Visible = True Then cboSearchType.Visible = False
    cboBurnDrives.Width = Me.ScaleWidth
    lblBurnStatus.Width = Me.ScaleWidth
    lvwSearch.Width = Me.ScaleWidth
    lvwSearch.Height = Me.ScaleHeight - 1200
    cmdSearch.Left = Me.ScaleWidth - cmdSearch.Width - 40
    If Err.Number <> 0 Then
        Err.Clear
    End If
    picResizeVerticle.Top = cboPath.Top + cboPath.Height
    If Err.Number <> 0 Then Err.Clear
    frmMain.Refresh
    InitResizers
    If lFullscreen = True Then
        ctlMovie1.Width = Me.ScaleWidth
        ctlMovie1.Height = Me.ScaleHeight
        ctlMovie1.setSize 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, Me.ScaleHeight / Screen.TwipsPerPixelY
    Else
        ctlMovie1.Width = Me.ScaleWidth
        ctlMovie1.Height = Me.ScaleHeight - 1200
        ctlMovie1.setSize 0, 0, Me.ScaleWidth / Screen.TwipsPerPixelX, Me.ScaleHeight / Screen.TwipsPerPixelY - 93
    End If
End If
LockWindowUpdate 0
If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Local Error Resume Next
EndProgram
End Sub

Private Sub fraFunction_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
If Button = 2 Then PopupMenu mnuPB
End Sub

Private Sub lvwCD_DblClick()
On Local Error GoTo ErrHandler
Dim msg As String
Select Case LCase(lvwCD.SelectedItem.SubItems(5))
Case "idle"
    cmdPlayCD_Click
Case "copied"
    msg = dirCopyTo.Path & "\" & lvwCD.SelectedItem.SubItems(1) & " - " & lvwCD.SelectedItem.SubItems(3) & ".wav"
    If DoesFileExist(msg) = True Then OpenMediaFile msg, True
Case "converted"
    msg = dirCopyTo.Path & "\" & lvwCD.SelectedItem.SubItems(1) & " - " & lvwCD.SelectedItem.SubItems(3) & ".mp3"
    If DoesFileExist(msg) = True Then OpenMediaFile msg, True
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwCD_DblClick()", Err.Description, Err.Number
End Sub

Private Sub lvwPending_DblClick()
On Local Error GoTo ErrHandler
If ProcessEntry(lvwPending.SelectedItem.Key, lvwPending.SelectedItem.SubItems(2), False) = True Then
    ProcessEntry lvwPending.SelectedItem.Key, lvwPending.SelectedItem.SubItems(2), True
    lvwPending.ListItems.Remove lvwPending.SelectedItem.index
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwPending_DblClick()", Err.Description, Err.Number
End Sub

Private Sub lvwPending_KeyDown(KeyCode As Integer, Shift As Integer)
On Local Error GoTo ErrHandler
Select Case KeyCode
Case 46
    lvwPending.ListItems.Remove lvwPending.SelectedItem.index
    If Err.Number <> 0 Then Err.Clear
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwPending_KeyDown(KeyCode As Integer, Shift As Integer)", Err.Description, Err.Number
End Sub

Private Sub lvwPending_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
If Button = 2 Then
    If Len(lvwPending.SelectedItem.Text) <> 0 Then
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
        Select Case LCase(Right(lvwPending.SelectedItem.Key, 4))
        Case ".mp3"
            mnuWavetoWMA123.Enabled = False
            mnuWaveToCDA123.Enabled = False
            mnuWaveToMp3123.Enabled = False
            mnuMP3ToWave123.Enabled = True
            mnuWMAToWave123.Enabled = False
        Case ".wav"
            mnuWavetoWMA123.Enabled = True
            mnuWaveToCDA123.Enabled = True
            mnuWaveToMp3123.Enabled = True
            mnuMP3ToWave123.Enabled = False
            mnuWMAToWave123.Enabled = False
        Case ".wma"
            mnuWavetoWMA123.Enabled = False
            mnuWaveToCDA123.Enabled = False
            mnuWaveToMp3123.Enabled = False
            mnuMP3ToWave123.Enabled = False
            mnuWMAToWave123.Enabled = True
        Case Else
            mnuWavetoWMA123.Enabled = False
            mnuWaveToCDA123.Enabled = False
            mnuWaveToMp3123.Enabled = False
            mnuMP3ToWave123.Enabled = False
            mnuWMAToWave123.Enabled = False
        End Select
        PopupMenu mnuPending
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub lvwPending_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lvwPending_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
AddToPending DraggedKnot.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwPending_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lvwBurn_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 2 Then
    PopupMenu mnulvwProcesses
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwPending_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub lvwSearch_DblClick()
On Local Error GoTo ErrHandler
OpenMediaFile lvwSearch.SelectedItem.SubItems(3) & lvwSearch.SelectedItem.SubItems(4), True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub lvwPending_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub mnuAboutAudiogen2_Click()
On Local Error GoTo ErrHandler
frmAbout.Show 1, frmMain
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAboutAudiogen2_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddAlltoBatch_Click()
On Local Error GoTo ErrHandler
Dim I As Integer
For I = 1 To tvwFiles.Nodes.Count
    If Len(tvwFiles.Nodes(I).Key) <> 0 Then
        AddToPending tvwFiles.Nodes(I).Key
    End If
Next I
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddAlltoBatch_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddAllToBurnQue1_Click()
On Local Error GoTo ErrHandler
Dim I As Integer
For I = 1 To lvwPending.ListItems.Count
    If LCase(Right(lvwPending.ListItems(I).Text, 4)) = ".wav" Then
        AddToBurnQue lvwPending.ListItems(I).Key
        lvwPending.ListItems.Remove I
    End If
Next I
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddAllToBurnQue1_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddBurnToBatch_Click()
On Local Error GoTo ErrHandler
AddToPending tvwFiles.SelectedItem.Key, True, False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddBurnToBatch_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddFavorite_Click()
On Local Error GoTo ErrHandler
frmAddFavorite.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddFavorite_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddFile2_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lTag As ID3Tag
msg = OpenDialog(Me, "Supported Files|" & lFileFormats.fSupportedTypes & "|", "Select File", CurDir)
If Len(msg) <> 0 Then
    lTag = RenderMp3Tag(msg)
Else
    AddToTreeView tvwPlaylist, tvwPlaylist.SelectedItem.Key, tvwChild, msg, GetFileTitle(msg)
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddFile2_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddFileToArtist_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lTag As ID3Tag
msg = OpenDialog(Me, "Supported Files|" & lFileFormats.fSupportedTypes & "|", "Select File", CurDir)
If Len(msg) <> 0 Then
    lTag = RenderMp3Tag(msg)
    AddToTreeView tvwPlaylist, tvwPlaylist.SelectedItem.Key, tvwChild, msg, lTag.Title
Else
    AddToTreeView tvwPlaylist, tvwPlaylist.SelectedItem.Key, tvwChild, msg, GetFileTitle(msg)
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddFileToArtist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddFileToFavorites_Click()
On Local Error GoTo ErrHandler
Dim I As Integer
I = FindTreeViewIndexByFileTitle(tvwFiles.SelectedItem.Text, frmMain.tvwPlaylist)
If I <> 0 Then
    AddFavorite tvwPlaylist.Nodes(I).Text
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddPlaylistItemToFavorites_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddPlaylistItemToFavorites_Click()
On Local Error GoTo ErrHandler
AddFavorite tvwPlaylist.SelectedItem.Text
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddPlaylistItemToFavorites_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddPlayToBatch_Click()
On Local Error GoTo ErrHandler
AddToPending tvwFiles.SelectedItem.Key, False, True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddPlayToBatch_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddStation_Click()
On Local Error GoTo ErrHandler
PromptAddInternetRadio
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddStation_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddStationToFavorites_Click()
On Local Error Resume Next
If Len(tvwFunctions.SelectedItem.Text) <> 0 Then
    AddFavorite tvwFunctions.SelectedItem.Text
Else
    frmAddFavorite.Show 1
End If
Exit Sub
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub mnuAddStationToFavorites_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddToBatch_Click()
On Local Error GoTo ErrHandler
If DoesFileExist(tvwFiles.SelectedItem.Key) = True And Len(tvwFiles.SelectedItem.Key) <> 0 Then AddToPending tvwFiles.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddToBatch_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddToBurnQue_Click()
On Local Error GoTo ErrHandler
AddToBurnQue tvwPlaylist.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddToBurnQue_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddToBurnQue1_Click()
On Local Error GoTo ErrHandler
If DoesFileExist(lvwPending.SelectedItem.Key) = True And Right(lvwPending.SelectedItem.Key, 4) = ".wav" Then
    AddToBurnQue lvwPending.SelectedItem.Key
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddToBurnQue1_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddToBurnQue12345_Click()
On Local Error GoTo ErrHandler
AddToBurnQue tvwPlaylist.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddToBurnQue12345_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAddToBurnQue2_Click()
On Local Error GoTo ErrHandler
AddToBurnQue tvwFiles.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAddToBurnQue2_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuAlphabetic_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylist
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuAlphabetic_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuArrangeByAlbum_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByAlbum
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuArrangeByAlbum_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuArrangeByArtist_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByArtist
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuArrangeByArtist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuArrangeByArtistAndAlbum_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByArtistAndAlbum
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuArrangeByArtistAndAlbum_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuArrangeByFormat_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByType
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuArrangeByFormat_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuCDRipper_Click()
frmCDRipper.Show 1
End Sub

Private Sub mnuChaneTitle_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lMP3 As MP3File, lTag As ID3Tag
If DoesFileExist(tvwPlaylist.SelectedItem.Key) = True And Right(LCase(tvwPlaylist.SelectedItem.Key), 4) = ".mp3" Then
    lTag = RenderMp3Tag(tvwPlaylist.SelectedItem.Key): DoEvents
    msg = InputBox("Enter new Title:", App.Title, lTag.Title)
    If Len(msg) <> 0 Then
        lTag.Title = msg
        lMP3.HasIDv2 = ReadID3v2(tvwPlaylist.SelectedItem.Key, lMP3.IDv2)
        If lMP3.HasIDv2 = True Then
            WriteID3v2 tvwPlaylist.SelectedItem.Key, lTag
        Else
            lMP3.HasIDv1 = ReadID3v1(tvwPlaylist.SelectedItem.Key, lMP3.IDv1)
            If lMP3.HasIDv1 = True Then
                WriteID3v1 tvwPlaylist.SelectedItem.Key, lTag
            Else
                WriteID3v2 tvwPlaylist.SelectedItem.Key, lTag
            End If
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuChaneTitle_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuChangeAlbum_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lType As Integer, lMP3 As MP3File, lTag As ID3Tag, I As Integer
If Right(LCase(tvwPlaylist.SelectedItem.Key), 4) = ".mp3" Then
    lTag = RenderMp3Tag(tvwPlaylist.SelectedItem.Key)
    msg = InputBox("Enter new Album:", App.Title, lTag.Album)
    If Len(msg) <> 0 Then
        If DoesFileExist(tvwPlaylist.SelectedItem.Key) = True Then
            lTag.Album = msg
            lMP3.HasIDv2 = ReadID3v2(tvwPlaylist.SelectedItem.Key, lMP3.IDv2)
            If lMP3.HasIDv2 = True Then
                lType = 2
            Else
                lMP3.HasIDv1 = ReadID3v1(tvwPlaylist.SelectedItem.Key, lMP3.IDv1)
                If lMP3.HasIDv1 = True Then
                    lType = 1
                Else
                    lType = 0
                End If
            End If
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuChangeAlbum_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuChangeArtist_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, lMP3 As MP3File, lTag As ID3Tag, I As Integer, lFile As String, lFileTitle As String
lFile = tvwPlaylist.SelectedItem.Key
lFileTitle = tvwPlaylist.SelectedItem.Text
If Err.Number = 91 Then
    Err.Clear
    Exit Sub
End If
If DoesFileExist(lFile) = True And Right(LCase(lFile), 4) = ".mp3" Then
    lTag = RenderMp3Tag(lFile): DoEvents
    msg2 = Parse(GetFileTitle(tvwPlaylist.SelectedItem.Key), "(", ")")
    If Len(msg2) = 0 Then
        msg2 = lTag.Artist
    Else
        If Len(lTag.Title) = 0 Then
            lTag.Title = lFileTitle
        End If
    End If
    msg = InputBox("Enter new Artist:", App.Title, msg2)
    If Len(msg) <> 0 Then
        lTag.Artist = msg
        lMP3.HasIDv2 = ReadID3v2(lFile, lMP3.IDv2)
        If lMP3.HasIDv2 = True Then
            WriteID3v2 lFile, lTag
        Else
            lMP3.HasIDv1 = ReadID3v1(lFile, lMP3.IDv1)
            If lMP3.HasIDv1 = True Then
                WriteID3v1 lFile, lTag
            Else
                WriteID3v2 lFile, lTag
            End If
        End If
        DoEvents
        I = FindTreeViewIndex(tvwPlaylist.SelectedItem.Text, tvwPlaylist)
        tvwPlaylist.Nodes.Remove I
        If DoesTreeViewItemExist(msg, tvwPlaylist) = True Then
            If Len(lTag.Title) <> 0 Then
                lFileTitle = lTag.Title
            Else
                If Len(lFileTitle) = 0 Then lFileTitle = GetFileTitle(lFile)
            End If
            AddToTreeView tvwPlaylist, msg, tvwChild, lFile, lFileTitle
        Else
            If Len(lTag.Title) <> 0 Then
                lFileTitle = lTag.Title
            Else
                If Len(lFileTitle) = 0 Then lFileTitle = GetFileTitle(lFile)
            End If
            AddToTreeView tvwPlaylist, "MP3", tvwChild, msg, msg, 8
            AddToTreeView tvwPlaylist, msg, tvwChild, lFile, lFileTitle
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuChangeArtist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuClear_Click()
On Local Error GoTo ErrHandler
lvwPending.ListItems.Clear
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuChangeArtist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuClearFiles_Click()
On Local Error GoTo ErrHandler
tvwFiles.Nodes.Clear
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuClearFiles_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuClearPlaylist_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.ResetPlaylist
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuClearPlaylist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuClearProcesses_Click()
On Local Error GoTo ErrHandler
lvwBurn.ListItems.Clear
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuClearPlaylist_Click()", Err.Description, Err.Number
End Sub


Private Sub mnuContract_Click()
On Local Error GoTo ErrHandler
tvwPlaylist.SelectedItem.Expanded = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuContract_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuConvertWave_Click()
On Local Error GoTo ErrHandler
If ProcessEntry(tvwPlaylist.SelectedItem.Key, ReturnProcessType(tvwPlaylist.SelectedItem.Key), False) = True Then ProcessEntry tvwPlaylist.SelectedItem.Key, ReturnProcessType(tvwPlaylist.SelectedItem.Key), True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuConvertWave_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuDecode_Click()
On Local Error GoTo ErrHandler
If ProcessEntry(tvwPlaylist.SelectedItem.Key, ReturnProcessType(tvwPlaylist.SelectedItem.Key), False) = True Then ProcessEntry tvwPlaylist.SelectedItem.Key, ReturnProcessType(tvwPlaylist.SelectedItem.Key), True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuDecode_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuDecodeFile_Click()
On Local Error GoTo ErrHandler
If ProcessEntry(tvwFiles.SelectedItem.Key, ReturnProcessType(tvwFiles.SelectedItem.Key), False) = True Then ProcessEntry tvwFiles.SelectedItem.Key, ReturnProcessType(tvwFiles.SelectedItem.Key), True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuDecodeFile_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuDecodeFile1_Click()
On Local Error GoTo ErrHandler
If ProcessEntry(tvwFiles.SelectedItem.Key, ReturnProcessType(tvwFiles.SelectedItem.Key), False) = True Then
    ProcessEntry tvwFiles.SelectedItem.Key, ReturnProcessType(tvwFiles.SelectedItem.Key), True
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuDecodeFile1_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuDelete_Click()
On Local Error GoTo ErrHandler
Dim mbox As VbMsgBoxResult
mbox = MsgBox("Are you sure you wish to delete " & tvwFiles.SelectedItem.Text & "?", vbYesNo + vbQuestion, App.Title)
If mbox = vbYes Then
    If DoesFileExist(tvwFiles.SelectedItem.Key) = True Then
        Kill tvwFiles.SelectedItem.Key
        tvwFiles.Nodes.Remove tvwFiles.SelectedItem.index
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuDelete_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuEditFavorites_Click()
On Local Error GoTo ErrHandler
frmEditFavorites.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuEditFavorites_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuEditInternetRadio_Click()
On Local Error GoTo ErrHandler
frmEditInternetRadio.Show 0, frmMain
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuEditInternetRadio_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuEditTag_Click()
On Local Error GoTo ErrHandler
Dim msg As String, lID3 As ID3Tag, b As Boolean
msg = tvwPlaylist.SelectedItem.Key
If Len(msg) <> 0 Then
    If DoesFileExist(msg) = True Then
        cboID3Type.Clear
        txtID3File.Text = msg
        b = ReadID3v2(msg, lID3)
        If b = True Then
            cboID3Type.AddItem "ID3 Version 2"
            cboID3Type.ListIndex = 0
            txtArtist.Enabled = True
            txtAlbum.Enabled = True
            txtComments.Enabled = True
            txtTitle.Enabled = True
            txtArtist.Text = lID3.Artist
            txtAlbum.Text = lID3.Album
            txtTitle.Text = lID3.Title
            txtComments.Text = lID3.Comment
        Else
            b = ReadID3v1(msg, lID3)
            If b = True Then
                cboID3Type.AddItem "ID3 Version 1"
                cboID3Type.ListIndex = 0
                txtArtist.Enabled = True
                txtAlbum.Enabled = True
                txtComments.Enabled = True
                txtTitle.Enabled = True
                txtArtist.Text = lID3.Artist
                txtAlbum.Text = lID3.Album
                txtTitle.Text = lID3.Title
                txtComments.Text = lID3.Comment
            Else
                cboID3Type.AddItem "No Tag Detected"
                cboID3Type.ListIndex = 0
                txtArtist.Enabled = False
                txtAlbum.Enabled = False
                txtComments.Enabled = False
                txtTitle.Enabled = False
7                txtArtist.Text = ""
                txtAlbum.Text = ""
                txtComments.Text = ""
                txtTitle.Text = ""
                txtComments.Text = ""
            End If
        End If
    End If
End If
ResetFrames
fraFunction(6).Visible = True
tblTop.Buttons(7).Value = tbrPressed
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuEditTag_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuEncodeToMp3_Click()
On Local Error GoTo ErrHandler
If LCase(Right(tvwFiles.SelectedItem.Key, 4)) = ".wav" Then
    If ProcessEntry(tvwFiles.SelectedItem.Key, "Wave to MP3", False) = True Then
        ProcessEntry tvwFiles.SelectedItem.Key, "Wave to MP3", True
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuEncodeToMp3_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuEncodeToWMA_Click()
On Local Error GoTo ErrHandler
If LCase(Right(tvwFiles.SelectedItem.Key, 4)) = ".wav" Then
    If ProcessEntry(tvwFiles.SelectedItem.Key, "Wave to WMA", False) = True Then
        ProcessEntry tvwFiles.SelectedItem.Key, "Wave to WMA", True
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuEncodeToWMA_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuExit_Click()
On Local Error GoTo ErrHandler
Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuExit_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuExpand_Click()
On Local Error GoTo ErrHandler
tvwPlaylist.SelectedItem.Expanded = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuExpand_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuFavorite_Click(index As Integer)
On Local Error GoTo ErrHandler
Dim I As Integer
I = FindTreeViewIndex(mnuFavorite(index).Caption, frmMain.tvwPlaylist)
If I <> 0 Then
    OpenMediaFile tvwPlaylist.Nodes(I).Key, True
Else
    If Left(LCase(mnuFavorite(index).Caption), 7) = "http://" Then
        PlayInternetRadio frmMain.ctlRadio1, mnuFavorite(index).Caption
    Else
        PlayInternetRadio frmMain.ctlRadio1, ReturnInternetRadioAddress(mnuFavorite(index).Caption)
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuFavorite_Click(Index As Integer)", Err.Description, Err.Number
End Sub

Private Sub mnuForumWeb_Click()
On Local Error GoTo ErrHandler
Surf "http://www.tnexgen.com/forum/", Me.hwnd
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuForumWeb_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuForward_Click()
On Local Error GoTo ErrHandler
tmrFastForward.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRewindM_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuFullScreenM_Click()
On Local Error GoTo ErrHandler
ctlMovie1.FullScreen
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuFullScreenM_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuHide_Click()
On Local Error GoTo ErrHandler
If Len(tvwFiles.SelectedItem.Key) <> 0 And DoesFileExist(tvwFiles.SelectedItem.Key) = True Then
    tvwFiles.Nodes.Remove tvwFiles.SelectedItem.index
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuHide_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuHomeWeb_Click()
On Local Error GoTo ErrHandler
Surf "http://www.tnexgen.com", Me.hwnd
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuHomeWeb_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuMP3ToWave123_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "MP3 to Wave"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuMP3ToWave123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuOpenCDDoor_Click()
On Local Error GoTo ErrHandler
ctlMovie1.OpenCDDoor
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuMP3ToWave123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuOpenMovie_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, I As Integer
msg = OpenDialog(frmMain, "Supported|*.m4a;*.avi;*.mpg;*.mpeg;*.mpe;*.mp3;*.mp2;*.mp1;*.wav;*.aif;*.aiff;*.aifc;*.au;*.mv1;*.mov;*.mpa;*.qt;*.snd;*.mpm;*.mpv;*.enc;*.mid;*.rmi;*.vob;*.wma;*.wmv|", App.Title, CurDir)
If Len(msg) <> 0 Then
    msg2 = msg
    msg2 = GetFileTitle(msg2)
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuMP3ToWave123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuPauseMovieM_Click()
On Local Error GoTo ErrHandler
ctlMovie1.PauseMovie
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuPauseMovieM_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuPlayFile_Click()
On Local Error GoTo ErrHandler
OpenMediaFile tvwFiles.SelectedItem.Key, True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuPlayFile_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuPlayMovie_Click()
On Local Error GoTo ErrHandler
ctlMovie1.PlayMovie
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuPlayMovie_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuPlayStation_Click()
On Local Error Resume Next
Dim msg As String, I As Integer
msg = tvwFunctions.SelectedItem.Text
If Err.Number = 91 Then
    I = AddtoInternetRadio(InputBox("Enter name of Radio station"), InputBox("Enter url of radio station:"), False)
    Err.Clear
Else
    msg = ""
    For I = 1 To ReturnInternetRadioCount()
        If Trim(LCase(ReturnInternetRadioName(I))) = Trim(LCase(tvwFunctions.SelectedItem.Text)) Then
            msg = ReturnInternetRadioAddress(tvwFunctions.SelectedItem.Text)
            If Len(msg) <> 0 Then
                PlayInternetRadio ctlRadio1, msg
            End If
            Exit For
        End If
    Next I
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub mnuPlayStation_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuPlayThisFile_Click()
On Local Error GoTo ErrHandler
OpenMediaFile tvwPlaylist.SelectedItem.Key, True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuPlayThisFile_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProcessAllEntrys_Click()
On Local Error GoTo ErrHandler
Dim I As Integer, c As Integer
If lvwPending.ListItems.Count <> 0 And lvwPending.ListItems.Count <> 1 Then
    c = lvwPending.ListItems.Count + 1
    For I = 1 To c
        If I < c Then
            If ProcessEntry(lvwPending.ListItems(I).Key, lvwPending.ListItems(I).SubItems(2), False) = True Then
                ProcessEntry lvwPending.ListItems(I).Key, lvwPending.ListItems(I).SubItems(2), True
                lvwPending.ListItems.Remove lvwPending.ListItems(I).index
            End If
        End If
    Next I
ElseIf lvwPending.ListItems.Count = 1 Then
    If ProcessEntry(lvwPending.ListItems(1).Key, lvwPending.ListItems(1).SubItems(2), False) Then
        ProcessEntry lvwPending.ListItems(1).Key, lvwPending.ListItems(1).SubItems(2), True
        frmMain.lvwPending.ListItems.Remove 1
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProcessAllEntrys_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProcessNow_Click()
On Local Error GoTo ErrHandler
tmrProcessPending.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProcessNow_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProcessQUe_Click()
On Local Error GoTo ErrHandler
tmrProcessPending.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProcessQUe_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProcessThisItemOnly_Click()
On Local Error GoTo ErrHandler
ProcessEntry lvwPending.SelectedItem.Key, lvwPending.SelectedItem.SubItems(2), True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProporties_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProporties_Click()
On Local Error GoTo ErrHandler
ShowFileProperties Me.hwnd, tvwFiles.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProporties_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProporties2_Click()
On Local Error GoTo ErrHandler
ShowFileProperties Me.hwnd, tvwFunctions.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProporties2_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuProporties3_Click()
On Local Error GoTo ErrHandler
ShowFileProperties Me.hwnd, tvwPlaylist.SelectedItem.Key
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuProporties2_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRemove_Click()
On Local Error GoTo ErrHandler
Dim mbox As VbMsgBoxResult
If Len(lvwPending.SelectedItem.Text) <> 0 Then
    mbox = MsgBox("Are you sure you wish to remove '" & lvwPending.SelectedItem.Text & "'?", vbYesNo + vbQuestion)
    If mbox = vbYes Then lvwPending.ListItems.Remove lvwPending.SelectedItem.index
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRemoveArtist_Click()
On Local Error GoTo ErrHandler
tvwPlaylist.Nodes.Remove tvwPlaylist.SelectedItem.index
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRemoveEntry_Click()
On Local Error GoTo ErrHandler
lvwBurn.ListItems.Remove lvwBurn.SelectedItem.index
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRemoveFile_Click()
On Local Error GoTo ErrHandler
tvwPlaylist.Nodes.Remove tvwPlaylist.SelectedItem.index
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRemoveStation_Click()
On Local Error GoTo ErrHandler
DeleteTreeViewInternetRadio tvwFunctions, tvwFunctions.SelectedItem.Text
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRemove_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuResumeMovie_Click()
On Local Error GoTo ErrHandler
ctlMovie1.ResumeMovie
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuResumeMovie_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuRewindM_Click()
On Local Error GoTo ErrHandler
tmrRewind.Enabled = True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuRewindM_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuSelectedEntrys_Click()
On Local Error GoTo ErrHandler
Dim I As Integer
For I = 1 To lvwPending.ListItems.Count
    If I < lvwPending.ListItems.Count Then
        If lvwPending.ListItems(I).Selected = True Then
            If ProcessEntry(lvwPending.ListItems(I).Key, lvwPending.ListItems(I).SubItems(2), False) = True Then
                frmMain.lvwPending.ListItems.Remove I
            End If
        End If
    End If
Next I
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuSelectedEntrys_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuSaveFilesAsPlaylist_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, I As Integer
For I = 1 To frmMain.tvwFiles.Nodes.Count
    If DoesFileExist(frmMain.tvwFiles.Nodes(I).Key) = True Then
        If Len(msg) <> 0 Then
            msg = msg & vbCrLf & frmMain.tvwFiles.Nodes(I).Key
        Else
            msg = frmMain.tvwFiles.Nodes(I).Key
        End If
    End If
Next I
msg2 = SaveDialog(Me, "M3U Files (*.m3u)|*.m3u|All Files (*.*)|*.*|", "Save as Playlist", CurDir)
If Len(msg2) <> 0 Then
    msg2 = Left(msg2, Len(msg2) - 1) & ".m3u"
    SaveFile msg2, msg
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuSaveFilesAsPlaylist_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuShowContainingFolder_Click()
On Local Error GoTo ErrHandler
Dim msg As String, msg2 As String, I As Integer
msg = tvwFiles.SelectedItem.Key
If Len(msg) <> 0 Then
    tvwFiles.Visible = False
    tvwFiles.Nodes.Clear
    tvwFiles.Visible = True
    msg2 = GetFileTitle(msg)
    msg2 = Left(msg, Len(msg) - Len(msg2))
    cboPath.AddItem msg2
    cboPath.ListIndex = FindComboBoxIndex(cboPath, msg2)
    cboPath_KeyPress 13
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuShowContainingFolder_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuSortByComments_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByComment
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuSortByComments_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuSortByYear_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistByYear
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuSortByYear_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuSortParentNode_Click()
On Local Error GoTo ErrHandler
Dim e As clsEditPlaylistEntry
Set e = New clsEditPlaylistEntry
e.SortPlaylistNode tvwPlaylist.SelectedItem.Parent.index
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuSortParentNode_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuStopMovieM_Click()
On Local Error GoTo ErrHandler
StopPlayback lblFilename.Tag
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuStopMovieM_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuStopStation_Click()
On Local Error GoTo ErrHandler
StopInternetRadio frmMain.ctlRadio1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuStopStation_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuThisEntry_Click()
On Local Error GoTo ErrHandler
If ProcessEntry(lvwPending.SelectedItem.Key, lvwPending.SelectedItem.SubItems(2), False) = True Then
    ProcessEntry lvwPending.SelectedItem.Key, lvwPending.SelectedItem.SubItems(2), True
    lvwPending.ListItems.Remove lvwPending.SelectedItem.index
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuThisEntry_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuToWaveToCDA_Click()

End Sub

Private Sub mnuToPlay_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "Play"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuToPlay_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuViewErrors_Click()
On Local Error GoTo ErrHandler
frmErrorReview.Show 1
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuViewErrors_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuWaveToCDA123_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "Wave to CDA"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuWaveToCDA123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuWaveToMp3123_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "Wave to MP3"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuWaveToMp3123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuWavetoWMA123_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "Wave to WMA"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuWavetoWMA123_Click()", Err.Description, Err.Number
End Sub

Private Sub mnuWMAToWave123_Click()
On Local Error GoTo ErrHandler
lvwPending.SelectedItem.SubItems(2) = "WMA to Wave"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub mnuWMAToWave123_Click()", Err.Description, Err.Number
End Sub

Private Sub picResizeHorrizontal_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    picDrag.Height = 80
    picDrag.Width = picResizeHorrizontal.Width
    picDrag.Top = picResizeHorrizontal.Top
    picDrag.Left = picResizeHorrizontal.Left
    picDrag.Visible = True
    lResizeVert = False
    picResizeHorrizontal.Visible = False
    tmrDragBlackLine.Enabled = True
    FormDrag picResizeHorrizontal.hwnd
    tmrDragBlackLine.Enabled = False
    picDrag.Visible = False
    picWhite.Visible = False
    picResizeHorrizontal.Visible = True
    picResizeHorrizontal.Left = 40
    lvwPending.Top = picResizeHorrizontal.Top
    lvwPending.Height = Me.Height - lvwPending.Top - 1650
    tvwFiles.Height = Me.Height - lvwPending.Height - 2000
    tvwFunctions.Height = Me.Height - lvwPending.Height - 2000
    picResizeVerticle.Height = tvwFiles.Height
    frmMain.Refresh
Else
    picDrag.Visible = False
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub picResizeHorrizontal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub picResizeVerticle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then
    lResizeVert = True
    picResizeVerticle.Visible = False
    picDrag.Height = tvwFunctions.Height - 40
    picDrag.Width = picResizeVerticle.Width
    picDrag.Top = picResizeVerticle.Top + 80
    picDrag.Visible = True
    picWhite.Visible = True
    picWhite.Width = picResizeVerticle.Width + 80
    picWhite.Height = picResizeVerticle.Height + 10
    picWhite.Left = picResizeVerticle.Left - 50
    picWhite.Top = picResizeVerticle.Top + 50
    tmrDragBlackLine.Enabled = True
    FormDrag picResizeVerticle.hwnd
    tvwFunctions.Width = picResizeVerticle.Left
    tvwFiles.Left = picResizeVerticle.Left + picResizeVerticle.Width
    tvwFiles.Width = Me.ScaleWidth - (tvwFunctions.Width + 100)
    picResizeVerticle.Top = cboPath.Top + cboPath.Height
    picResizeVerticle.Visible = True
    tmrDragBlackLine.Enabled = False
    picDrag.Visible = False
    picWhite.Visible = False
    picResizeVerticle.Visible = True
Else
    picDrag.Visible = False
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub picResizeVerticle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub sldProgress_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
lProgressClicked = True
End Sub

Private Sub sldProgress_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error Resume Next
Select Case LCase(Right(lblFilename.Caption, 4))
Case ".mp3"
    ctlMP3Player.Seek sldProgress.Value * lFrameCount / 100
Case Else
    ctlMovie1.ChangeMoviePosition sldProgress.Value
End Select
lProgressClicked = False
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub sldProgress_Click()", Err.Description, Err.Number
End Sub

Private Sub sldVolume_Click()
On Local Error GoTo ErrHandler
Select Case Right(LCase(lblFilename.Tag), 4)
Case ".mp3"
    ctlMP3Player.SetVolume sldVolume.Value, sldVolume.Value
Case Else
    ctlMovie1.SetVolume sldVolume.Value * 10
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub sldVolume_Click()", Err.Description, Err.Number
End Sub

Private Sub sldVolume_Scroll()
On Local Error GoTo ErrHandler
sldVolume_Click
ErrHandler:
    ProcessRuntimeError "Private Sub sldVolume_Scroll()", Err.Description, Err.Number
End Sub

Private Sub tblTop_ButtonClick(ByVal Button As MSComctlLib.Button)
On Local Error GoTo ErrHandler
ResetFrames
fraFunction(Button.index - 1).Visible = True
DoEvents
Select Case Button.index
Case 3
    If cboDrives.ListCount = 1 And Len(cboDrives.Text) = 0 Then
        cboDrives.ListIndex = 0
    End If
Case 4
    If cboBurnDrives.ListCount = 0 Then
        showRecorders
    Else
        If Len(cboBurnDrives.Text) = 0 Then
            cboBurnDrives.ListIndex = 0
        End If
    End If
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tblTop_ButtonClick(ByVal Button As MSComctlLib.Button)", Err.Description, Err.Number
End Sub

Private Sub tmrCopyAll_Timer()
On Local Error Resume Next
If lRipping = False Then
    Dim I As Integer, msg As String, b As Boolean, o As Boolean
    For I = 1 To lvwCD.ListItems.Count
        If lvwCD.ListItems(I).Checked = True Then
            Select Case LCase(Trim(lvwCD.ListItems(I).SubItems(5)))
            Case "copied"
                msg = dirCopyTo.Path & "\" & lvwCD.ListItems(I).SubItems(1) & " - " & Left(lvwCD.ListItems(I).SubItems(3), Len(lvwCD.ListItems(I).SubItems(3))) & ".mp3"
                If DoesFileExist(msg) = False Then
                    lvwCD.ListItems(I).SubItems(5) = "Converting ..."
                    lvwCD.ListItems(I).ForeColor = vbBlue
                    lvwCD.ListItems(I).ListSubItems(1).ForeColor = vbBlue
                    lvwCD.ListItems(I).ListSubItems(2).ForeColor = vbBlue
                    lvwCD.ListItems(I).ListSubItems(3).ForeColor = vbBlue
                    lvwCD.ListItems(I).ListSubItems(4).ForeColor = vbBlue
                    lvwCD.ListItems(I).ListSubItems(5).ForeColor = vbBlue
                    tmrCopyAll.Enabled = False
                    ProcessEntry Left(msg, Len(msg) - 4) & ".wav", "Wave to MP3", True
                    o = True
                End If
                Exit For
            Case "converted"
            Case "idle"
                msg = dirCopyTo.Path & "\" & lvwCD.ListItems(I).SubItems(1) & " - " & Left(lvwCD.ListItems(I).SubItems(3), Len(lvwCD.ListItems(I).SubItems(3))) & ".wav"
                If DoesFileExist(msg) = False Then
                    lRipIndex = I
                    If Err.Number <> 0 Then Err.Clear
                    tmrCopyAll.Enabled = False
                    b = RipTrack(CLng(lRipIndex), msg)
                    o = True
                End If
                Exit For
            End Select
        End If
    Next I
    If o = False Then tmrCopyAll.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrCopyAll_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrDelayAddToPlaylist_Timer()
On Local Error Resume Next
Dim I As Integer
Select Case lstAddToPlaylist.ListCount
Case 0
    tmrDelayAddToPlaylist.Enabled = False
Case -1
    tmrDelayAddToPlaylist.Enabled = False
Case Else
    AddToPlaylist lstAddToPlaylist.List(0)
    lstAddToPlaylist.RemoveItem 0
End Select
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrDelayAddToPlaylist_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrDelayExpandPlaylist_Timer()
On Local Error Resume Next
Dim I As Integer
For I = 1 To frmMain.tvwPlaylist.Nodes.Count
    frmMain.tvwPlaylist.Nodes(I).Expanded = True
Next I
tmrDelayExpandPlaylist.Enabled = False
Me.SetFocus
End Sub

Private Sub tmrDelayLoadTV_Timer()
On Local Error Resume Next
tmrDelayLoadTV.Enabled = False
LoadTVFromFile tvwPlaylist, lIniFiles.iPlaylistTreeView
tvwPlaylist.Nodes(1).Selected = True
frmMain.Visible = True
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrDelayLoadTV_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrDragBlackLine_Timer()
On Local Error Resume Next
If picDrag.Visible = True Then
    If lResizeVert = True Then
        picDrag.Left = picResizeVerticle.Left
    Else
        picDrag.Top = picResizeHorrizontal.Top
    End If
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrDragBlackLine_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrFastForward_Timer()
On Local Error Resume Next
ctlMovie1.ForwardFrames 80
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrFastForward_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrProcessPending_Timer()
On Local Error Resume Next
Dim I As Integer
If lBusy = False Then
    If lvwPending.ListItems.Count <> 0 Then
        For I = 1 To lvwPending.ListItems.Count
            Select Case LCase(lvwPending.ListItems(I).SubItems(2))
            Case "play"
                lBusy = True
                ProcessEntry lvwPending.ListItems(I).Key, "Play", True
                tmrProcessPending.Enabled = False
                Exit For
            Case "wave to wma"
                If LCase(Right(lvwPending.ListItems(I).Key, 4)) = ".wav" Then
                    lBusy = True
                    ProcessEntry lvwPending.ListItems(I).Key, "Wave to WMA", True
                    tmrProcessPending.Enabled = False
                    Exit For
                End If
            Case "wma to wave"
                If LCase(Right(lvwPending.ListItems(I).Key, 4)) = ".wma" Then
                    lBusy = True
                    ProcessEntry lvwPending.ListItems(I).Key, "WMA to Wave", True
                    Exit For
                End If
            Case "mp3 to wave"
                If LCase(Right(lvwPending.ListItems(I).Key, 4)) = ".mp3" Then
                    lBusy = True
                    ProcessEntry lvwPending.ListItems(I).Key, "MP3 to Wave", True
                    tmrProcessPending.Enabled = False
                    Exit For
                End If
            Case "wave to mp3"
                If LCase(Right(lvwPending.ListItems(I).Key, 4)) = ".wav" Then
                    lBusy = True
                    ProcessEntry lvwPending.ListItems(I).Key, "Wave to MP3", True
                    tmrProcessPending.Enabled = False
                    Exit For
                End If
            Case "wave to cda"
                If LCase(Right(lvwPending.ListItems(I).Key, 4)) = ".wav" Then
                    AddToBurnQue lvwPending.ListItems(I).Key
                    lvwPending.ListItems.Remove I
                    Exit For
                End If
            End Select
        Next I
    ElseIf lvwPending.ListItems.Count = 0 Then
        tmrProcessPending.Enabled = False
    ElseIf lvwPending.ListItems.Count = -1 Then
        tmrProcessPending.Enabled = False
    End If
Else
    tmrProcessPending.Enabled = False
    Exit Sub
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrProcessPending_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrProgress_Timer()
On Local Error Resume Next
Dim I As Integer
If LCase(Right(lblFilename.Tag, 4)) = ".mp3" Then
    tmrProgress.Enabled = False
    Exit Sub
End If
If Len(lblFilename.Tag) <> 0 Then
    If lProgressClicked = False Then
        frmMain.sldProgress.Value = ctlMovie1.ReturnCurrentPosition
        I = FindListViewIndexByKey(lvwPending, lblFilename.Tag)
        If frmMain.sldProgress.Value <> 0 Then
            If I <> 0 Then
                If lvwPending.ListItems(I).SubItems(4) <> Int(frmMain.sldProgress.Value) & "%" Then
                    If Int(frmMain.sldProgress.Value) = frmMain.sldProgress.Max Then
                        lvwPending.ListItems(I).SubItems(4) = "Finished"
                        lvwPending.ListItems.Remove I
                        StopPlayback lblFilename.Tag
                        tmrProgress.Enabled = False
                        frmMain.SetBusy False
                        frmMain.tmrProcessPending.Enabled = True
                        lBusy = False
                    Else
                        If lvwPending.ListItems(I).SubItems(4) <> "Play " & Int(frmMain.sldProgress.Value * 100 / frmMain.sldProgress.Max) & "%" Then lvwPending.ListItems(I).SubItems(4) = "Play " & Int(frmMain.sldProgress.Value * 100 / frmMain.sldProgress.Max) & "%"
                    End If
                End If
            End If
        End If
    End If
Else
    tmrProgress.Enabled = False
End If
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrProgress_Timer()", Err.Description, Err.Number
End Sub

Private Sub tmrRewind_Timer()
On Local Error Resume Next
ctlMovie1.RewindFrames 80
If Err.Number <> 0 Then ProcessRuntimeError "Private Sub tmrRewind_Timer()", Err.Description, Err.Number
End Sub

Private Sub tvwFiles_Click()
On Local Error GoTo ErrHandler
If LCase(tvwFunctions.SelectedItem.Text) = "bitrate" Then
    SetMP3Bitrate tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "sample rate" Then
    SetMP3SampleRate tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "channels" Then
    SetMP3Channels tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "auto delete wave" Then
    SetAutoDeleteWave tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "attributes" Then
    SetAttributes tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "auto eject" Then
    SetAutoEject tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "test mode" Then
    SetTestMode tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "auto normalize" Then
    SetAutoNormalize tvwFiles.SelectedItem.Text
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "cd speed" Then
    SetCDSpeed Int(tvwFiles.SelectedItem.Text)
ElseIf LCase(tvwFunctions.SelectedItem.Text) = "rip path" Then
    If LCase(tvwFiles.SelectedItem.Text) = "change" Then
        Dim d As New frmDirSelect
        Set d = New frmDirSelect
        d.Show 1
        If Len(ReturnSelectedDirectory()) <> 0 Then
            SetRipPath ReturnSelectedDirectory
            FunctionsTreeView
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFiles_Click()", Err.Description, Err.Number
End Sub

Public Sub ShowCDSpeed()
On Local Error GoTo ErrHandler
Dim I As Integer, cD As cDrive, n As Integer, p As Integer
If frmMain.cboDrives.ListIndex = -1 Then
    frmMain.cboDrives.ListIndex = 0
End If
I = frmMain.cboDrives.ListIndex + 1
frmMain.tvwFiles.Sorted = False
If I > 0 Then
    Set cD = m_cRip.CDDrive(I)
    If (cD.IsUnitReady) Then
        n = cD.CDSpeed
        Do Until p = n
            tvwFiles.Nodes.Add , , , Str(p + 1), 12
            p = p + 1
        Loop
        If ReturnCDSpeed <> 0 Then
            tvwFiles.Nodes(FindTreeViewIndex(ReturnCDSpeed(), frmMain.tvwFiles)).Selected = True
        Else
            tvwFiles.Nodes(FindTreeViewIndex(cD.CDSpeed, frmMain.tvwFiles)).Selected = True
        End If
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Public Sub ShowCDSpeed()", Err.Description, Err.Number
End Sub

Private Sub tvwFiles_DblClick()
On Local Error GoTo ErrHandler
Dim b As VbMsgBoxResult
Select Case LCase(Right(tvwFiles.SelectedItem.Key, 4))
Case ".m3u"
    b = MsgBox("Add contents of '" & tvwFiles.SelectedItem.Text & "' to playlist?", vbYesNo + vbQuestion)
    If b = vbYes Then
        LoadPlaylist tvwFiles.SelectedItem.Key, True
    End If
Case Else
    OpenMediaFile tvwFiles.SelectedItem.Key, True
End Select
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFiles_DblClick()", Err.Description, Err.Number
End Sub

Private Sub tvwFiles_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
Dim knot As Node
If Button = 1 Then
    Set knot = tvwFiles.HitTest(x, y)
    If Not (knot Is Nothing) Then
        knot.Selected = True
        Set DraggedKnot = knot
    End If
ElseIf Button = 2 Then
    Select Case LCase(Right(tvwFiles.SelectedItem.Text, 4))
    Case ".mp3"
        mnuAddToBatch.Caption = "Add Decode to Que"
    Case ".wav"
        mnuAddToBatch.Caption = "Add Encode MP3 to Que"
    Case ".wma"
        mnuAddToBatch.Caption = "Add Decode to Que"
    Case Else
        mnuAddToBatch.Enabled = False
        mnuAddToBatch.Caption = "Add to Burn Que"
    End Select
    PopupMenu mnutvwFile
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub tvwFunctions_DblClick()
On Local Error GoTo ErrHandler
If Left(tvwFunctions.SelectedItem.FullPath, 15) = "Internet Radio\" Then PlayInternetRadio ctlRadio1, ReturnInternetRadioAddress(tvwFunctions.SelectedItem.Text)
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFunctions_DblClick()", Err.Description, Err.Number
End Sub

Private Sub tvwFunctions_KeyPress(KeyAscii As Integer)
On Local Error GoTo ErrHandler
If KeyAscii = 27 Then Unload Me
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFunctions_KeyPress(KeyAscii As Integer)", Err.Description, Err.Number
End Sub

Private Sub tvwFunctions_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 2 Then
    Select Case LCase(tvwFunctions.SelectedItem.Parent.Text)
    Case "internet radio"
        PopupMenu mnuRadio
    End Select
ElseIf Button = 1 Then
End If
Exit Sub
ErrHandler:
    Err.Clear
End Sub

Private Sub tvwFunctions_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 1 Then FunctionsTreeView
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwFunctions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub tvwPlaylist_AfterLabelEdit(Cancel As Integer, NewString As String)
On Local Error GoTo ErrHandler
TreeViewLabelEdit tvwPlaylist, lLabelEdit, NewString
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwPlaylist_AfterLabelEdit(Cancel As Integer, NewString As String)", Err.Description, Err.Number
End Sub

Private Sub tvwPlaylist_BeforeLabelEdit(Cancel As Integer)
On Local Error GoTo ErrHandler
lLabelEdit = tvwPlaylist.SelectedItem.Text
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwPlaylist_BeforeLabelEdit(Cancel As Integer)", Err.Description, Err.Number
End Sub

Private Sub tvwPlaylist_DblClick()
On Local Error GoTo ErrHandler
OpenMediaFile tvwPlaylist.SelectedItem.Key, True
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwPlaylist_DblClick()", Err.Description, Err.Number
End Sub

Private Sub tvwPlaylist_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Local Error GoTo ErrHandler
If Button = 2 Then
    If LCase(Right(frmMain.tvwPlaylist.SelectedItem.Key, 4)) = ".mp3" Then
        If DoesFileExist(frmMain.tvwPlaylist.SelectedItem.Key) = True Then
            PopupMenu mnutvwPlaylistMenu
        End If
    ElseIf LCase(Right(frmMain.tvwPlaylist.SelectedItem.Key, 4)) = ".wav" Then
        If DoesFileExist(frmMain.tvwPlaylist.SelectedItem.Key) = True Then
            PopupMenu mnuWaveFileMaint
        End If
    Else
        PopupMenu mnutvwPlaylistMenu2
    End If
End If
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub wskFreeDB_Close()
On Local Error GoTo ErrHandler
wskFreeDB.Close
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub tvwPlaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)", Err.Description, Err.Number
End Sub

Private Sub wskFreeDB_Connect()
On Local Error GoTo ErrHandler
wskFreeDB.SendData "cddb hello guidex@tnexgen.com " & wskFreeDB.LocalHostName & " Audiogen 2." & App.Minor & vbCrLf: DoEvents
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub wskFreeDB_Connect()", Err.Description, Err.Number
End Sub

Private Sub wskFreeDB_DataArrival(ByVal bytesTotal As Long)
On Local Error GoTo ErrHandler
Dim msg As String
wskFreeDB.GetData msg, vbString
ProcessFreeDBString msg, wskFreeDB
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub wskFreeDB_DataArrival(ByVal bytesTotal As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_BlockProgress(ByVal nCurrentBlock As Long, ByVal nTotalBlocks As Long)
On Local Error GoTo ErrHandler
DoEvents
If lBurnTrack + 1 <> 0 Then lvwBurn.ListItems(lBurnTrack + 1).SubItems(1) = Int(Trim(nCurrentBlock) * 100 / nTotalBlocks) & "%"
prgBurn.Value = nCurrentBlock
prgBurn.Max = nTotalBlocks
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_BlockProgress(ByVal nCurrentBlock As Long, ByVal nTotalBlocks As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_BurnComplete(ByVal Status As Long)
On Local Error GoTo ErrHandler
lvwBurn.ListItems.Clear
Set frmMain.cmdBurn.PictureNormal = imgBurn1.Picture
Set frmMain.cmdBurn.PictureOver = imgBurn2.Picture
showStatus "Burn Complete"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_BurnComplete(ByVal status As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_ClosingDisc(ByVal nEstimatedSeconds As Long)
On Local Error GoTo ErrHandler
showStatus "Closing Disc: " & nEstimatedSeconds
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_ClosingDisc(ByVal nEstimatedSeconds As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_EraseComplete(ByVal Status As Long)
On Local Error GoTo ErrHandler
showStatus "Erase Complete"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_EraseComplete(ByVal status As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_PnPActivity()
On Local Error GoTo ErrHandler
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_PnPActivity()", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_PreparingBurn(ByVal nEstimatedSeconds As Long)
On Local Error GoTo ErrHandler
showStatus "Burning in: " & nEstimatedSeconds & " second(s)"
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_PreparingBurn(ByVal nEstimatedSeconds As Long)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_QueryCancel(bCancel As Boolean)
On Local Error GoTo ErrHandler
DoEvents
bCancel = m_bCancel
Exit Sub
ErrHandler:
    ProcessRuntimeError "Private Sub m_cDiscMaster_QueryCancel(bCancel As Boolean)", Err.Description, Err.Number
End Sub

Private Sub m_cDiscMaster_TrackProgress(ByVal nCurrentTrack As Long, ByVal nTotalTracks As Long)
On Local Error GoTo ErrHandler
If (nCurrentTrack = nTotalTracks) Then
    showStatus "Completed burning tracks."
Else
    Dim I As Integer
    If nCurrentTrack + 1 <> 0 Then
        If LCase(lvwBurn.ListItems(nCurrentTrack + 1).Text) <> "burning" Then lvwBurn.ListItems(nCurrentTrack + 1).Text = "Burning"
    End If
    showStatus "Burning " & Trim(Str(nCurrentTrack + 1)) & " of " & Trim(Str(nTotalTracks))
    m_CurrentTrack = nCurrentTrack
    lBurnTrack = nCurrentTrack
    If lBurnTrack > 1 Then
        For I = 1 To nCurrentTrack
            If LCase(lvwBurn.ListItems(I).Text) <> "burned" Then lvwBurn.ListItems(I).Text = "Burned"
        Next I
    End If
End If
Exit Sub
ErrHandler:
End Sub
