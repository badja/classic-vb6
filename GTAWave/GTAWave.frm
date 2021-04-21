VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGTAWave 
   Caption         =   "GTA Wave"
   ClientHeight    =   5550
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11040
   Icon            =   "GTAWave.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   11040
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "ilsToolbar"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   25
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Description     =   "Open File"
            Object.ToolTipText     =   "Open GTA Sound File"
            Object.Tag             =   ""
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "rungta"
            Description     =   "Run GTA"
            Object.ToolTipText     =   "Run GTA"
            Object.Tag             =   ""
            ImageKey        =   "rungta"
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "rungta2"
            Description     =   "Run GTA2"
            Object.ToolTipText     =   "Run GTA2"
            Object.Tag             =   ""
            ImageKey        =   "rungta2"
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "edit"
            Description     =   "Open Sound"
            Object.ToolTipText     =   "Open Sound"
            Object.Tag             =   ""
            ImageKey        =   "edit"
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "play"
            Description     =   "Play"
            Object.ToolTipText     =   "Play Sound"
            Object.Tag             =   ""
            ImageKey        =   "play"
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "playloop"
            Description     =   "Play Loop"
            Object.ToolTipText     =   "Play Loop"
            Object.Tag             =   ""
            ImageKey        =   "loop"
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "clear"
            Description     =   "Clear"
            Object.ToolTipText     =   "Clear Sound"
            Object.Tag             =   ""
            ImageKey        =   "clear"
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "pitch"
            Description     =   "Pitch"
            Object.ToolTipText     =   "Change Pitch"
            Object.Tag             =   ""
            ImageKey        =   "pitch"
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "variation"
            Description     =   "Pitch Variation"
            Object.ToolTipText     =   "Change Pitch Variation Range - GTA2 only"
            Object.Tag             =   ""
            ImageKey        =   "variation"
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "looppos"
            Description     =   "Loop Start/End Points"
            Object.ToolTipText     =   "Change Loop Start/End Points - GTA2 only"
            Object.Tag             =   ""
            ImageKey        =   "looppos"
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "import"
            Description     =   "Import"
            Object.ToolTipText     =   "Import Sound"
            Object.Tag             =   ""
            ImageKey        =   "import"
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "export"
            Description     =   "Export"
            Object.ToolTipText     =   "Export Sound"
            Object.Tag             =   ""
            ImageKey        =   "export"
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "restore"
            Description     =   "Restore"
            Object.ToolTipText     =   "Restore Sound"
            Object.Tag             =   ""
            ImageKey        =   "restore"
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "restoreall"
            Description     =   "Restore All"
            Object.ToolTipText     =   "Restore All Sounds"
            Object.Tag             =   ""
            ImageKey        =   "restoreall"
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "autoplay"
            Description     =   "Auto Play"
            Object.ToolTipText     =   "Auto Play - Play sounds when selected"
            Object.Tag             =   ""
            ImageKey        =   "autoplay"
            Value           =   1
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "random"
            Description     =   "Random Variation"
            Object.ToolTipText     =   "Random Pitch Variation - GTA2 only"
            Object.Tag             =   ""
            ImageKey        =   "random"
            Value           =   1
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "synchronous"
            Description     =   "Synchronous"
            Object.ToolTipText     =   "Synchronous - Pause while playing sounds"
            Object.Tag             =   ""
            ImageKey        =   "synchronous"
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "cutoff"
            Description     =   "Cut Off"
            Object.ToolTipText     =   "Cut Off - Don't wait for current sound to finish"
            Object.Tag             =   ""
            ImageKey        =   "cutoff"
            Value           =   1
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "stop"
            Description     =   "Stop Playing"
            Object.ToolTipText     =   "Stop Playing"
            Object.Tag             =   ""
            ImageKey        =   "stop"
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrExternalEdit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10440
      Top             =   4800
   End
   Begin ComctlLib.ListView lvwSounds 
      Height          =   4815
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   8493
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      _Version        =   327682
      Icons           =   "ilsIcons"
      SmallIcons      =   "ilsSmallIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      OLEDropMode     =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Description"
         Object.Width           =   5743
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Sample Rate"
         Object.Width           =   1428
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Variation"
         Object.Width           =   1190
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Loop Start"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Loop End"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Bits"
         Object.Width           =   264
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Channels"
         Object.Width           =   979
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Index"
         Object.Width           =   502
      EndProperty
   End
   Begin ComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   5295
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2963
            MinWidth        =   2963
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   2267
            MinWidth        =   2267
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   13679
            TextSave        =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgImport 
      Left            =   10440
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Import"
      Filter          =   "Sound (*.wav)|*.wav|All files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   10440
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Export"
      Filter          =   "Sound (*.wav)|*.wav|All files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   10440
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "GTA Sounds (*.sdt)|*.sdt"
   End
   Begin ComctlLib.ImageList ilsIcons 
      Left            =   10440
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0442
            Key             =   "bigSound"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilsSmallIcons 
      Left            =   10440
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":075C
            Key             =   "smallSound"
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ilsToolbar 
      Left            =   10440
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   19
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0A76
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0B88
            Key             =   "import"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0C9A
            Key             =   "export"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0DAC
            Key             =   "play"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0EBE
            Key             =   "edit"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":0FD0
            Key             =   "clear"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":10E2
            Key             =   "autoplay"
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":11F4
            Key             =   "synchronous"
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":1306
            Key             =   "loop"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":1418
            Key             =   "cutoff"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":152A
            Key             =   "rungta"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":187C
            Key             =   "pitch"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":198E
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":1AA0
            Key             =   "restore"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":1DF2
            Key             =   "restoreall"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":1F04
            Key             =   "variation"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":2016
            Key             =   "random"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":2128
            Key             =   "looppos"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "GTAWave.frx":223A
            Key             =   "rungta2"
         EndProperty
      EndProperty
   End
   Begin VB.OLE oleEdit 
      Height          =   495
      Left            =   10440
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuFileSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileBackup 
         Caption         =   "GTA &Backup Wizard..."
      End
      Begin VB.Menu mnuFileBackup2 
         Caption         =   "GTA2 B&ackup Wizard..."
      End
      Begin VB.Menu mnuFileSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditRunGTA 
         Caption         =   "&Run GTA"
      End
      Begin VB.Menu mnuEditRunGTA2 
         Caption         =   "R&un GTA2"
      End
      Begin VB.Menu mnuEditSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Enabled         =   0   'False
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvert 
         Caption         =   "&Invert Selection"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuEditSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuSound 
      Caption         =   "&Sound"
      Begin VB.Menu mnuSoundOpen 
         Caption         =   "&Open"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundPlay 
         Caption         =   "&Play"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundPlayLoop 
         Caption         =   "Play &Loop"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundSeparator3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSoundClear 
         Caption         =   "&Clear"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundPitch 
         Caption         =   "Pi&tch..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundVariation 
         Caption         =   "Pitch &Variation Range..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundLoopPos 
         Caption         =   "Loop Start/End Poi&nts..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSoundImport 
         Caption         =   "&Import..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundExport 
         Caption         =   "&Export..."
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSoundRestore 
         Caption         =   "&Restore"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSoundRestoreAll 
         Caption         =   "Restore &All"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuPlay 
      Caption         =   "&Play"
      Begin VB.Menu mnuPlayAutoPlay 
         Caption         =   "&Auto Play"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlayRandom 
         Caption         =   "&Random Pitch Variation"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlaySeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlaySynchronous 
         Caption         =   "&Synchronous"
      End
      Begin VB.Menu mnuPlayCutOff 
         Caption         =   "&Cut Off"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuPlaySeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlayStop 
         Caption         =   "S&top Playing"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpReadme 
         Caption         =   "&View ReadMe.txt"
      End
      Begin VB.Menu mnuHelpSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
End
Attribute VB_Name = "frmGTAWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private blnChanged As Boolean
Private blnDontRepeat As Boolean
Private blnReOpen As Boolean
Private blnReplaceError As Boolean
Private blnReplaceWarn As Boolean
Private itmEdit As ListItem
Private intLastSort As Integer
Private sngOldWidth As Single
Private strBackupFile As String
Private strEditFile As String
Private strTempPath As String

Private Sub ReplaceSound(itmX As ListItem, strSound As String, strSampleRate As String)
    Dim strFileStem As String
    Dim intResponse As Integer
    Dim intIndex As Integer, intLastIndex As Integer
    Dim lngBegin As Long, lngLength As Long
    Dim lngLast As Long, lngLastLength As Long
    Dim strInfoFile As String, strRawFile As String
    Dim strBefore As String, strAfter As String
    Dim intI As Integer
    Dim lngCurrent As Long
    
    On Error GoTo ErrorHandler
    strFileStem = Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4)
    If Not blnReplaceWarn And strBackupFile = "" Then
        If intGTAVersion = 1 Then
            intResponse = MsgBox("WARNING! The backups for this file could not be found. Changing a sound will overwrite it permanently." & strNewLine & strNewLine & "You MUST run the GTA Backup Wizard from the File menu if you want to restore the original GTA sounds. You should run this Wizard even if you have already made backups." & strNewLine & strNewLine & "Press OK if you still want to modify the sound.", vbOKCancel + vbExclamation + vbDefaultButton2)
        Else
            intResponse = MsgBox("WARNING! The backups for this file could not be found. Changing a sound will overwrite it permanently." & strNewLine & strNewLine & "You MUST run the GTA2 Backup Wizard from the File menu if you want to restore the original GTA2 sounds. You should run this Wizard even if you have already made backups." & strNewLine & strNewLine & "Press OK if you still want to modify the sound.", vbOKCancel + vbExclamation + vbDefaultButton2)
        End If
    End If
    
    If intResponse = vbOK Or blnReplaceWarn Or strBackupFile <> "" Then
        Screen.MousePointer = vbHourglass
        blnReplaceWarn = True
        intIndex = itmX.Tag
        intLastIndex = lvwSounds.ListItems.Count - 1
        lngBegin = AddHex(Left(strInfo(intIndex), 4))
        lngLength = AddHex(Mid(strInfo(intIndex), 5, 4))
        lngLast = AddHex(Left(strInfo(intLastIndex), 4))
        lngLastLength = AddHex(Mid(strInfo(intLastIndex), 5, 4))
        strInfoFile = dlgOpen.filename
        strRawFile = Left(dlgOpen.filename, Len(dlgOpen.filename) - 4) & ".raw"
        
        If Dir(strRawFile) = "" Then
            MsgBox "Cannot find '" & strRawFile & "'", vbCritical
        Else
            If GetAttr(strInfoFile) And vbReadOnly Or GetAttr(strRawFile) And vbReadOnly Then
                Screen.MousePointer = vbDefault
                If MsgBox("'" & dlgOpen.FileTitle & "' and/or '" & Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4) & ".RAW' are read-only. To continue, they must be made writable." & strNewLine & strNewLine & "Do you want to make them writable?", vbYesNo Or vbQuestion) = vbYes Then
                    Screen.MousePointer = vbHourglass
                    SetAttr strInfoFile, vbNormal
                    SetAttr strRawFile, vbNormal
                Else
                    blnReplaceError = True
                    Exit Sub
                End If
            End If
        
            Open strRawFile For Binary Access Read As #1
            strBefore = Space(lngBegin)
            Get #1, , strBefore
            strAfter = Space(lngLast + lngLastLength - lngBegin - lngLength)
            Get #1, lngBegin + lngLength + 1, strAfter
            Close #1
            
            SafeKill strRawFile
            Open strRawFile For Binary Access Write As #1
            Put #1, , strBefore
            Put #1, , strSound
            Put #1, , strAfter
            Close #1
            
            Mid(strInfo(intIndex), 5, 4) = MakeHexString(Len(strSound), 4)
            itmX.SubItems(1) = Format(AddHex(Mid(strInfo(intIndex), 5, 4)), "#,0") & " bytes"
            If strSampleRate <> "" Then Mid(strInfo(intIndex), 9, 4) = strSampleRate
            itmX.SubItems(2) = Format(AddHex(Mid(strInfo(intIndex), 9, 4)), "#,0") & " Hz"
            
            If intGTAVersion = 2 Then
                If AddHex(Mid(strInfo(intIndex), 13, 4)) >= AddHex(Mid(strInfo(intIndex), 9, 4)) Then
                    Mid(strInfo(intIndex), 13, 4) = String(4, Chr(0))
                    itmX.SubItems(3) = "± 0 Hz"
                End If
                If AddHex(Mid(strInfo(intIndex), 17, 4)) > AddHex(Mid(strInfo(intIndex), 5, 4)) Then
                    Mid(strInfo(intIndex), 17, 4) = String(4, Chr(0))
                    itmX.SubItems(4) = "0 bytes"
                End If
                If AddHex(Mid(strInfo(intIndex), 21, 4)) > AddHex(Mid(strInfo(intIndex), 5, 4)) Then
                    Mid(strInfo(intIndex), 21, 4) = String(4, Chr(255))
                    itmX.SubItems(5) = "end of sound"
                End If
            End If
            
            Open strInfoFile For Binary Access Write As #1
            Put #1, 12 * intGTAVersion * intIndex + 1, strInfo(intIndex)
            lngCurrent = AddHex(Left(strInfo(intIndex), 4)) + Len(strSound)
            
            For intI = intIndex + 1 To intLastIndex
                Mid(strInfo(intI), 1, 4) = MakeHexString(lngCurrent, 4)
                Put #1, 12 * intGTAVersion * intI + 1, strInfo(intI)
                lngCurrent = lngCurrent + AddHex(Mid(strInfo(intI), 5, 4))
            Next intI
            
            Close #1
            UpdateSize
            UpdateSpace
            Screen.MousePointer = vbDefault
        End If
    End If
    
    Exit Sub

ErrorHandler:
    Screen.MousePointer = vbDefault
    Close #1
    If Err = 75 Then MsgBox "The file is read-only", vbCritical Else MsgBox "File error", vbCritical
    blnReplaceError = True
End Sub

Private Sub Form_Load()
    Initialise
    If GetSetting("GTA Wave", "Options", "RunWizard", 1) = 1 Then mnuFileBackup_Click
    If GetSetting("GTA Wave", "Options", "RunWizard2", 1) = 1 Then mnuFileBackup2_Click
End Sub

Private Sub Initialise()
    Dim lngLength As Long
    
    LoadReg
    dlgOpen.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgImport.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgExport.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    blnLastConst = True
    blnLastScaleVar = True
    lngPlayLoopStart = -1
    sngScale = 100
    strNewLine = Chr(13) & Chr(10)
    strTempPath = Space(256)
    lngLength = GetTempPath(256, strTempPath)
    strTempPath = Left(strTempPath, lngLength)
    If Right(strTempPath, 1) <> "\" Then strTempPath = strTempPath & "\"
    strEditFile = strTempPath & "~GTAWave.wav"
    strTempFile = strTempPath & "~GTAWave.tmp"
    SafeKill strTempFile
    SafeKill strEditFile
End Sub

Private Sub Form_Resize()
    Dim intI As Integer
    Dim intOffset As Integer
    Dim sngWidth As Single
    Dim sngTHeight As Single
    
    If WindowState <> vbMinimized Then
    
        If intGTAVersion = 1 Then intOffset = 2010 Else intOffset = 2865
        For intI = 2 To 9
            sngWidth = sngWidth + lvwSounds.ColumnHeaders(intI).Width
        Next intI
        
        If Width < sngOldWidth And Width - 120 - sngWidth - intOffset >= 0 Then lvwSounds.ColumnHeaders(1).Width = Width - 120 - sngWidth - intOffset
        
        If tlbToolbar.Visible Then sngTHeight = tlbToolbar.Height Else sngTHeight = 0
        
        If Height - sngTHeight - 1005 < 0 Then
            lvwSounds.Height = 0
            Height = sngTHeight + 1005
        Else
            lvwSounds.Move 0, sngTHeight + 60, Width - 120, Height - sngTHeight - 1005
        End If
        
        If Width >= sngOldWidth And Width - 120 - sngWidth - intOffset >= 0 Then lvwSounds.ColumnHeaders(1).Width = Width - 120 - sngWidth - intOffset
        sngOldWidth = Width
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPlaying
    SafeKill strTempFile
    SafeKill strEditFile
    SaveReg
    End
End Sub

Private Sub UpdateSize()
    Dim lngSize As Long
    Dim itmX As ListItem
    
    If CountSelected = 0 Then
        lngSize = FileLen(Left(dlgOpen.filename, Len(dlgOpen.filename) - 4) & ".raw")
    Else
    
        For Each itmX In lvwSounds.ListItems
            If itmX.Selected Then lngSize = lngSize + AddHex(Mid(strInfo(itmX.Tag), 5, 4))
        Next itmX
    
    End If
        
    staStatus.Panels(2).Text = Format(lngSize, "#,0") & " bytes"
End Sub

Private Sub UpdateSpace()
    Dim lngLength As Long
    
    lngLength = FileLen(Left(dlgOpen.filename, Len(dlgOpen.filename) - 4) & ".raw")
    
    If intGTAVersion = 1 Then
        If LCase(Left(dlgOpen.FileTitle, 5)) = "level" Then
            
            If lngLength <= 1048576 Then
                staStatus.Panels(3).Text = Format(1048576 - lngLength, "#,0") & " bytes under limit"
            Else
                staStatus.Panels(3).Text = Format(lngLength - 1048576, "#,0") & " bytes over limit"
            End If
        
        Else
            staStatus.Panels(3).Text = "no limit"
        End If
    Else
        If lngLength <= 6100000 Then
            staStatus.Panels(3).Text = Format(6100000 - lngLength, "#,0") & " bytes under limit"
        Else
            staStatus.Panels(3).Text = Format(lngLength - 6100000, "#,0") & " bytes over limit"
        End If
    End If
End Sub

Private Sub lvwSounds_BeforeLabelEdit(Cancel As Integer)
    blnDontRepeat = True
    Cancel = True
End Sub

Private Sub lvwSounds_Click()
    UpdateAll
End Sub

Private Sub lvwSounds_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim itmX As ListItem
    Dim lngEndPos As Long
    
    For Each itmX In lvwSounds.ListItems
        
        Select Case ColumnHeader.Index
            Case 2
                itmX.SubItems(1) = Format(AddHex(Mid(strInfo(itmX.Tag), 5, 4)), "0000000")
            Case 3
                itmX.SubItems(2) = Format(AddHex(Mid(strInfo(itmX.Tag), 9, 4)), "0000000")
            Case 4
                If intGTAVersion = 2 Then itmX.SubItems(3) = Format(AddHex(Mid(strInfo(itmX.Tag), 13, 4)), "0000000")
            Case 5
                If intGTAVersion = 2 Then itmX.SubItems(4) = Format(AddHex(Mid(strInfo(itmX.Tag), 17, 4)), "0000000")
            Case 6
                If intGTAVersion = 2 Then itmX.SubItems(5) = Format(AddHex(Mid(strInfo(itmX.Tag), 21, 4)), "0000000")
            Case 7
                itmX.SubItems(6) = Format(itmX.SubItems(6), "00")
            Case 9
                itmX.SubItems(8) = Format(itmX.Tag + 1, "000")
        End Select
    
    Next itmX
    
    lvwSounds.SortKey = ColumnHeader.Index - 1
    
    If intLastSort = ColumnHeader.Index Then
        lvwSounds.SortOrder = -Not -lvwSounds.SortOrder
    Else
        lvwSounds.SortOrder = lvwAscending
    End If
    
    lvwSounds.Sorted = True

    For Each itmX In lvwSounds.ListItems
        
        Select Case ColumnHeader.Index
            Case 2
                itmX.SubItems(1) = Format(AddHex(Mid(strInfo(itmX.Tag), 5, 4)), "#,0") & " bytes"
            Case 3
                itmX.SubItems(2) = Format(AddHex(Mid(strInfo(itmX.Tag), 9, 4)), "#,0") & " Hz"
            Case 4
                If intGTAVersion = 2 Then itmX.SubItems(3) = "± " & Format(AddHex(Mid(strInfo(itmX.Tag), 13, 4)), "#,0") & " Hz"
            Case 5
                If intGTAVersion = 2 Then itmX.SubItems(4) = Format(AddHex(Mid(strInfo(itmX.Tag), 17, 4)), "#,0") & " bytes"
            Case 6
                If intGTAVersion = 2 Then
                    lngEndPos = AddHex(Mid(strInfo(itmX.Tag), 21, 4))
                    If lngEndPos = -1 Then
                        itmX.SubItems(5) = "end of sound"
                    Else
                        itmX.SubItems(5) = Format(lngEndPos, "#,0") & " bytes"
                    End If
                End If
            Case 7
                itmX.SubItems(6) = Format(itmX.SubItems(6), "0")
            Case 9
                itmX.SubItems(8) = Format(itmX.Tag + 1, "0")
        End Select
    
    Next itmX
    
    lvwSounds.Sorted = False
    intLastSort = ColumnHeader.Index
End Sub

Private Sub lvwSounds_DblClick()
    Select Case GetSetting("GTA Wave", "Options", "DoubleClick", 1)
        Case 1
            If mnuSoundOpen.Enabled Then mnuSoundOpen_Click
        Case 2
            If mnuSoundPlay.Enabled Then mnuSoundPlay_Click
        Case 3
            If mnuSoundPlayLoop.Enabled Then mnuSoundPlayLoop_Click
    End Select
End Sub

Private Sub lvwSounds_ItemClick(ByVal Item As ComctlLib.ListItem)
    If blnDontRepeat Then
        blnDontRepeat = False
    Else
        UpdateAll
        If mnuPlayAutoPlay.Checked And CountSelected = 1 And Item Is lvwSounds.SelectedItem Then mnuSoundPlay_Click
    End If
End Sub

Private Sub lvwSounds_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 And blnEditing And -GetSetting("GTA Wave", "Options", "Editor", 0) Then
        tmrExternalEdit.Enabled = False
        SafeKill strEditFile
        blnEditing = False
        mnuFileOpen.Enabled = True
        tlbToolbar.Buttons("open").Enabled = True
        mnuFileClose.Enabled = True
        mnuSoundRestoreAll.Enabled = True
        tlbToolbar.Buttons("restoreall").Enabled = True
        UpdateSpace
        UpdateAll
    End If
End Sub

Private Sub lvwSounds_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CountSelected > 0 And Button = vbLeftButton Then lvwSounds.OLEDrag
End Sub

Private Sub lvwSounds_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        
        If CountSelected = 0 Then
            PopupMenu mnuPlay, vbPopupMenuRightButton
        Else
            Select Case GetSetting("GTA Wave", "Options", "DoubleClick", 1)
                Case 0
                    PopupMenu mnuSound, vbPopupMenuRightButton
                Case 1
                    PopupMenu mnuSound, vbPopupMenuRightButton, , , mnuSoundOpen
                Case 2
                    PopupMenu mnuSound, vbPopupMenuRightButton, , , mnuSoundPlay
                Case 3
                    PopupMenu mnuSound, vbPopupMenuRightButton, , , mnuSoundPlayLoop
            End Select
        End If
    
    End If
End Sub

Private Sub lvwSounds_OLECompleteDrag(Effect As Long)
    Dim itmX As ListItem
    Dim strFileName As String
    
    If Effect = vbDropEffectNone Then
        
        For Each itmX In lvwSounds.ListItems
            
            If itmX.Selected Then
                strFileName = strTempPath & itmX.Text & ".wav"
                SafeKill strFileName
            End If
            
        Next itmX
        
    End If
End Sub

Private Sub lvwSounds_OLEDragDrop(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim itmX As ListItem
    
    Set itmX = lvwSounds.HitTest(x, y)
    If Not (itmX Is Nothing) Then
        If MsgBox("Are you sure you want to replace '" & itmX.Text & "' with '" & Data.Files(1) & "'?", vbYesNo + vbQuestion) = vbYes Then Import Data.Files(1), itmX
    End If
    Set lvwSounds.DropHighlight = Nothing
End Sub

Private Sub lvwSounds_OLEDragOver(Data As ComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    Dim itmTarget As ListItem
    
    Set itmTarget = lvwSounds.HitTest(x, y)
    
    If Data.GetFormat(vbCFFiles) Then
        
        If Data.Files.Count = 1 And Not blnEditing Then
            Set lvwSounds.DropHighlight = itmTarget
            If itmTarget Is Nothing Then
                Effect = vbDropEffectNone
            Else
                itmTarget.EnsureVisible
                Effect = vbDropEffectCopy And Effect
            End If
        Else
            Effect = vbDropEffectNone
        End If
        
    Else
        Effect = vbDropEffectNone
    End If
End Sub

Private Sub lvwSounds_OLEStartDrag(Data As ComctlLib.DataObject, AllowedEffects As Long)
    Dim strStem As String
    Dim itmX As ListItem
    Dim strFileName As String
    
    On Error GoTo ErrorHandler
    AllowedEffects = vbDropEffectMove
    Data.SetData , vbCFFiles
    
    For Each itmX In lvwSounds.ListItems
        
        If itmX.Selected Then
            strFileName = strTempPath & SafeName(itmX.Text) & ".wav"
            CreateFile strFileName, itmX.Tag, dlgOpen.filename, False, False
            Data.Files.Add (strFileName)
        End If
        
    Next itmX
    
    Exit Sub

ErrorHandler:
    MsgBox "File error", vbCritical
End Sub

Private Sub mnuEditInvert_Click()
    Dim itmX As ListItem
    
    For Each itmX In lvwSounds.ListItems
        itmX.Selected = Not itmX.Selected
    Next itmX

    UpdateAll
End Sub

Private Sub mnuEditOptions_Click()
    frmOptions.Show vbModal
    FindBackup
End Sub

Private Sub mnuEditRunGTA_Click()
    Dim dblDummy As Double
    Dim strProgram As String
    
    On Error GoTo CannotRun
    strProgram = GetSetting("GTA Wave", "Options", "GTAProgFile")
    
    If strProgram = "" Then
        MsgBox "No GTA program file specified. Set the GTA Program File in Options under the Edit menu.", vbExclamation
    Else
        ChDir (GetPath(strProgram))
        dblDummy = Shell(strProgram, vbNormalFocus)
    End If
    
    Exit Sub
    
CannotRun:
    MsgBox "Cannot run '" & strProgram & "'. Check the GTA Program File in Options under the Edit menu.", vbExclamation
End Sub

Private Sub mnuEditRunGTA2_Click()
    Dim dblDummy As Double
    Dim strProgram As String
    
    On Error GoTo CannotRun
    strProgram = GetSetting("GTA Wave", "Options", "GTAProgFile2")
    
    If strProgram = "" Then
        MsgBox "No GTA2 program file specified. Set the GTA2 Program File in Options under the Edit menu.", vbExclamation
    Else
        ChDir (GetPath(strProgram))
        dblDummy = Shell(strProgram, vbNormalFocus)
    End If
    
    Exit Sub
    
CannotRun:
    MsgBox "Cannot run '" & strProgram & "'. Check the GTA2 Program File in Options under the Edit menu.", vbExclamation
End Sub

Private Sub mnuEditSelectAll_Click()
    Dim itmX As ListItem
    
    For Each itmX In lvwSounds.ListItems
        itmX.Selected = True
    Next itmX
    
    UpdateAll
End Sub

Private Sub mnuEditToolbar_Click()
    tlbToolbar.Visible = Not tlbToolbar.Visible
    mnuEditToolbar.Checked = Not mnuEditToolbar.Checked
    Form_Resize
End Sub

Private Sub mnuFileBackup2_Click()
    intWizardVersion = 2
    frmBackupWizard.Show vbModal
    FindBackup
End Sub

Private Sub mnuHelpAbout_Click()
    frmAboutBox.Show 1
End Sub

Private Sub mnuFileBackup_Click()
    intWizardVersion = 1
    frmBackupWizard.Show vbModal
    FindBackup
End Sub

Private Sub mnuFileClose_Click()
    CloseFile
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuHelpReadme_Click()
    Dim dblTemp As Double
    Dim strReadMeFile As String
    
    strReadMeFile = App.Path
    If Right(strReadMeFile, 1) <> "\" Then strReadMeFile = strReadMeFile & "\"
    strReadMeFile = strReadMeFile & "ReadMe.txt"
    
    If Dir(strReadMeFile) = "" Then
        MsgBox "The ReadMe.txt file is missing.", vbExclamation
    Else
        dblTemp = Shell("notepad.exe " & strReadMeFile, vbNormalFocus)
    End If
End Sub

Private Sub mnuPlayRandom_Click()
    mnuPlayRandom.Checked = Not mnuPlayRandom.Checked
    If mnuPlayRandom.Checked Then tlbToolbar.Buttons("random").Value = tbrPressed Else tlbToolbar.Buttons("random").Value = tbrUnpressed
End Sub

Private Sub mnuPlayStop_Click()
    StopPlaying
End Sub

Private Sub mnuSoundExport_Click()
    Dim itmX As ListItem
    
    On Error GoTo ErrorHandler
    Set itmX = FindSelected
    dlgExport.filename = SafeName(itmX.Text)
    dlgExport.ShowSave
    dlgExport.InitDir = ""
    CreateFile dlgExport.filename, itmX.Tag, dlgOpen.filename, False, False
    Exit Sub
    
ErrorHandler:
End Sub

Private Sub mnuSoundImport_Click()
    Dim itmX As ListItem
    
    On Error GoTo ErrorHandler
    Set itmX = FindSelected
    dlgImport.ShowOpen
    dlgImport.InitDir = ""
    If MsgBox("Are you sure you want to replace '" & itmX.Text & "' with '" & dlgImport.filename & "'?", vbYesNo + vbQuestion) = vbYes Then Import dlgImport.filename, itmX

ErrorHandler:
End Sub

Private Sub mnuFileOpen_Click()
    Dim strBackup1 As String, strBackup2 As String
    Dim strRawFile As String, strDesFile As String
    Dim blnLevel0 As Boolean
    Dim intI As Integer
    Dim strData As String
    Dim itmX As ListItem
    Dim strDes As String
    Dim lngEndPos As Long
    
    On Error GoTo ErrorHandler
    If blnReOpen Then blnReOpen = False Else dlgOpen.ShowOpen
    dlgOpen.InitDir = ""
    strBackup1 = GetSetting("GTA Wave", "Options", "BackupDir")
    strBackup2 = GetSetting("GTA Wave", "Options", "BackupDir2")
    If Right(strBackup1, 1) <> "\" Then strBackup1 = strBackup1 & "\"
    If Right(strBackup2, 1) <> "\" Then strBackup2 = strBackup2 & "\"
    
    If GetPath(dlgOpen.filename) = strBackup1 Or GetPath(dlgOpen.filename) = strBackup2 Then
        MsgBox "This file is a backup. You must open a file inside GTADATA\AUDIO in your GTA folder, or a file inside DATA\AUDIO in you GTA2 folder." & strNewLine & strNewLine & "You can change where GTA Wave looks for your backups by running the appropriate Backup Wizard from the File menu.", vbCritical
        Exit Sub
    End If
    
    If oleEdit.AppIsRunning Then oleEdit.Close
    strRawFile = Left(dlgOpen.filename, Len(dlgOpen.filename) - 4) & ".raw"
    strDesFile = App.Path
    If Right(strDesFile, 1) <> "\" Then strDesFile = strDesFile & "\"
    
    If LCase(Left(dlgOpen.FileTitle, 5)) = "level" And LCase(dlgOpen.FileTitle) <> "level000.sdt" Then
        strDesFile = strDesFile & "level.sdf"
    Else
        If LCase(dlgOpen.FileTitle) = "bil.sdt" And FileLen(dlgOpen.filename) = 5136 Then
            strDesFile = strDesFile & "bildemo.sdf"
        Else
            strDesFile = strDesFile & Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4) & ".sdf"
        End If
        If LCase(dlgOpen.FileTitle) = "level000.sdt" Then blnLevel0 = True
    End If
    
    If Dir(strRawFile) = "" Then
        MsgBox "Cannot find '" & strRawFile & "'", vbCritical
        CloseFile
    Else
        Open dlgOpen.filename For Binary Access Read As #1
        Open strDesFile For Input Access Read As #2
        
        If LCase(dlgOpen.FileTitle) = "level000.sdt" Or LCase(dlgOpen.FileTitle) = "level001.sdt" Or LCase(dlgOpen.FileTitle) = "level002.sdt" Or LCase(dlgOpen.FileTitle) = "level003.sdt" Or LCase(dlgOpen.FileTitle) = "misbrief.sdt" Or LCase(dlgOpen.FileTitle) = "vocalcom.sdt" Then intGTAVersion = 1 Else intGTAVersion = 2
        
        lvwSounds.ListItems.Clear
        strData = Space(12 * intGTAVersion)
        Get #1, , strData
        
        Do While Not EOF(1)
            ReDim Preserve strInfo(intI)
            strInfo(intI) = strData
            Line Input #2, strDes
            If strDes = "" Then strDes = "Sound " & intI + 1
            Set itmX = lvwSounds.ListItems.Add(, , strDes, "bigSound", "smallSound")
            itmX.SubItems(1) = Format(AddHex(Mid(strInfo(intI), 5, 4)), "#,0") & " bytes"
            itmX.SubItems(2) = Format(AddHex(Mid(strInfo(intI), 9, 4)), "#,0") & " Hz"
            
            If intGTAVersion = 1 Then
                lvwSounds.ColumnHeaders(4).Width = 0
                lvwSounds.ColumnHeaders(5).Width = 0
                lvwSounds.ColumnHeaders(6).Width = 0
            Else
                lvwSounds.ColumnHeaders(4).Width = 674.64
                lvwSounds.ColumnHeaders(5).Width = 899.71
                lvwSounds.ColumnHeaders(6).Width = 899.71
                itmX.SubItems(3) = "± " & Format(AddHex(Mid(strInfo(intI), 13, 4)), "#,0") & " Hz"
                itmX.SubItems(4) = Format(AddHex(Mid(strInfo(intI), 17, 4)), "#,0") & " bytes"
                lngEndPos = AddHex(Mid(strInfo(intI), 21, 4))
                If lngEndPos = -1 Then
                    itmX.SubItems(5) = "end of sound"
                Else
                    itmX.SubItems(5) = Format(lngEndPos, "#,0") & " bytes"
                End If
            End If
            
            If intGTAVersion = 1 Then
                If blnLevel0 Then
                    itmX.SubItems(6) = "16"
                    If intI < 3 Then itmX.SubItems(7) = "stereo" Else itmX.SubItems(7) = "mono"
                Else
                    itmX.SubItems(6) = "8"
                    itmX.SubItems(7) = "mono"
                End If
            Else
                If intI >= 69 And intI <= 136 Then itmX.SubItems(6) = "8" Else itmX.SubItems(6) = "16"
                itmX.SubItems(7) = "mono"
            End If

            itmX.SubItems(8) = intI + 1
            itmX.Tag = intI
            intI = intI + 1
            Get #1, , strData
        Loop
    
        Close
        
        If intI = 0 Then
            MsgBox "Cannot read '" & dlgOpen.filename & "'", vbCritical
            CloseFile
        Else
            Caption = Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4) & " - GTA Wave"
            mnuFileClose.Enabled = True
            mnuEditSelectAll.Enabled = True
            mnuEditInvert.Enabled = True
            mnuSoundRestoreAll.Enabled = True
            tlbToolbar.Buttons("restoreall").Enabled = True
            lvwSounds.Enabled = True
            staStatus.Enabled = True
            staStatus.Panels(1).Text = ""
            UpdateCount
            UpdateSize
            UpdateSpace
            UpdateAll
            blnReplaceWarn = False
            intLastSort = 0
            FindBackup
            Form_Resize
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Select Case Err
        Case 52, 62
            strDes = ""
            Resume Next
        Case 53
            Resume Next
        Case Is <> cdlCancel
            MsgBox "File error", vbCritical
            Close
    End Select
End Sub

Private Sub mnuPlayAutoPlay_Click()
    mnuPlayAutoPlay.Checked = Not mnuPlayAutoPlay.Checked
    If mnuPlayAutoPlay.Checked Then tlbToolbar.Buttons("autoplay").Value = tbrPressed Else tlbToolbar.Buttons("autoplay").Value = tbrUnpressed
End Sub

Private Sub mnuSoundClear_Click()
    Dim intCount As Integer, intResponse As Integer
    Dim itmX As ListItem
    
    intCount = CountSelected
    
    If intCount = 1 Then
        intResponse = MsgBox("Are you sure want to clear '" & FindSelected().Text & "'?", vbYesNo + vbQuestion)
    Else
        intResponse = MsgBox("Are you sure want to clear these " & intCount & " sounds?", vbYesNo + vbQuestion)
    End If
    
    If intResponse = vbYes Then
    
        For Each itmX In lvwSounds.ListItems
            If itmX.Selected Then ReplaceSound itmX, "", ""
            
            If blnReplaceError Then
                blnReplaceError = False
                Exit Sub
            End If
            
        Next itmX
        
    End If
End Sub

Private Sub mnuPlayCutOff_Click()
    mnuPlayCutOff.Checked = Not mnuPlayCutOff.Checked
    If mnuPlayCutOff.Checked Then tlbToolbar.Buttons("cutoff").Value = tbrPressed Else tlbToolbar.Buttons("cutoff").Value = tbrUnpressed
End Sub

Private Sub mnuSoundLoopPos_Click()
    Dim itmX As ListItem
    Dim strInfoFile As String
    Dim strLoopStart As String
    Dim lngLoopEnd As Long
    Dim strLoopEnd As String
    
    On Error GoTo ErrorHandler:
    intPlayIndex = FindSelected.Tag
    lngCurSize = AddHex(Mid(strInfo(intPlayIndex), 5, 4))
    
    If CountSelected = 1 Then
        frmLoopPos.txtLoopStart = AddHex(Mid(strInfo(intPlayIndex), 17, 4))
        lngLoopEnd = AddHex(Mid(strInfo(intPlayIndex), 21, 4))
        If lngLoopEnd = -1 Then
            frmLoopPos.optSetEnd.Value = True
        Else
            frmLoopPos.txtLoopEnd = lngLoopEnd
        End If
    Else
        frmLoopPos.cmdPlay(0).Enabled = False
        frmLoopPos.cmdPlay(1).Enabled = False
        frmLoopPos.cmdStop.Enabled = False
    End If
    
    blnCancel = True
    frmLoopPos.txtLoopStart.SelLength = Len(frmLoopPos.txtLoopStart.Text)
    frmLoopPos.Show vbModal
    
    If Not blnCancel Then
    
        strInfoFile = dlgOpen.filename
        
        If GetAttr(strInfoFile) And vbReadOnly Then
            If MsgBox("'" & dlgOpen.FileTitle & "' is read-only. To continue, it must be made writable." & strNewLine & strNewLine & "Do you want to make it writable?", vbYesNo Or vbQuestion) = vbYes Then
                SetAttr strInfoFile, vbNormal
            Else
                lngPlayLoopStart = -1
                Exit Sub
            End If
        End If
        
        Open strInfoFile For Binary Access Write As #1
        
        For Each itmX In lvwSounds.ListItems
            
            If itmX.Selected Then
                If lngPlayLoopStart > AddHex(Mid(strInfo(itmX.Tag), 5, 4)) Then
                    MsgBox "The new loop start point is too large for sound " & itmX.Tag + 1 & " (" & itmX.Text & "). Value not changed.", vbExclamation
                Else
                    strLoopStart = MakeHexString(lngPlayLoopStart, 4)
                    Put #1, 24 * itmX.Tag + 17, strLoopStart
                    Mid(strInfo(itmX.Tag), 17, 4) = strLoopStart
                    itmX.SubItems(4) = Format(lngPlayLoopStart, "#,0") & " bytes"
                
                    If lngPlayLoopEnd > AddHex(Mid(strInfo(itmX.Tag), 5, 4)) Then
                        MsgBox "The new loop end point is too large for sound " & itmX.Tag + 1 & " (" & itmX.Text & "). Value not changed.", vbExclamation
                    Else
                        strLoopEnd = MakeHexString(lngPlayLoopEnd, 4)
                        Put #1, 24 * itmX.Tag + 21, strLoopEnd
                        Mid(strInfo(itmX.Tag), 21, 4) = strLoopEnd
                        If lngPlayLoopEnd = -1 Then
                            itmX.SubItems(5) = "end of sound"
                        Else
                            itmX.SubItems(5) = Format(lngPlayLoopEnd, "#,0") & " bytes"
                        End If
                    End If
                End If
            End If
            
        Next itmX
        
        Close #1

    End If
    
    lngPlayLoopStart = -1
    Exit Sub
    
ErrorHandler:
    lngPlayLoopStart = -1
    Close #1
    If Err = 75 Then MsgBox "The file is read-only", vbCritical Else MsgBox "File error", vbCritical
End Sub

Private Sub mnuSoundVariation_Click()
    Dim itmX As ListItem
    Dim strInfoFile As String
    Dim lngVariation As Long, strVariation As String
    
    On Error GoTo ErrorHandler:
    intPlayIndex = FindSelected.Tag
    lngCurRate = AddHex(Mid(strInfo(intPlayIndex), 9, 4))
    lngCurVariation = AddHex(Mid(strInfo(intPlayIndex), 13, 4))
    
    If CountSelected = 1 Then
        frmVariation.txtConstant = lngCurVariation
    Else
        frmVariation.cmdPlay(0).Enabled = False
        frmVariation.cmdPlay(1).Enabled = False
        frmVariation.cmdPlay(2).Enabled = False
    End If
    
    frmVariation.txtScale = sngScale
    If blnLastConst Then frmVariation.optConstant.Value = True
    blnCancel = True
    frmVariation.Show vbModal
    
    If Not blnCancel Then
    
        strInfoFile = dlgOpen.filename
        
        If GetAttr(strInfoFile) And vbReadOnly Then
            If MsgBox("'" & dlgOpen.FileTitle & "' is read-only. To continue, it must be made writable." & strNewLine & strNewLine & "Do you want to make it writable?", vbYesNo Or vbQuestion) = vbYes Then
                SetAttr strInfoFile, vbNormal
            Else
                lngPlayVariation = 0
                Exit Sub
            End If
        End If
        
        Open strInfoFile For Binary Access Write As #1
        
        For Each itmX In lvwSounds.ListItems
            
            If itmX.Selected Then
                If blnScale Then lngPlayVariation = sngScale / 100 * AddHex(Mid(strInfo(itmX.Tag), 13, 4))
                
                If lngPlayVariation >= AddHex(Mid(strInfo(itmX.Tag), 9, 4)) Then
                    MsgBox "The new pitch variation range is too large for sound " & itmX.Tag + 1 & " (" & itmX.Text & "). Value not changed.", vbExclamation
                Else
                    strVariation = MakeHexString(lngPlayVariation, 4)
                    Put #1, 24 * itmX.Tag + 13, strVariation
                    Mid(strInfo(itmX.Tag), 13, 4) = strVariation
                    itmX.SubItems(3) = "± " & Format(lngPlayVariation, "#,0") & " Hz"
                End If
            End If
            
        Next itmX
        
        Close #1

    End If
    
    lngPlayVariation = 0
    Exit Sub
    
ErrorHandler:
    lngPlayVariation = 0
    Close #1
    If Err = 75 Then MsgBox "The file is read-only", vbCritical Else MsgBox "File error", vbCritical
End Sub

Private Sub mnuSoundOpen_Click()
    Dim blnExternal As Boolean

    On Error GoTo ErrorHandler

    blnEditing = True
    mnuFileOpen.Enabled = False
    tlbToolbar.Buttons("open").Enabled = False
    mnuFileClose.Enabled = False
    mnuSoundImport.Enabled = False
    tlbToolbar.Buttons("import").Enabled = False
    mnuSoundClear.Enabled = False
    tlbToolbar.Buttons("clear").Enabled = False
    mnuSoundOpen.Enabled = False
    tlbToolbar.Buttons("edit").Enabled = False
    mnuSoundPitch.Enabled = False
    tlbToolbar.Buttons("pitch").Enabled = False
    mnuSoundVariation.Enabled = False
    tlbToolbar.Buttons("variation").Enabled = False
    mnuSoundLoopPos.Enabled = False
    tlbToolbar.Buttons("looppos").Enabled = False
    mnuSoundRestore.Enabled = False
    tlbToolbar.Buttons("restore").Enabled = False
    mnuSoundRestoreAll.Enabled = False
    tlbToolbar.Buttons("restoreall").Enabled = False
    Set itmEdit = FindSelected

    If -GetSetting("GTA Wave", "Options", "Editor", 0) Then
        blnExternal = True
        CreateFile strEditFile, itmEdit.Tag, dlgOpen.filename, False, False
        staStatus.Panels(3).Text = "Sound currently being externally edited. Press Esc to cancel."
        dblExternalTaskID = Shell(GetSetting("GTA Wave", "Options", "EditorProgram", "") & " " & strEditFile, vbNormalFocus)
        varExternalDate = FileDateTime(strEditFile)
        tmrExternalEdit = True
    Else
        oleEdit.HostName = Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4)
        CreateFile strEditFile, itmEdit.Tag, dlgOpen.filename, False, False
        oleEdit.CreateEmbed strEditFile
        oleEdit.DoVerb vbOLEShow
    End If
    
    Exit Sub

ErrorHandler:
    blnEditing = False
    mnuFileOpen.Enabled = True
    tlbToolbar.Buttons("open").Enabled = True
    mnuFileClose.Enabled = True
    mnuSoundRestoreAll.Enabled = True
    tlbToolbar.Buttons("restoreall").Enabled = True
    UpdateSpace
    UpdateAll
    If blnExternal Then
        MsgBox "Cannot open external editor '" & GetSetting("GTA Wave", "Options", "EditorProgram", "") & "'. Set the sound editor in Options under the Edit menu.", vbCritical
    Else
        MsgBox "Cannot open the editing program associated with WAV files. Try reinstalling Sound Recorder by opening Add/Remove Programs in the Control Panel.", vbCritical
    End If
    Exit Sub
End Sub

Private Sub mnuSoundPitch_Click()
    Dim itmX As ListItem
    Dim strInfoFile As String, strSampleRate As String
    Dim lngVariation As Long, strVariation As String
    
    On Error GoTo ErrorHandler:
    intPlayIndex = FindSelected.Tag
    lngCurRate = AddHex(Mid(strInfo(intPlayIndex), 9, 4))
    
    If CountSelected = 1 Then
        frmPitch.txtConstant = lngCurRate
    Else
        frmPitch.cmdPlay.Enabled = False
    End If
    
    frmPitch.txtScale = sngScale
    If blnLastConst Then frmPitch.optConstant.Value = True
    If blnLastScaleVar Then frmPitch.chkVariation.Value = 1 Else frmPitch.chkVariation.Value = 0
    blnCancel = True
    frmPitch.Show vbModal
    
    If Not blnCancel Then
    
        strInfoFile = dlgOpen.filename
        
        If GetAttr(strInfoFile) And vbReadOnly Then
            If MsgBox("'" & dlgOpen.FileTitle & "' is read-only. To continue, it must be made writable." & strNewLine & strNewLine & "Do you want to make it writable?", vbYesNo Or vbQuestion) = vbYes Then
                SetAttr strInfoFile, vbNormal
            Else
                lngPlayRate = 0
                Exit Sub
            End If
        End If
        
        Open strInfoFile For Binary Access Write As #1
        
        For Each itmX In lvwSounds.ListItems
            
            If itmX.Selected Then
                If blnScale Then lngPlayRate = sngScale / 100 * AddHex(Mid(strInfo(itmX.Tag), 9, 4))
                strSampleRate = MakeHexString(lngPlayRate, 4)
                Put #1, 12 * intGTAVersion * itmX.Tag + 9, strSampleRate
                Mid(strInfo(itmX.Tag), 9, 4) = strSampleRate
                itmX.SubItems(2) = Format(lngPlayRate, "#,0") & " Hz"
                
                If blnScale And blnScaleVar And intGTAVersion = 2 Then
                    lngVariation = sngScale / 100 * AddHex(Mid(strInfo(itmX.Tag), 13, 4))
                    strVariation = MakeHexString(lngVariation, 4)
                    Put #1, 24 * itmX.Tag + 13, strVariation
                    Mid(strInfo(itmX.Tag), 13, 4) = strVariation
                    itmX.SubItems(3) = "± " & Format(lngVariation, "#,0") & " Hz"
                End If
            
            End If
            
        Next itmX
        
        Close #1

    End If
    
    lngPlayRate = 0
    Exit Sub
    
ErrorHandler:
    lngPlayRate = 0
    Close #1
    If Err = 75 Then MsgBox "The file is read-only", vbCritical Else MsgBox "File error", vbCritical
End Sub

Private Sub mnuSoundPlay_Click()
    Dim lngSuccess As Long
    Dim lngUFlags As Long
    
    On Error GoTo CannotKill
    
    lngUFlags = SND_NODEFAULT
    If blnLooping Then StopPlaying
    
    If Not mnuPlaySynchronous.Checked Then
        lngUFlags = lngUFlags Or SND_ASYNC
        If mnuPlayCutOff.Checked Then StopPlaying
    End If
    
    If Dir(strTempFile) <> "" Then Kill strTempFile
    CreateFile strTempFile, lvwSounds.SelectedItem.Tag, dlgOpen.filename, False, mnuPlayRandom.Checked
    If mnuPlaySynchronous.Checked Then Screen.MousePointer = vbHourglass
    lngSuccess = sndPlaySound(strTempFile, lngUFlags)
    If mnuPlaySynchronous.Checked Then Screen.MousePointer = vbDefault

CannotKill:
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuPlaySynchronous_Click()
    mnuPlaySynchronous.Checked = Not mnuPlaySynchronous.Checked
    If mnuPlaySynchronous.Checked Then tlbToolbar.Buttons("synchronous").Value = tbrPressed Else tlbToolbar.Buttons("synchronous").Value = tbrUnpressed
    
    If mnuPlaySynchronous.Checked Then
        mnuPlayCutOff.Enabled = False
        tlbToolbar.Buttons("cutoff").Enabled = False
        StopPlaying
    Else
        mnuPlayCutOff.Enabled = True
        tlbToolbar.Buttons("cutoff").Enabled = True
    End If
End Sub

Private Sub mnuSoundPlayLoop_Click()
    Dim lngSuccess As Long
    Dim lngUFlags As Long
    
    On Error GoTo CannotKill
    
    lngUFlags = SND_NODEFAULT Or SND_ASYNC Or SND_LOOP
    If mnuPlayCutOff.Checked Or blnLooping Then StopPlaying
    If Dir(strTempFile) <> "" Then Kill strTempFile
    CreateFile strTempFile, lvwSounds.SelectedItem.Tag, dlgOpen.filename, True, mnuPlayRandom.Checked
    lngSuccess = sndPlaySound(strTempFile, lngUFlags)
    blnLooping = True

CannotKill:
End Sub

Private Sub mnuSoundRestore_Click()
    Dim intCount As Integer, intResponse As Integer
    Dim itmX As ListItem
    Dim strData As String, strTemp As String
    Dim lngEndPos As Long
    
    On Error GoTo ErrorHandler:
    
    If strBackupFile = "" Then
        MsgBox "The backups for this file could not be found. Please run the appropriate Backup Wizard from the File menu.", vbCritical
    Else
        intCount = CountSelected
        
        If intCount = 1 Then
            intResponse = MsgBox("Are you sure want to restore '" & FindSelected().Text & "'?", vbYesNo + vbQuestion)
        Else
            intResponse = MsgBox("Are you sure want to restore these " & intCount & " sounds?", vbYesNo + vbQuestion)
        End If
        
        If intResponse = vbYes Then
            StopPlaying
            strData = Space(12 * intGTAVersion)
            Open strBackupFile & ".sdt" For Binary Access Read As #2
            
            For Each itmX In lvwSounds.ListItems
            
                If itmX.Selected Then
                    Get #2, 12 * intGTAVersion * itmX.Tag + 1, strData
                    strTemp = strInfo(itmX.Tag)
                    strInfo(itmX.Tag) = strData
                    CreateFile strTempFile, itmX.Tag, strBackupFile & ".sdt", False, False
                    If intGTAVersion = 2 Then Mid(strTemp, 13, 12) = Right(strData, 12)
                    strInfo(itmX.Tag) = strTemp
                    Import strTempFile, itmX
                    
                    If intGTAVersion = 2 Then
                        itmX.SubItems(3) = "± " & Format(AddHex(Mid(strTemp, 13, 4)), "#,0") & " Hz"
                        itmX.SubItems(4) = Format(AddHex(Mid(strTemp, 17, 4)), "#,0") & " bytes"
                        lngEndPos = AddHex(Mid(strTemp, 21, 4))
                        If lngEndPos = -1 Then
                            itmX.SubItems(5) = "end of sound"
                        Else
                            itmX.SubItems(5) = Format(lngEndPos, "#,0") & " bytes"
                        End If
                    End If
                End If
                
            Next itmX
            
            Close #2
        
        End If
    
    End If
    
    Exit Sub
    
ErrorHandler:
    Close #2
    MsgBox "File error", vbCritical
End Sub

Private Sub mnuSoundRestoreAll_Click()
    Dim strInfoFile As String, strRawFile As String

    On Error GoTo ErrorHandler:
    
    If strBackupFile = "" Then
        MsgBox "The backups for this file could not be found. Please run the appropriate Backup Wizard from the File menu.", vbCritical
    ElseIf MsgBox("Are you sure want to restore all the sounds?", vbYesNo + vbQuestion) = vbYes Then
        
        strInfoFile = dlgOpen.filename
        strRawFile = Left(dlgOpen.filename, Len(dlgOpen.filename) - 4) & ".raw"
        
        If GetAttr(strInfoFile) And vbReadOnly Or GetAttr(strRawFile) And vbReadOnly Then
            If MsgBox("'" & dlgOpen.FileTitle & "' and/or '" & Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4) & ".RAW' are read-only. To continue, they must be made writable." & strNewLine & strNewLine & "Do you want to make them writable?", vbYesNo Or vbQuestion) = vbYes Then
                SetAttr strInfoFile, vbNormal
                SetAttr strRawFile, vbNormal
            Else
                Exit Sub
            End If
        End If
        
        Screen.MousePointer = vbHourglass
        FileCopy strBackupFile & ".sdt", strInfoFile
        FileCopy strBackupFile & ".raw", strRawFile
        blnReOpen = True
        mnuFileOpen_Click
        Screen.MousePointer = vbDefault
    End If
    
    Exit Sub
    
ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox "File error", vbCritical
End Sub

Private Sub oleEdit_Updated(Code As Integer)
    On Error GoTo ErrorHandler
    If Code = vbOLEChanged Then blnChanged = True
    
    If Code = vbOLEClosed Then
        SafeKill strEditFile
        
        If blnChanged Then
            
            If MsgBox("Do you want to update '" & itmEdit.Text & "' in " & oleEdit.HostName & "?", vbYesNo + vbQuestion) = vbYes Then
                Open strEditFile For Binary Access Write As #1
                oleEdit.SaveToFile 1
                Close #1
                Import strEditFile, itmEdit
                SafeKill strEditFile
            End If
        
        End If
        
        blnChanged = False
        blnEditing = False
        mnuFileOpen.Enabled = True
        tlbToolbar.Buttons("open").Enabled = True
        mnuFileClose.Enabled = True
        mnuSoundRestoreAll.Enabled = True
        tlbToolbar.Buttons("restoreall").Enabled = True
        UpdateAll
    End If
    
    Exit Sub
    
ErrorHandler:
    Close #1
    If Err = 75 Then MsgBox "The file is read-only", vbCritical Else MsgBox "File error", vbCritical
    blnChanged = False
    blnEditing = False
    mnuFileOpen.Enabled = True
    tlbToolbar.Buttons("open").Enabled = True
    mnuFileClose.Enabled = True
    mnuSoundRestoreAll.Enabled = True
    tlbToolbar.Buttons("restoreall").Enabled = True
    UpdateAll
End Sub

Private Sub Import(strFileName As String, itmX As ListItem)
    Dim intIndex As Integer
    Dim strData As String
    Dim lngChannels As Long
    Dim strSampleRate As String
    Dim lngBits As Long
    
    On Error GoTo ErrorHandler
    
    Screen.MousePointer = vbHourglass
    intIndex = itmX.Tag
    
    Open strFileName For Binary Access Read As #1
    
    strData = Space(4)
    Get #1, , strData
    
    Do While strData <> "RIFF"
        If EOF(1) Then Error 3
        Get #1, , strData
    Loop
    
    Get #1, , strData
    Get #1, , strData
    Get #1, , strData
    If strData <> "fmt " Then Error 3
    
    Get #1, , strData
    strData = Space(AddHex(strData))
    Get #1, , strData
    If AddHex(Left(strData, 2)) <> 1 Then Error 3
    lngChannels = AddHex(Mid(strData, 3, 2))
    strSampleRate = Mid(strData, 5, 4)
    lngBits = AddHex(Mid(strData, 15, 2))
    
    If intGTAVersion = 1 Then
        If LCase(dlgOpen.FileTitle) = "level000.sdt" Then
            If lngBits <> 16 Or intIndex <= 2 And lngChannels <> 2 Or intIndex >= 3 And lngChannels <> 1 Then Error 3
        Else
            If lngBits <> 8 Then Error 3
        End If
    Else
        If lngChannels <> 1 Then Error 3
    
        If intIndex >= 69 And intIndex <= 136 Then
            If lngBits <> 8 Then Error 3
        Else
            If lngBits <> 16 Then Error 3
        End If
    End If
    
    strData = Space(4)
    Get #1, , strData
    
    Do While strData <> "data"
        If EOF(1) Then Error 3
        Get #1, , strData
        strData = Space(AddHex(strData))
        Get #1, , strData
        strData = Space(4)
        Get #1, , strData
    Loop
            
    Get #1, , strData
    strData = Space(AddHex(strData))
    Get #1, , strData
    
    Close #1
    Screen.MousePointer = vbDefault
    ReplaceSound itmX, strData, strSampleRate
    Exit Sub
    
ErrorHandler:
    Screen.MousePointer = vbDefault
    MsgBox "File error. Make sure the sound is a " & itmX.SubItems(6) & "-bit " & itmX.SubItems(7) & " WAV file.", vbCritical
    Close #1
End Sub

Private Sub CloseFile()
    Caption = "GTA Wave"
    lvwSounds.ListItems.Clear
    lvwSounds.Enabled = False
    mnuFileClose.Enabled = False
    mnuEditSelectAll.Enabled = False
    mnuEditInvert.Enabled = False
    mnuSoundPlay.Enabled = False
    tlbToolbar.Buttons("play").Enabled = False
    mnuSoundPlayLoop.Enabled = False
    tlbToolbar.Buttons("playloop").Enabled = False
    mnuSoundExport.Enabled = False
    tlbToolbar.Buttons("export").Enabled = False
    mnuSoundImport.Enabled = False
    tlbToolbar.Buttons("import").Enabled = False
    mnuSoundClear.Enabled = False
    tlbToolbar.Buttons("clear").Enabled = False
    mnuSoundOpen.Enabled = False
    tlbToolbar.Buttons("edit").Enabled = False
    mnuSoundPitch.Enabled = False
    tlbToolbar.Buttons("pitch").Enabled = False
    mnuSoundVariation.Enabled = False
    tlbToolbar.Buttons("variation").Enabled = False
    mnuSoundLoopPos.Enabled = False
    tlbToolbar.Buttons("looppos").Enabled = False
    mnuSoundRestore.Enabled = False
    tlbToolbar.Buttons("restore").Enabled = False
    mnuSoundRestoreAll.Enabled = False
    tlbToolbar.Buttons("restoreall").Enabled = False
    staStatus.Enabled = False
    staStatus.Panels(1).Text = ""
    staStatus.Panels(2).Text = ""
    staStatus.Panels(3).Text = ""
    intGTAVersion = 0
End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.Key
        Case "open"
            mnuFileOpen_Click
        Case "rungta"
            mnuEditRunGTA_Click
        Case "rungta2"
            mnuEditRunGTA2_Click
        Case "edit"
            mnuSoundOpen_Click
        Case "play"
            mnuSoundPlay_Click
        Case "playloop"
            mnuSoundPlayLoop_Click
        Case "clear"
            mnuSoundClear_Click
        Case "pitch"
            mnuSoundPitch_Click
        Case "variation"
            mnuSoundVariation_Click
        Case "looppos"
            mnuSoundLoopPos_Click
        Case "import"
            mnuSoundImport_Click
        Case "export"
            mnuSoundExport_Click
        Case "restore"
            mnuSoundRestore_Click
        Case "restoreall"
            mnuSoundRestoreAll_Click
        Case "autoplay"
            mnuPlayAutoPlay_Click
        Case "synchronous"
            mnuPlaySynchronous_Click
        Case "random"
            mnuPlayRandom_Click
        Case "cutoff"
            mnuPlayCutOff_Click
        Case "stop"
            mnuPlayStop_Click
    End Select
End Sub

Private Function CountSelected() As Integer
    Dim itmX As ListItem
    Dim intCount As Integer
    
    For Each itmX In lvwSounds.ListItems
        If itmX.Selected Then intCount = intCount + 1
    Next itmX
    
    CountSelected = intCount
End Function

Private Function FindSelected() As ListItem
    Dim intI As Integer
    
    Do While Not lvwSounds.ListItems(intI + 1).Selected
        intI = intI + 1
    Loop
    
    Set FindSelected = lvwSounds.ListItems(intI + 1)
End Function

Private Function UpdateCount()
    Dim intCount As Integer
    
    intCount = CountSelected
    
    If intCount = 0 Then
        staStatus.Panels(1) = lvwSounds.ListItems.Count & " sound(s)"
    Else
        staStatus.Panels(1) = intCount & " sound(s) selected"
    End If
End Function

Private Sub UpdateAll()
    Dim intCount As Integer
    
    UpdateCount
    UpdateSize
    intCount = CountSelected
    
    If intCount = 1 Then
        mnuSoundPlay.Enabled = True
        tlbToolbar.Buttons("play").Enabled = True
        mnuSoundPlayLoop.Enabled = True
        tlbToolbar.Buttons("playloop").Enabled = True
        mnuSoundExport.Enabled = True
        tlbToolbar.Buttons("export").Enabled = True
        
        If Not blnEditing Then
            mnuSoundImport.Enabled = True
            tlbToolbar.Buttons("import").Enabled = True
            mnuSoundClear.Enabled = True
            tlbToolbar.Buttons("clear").Enabled = True
            mnuSoundOpen.Enabled = True
            tlbToolbar.Buttons("edit").Enabled = True
            mnuSoundPitch.Enabled = True
            tlbToolbar.Buttons("pitch").Enabled = True
            
            If intGTAVersion = 2 Then
                mnuSoundVariation.Enabled = True
                tlbToolbar.Buttons("variation").Enabled = True
                mnuSoundLoopPos.Enabled = True
                tlbToolbar.Buttons("looppos").Enabled = True
            End If
            
            mnuSoundRestore.Enabled = True
            tlbToolbar.Buttons("restore").Enabled = True
        End If
        
    Else
        mnuSoundExport.Enabled = False
        tlbToolbar.Buttons("export").Enabled = False
        mnuSoundImport.Enabled = False
        tlbToolbar.Buttons("import").Enabled = False
        mnuSoundOpen.Enabled = False
        tlbToolbar.Buttons("edit").Enabled = False
    
        If intCount = 0 Then
            mnuSoundPlay.Enabled = False
            tlbToolbar.Buttons("play").Enabled = False
            mnuSoundPlayLoop.Enabled = False
            tlbToolbar.Buttons("playloop").Enabled = False
            mnuSoundClear.Enabled = False
            tlbToolbar.Buttons("clear").Enabled = False
            mnuSoundPitch.Enabled = False
            tlbToolbar.Buttons("pitch").Enabled = False
            mnuSoundVariation.Enabled = False
            tlbToolbar.Buttons("variation").Enabled = False
            mnuSoundLoopPos.Enabled = False
            tlbToolbar.Buttons("looppos").Enabled = False
            mnuSoundRestore.Enabled = False
            tlbToolbar.Buttons("restore").Enabled = False
        Else
            mnuSoundPlay.Enabled = True
            tlbToolbar.Buttons("play").Enabled = True
            mnuSoundPlayLoop.Enabled = True
            tlbToolbar.Buttons("playloop").Enabled = True
            
            If Not blnEditing Then
                mnuSoundClear.Enabled = True
                tlbToolbar.Buttons("clear").Enabled = True
                mnuSoundPitch.Enabled = True
                tlbToolbar.Buttons("pitch").Enabled = True
                
                If intGTAVersion = 2 Then
                    mnuSoundVariation.Enabled = True
                    tlbToolbar.Buttons("variation").Enabled = True
                    mnuSoundLoopPos.Enabled = True
                    tlbToolbar.Buttons("looppos").Enabled = True
                End If
                
                mnuSoundRestore.Enabled = True
                tlbToolbar.Buttons("restore").Enabled = True
            End If
                
        End If
        
    End If
End Sub

Private Function SafeName(strFileName As String) As String
    Dim strReplace As String
    Dim intI As Integer, intPos As Integer
    Dim strNewName As String
    
    strReplace = "\/:*?" & Chr(34) & "<>|----.'---"
    
    For intI = 1 To Len(strFileName)
        intPos = InStr(strReplace, Mid(strFileName, intI, 1))
        
        If intPos < 10 And intPos <> 0 Then
            strNewName = strNewName & Mid(strReplace, intPos + 9, 1)
        Else
            strNewName = strNewName & Mid(strFileName, intI, 1)
        End If
    Next intI
    
    SafeName = strNewName
End Function

Private Sub SaveReg()
    Dim strOpen As String
    Dim strImport As String
    Dim strExport As String
    
    strOpen = GetPath(dlgOpen.filename)
    strImport = GetPath(dlgImport.filename)
    strExport = GetPath(dlgExport.filename)
    If strOpen = "" Then strOpen = dlgOpen.InitDir
    If strImport = "" Then strImport = dlgImport.InitDir
    If strExport = "" Then strExport = dlgExport.InitDir
    SaveSetting "GTA Wave", "Directories", "Open", strOpen
    SaveSetting "GTA Wave", "Directories", "Import", strImport
    SaveSetting "GTA Wave", "Directories", "Export", strExport
    
    SaveSetting "GTA Wave", "Options", "Toolbar", -mnuEditToolbar.Checked
    SaveSetting "GTA Wave", "Play", "AutoPlay", -mnuPlayAutoPlay.Checked
    SaveSetting "GTA Wave", "Play", "Synchronous", -mnuPlaySynchronous.Checked
    SaveSetting "GTA Wave", "Play", "Random", -mnuPlayRandom.Checked
    SaveSetting "GTA Wave", "Play", "CutOff", -mnuPlayCutOff.Checked
    
    If WindowState = vbNormal Then
        SaveSetting "GTA Wave", "Window", "Left", Left
        SaveSetting "GTA Wave", "Window", "Top", Top
        SaveSetting "GTA Wave", "Window", "Width", Width
        SaveSetting "GTA Wave", "Window", "Height", Height
    Else
        DeleteSetting "GTA Wave", "Window", "Left"
        DeleteSetting "GTA Wave", "Window", "Top"
        DeleteSetting "GTA Wave", "Window", "Width"
        DeleteSetting "GTA Wave", "Window", "Height"
    End If
End Sub

Private Sub LoadReg()
    Dim blnToolbar As Integer
    Dim blnAutoPlay As Integer, blnSynchronous As Integer
    Dim blnRandom As Integer, blnCutOff As Integer
    Dim strLeft As String, strTop As String
    Dim strWidth As String, strHeight As String
    
    dlgOpen.InitDir = GetSetting("GTA Wave", "Directories", "Open")
    dlgImport.InitDir = GetSetting("GTA Wave", "Directories", "Import")
    dlgExport.InitDir = GetSetting("GTA Wave", "Directories", "Export")
    
    blnToolbar = -GetSetting("GTA Wave", "Options", "Toolbar", 1)
    mnuEditToolbar.Checked = blnToolbar
    tlbToolbar.Visible = blnToolbar
    
    blnAutoPlay = -GetSetting("GTA Wave", "Play", "AutoPlay", 1)
    blnSynchronous = -GetSetting("GTA Wave", "Play", "Synchronous", 0)
    blnRandom = -GetSetting("GTA Wave", "Play", "Random", 1)
    blnCutOff = -GetSetting("GTA Wave", "Play", "CutOff", 1)
    If blnAutoPlay <> mnuPlayAutoPlay.Checked Then mnuPlayAutoPlay_Click
    If blnSynchronous <> mnuPlaySynchronous.Checked Then mnuPlaySynchronous_Click
    If blnRandom <> mnuPlayRandom.Checked Then mnuPlayRandom_Click
    If blnCutOff <> mnuPlayCutOff.Checked Then mnuPlayCutOff_Click
    
    strLeft = GetSetting("GTA Wave", "Window", "Left")
    strTop = GetSetting("GTA Wave", "Window", "Top")
    strWidth = GetSetting("GTA Wave", "Window", "Width")
    strHeight = GetSetting("GTA Wave", "Window", "Height")
    If strLeft <> "" Then Left = strLeft
    If strTop <> "" Then Top = strTop
    If strWidth <> "" Then Width = strWidth
    If strHeight <> "" Then Height = strHeight
End Sub

Private Sub FindBackup()
    On Error GoTo NoBackup
    
    Dim intI As Integer
    
    If intGTAVersion = 0 Then Exit Sub
    
    If intGTAVersion = 1 Then
        strBackupFile = GetSetting("GTA Wave", "Options", "BackupDir")
    Else
        strBackupFile = GetSetting("GTA Wave", "Options", "BackupDir2")
    End If
    
    If strBackupFile <> "" Then
        If Right(strBackupFile, 1) <> "\" Then strBackupFile = strBackupFile & "\"
        strBackupFile = strBackupFile & Left(dlgOpen.FileTitle, Len(dlgOpen.FileTitle) - 4)
        If Dir(strBackupFile & ".sdt") = "" Or Dir(strBackupFile & ".raw") = "" Then strBackupFile = ""
    End If
    
    Exit Sub
    
NoBackup:
    strBackupFile = ""
End Sub

Private Sub tmrExternalEdit_Timer()
    On Error GoTo ErrorHandler
    
    Dim intResponse As Integer
    Dim blnAlreadyClosed As Boolean
    
    If varExternalDate <> FileDateTime(strEditFile) Then
        AppActivate frmGTAWave.Caption
        intResponse = MsgBox("The sound being externally edited has changed. Import the new sound?" & strNewLine & strNewLine & "Click Yes to import the edited sound" & strNewLine & "Click No to abandon editing the sound" & strNewLine & "Click Cancel to continue editing the sound", vbQuestion + vbYesNoCancel)
        
        If intResponse = vbYes Then
            tmrExternalEdit.Enabled = False
            AppActivate dblExternalTaskID, True
            If Not blnAlreadyClosed Then SendKeys "%{F4}", True
            AppActivate frmGTAWave.Caption
            Import strEditFile, itmEdit
            SafeKill strEditFile
            blnEditing = False
            mnuFileOpen.Enabled = True
            tlbToolbar.Buttons("open").Enabled = True
            mnuFileClose.Enabled = True
            mnuSoundRestoreAll.Enabled = True
            tlbToolbar.Buttons("restoreall").Enabled = True
            UpdateSpace
            UpdateAll
        ElseIf intResponse = vbNo Then
            tmrExternalEdit.Enabled = False
            AppActivate dblExternalTaskID, True
            If Not blnAlreadyClosed Then SendKeys "%{F4}", True
            AppActivate frmGTAWave.Caption
            SafeKill strEditFile
            blnEditing = False
            mnuFileOpen.Enabled = True
            tlbToolbar.Buttons("open").Enabled = True
            mnuFileClose.Enabled = True
            mnuSoundRestoreAll.Enabled = True
            tlbToolbar.Buttons("restoreall").Enabled = True
            UpdateSpace
            UpdateAll
        Else
            varExternalDate = FileDateTime(strEditFile)
            AppActivate dblExternalTaskID
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    If Err = 5 And intResponse <> vbCancel Then
        blnAlreadyClosed = True
        Resume Next
    End If
    
    Select Case Err
        Case 5
            MsgBox "The external editor has already been closed. Sound not changed.", vbExclamation
            SafeKill strEditFile
        Case 75
            MsgBox "The file is read-only", vbCritical
        Case Else
            MsgBox "File error", vbCritical
    End Select
    
    tmrExternalEdit.Enabled = False
    blnEditing = False
    mnuFileOpen.Enabled = True
    tlbToolbar.Buttons("open").Enabled = True
    mnuFileClose.Enabled = True
    mnuSoundRestoreAll.Enabled = True
    tlbToolbar.Buttons("restoreall").Enabled = True
    UpdateSpace
    UpdateAll
End Sub
