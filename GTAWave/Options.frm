VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   Icon            =   "Options.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraEditor 
      Caption         =   "Sound Editor"
      Height          =   1575
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   5655
      Begin VB.CommandButton cmdEditor 
         Caption         =   "Browse..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   18
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox txtEditor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   3615
      End
      Begin VB.OptionButton optExternal 
         Caption         =   "&External editor:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton optDefault 
         Caption         =   "&Default program associated with WAV files"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.Frame fraFile 
      Caption         =   "GTA2 File Locations"
      Height          =   1695
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   5655
      Begin VB.TextBox txtBackup 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   9
         Top             =   600
         Width           =   3615
      End
      Begin VB.TextBox txtProgram 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1200
         Width           =   3615
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Browse..."
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdProgram 
         Caption         =   "Browse..."
         Height          =   375
         Index           =   1
         Left            =   4200
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblBackup 
         Caption         =   "B&ackup Folder:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblProgram 
         Caption         =   "G&TA2 Program File:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   7080
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Programs|*.bat;*.com;*.exe|All files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   25
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   24
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame fraDoubleClick 
      Caption         =   "On Double Click"
      Height          =   1815
      Left            =   6000
      TabIndex        =   19
      Top             =   120
      Width           =   1575
      Begin VB.OptionButton optPlayLoop 
         Caption         =   "Play &Loop"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optPlay 
         Caption         =   "&Play Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton optOpen 
         Caption         =   "&Open Sound"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optNothing 
         Caption         =   "Do &Nothing"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame fraFile 
      Caption         =   "GTA File Locations"
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CommandButton cmdProgram 
         Caption         =   "Browse..."
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdBackup 
         Caption         =   "Browse..."
         Height          =   375
         Index           =   0
         Left            =   4200
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtProgram 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox txtBackup 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   3615
      End
      Begin VB.Label lblProgram 
         Caption         =   "&GTA Program File:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lblBackup 
         Caption         =   "&Backup Folder:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBackup_Click(Index As Integer)
    intBrowseVersion = Index
    frmBackupFolder.Show vbModal
    
    If Not blnCancel Then txtBackup(Index).Text = strBackupDir
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEditor_Click()
    On Error GoTo Cancel
    dlgBrowse.filename = txtEditor.Text
    dlgBrowse.ShowOpen
    txtEditor.Text = dlgBrowse.filename
    Exit Sub
    
Cancel:
End Sub

Private Sub cmdOK_Click()
    Dim intDoubleClick As Integer
    
    SaveSetting "GTA Wave", "Options", "BackupDir", txtBackup(0).Text
    SaveSetting "GTA Wave", "Options", "BackupDir2", txtBackup(1).Text
    SaveSetting "GTA Wave", "Options", "GTAProgFile", txtProgram(0).Text
    SaveSetting "GTA Wave", "Options", "GTAProgFile2", txtProgram(1).Text
    SaveSetting "GTA Wave", "Options", "Editor", -optExternal
    SaveSetting "GTA Wave", "Options", "EditorProgram", txtEditor.Text
    
    If optNothing Then
        intDoubleClick = 0
    ElseIf optOpen Then
        intDoubleClick = 1
    ElseIf optPlay Then
        intDoubleClick = 2
    Else
        intDoubleClick = 3
    End If
    
    SaveSetting "GTA Wave", "Options", "DoubleClick", intDoubleClick
    Unload Me
End Sub

Private Sub cmdProgram_Click(Index As Integer)
    On Error GoTo Cancel
    dlgBrowse.filename = txtProgram(Index).Text
    dlgBrowse.ShowOpen
    txtProgram(Index).Text = dlgBrowse.filename
    Exit Sub
    
Cancel:
End Sub

Private Sub Form_Load()
    Dim intDoubleClick As Integer
    
    dlgBrowse.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    txtBackup(0).Text = GetSetting("GTA Wave", "Options", "BackupDir")
    txtBackup(1).Text = GetSetting("GTA Wave", "Options", "BackupDir2")
    txtProgram(0).Text = GetSetting("GTA Wave", "Options", "GTAProgFile")
    txtProgram(1).Text = GetSetting("GTA Wave", "Options", "GTAProgFile2")
    optExternal = -GetSetting("GTA Wave", "Options", "Editor", 0)
    txtEditor.Text = GetSetting("GTA Wave", "Options", "EditorProgram")
    intDoubleClick = GetSetting("GTA Wave", "Options", "DoubleClick", 1)
    
    If intDoubleClick = 0 Then
        optNothing.Value = True
    ElseIf intDoubleClick = 1 Then
        optOpen.Value = True
    ElseIf intDoubleClick = 2 Then
        optPlay.Value = True
    Else
        optPlayLoop.Value = True
    End If
    
    If blnEditing Then
        fraEditor.Enabled = False
        optDefault.Enabled = False
        optExternal.Enabled = False
        txtEditor.Enabled = False
        cmdEditor.Enabled = False
    End If
End Sub

Private Sub optDefault_Click()
    txtEditor.Enabled = False
    cmdEditor.Enabled = False
End Sub

Private Sub optDefault_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optExternal_Click()
    txtEditor.Enabled = True
    cmdEditor.Enabled = True
End Sub

Private Sub optExternal_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optNothing_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optOpen_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optPlay_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optPlayLoop_DblClick()
    cmdOK.Value = True
End Sub
