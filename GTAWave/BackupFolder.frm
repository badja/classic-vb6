VERSION 5.00
Begin VB.Form frmBackupFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Folder"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3840
   Icon            =   "BackupFolder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   3840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.DriveListBox drvBackup 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   3960
      Width           =   3615
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   4440
      Width           =   1215
   End
   Begin VB.DirListBox dirBackup 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.Label lblFolder 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label lblBackup 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frmBackupFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strInitDrive As String

Private Sub cmdCancel_Click()
    blnCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    blnCancel = False
    strBackupDir = dirBackup.Path
    Unload Me
End Sub

Private Sub dirBackup_Change()
    lblFolder.Caption = dirBackup.Path
End Sub

Private Sub drvBackup_Change()
    On Error GoTo Unavailable
    dirBackup.Path = drvBackup.Drive
    Exit Sub
    
Unavailable:
    MsgBox "Drive is unavailable", vbExclamation
    drvBackup.Drive = strInitDrive
    Resume Next
End Sub

Private Sub Form_Load()
    strInitDrive = drvBackup.Drive
    lblBackup.Caption = "Select the folder which contains the original GTA"
    If intBrowseVersion = 2 Then lblBackup.Caption = lblBackup.Caption & "2"
    lblBackup.Caption = lblBackup.Caption & " sound files:"
    lblFolder.Caption = dirBackup.Path
End Sub
