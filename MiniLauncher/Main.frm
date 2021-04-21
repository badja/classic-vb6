VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MiniLauncher"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Run"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgAdd 
      Left            =   4080
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   4080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
      Filter          =   "Programs|*.bat;*.com;*.exe|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtParameters 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   4455
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse..."
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtProgram 
      Height          =   495
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblParameters 
      Caption         =   "Para&meters"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblProgram 
      Caption         =   "&Program"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    On Error GoTo Cancelled
    dlgAdd.ShowOpen
    txtParameters.Text = txtParameters.Text & Chr(34) & dlgAdd.filename & Chr(34) & " "
    Exit Sub
Cancelled:
End Sub

Private Sub cmdBrowse_Click()
    On Error GoTo Cancelled
    dlgBrowse.ShowOpen
    txtProgram.Text = dlgBrowse.filename
    Exit Sub
Cancelled:
End Sub

Private Sub cmdClear_Click()
    txtParameters.Text = ""
End Sub

Private Sub cmdRun_Click()
    Dim dblDummy As Double
    dblDummy = Shell(Chr(34) & txtProgram.Text & Chr(34) & " " & txtParameters.Text, vbNormalFocus)
End Sub

Private Sub Form_Load()
    dlgBrowse.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgAdd.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
End Sub
