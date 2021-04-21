VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form frmExtractor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extractor"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Extractor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   3960
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtSave 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog dlgOpen 
      Left            =   3360
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327681
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "&Extract"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Browse..."
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      ItemData        =   "Extractor.frx":0442
      Left            =   720
      List            =   "Extractor.frx":0452
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtSize 
      Height          =   285
      Left            =   2040
      TabIndex        =   10
      Text            =   "0"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtStart 
      Height          =   285
      Left            =   720
      TabIndex        =   7
      Text            =   "0"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtOpen 
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "&Save"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   495
   End
   Begin VB.Label lblSize 
      Caption         =   "S&ize"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label lblStart 
      Caption         =   "S&tart"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label lblFile 
      Caption         =   "&Open"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmExtractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSize_Click()
    If cmbSize.ListIndex = 3 Then txtSize.Enabled = False Else txtSize.Enabled = True
End Sub

Private Sub cmdExtract_Click()
    Dim lngStart As Long, lngSize As Long
    Dim strData As String

    On Error GoTo ErrorHandler
    
    lngStart = txtStart.Text
    
    Select Case cmbSize.ListIndex
        Case 0
            lngSize = txtSize.Text
        Case 1
            lngSize = txtSize.Text - lngStart + 1
        Case 2
            lngSize = txtSize.Text - lngStart
        Case 3
            lngSize = FileLen(txtOpen.Text) - lngStart
    End Select
    
    strData = Space(lngSize)
    Open txtOpen.Text For Binary Access Read As #1
    Get #1, lngStart + 1, strData
    Close #1
    If Dir(txtSave.Text) <> "" Then Kill txtSave.Text
    Open txtSave.Text For Binary Access Write As #1
    Put #1, , strData
    Close #1
    Exit Sub
    
ErrorHandler:
    MsgBox "Error", vbCritical
End Sub

Private Sub cmdOpen_Click()
    On Error GoTo Cancel
    dlgOpen.ShowOpen
    txtOpen.Text = dlgOpen.filename
    
Cancel:
End Sub

Private Sub cmdSave_Click()
    On Error GoTo Cancel
    dlgSave.ShowSave
    txtSave.Text = dlgSave.filename
    
Cancel:
End Sub

Private Sub Form_Load()
    cmbSize.ListIndex = 0
    dlgOpen.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    dlgSave.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
End Sub
