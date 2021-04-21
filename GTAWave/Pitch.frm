VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmPitch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Pitch"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "Pitch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtScale 
      Height          =   285
      Left            =   1080
      TabIndex        =   5
      Text            =   "100"
      Top             =   660
      Width           =   615
   End
   Begin VB.CheckBox chkVariation 
      Caption         =   "Scale pitch &variation range(s)"
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play"
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   10
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1215
   End
   Begin ComCtl2.UpDown updConstant 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   22050
      BuddyControl    =   "txtConstant"
      BuddyDispid     =   196614
      OrigLeft        =   3120
      OrigTop         =   60
      OrigRight       =   3360
      OrigBottom      =   435
      Increment       =   1000
      Max             =   44100
      Min             =   4000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtConstant 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Text            =   "22050"
      Top             =   180
      Width           =   615
   End
   Begin VB.OptionButton optScale 
      Caption         =   "&Scale to"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   975
   End
   Begin VB.OptionButton optConstant 
      Caption         =   "Set &constant sample rate of"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Value           =   -1  'True
      Width           =   2295
   End
   Begin ComCtl2.UpDown updScale 
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   660
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   100
      BuddyControl    =   "txtScale"
      BuddyDispid     =   196609
      OrigLeft        =   1800
      OrigTop         =   540
      OrigRight       =   2040
      OrigBottom      =   915
      Increment       =   10
      Max             =   500
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblScale 
      Caption         =   "% of current sample rate(s)"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   705
      Width           =   1935
   End
   Begin VB.Label lblConstant 
      Caption         =   "Hz"
      Height          =   255
      Left            =   3360
      TabIndex        =   4
      Top             =   225
      Width           =   255
   End
End
Attribute VB_Name = "frmPitch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optConstant.Value Then
        lngPlayRate = txtConstant.Text
        blnScale = False
        blnLastConst = True
    Else
        lngPlayRate = 1
        sngScale = txtScale.Text
        blnScale = True
        blnLastConst = False
    End If
    
    If chkVariation.Value Then
        blnScaleVar = True
        blnLastScaleVar = True
    Else
        blnScaleVar = False
        blnLastScaleVar = False
    End If
    
    If lngPlayRate > 0 Then
        blnCancel = False
        Unload Me
    Else
        MsgBox "Invalid sample rate", vbExclamation
    End If
End Sub

Private Sub cmdPlay_Click()
    On Error GoTo CannotKill
    
    Dim lngSuccess As Long
    
    If optConstant.Value Then
        lngPlayRate = txtConstant.Text
    Else
        lngPlayRate = txtScale.Text / 100 * lngCurRate
    End If
    
    If lngPlayRate > 0 Then
        StopPlaying
        If Dir(strTempFile) <> "" Then SafeKill strTempFile
        CreateFile strTempFile, intPlayIndex, frmGTAWave.dlgOpen.filename, False, False
        lngSuccess = sndPlaySound(strTempFile, SND_ASYNC)
    Else
        MsgBox "Invalid sample rate", vbExclamation
    End If
    
CannotKill:
End Sub

Private Sub Form_Load()
    If intGTAVersion = 1 Then chkVariation.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPlaying
End Sub

Private Sub optConstant_Click()
    chkVariation.Enabled = False
End Sub

Private Sub optConstant_DblClick()
    cmdOK.Value = True
End Sub

Private Sub optScale_Click()
    If intGTAVersion = 2 Then chkVariation.Enabled = True
End Sub

Private Sub optScale_DblClick()
    cmdOK.Value = True
End Sub

Private Sub txtConstant_Change()
    optConstant.Value = True
End Sub

Private Sub txtScale_Change()
    optScale.Value = True
End Sub

Private Sub updConstant_Change()
    txtConstant.SetFocus
End Sub

Private Sub updScale_Change()
    txtScale.SetFocus
End Sub
