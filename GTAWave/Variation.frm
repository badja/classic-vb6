VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmVariation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Pitch Variation Range"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   Icon            =   "Variation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4800
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
      OrigTop         =   720
      OrigRight       =   2040
      OrigBottom      =   1005
      Increment       =   10
      Max             =   500
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ComCtl2.UpDown updConstant 
      Height          =   285
      Left            =   3840
      TabIndex        =   3
      Top             =   180
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   327681
      Value           =   1000
      BuddyControl    =   "txtConstant"
      BuddyDispid     =   196610
      OrigLeft        =   3990
      OrigTop         =   180
      OrigRight       =   4230
      OrigBottom      =   465
      Increment       =   100
      Max             =   20000
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtConstant 
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Text            =   "1000"
      Top             =   180
      Width           =   615
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play &Highest"
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play &Lowest"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton optConstant 
      Caption         =   "Set &constant pitch variation range of ±"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Value           =   -1  'True
      Width           =   3015
   End
   Begin VB.OptionButton optScale 
      Caption         =   "&Scale to"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   690
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   11
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   1800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play &Mean"
      Height          =   375
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblConstant 
      Caption         =   "Hz"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   225
      Width           =   255
   End
   Begin VB.Label lblScale 
      Caption         =   "% of current pitch variation range(s)"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   705
      Width           =   2535
   End
End
Attribute VB_Name = "frmVariation"
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
        lngPlayVariation = txtConstant.Text
    Else
        lngPlayVariation = txtScale.Text / 100 * lngCurVariation
    End If
    
    If cmdPlay(0).Enabled And lngPlayVariation >= lngCurRate Then
        MsgBox "Pitch variation range is too large for sound", vbExclamation
    Else
    
        If optConstant.Value Then
            blnScale = False
            blnLastConst = True
        Else
            lngPlayVariation = 1
            sngScale = txtScale.Text
            blnScale = True
            blnLastConst = False
        End If
        
        If lngPlayVariation >= 0 Then
            blnCancel = False
            Unload Me
        Else
            MsgBox "Invalid pitch variation range", vbExclamation
        End If
    End If
End Sub

Private Sub cmdPlay_Click(Index As Integer)
    On Error GoTo CannotKill
    
    Dim lngSuccess As Long
    
    If txtConstant.Text >= 0 Then
        
        If optConstant.Value Then
            lngPlayVariation = txtConstant.Text
        Else
            lngPlayVariation = txtScale.Text / 100 * lngCurVariation
        End If
        
        If Index = 0 Then lngPlayVariation = -lngPlayVariation
        
        If Abs(lngPlayVariation) >= lngCurRate Then
            MsgBox "Pitch variation range is too large for sound", vbExclamation
        Else
            If Index = 1 Then lngPlayVariation = 0
            StopPlaying
            If Dir(strTempFile) <> "" Then SafeKill strTempFile
            CreateFile strTempFile, intPlayIndex, frmGTAWave.dlgOpen.filename, False, False
            lngSuccess = sndPlaySound(strTempFile, SND_ASYNC)
        End If
    Else
        MsgBox "Invalid pitch variation range", vbExclamation
    End If
    
CannotKill:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPlaying
End Sub

Private Sub optConstant_DblClick()
    cmdOK.Value = True
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
