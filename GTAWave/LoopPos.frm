VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.0#0"; "COMCT232.OCX"
Begin VB.Form frmLoopPos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Loop Start/End Points"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "LoopPos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLoopStart 
      Caption         =   "Loop start point"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtLoopStart 
         Height          =   285
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   360
         Width           =   735
      End
      Begin ComCtl2.UpDown updLoopStart 
         Height          =   285
         Left            =   2535
         TabIndex        =   3
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtLoopStart"
         BuddyDispid     =   196610
         OrigLeft        =   2760
         OrigTop         =   360
         OrigRight       =   3000
         OrigBottom      =   645
         Increment       =   1000
         Max             =   999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label txtSet 
         Caption         =   "&Set loop start point to"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   405
         Width           =   1695
      End
      Begin VB.Label txtFrom1 
         Caption         =   "bytes from start of sound(s)"
         Height          =   255
         Left            =   2880
         TabIndex        =   4
         Top             =   405
         Width           =   1935
      End
   End
   Begin VB.Frame fraLoopEnd 
      Caption         =   "Loop end point"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5175
      Begin VB.TextBox txtLoopEnd 
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "0"
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton optSetPos 
         Caption         =   "Set loop end point to"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   750
         Width           =   2055
      End
      Begin VB.OptionButton optSetEnd 
         Caption         =   "Set loop end point to end of sound(s)"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   3015
      End
      Begin ComCtl2.UpDown updLoopEnd 
         Height          =   285
         Left            =   2775
         TabIndex        =   9
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   327681
         BuddyControl    =   "txtLoopEnd"
         BuddyDispid     =   196614
         OrigLeft        =   3000
         OrigTop         =   720
         OrigRight       =   3240
         OrigBottom      =   1005
         Increment       =   1000
         Max             =   999999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label txtFrom2 
         Caption         =   "bytes from start of sound(s)"
         Height          =   255
         Left            =   3120
         TabIndex        =   10
         Top             =   765
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "S&top Playing"
      Height          =   375
      Left            =   2760
      TabIndex        =   13
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play &Loop"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   14
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   15
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "&Play Sound"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "frmLoopPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    lngPlayLoopStart = txtLoopStart.Text
    lngPlayLoopEnd = txtLoopEnd.Text
    
    If lngPlayLoopStart < 0 Then
        MsgBox "Invalid loop start point", vbExclamation
    ElseIf cmdPlay(0).Enabled And lngPlayLoopStart > lngCurSize Then
        MsgBox "Loop start point is too large for sound", vbExclamation
    Else
        If optSetPos.Value Then
            If lngPlayLoopEnd < 0 Then
                MsgBox "Invalid loop end point", vbExclamation
                Exit Sub
            ElseIf lngPlayLoopEnd < lngPlayLoopStart Then
                MsgBox "Loop end point cannot be less than loop start point", vbExclamation
                Exit Sub
            ElseIf cmdPlay(0).Enabled And lngPlayLoopEnd > lngCurSize Then
                MsgBox "Loop end point is too large for sound", vbExclamation
                Exit Sub
            End If
        Else
            lngPlayLoopEnd = -1
        End If
        
        blnCancel = False
        Unload Me
    End If
End Sub

Private Sub cmdPlay_Click(Index As Integer)
    On Error GoTo CannotKill
    
    Dim lngSuccess As Long
    Dim lngUFlags As Long
    
    If txtLoopStart.Text >= 0 Then
        lngPlayLoopStart = txtLoopStart.Text
        
        If lngPlayLoopStart > lngCurSize Then
            MsgBox "Loop start point is too large for sound", vbExclamation
        Else
            lngPlayLoopEnd = txtLoopEnd.Text
            
            If optSetPos.Value = True Then
                If lngPlayLoopEnd < 0 Then
                    MsgBox "Invalid loop end point", vbExclamation
                    Exit Sub
                ElseIf lngPlayLoopEnd < lngPlayLoopStart Then
                    MsgBox "Loop end point cannot be less than loop start point", vbExclamation
                    Exit Sub
                ElseIf lngPlayLoopEnd > lngCurSize Then
                    MsgBox "Loop end point is too large for sound", vbExclamation
                    Exit Sub
                End If
            Else
                lngPlayLoopEnd = -1
            End If
            
            lngUFlags = SND_NODEFAULT Or SND_ASYNC
            If Index = 1 Then lngUFlags = lngUFlags Or SND_ASYNC Or SND_LOOP
            StopPlaying
            If Dir(strTempFile) <> "" Then SafeKill strTempFile
            CreateFile strTempFile, intPlayIndex, frmGTAWave.dlgOpen.filename, -Index, False
            lngSuccess = sndPlaySound(strTempFile, lngUFlags)
        End If
    Else
        MsgBox "Invalid loop start point", vbExclamation
    End If
    
CannotKill:
End Sub

Private Sub cmdStop_Click()
    StopPlaying
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopPlaying
End Sub

Private Sub txtLoopEnd_Change()
    optSetPos.Value = True
End Sub
