VERSION 5.00
Begin VB.Form frmFileLocations 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Locations"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "FileLocations.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBaseSav 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   4455
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtAutoSav 
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   4455
   End
   Begin VB.TextBox txtVTI 
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4455
   End
   Begin VB.TextBox txtDevpac 
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtTASM 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblBaseSav 
      Caption         =   "Virtual TI &base saved state file:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label lblAutoSav 
      Caption         =   "Virtual TI &automatic saved state file:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label lblVTI 
      Caption         =   "&Virtual TI Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label lblDevpac 
      Caption         =   "&Devpac83 Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblTASM 
      Caption         =   "&TASM Directory:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFileLocations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Right(txtTASM.Text, 1) <> "\" Then txtTASM.Text = txtTASM.Text & "\"
    If Right(txtDevpac.Text, 1) <> "\" Then txtDevpac.Text = txtDevpac.Text & "\"
    If Right(txtVTI.Text, 1) <> "\" Then txtVTI.Text = txtVTI.Text & "\"
    SaveSetting ThisApp, SetKey, "TASMDir", txtTASM.Text
    SaveSetting ThisApp, SetKey, "Devpac83Dir", txtDevpac.Text
    SaveSetting ThisApp, SetKey, "VTIDir", txtVTI.Text
    SaveSetting ThisApp, SetKey, "AutoSaveFile", txtAutoSav.Text
    SaveSetting ThisApp, SetKey, "BaseSaveFile", txtBaseSav.Text
    Unload Me
End Sub

Private Sub Form_Load()
    txtTASM.Text = GetSetting(ThisApp, SetKey, "TASMDir")
    txtDevpac.Text = GetSetting(ThisApp, SetKey, "Devpac83Dir")
    txtVTI.Text = GetSetting(ThisApp, SetKey, "VTIDir")
    txtAutoSav.Text = GetSetting(ThisApp, SetKey, "AutoSaveFile")
    txtBaseSav.Text = GetSetting(ThisApp, SetKey, "BaseSaveFile")
End Sub
