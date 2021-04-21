VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form frmDifficulty 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Difficulty"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "Diffic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider sldDifficulty 
      Height          =   630
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1111
      _Version        =   327682
      LargeChange     =   1
      Min             =   2
      Max             =   8
      SelStart        =   5
      Value           =   5
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label lblHard 
      Alignment       =   1  'Right Justify
      Caption         =   "Hard"
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblEasy 
      Caption         =   "Easy"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmDifficulty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmDifficulty.Hide
End Sub

Private Sub cmdOK_Click()
    Amount = sldDifficulty.Value * 5
    Speed = sldDifficulty.Value * 15
    frmDifficulty.Hide
End Sub

