VERSION 5.00
Begin VB.Form frmAboutBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Dodger"
   ClientHeight    =   2040
   ClientLeft      =   4395
   ClientTop       =   2925
   ClientWidth     =   5055
   Icon            =   "AboutBox.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmail 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "jaigar@yahoo.com"
      Top             =   1080
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      MaskColor       =   &H00000000&
      TabIndex        =   1
      Top             =   1560
      Width           =   975
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   525
      Left            =   255
      Picture         =   "AboutBox.frx":000C
      ScaleHeight     =   525
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Copyright 1998, James Grucza"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   5
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Version 1.0"
      Height          =   210
      Index           =   2
      Left            =   1200
      TabIndex        =   4
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      Caption         =   "Dodger"
      Height          =   210
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      Caption         =   "E-mail:"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   2
      Top             =   1110
      Width           =   495
   End
End
Attribute VB_Name = "frmAboutBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
