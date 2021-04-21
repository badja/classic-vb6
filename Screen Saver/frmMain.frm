VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9855
   LinkTopic       =   "Form1"
   ScaleHeight     =   8205
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox picScreen 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   120
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   240
      Width           =   9600
      Begin VB.Label lblIntro 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":0000
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   6975
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   9375
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    'lblIntro.Top = 7200
    picScreen.Move Screen.Width / 2 - 4800, Screen.Height / 2 - 3600
    frmMain.Move 0, 0, Screen.Width, Screen.Height
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseMoveCount = MouseMoveCount + 1
    If MouseMoveCount > 10 Then
      MouseMoveCount = 0
      Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsPasswordEnabled = 1 Then
      ' show the cursor
      Call ShowCursor(True)
      ' show the password entry box
      If VerifyScreenSavePwd(Me.hwnd) = False Then
        ' incorrect password - cancel unload
        Cancel = True
        ' hide cursor
        Call ShowCursor(False)
        Exit Sub
      End If
    End If
    
    'got this far?  Then clean up and exit
    Call EnableCtrlAltDelete(True)
    Call ShowCursor(True)
End Sub

Private Sub Timer1_Timer()
    'lblIntro.Top = lblIntro.Top - 30
    'If lblIntro.Top = -11520 Then lblIntro.Top = 7200
End Sub
