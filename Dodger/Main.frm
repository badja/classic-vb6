VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dodger"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   9510
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6135
   ScaleWidth      =   9510
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Image imgShip2 
      Height          =   210
      Left            =   3120
      MouseIcon       =   "Main.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":074C
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Image imgShip1 
      Height          =   210
      Left            =   5760
      MouseIcon       =   "Main.frx":0B45
      MousePointer    =   99  'Custom
      Picture         =   "Main.frx":0E4F
      Top             =   4200
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Dodger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   9255
   End
   Begin VB.Shape shpObstacle 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   375
      Index           =   0
      Left            =   480
      Shape           =   1  'Square
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuGameNew 
         Caption         =   "&New Game"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGamePause 
         Caption         =   "&Pause"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuGameEnd 
         Caption         =   "&End Game"
         Enabled         =   0   'False
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuGameSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuGameDifficulty 
         Caption         =   "&Difficulty..."
      End
      Begin VB.Menu mnuGameSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Size As Integer
Dim GoLeft As Integer, GoRight As Integer, GoUp As Integer, GoDown As Integer
Dim GoLeft2 As Integer, GoRight2 As Integer, GoUp2 As Integer, GoDown2 As Integer
Dim XPos As Integer, YPos As Integer
Dim XPos2 As Integer, YPos2 As Integer
Option Explicit

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KEY_LEFT
            If imgShip1.Visible = False And Timer1.Enabled = True Then imgShip1.Visible = True
            GoLeft = True
        Case KEY_RIGHT
            If imgShip1.Visible = False And Timer1.Enabled = True Then imgShip1.Visible = True
            GoRight = True
        Case KEY_UP
            If imgShip1.Visible = False And Timer1.Enabled = True Then imgShip1.Visible = True
            GoUp = True
        Case KEY_DOWN
            If imgShip1.Visible = False And Timer1.Enabled = True Then imgShip1.Visible = True
            GoDown = True
    End Select
    Select Case KeyCode
        Case Asc("A")
            If imgShip2.Visible = False And Timer1.Enabled = True Then imgShip2.Visible = True
            GoLeft2 = True
        Case Asc("D")
            If imgShip2.Visible = False And Timer1.Enabled = True Then imgShip2.Visible = True
            GoRight2 = True
        Case Asc("W")
            If imgShip2.Visible = False And Timer1.Enabled = True Then imgShip2.Visible = True
            GoUp2 = True
        Case Asc("X")
            If imgShip2.Visible = False And Timer1.Enabled = True Then imgShip2.Visible = True
            GoDown2 = True
        Case Asc("S")
            If imgShip2.Visible = False And Timer1.Enabled = True Then imgShip2.Visible = True
            GoDown2 = True
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case KEY_LEFT
            GoLeft = False
        Case KEY_RIGHT
            GoRight = False
        Case KEY_UP
            GoUp = False
        Case KEY_DOWN
            GoDown = False
    End Select
    Select Case KeyCode
        Case Asc("A")
            GoLeft2 = False
        Case Asc("D")
            GoRight2 = False
        Case Asc("W")
            GoUp2 = False
        Case Asc("X")
            GoDown2 = False
        Case Asc("S")
            GoDown2 = False
    End Select
End Sub

Private Sub Form_Load()
    Randomize
    Amount = 25
    Speed = 75
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub imgShip1_Click()
    mnuGameEnd_Click
    MsgBox "Player 1 has been shot down"
    mnuGameNew_Click
    Exit Sub
End Sub

Private Sub imgShip2_Click()
    mnuGameEnd_Click
    MsgBox "Player 2 has been shot down"
    mnuGameNew_Click
    Exit Sub
End Sub

Private Sub mnuGameAbout_Click()
    frmAboutBox.Show 1
End Sub

Private Sub mnuGameEnd_Click()
    Timer1.Enabled = False
    lblInfo.Caption = "Game Over"
            lblInfo.Visible = True
    Dim All As Integer
    For All = 1 To Amount - 1
        Unload shpObstacle(All)
        shpObstacle(0).Visible = False
    Next All
    mnuGameNew.Enabled = True
    mnuGamePause.Enabled = False
    mnuGameEnd.Enabled = False
    mnuGameAbout.Enabled = True
    mnuGameDifficulty.Enabled = True
    imgShip1.Visible = False
    imgShip2.Visible = False
End Sub

Private Sub mnuGameExit_Click()
    End
End Sub

Private Sub mnuGameNew_Click()
    Dim All As Integer
    XPos = ScaleWidth / 3 * 2
    YPos = ScaleHeight / 2
    XPos2 = ScaleWidth / 3
    YPos2 = ScaleHeight / 2
    lblInfo.Visible = False
    GoLeft = False
    GoRight = False
    GoUp = False
    GoDown = False
    GoLeft2 = False
    GoRight2 = False
    GoUp2 = False
    GoDown2 = False
    For All = 0 To Amount - 1
        If All > 0 Then Load shpObstacle(All)
        shpObstacle(All).FillColor = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
        Size = Int((2000 + 1) * Rnd)
        shpObstacle(All).Move Int((ScaleWidth + 1) * Rnd), Int((ScaleHeight + 1) * Rnd), Size, Size
        shpObstacle(All).Visible = True
    Next All
    mnuGameNew.Enabled = False
    mnuGamePause.Enabled = True
    mnuGameEnd.Enabled = True
    mnuGameAbout.Enabled = False
    mnuGameDifficulty.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub mnuGameDifficulty_Click()
    frmDifficulty.Show 1
End Sub

Private Sub mnuGamePause_Click()
    If mnuGamePause.Checked = False Then
        Timer1.Enabled = False
        mnuGamePause.Checked = True
    Else
        Timer1.Enabled = True
        mnuGamePause.Checked = False
    End If
End Sub

Private Sub Timer1_Timer()
    Dim All As Integer
    Dim Ecc As Single
    Dim MoveSpeed As Integer
    
    MoveSpeed = 150
    If GoLeft = True Then XPos = XPos - MoveSpeed
    If GoRight = True Then XPos = XPos + MoveSpeed
    If GoUp = True Then YPos = YPos - MoveSpeed
    If GoDown = True Then YPos = YPos + MoveSpeed
    imgShip1.Move XPos, YPos
    
    If GoLeft2 = True Then XPos2 = XPos2 - MoveSpeed
    If GoRight2 = True Then XPos2 = XPos2 + MoveSpeed
    If GoUp2 = True Then YPos2 = YPos2 - MoveSpeed
    If GoDown2 = True Then YPos2 = YPos2 + MoveSpeed
    imgShip2.Move XPos2, YPos2
    
    Ecc = 0.04
    For All = 0 To Amount - 1
        shpObstacle(All).Move shpObstacle(All).Left + Ecc * (shpObstacle(All).Left + shpObstacle(All).Width / 2 - ScaleWidth / 2), shpObstacle(All).Top + Ecc * (shpObstacle(All).Top + shpObstacle(All).Height / 2 - ScaleHeight / 2), shpObstacle(All).Width + Speed, shpObstacle(All).Height + Speed
        If shpObstacle(All).Width > 2000 Then
            shpObstacle(All).Move Int((ScaleWidth + 1) * Rnd), Int((ScaleHeight + 1) * Rnd), 0, 0
            shpObstacle(All).ZOrder 1
            shpObstacle(All).FillColor = RGB(Int(256 * Rnd), Int(256 * Rnd), Int(256 * Rnd))
        End If
        If imgShip1.Visible = True And imgShip1.Top < 0 Or imgShip1.Left < 0 Or imgShip1.Top > ScaleHeight - imgShip1.Height Or imgShip1.Left > ScaleWidth - imgShip1.Width Then
            mnuGameEnd_Click
            MsgBox "Player 1 has been sucked into the black hole!"
            mnuGameNew_Click
            Exit Sub
        End If
        If shpObstacle(All).Width > 1000 And imgShip1.Visible = True And _
            imgShip1.Left >= shpObstacle(All).Left - imgShip1.Width And _
            imgShip1.Top >= shpObstacle(All).Top - imgShip1.Height And _
            imgShip1.Left <= shpObstacle(All).Left + shpObstacle(All).Width And _
            imgShip1.Top <= shpObstacle(All).Top + shpObstacle(All).Height Then
            mnuGameEnd_Click
            MsgBox "Player 1 has been hit!"
            mnuGameNew_Click
            Exit Sub
        End If
        
        If imgShip2.Visible = True And imgShip2.Top < 0 Or imgShip2.Left < 0 Or imgShip2.Top > ScaleHeight - imgShip2.Height Or imgShip2.Left > ScaleWidth - imgShip2.Width Then
            mnuGameEnd_Click
            MsgBox "Player 2 has been sucked into the black hole!"
            mnuGameNew_Click
            Exit Sub
        End If
        If shpObstacle(All).Width > 1000 And imgShip2.Visible = True And _
            imgShip2.Left >= shpObstacle(All).Left - imgShip2.Width And _
            imgShip2.Top >= shpObstacle(All).Top - imgShip2.Height And _
            imgShip2.Left <= shpObstacle(All).Left + shpObstacle(All).Width And _
            imgShip2.Top <= shpObstacle(All).Top + shpObstacle(All).Height Then
            mnuGameEnd_Click
            MsgBox "Player 2 has been hit!"
            mnuGameNew_Click
            Exit Sub
        End If
    Next All
    
End Sub
