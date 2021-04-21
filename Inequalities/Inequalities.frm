VERSION 5.00
Begin VB.Form frmInequalities 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   425
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   679
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGraph 
      Caption         =   "&Graph"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox picViewport 
      Height          =   5775
      Left            =   120
      ScaleHeight     =   381
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   661
      TabIndex        =   0
      Top             =   600
      Width           =   9975
      Begin VB.PictureBox picCanvas 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2175
         Left            =   120
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   273
         TabIndex        =   1
         Top             =   120
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmInequalities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sngPixelsPerUnit As Single
Private sngCentreX As Single, sngCentreY As Single
Private sngCanvasWidth As Single, sngCanvasHeight As Single
Private sngViewportWidth As Single, sngViewportHeight As Single
Private sngXMin As Single, sngYMin As Single
Private sngXMax As Single, sngYMax As Single
Private blnMoving As Boolean
Private intMoveX As Integer
Private intMoveY As Integer
Private blnCancelMove As Boolean

Private Sub cmdGraph_Click()
    Dim i As Integer, j As Integer
    Dim intX As Integer, intY As Integer
    Dim intXMin As Integer, intYMin As Integer
    Dim X As Single, Y As Single
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    intXMin = -picCanvas.Left
    intYMin = -picCanvas.Top
    intY = intYMin
    picCanvas.Visible = False
    
    For i = 0 To picViewport.ScaleHeight - 1
        Y = sngYMax - sngViewportHeight * i / (picViewport.ScaleHeight - 1)
        intX = intXMin
        For j = 0 To picViewport.ScaleWidth - 1
            X = sngXMin + sngViewportWidth * j / (picViewport.ScaleWidth - 1)
            'If X * Cos(Y) + Y * Cos(X) < Y * Exp(X) - X * Exp(Y) Then picCanvas.PSet (intX, intY)
            'If Cos(X) < Cos(Y) Then picCanvas.PSet (intX, intY)
            'If Exp(X ^ (Y / (2 * Sin(Y))) / (1 - Cos(X) ^ 2)) < Sin(Exp(1 / (Cos(X) + (Y / 2)))) / Y Then picCanvas.PSet (intX, intY)
            If Sin(Y * (X + Y) * Sin(Y) * (1 - Cos(Y))) < Sin(Cos((Y - X) * (Cos(Y) + X))) / Y Then picCanvas.PSet (intX, intY)
            intX = intX + 1
        Next j
        intY = intY + 1
    Next i
    
    picCanvas.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    sngPixelsPerUnit = 16 * 2
    sngCentreX = 0
    sngCentreY = 0
    picCanvas.Width = 1024
    picCanvas.Height = 768
    UpdateVars
    picCanvas.Move Int((sngViewportWidth / 2 - sngCanvasWidth / 2 - sngCentreX) * sngPixelsPerUnit), Int((sngViewportHeight / 2 - sngCanvasHeight / 2 + sngCentreY) * sngPixelsPerUnit)
End Sub

Private Sub picCanvas_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picCanvas.MousePointer = vbSizePointer
    intMoveX = X + picCanvas.Left
    intMoveY = Y + picCanvas.Top
    blnMoving = True
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If blnCancelMove Then
        blnCancelMove = False
    ElseIf blnMoving Then
        blnCancelMove = True
        picCanvas.Move Int((sngViewportWidth / 2 - sngCanvasWidth / 2 - sngCentreX) * sngPixelsPerUnit + X + picCanvas.Left - intMoveX), Int((sngViewportHeight / 2 - sngCanvasHeight / 2 + sngCentreY) * sngPixelsPerUnit + Y + picCanvas.Top - intMoveY)
    End If
End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnMoving = False
    sngCentreX = sngCentreX - (X + picCanvas.Left - intMoveX) / sngPixelsPerUnit
    sngCentreY = sngCentreY + (Y + picCanvas.Top - intMoveY) / sngPixelsPerUnit
    UpdateVars
    picCanvas.MousePointer = vbDefault
End Sub

Private Sub UpdateVars()
    sngCanvasWidth = picCanvas.Width / sngPixelsPerUnit
    sngCanvasHeight = picCanvas.Height / sngPixelsPerUnit
    sngViewportWidth = picViewport.ScaleWidth / sngPixelsPerUnit
    sngViewportHeight = picViewport.ScaleHeight / sngPixelsPerUnit
    sngXMin = sngCentreX - sngViewportWidth / 2
    sngYMin = sngCentreY - sngViewportHeight / 2
    sngXMax = sngCentreX + sngViewportWidth / 2
    sngYMax = sngCentreY + sngViewportHeight / 2
End Sub
