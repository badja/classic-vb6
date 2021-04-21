VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMandelbrot 
   Caption         =   "Mandelbrot Set"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Slider sldMaxIterations 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3836
      _Version        =   393216
      Orientation     =   1
      Max             =   100
      SelStart        =   25
      TickFrequency   =   10
      Value           =   25
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.PictureBox picView 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ClipControls    =   0   'False
      Height          =   3015
      Left            =   1560
      ScaleHeight     =   197
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   197
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label lblMaxIterations 
      Caption         =   "&Max. iterations"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmMandelbrot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type Complex
    Real As Single
    Imag As Single
End Type

Private Function InSet(c As Complex) As Boolean
    Dim z As Complex
    Dim blnInfinity As Boolean
    Dim intIterations As Integer
    
    Do Until blnInfinity Or intIterations = 100
        z = ComplexSum(ComplexProduct(z, z), c)
        blnInfinity = ToInfinity(z)
        intIterations = intIterations + 1
    Loop
    
    InSet = Not blnInfinity
End Function

Private Function GetColor(c As Complex) As Long
    Dim z As Complex
    Dim intIterations As Integer
    Dim intMaxIterations As Integer
    Dim blnInfinity As Boolean
    
    intMaxIterations = sldMaxIterations.Value
    blnInfinity = ToInfinity(z)
    
    Do Until blnInfinity Or intIterations = intMaxIterations
        z = ComplexSum(ComplexProduct(z, z), c)
        blnInfinity = ToInfinity(z)
        intIterations = intIterations + 1
    Loop
    
    If blnInfinity Then GetColor = intIterations * 255 / intMaxIterations
End Function

Private Function ComplexSum(a As Complex, b As Complex) As Complex
    ComplexSum.Real = a.Real + b.Real
    ComplexSum.Imag = a.Imag + b.Imag
End Function

Private Function ComplexProduct(a As Complex, b As Complex) As Complex
    ComplexProduct.Real = a.Real * b.Real - a.Imag * b.Imag
    ComplexProduct.Imag = a.Real * b.Imag + a.Imag * b.Real
End Function

Private Function ToInfinity(z As Complex) As Boolean
    ToInfinity = Abs(z.Real) > 2 Or Abs(z.Imag) > 2
End Function

Private Sub cmdGenerate_Click()
    Dim i As Integer, j As Integer
    Dim z As Complex
    Dim intWidth As Integer, intHeight As Integer
    
    Screen.MousePointer = vbHourglass
    intWidth = picView.ScaleWidth
    intHeight = picView.ScaleHeight
    
    For i = 0 To intHeight - 1
        For j = 0 To intWidth - 1
            z.Real = -2 + j * 3 / intWidth
            z.Imag = -1.5 + i * 3 / intHeight
            'If InSet(z) Then picView.PSet (j, i)
            picView.PSet (j, i), GetColor(z)
        Next j
    Next i
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    Dim intXDiff As Integer, intYDiff As Integer
    
    On Error Resume Next
    intXDiff = Width - picView.Width - 1905
    intYDiff = Height - picView.Height - 705
    
    If intXDiff < intYDiff Then
        picView.Move 1560, 120, Width - 1905, Width - 1905
    Else
        picView.Move 1560, 120, Height - 705, Height - 705
    End If
End Sub
