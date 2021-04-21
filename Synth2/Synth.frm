VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSynth 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Synth"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   528
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "S&catter"
      Height          =   375
      Index           =   3
      Left            =   6840
      TabIndex        =   6
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "S&quare"
      Height          =   375
      Index           =   2
      Left            =   5760
      TabIndex        =   5
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "S&ine"
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   4
      Top             =   4200
      Width           =   975
   End
   Begin VB.Timer tmrPlayWave 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "S&awtooth"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   3
      Top             =   4200
      Width           =   975
   End
   Begin VB.PictureBox picWave 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404000&
      ForeColor       =   &H0000FF00&
      Height          =   3900
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   0
      Top             =   120
      Width           =   7695
   End
   Begin MSComctlLib.Slider sldSmoothing 
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   4080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   8
      Max             =   64
      SelStart        =   16
      TickFrequency   =   8
      Value           =   16
   End
   Begin VB.Label lblSmoothing 
      Caption         =   "&Smoothing"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   855
   End
End
Attribute VB_Name = "frmSynth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bytSamples(511) As Byte
Private blnWaveChanged

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const PI = 3.14159265359

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Private Type WaveHeader
    Riff As String * 4
    LenFileMinus8 As Long
    Wave As String * 4
    Fmt As String * 4
    LenFmtData As Long
    FormatTag As Integer
    Channels As Integer
    SampleRate As Long
    BytesPerSecond As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    Data As String * 4
    LenDataBlock As Long
End Type

'*****************************************************
' Purpose:  Writes the standard wave file header to an
'           open file.
' Inputs:
'   intFileNumber:  the file number of the open file
'   lngLenData:     the length of the sound data
'   intFormat:      the format of the sound data 1=PCM
'   intChannels:    the number of channels, 1 = mono
'   lngSampleRate:  the sample rate, eg. 44100
'   intBits:        the number of bits/sample, 8 or 16
'*****************************************************

Private Sub PutWaveHeader(intFileNumber As Integer, lngLenData As Long, intFormat As Integer, intChannels As Integer, lngSampleRate As Long, intBits As Integer)
    Dim udtWaveHeader As WaveHeader
    
    With udtWaveHeader
        .Riff = "RIFF"
        .Wave = "WAVE"
        .Fmt = "fmt "
        .LenFmtData = 16
        .Data = "data"
        
        .LenFileMinus8 = lngLenData + 36
        .FormatTag = intFormat
        .Channels = intChannels
        .SampleRate = lngSampleRate
        .BlockAlign = intChannels * intBits / 8
        .BytesPerSecond = lngSampleRate * .BlockAlign
        .BitsPerSample = intBits
        .LenDataBlock = lngLenData
    End With
    
    Put intFileNumber, , udtWaveHeader
End Sub

'*****************************************************
' Purpose:  Reads a standard wave file header from an
'           open file.
' Effects:  Raises error number 513 if the header is
'           not a valid standard wave file header.
' Inputs:
'   intFileNumber:  the file number of the open file
' Returns:  A user-defined type containing all of the
'           data in the wave file header.
'*****************************************************

Private Function GetWaveHeader(intFileNumber As Integer) As WaveHeader
    Dim udtWaveHeader As WaveHeader

    Get intFileNumber, , udtWaveHeader
    GetWaveHeader = udtWaveHeader
    With udtWaveHeader
        If .Riff <> "RIFF" Or _
        .Wave <> "WAVE" Or _
        .Fmt <> "fmt " Or _
        .LenFmtData <> 16 Or _
        .Data <> "data" Then Err.Raise 513, , "Not a valid wave file"
    End With
End Function

Private Sub KillSound()
    Dim lngSuccess As Long
    lngSuccess = sndPlaySound(vbNullString, 0)
End Sub

Private Sub cmdReset_Click(Index As Integer)
    ResetWave (Index)
End Sub

Private Sub Form_Load()
    ResetWave 0
    
'    Dim udtWaveHeader As WaveHeader
'    On Error Resume Next
'    Open "c:\windows\desktop\temp.wav" For Binary Access Read As 1
'    udtWaveHeader = GetWaveHeader(1)
'    If Err.Number Then MsgBox Err.Description, vbCritical
'    Close 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillSound
End Sub

Private Sub picWave_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then ChangePoint X, Y
End Sub

Private Sub picWave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then ChangePoint X, Y
End Sub

Private Sub ChangePoint(X As Single, Y As Single)
    Dim i As Integer, j As Integer
    Dim intSmooth As Integer
    Dim sngDistance As Single
    
    If Y < 0 Or Y > 255 Then Exit Sub
    intSmooth = sldSmoothing.Value
    For i = X - intSmooth To X + intSmooth
        j = i
        If j < 0 Then j = j + 512
        If j >= 512 Then j = j - 512
        picWave.PSet (j, 255 - bytSamples(j)), picWave.BackColor
        If intSmooth = 0 Then
            bytSamples(j) = (255 - Y)
        Else
            bytSamples(j) = bytSamples(j) + (255 - Y - bytSamples(j)) * Exp(-((i - X) / intSmooth * 2) ^ 2)
        End If
        picWave.PSet (j, 255 - bytSamples(j))
    Next i
    
    blnWaveChanged = True
End Sub

Private Sub ResetWave(intStyle As Integer)
    Dim i As Integer
    
    picWave.Cls
    For i = 0 To 511
        Select Case intStyle
            Case 0
                bytSamples(i) = i Mod 256
            Case 1
                bytSamples(i) = 127.5 * (Sin(i / 512 * 4 * PI) + 1)
            Case 2
                bytSamples(i) = i Mod 256 < 128
            Case 3
                bytSamples(i) = Int(256 * Rnd)
        End Select
        picWave.PSet (i, 255 - bytSamples(i))
    Next i
    blnWaveChanged = True
End Sub

Private Sub PlayWave()
    Dim lngUFlags As Long
    Dim lngSuccess As Long

    KillSound
    Open "c:\windows\desktop\temp.wav" For Binary Access Write As 1
    PutWaveHeader 1, 512, 1, 1, 44100, 8
    Put #1, , bytSamples
    Close 1
    lngUFlags = SND_NODEFAULT Or SND_ASYNC Or SND_LOOP
    lngSuccess = sndPlaySound("c:\windows\desktop\temp.wav", lngUFlags)
End Sub

Private Sub tmrPlayWave_Timer()
    If blnWaveChanged Then
        PlayWave
        blnWaveChanged = False
    End If
End Sub
