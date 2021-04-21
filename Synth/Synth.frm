VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSynth 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   11
      Left            =   3240
      TabIndex        =   19
      Top             =   4920
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   10
      Left            =   3240
      TabIndex        =   18
      Top             =   5520
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   9
      Left            =   3240
      TabIndex        =   17
      Top             =   6120
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   8
      Left            =   3240
      TabIndex        =   16
      Top             =   6720
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   7
      Left            =   3240
      TabIndex        =   15
      Top             =   4320
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   6
      Left            =   3240
      TabIndex        =   14
      Top             =   3720
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   5
      Left            =   3240
      TabIndex        =   13
      Top             =   3120
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   4
      Left            =   3240
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   2
      Left            =   3240
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   1
      Left            =   3240
      TabIndex        =   9
      Top             =   720
      Width           =   975
   End
   Begin VB.CheckBox chkSign 
      Caption         =   "Negative"
      Height          =   495
      Index           =   0
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   975
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   13
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   6
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   7
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   8
      Left            =   120
      TabIndex        =   20
      Top             =   4920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   5520
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Top             =   6120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
   Begin MSComctlLib.Slider sliAmplitude 
      Height          =   495
      Index           =   11
      Left            =   120
      TabIndex        =   23
      Top             =   6720
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
      _Version        =   393216
      LargeChange     =   2
      Min             =   1
      Max             =   13
      SelStart        =   12
      Value           =   13
   End
End
Attribute VB_Name = "frmSynth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_LOOP = &H8
Private Const PI = 3.14159265359
Private Const NUM_HARMONICS = 12

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

Private Sub WriteWaveHeader(intFileNumber As Integer, lngLenData As Long, intFormat As Integer, intChannels As Integer, lngSampleRate As Long, intBits As Integer)
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

Private Sub KillSound()
    Dim lngSuccess As Long
    lngSuccess = sndPlaySound(vbNullString, 0)
End Sub

Private Sub chkSign_Click(Index As Integer)
    PlayWaveform
End Sub

Private Sub Form_Unload(Cancel As Integer)
    KillSound
End Sub

Private Sub PlayWaveform()
    Dim sngWaveform(255) As Single
    Dim i As Integer, j As Integer
    Dim sngOrdinate As Single
    Dim sngMax As Single
    Dim strByte As String
    Dim lngUFlags As Long
    Dim lngSuccess As Long
    
    For i = 0 To 255
        sngOrdinate = 0
        For j = 0 To NUM_HARMONICS - 1
            'sngOrdinate = sngOrdinate + sliAmplitude(j).Value * Sin((j + 1) * i / 256 * 2 * PI)
            If sliAmplitude(j).Value < sliAmplitude(j).Max Then sngOrdinate = sngOrdinate - (2 * chkSign(j).Value - 1) * Sin((j + 1) * i / 256 * 2 * PI) / sliAmplitude(j).Value
        Next j
        sngWaveform(i) = sngOrdinate
        If Abs(sngOrdinate) > sngMax Then sngMax = Abs(sngOrdinate)
    Next i
    
    If sngMax = 0 Then sngMax = 1
    KillSound
    Open "c:\windows\desktop\temp.wav" For Binary Access Write As 1
    WriteWaveHeader 1, 256, 1, 1, 22050, 8
    
    For i = 0 To 255
        strByte = Chr(sngWaveform(i) / sngMax * 127.5 + 127.5)
'        strByte = Chr(sngWaveform(i) / 8 + 128)
        Put #1, , strByte
    Next i
    
    Close 1
    lngUFlags = SND_NODEFAULT Or SND_ASYNC Or SND_LOOP
    lngSuccess = sndPlaySound("c:\windows\desktop\temp.wav", lngUFlags)
End Sub

Private Sub sliAmplitude_Change(Index As Integer)
    PlayWaveform
End Sub
