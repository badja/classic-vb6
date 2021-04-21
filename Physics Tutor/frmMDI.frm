VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Physics Tutor"
   ClientHeight    =   7230
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9810
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileChangeUser 
         Caption         =   "&Change User..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTopics 
      Caption         =   "To&pics"
      Begin VB.Menu mnuTopicsProgressionChart 
         Caption         =   "&Progression Chart"
      End
      Begin VB.Menu mnuTopicsSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTopicsWaves 
         Caption         =   "&Waves"
         Begin VB.Menu mnuTopicsWavesBackground 
            Caption         =   "&Background Information"
         End
         Begin VB.Menu mnuTopicsWavesSinglePulse 
            Caption         =   "Single &Pulse"
         End
         Begin VB.Menu mnuTopicsWavesStanding 
            Caption         =   "&Standing"
         End
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuToolsFrequencySpectrum 
         Caption         =   "&Frequency Spectrum"
      End
      Begin VB.Menu mnuToolsVectors 
         Caption         =   "&Vectors"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents..."
      End
      Begin VB.Menu mnuHelpIndex 
         Caption         =   "&Index..."
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Physics Tutor..."
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
