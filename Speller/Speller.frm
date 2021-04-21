VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSpeller 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Speller"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   Icon            =   "Speller.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4920
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optMode 
      Caption         =   "&Hangman"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
      Scrolling       =   1
   End
   Begin VB.OptionButton optMode 
      Caption         =   "&Anagram"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton optMode 
      Caption         =   "&Wildcard"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.OptionButton optMode 
      Caption         =   "&Correct spelling"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.ListBox lstSuggestions 
      Height          =   2400
      Left            =   2520
      TabIndex        =   7
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtWord 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "frmSpeller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private appWord As Application
Private dblProgress As Double
Private blnCancel  As Boolean

Private Sub cmdCancel_Click()
    blnCancel = True
End Sub

Private Sub cmdGo_Click()
    Dim varMode As Variant
    
    Screen.MousePointer = vbHourglass
    varMode = Switch(optMode(0).Value, wdSpellword, optMode(1).Value, wdWildcard, optMode(2).Value, wdAnagram, optMode(3).Value, wdWildcard)
    lstSuggestions.Clear
    
    If varMode = wdWildcard Then
        dblProgress = 0
        blnCancel = False
        cmdCancel.Enabled = True
        GetAllSuggestions txtWord.Text, lstSuggestions, 0
        cmdCancel.Enabled = False
        prgProgress.Value = 0
    Else
        GetSuggestions txtWord.Text, varMode, lstSuggestions
    End If
    
    If lstSuggestions.ListCount = 0 And Not blnCancel Then
        lstSuggestions.AddItem "(no suggestions)"
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub GetAllSuggestions(ByVal strWord As String, lstListBox As Control, intDepth As Integer)
    Dim i As Integer
    Dim bytPrefix As Byte
    Dim intDummy As Integer
    
    If blnCancel = True Then Exit Sub
    
    Do While Mid(strWord, i + 1, 1) = "?" And i < Len(strWord)
        i = i + 1
    Loop
    
    If i > 0 Then
        For bytPrefix = Asc("a") To Asc("z")
            Mid(strWord, i, 1) = Chr(bytPrefix)
            GetAllSuggestions strWord, lstSuggestions, intDepth + 1
        Next bytPrefix
    Else
        GetSuggestions strWord, wdWildcard, lstSuggestions
        dblProgress = dblProgress + 100 / (26 ^ intDepth)
        If dblProgress > 100 Then dblProgress = 100
        prgProgress.Value = dblProgress
        intDummy = DoEvents()
    End If
End Sub

Private Sub GetSuggestions(strWord As String, varMode As Variant, lstListBox As Control)
    Dim sugList As SpellingSuggestions
    Dim sug As SpellingSuggestion
    Dim i As Integer
    
'    On Error GoTo ErrorHandler
    Set sugList = appWord.GetSpellingSuggestions(strWord, , , , Switch(optMode(0).Value, wdSpellword, optMode(1).Value, wdWildcard, optMode(2).Value, wdAnagram, optMode(3).Value, wdWildcard))
    
    For Each sug In sugList
        If optMode(3).Value Then
            For i = 1 To Len(sug.Name)
                If Mid(txtWord.Text, i, 1) = "?" And InStr(1, txtWord.Text, Mid(sug.Name, i, 1), vbTextCompare) > 0 Then Exit For
            Next i
            If i > Len(sug.Name) Then lstListBox.AddItem sug.Name
        Else
            lstListBox.AddItem sug.Name
        End If
    Next
    
'    Exit Sub
'ErrorHandler:
'    MsgBox Err.Description, vbCritical
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Set appWord = GetObject(, "Word.Application")
    If appWord Is Nothing Then
        Set appWord = CreateObject("Word.Application")
        If appWord Is Nothing Then
            MsgBox "Unable to Start Microsoft Word"
            Exit Sub
        End If
    End If
    appWord.Documents.Add
End Sub

Private Sub Form_Unload(Cancel As Integer)
    appWord.Documents.Close wdDoNotSaveChanges
    appWord.Quit
End Sub
