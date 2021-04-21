VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmTileEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TileEditor"
   ClientHeight    =   6240
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10560
   Icon            =   "TileEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   416
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   704
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrPreventEdit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   1440
   End
   Begin MSComDlg.CommonDialog dlgTileSet 
      Left            =   0
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Load Tile Set"
      Filter          =   "Bitmaps (*.bmp)|*.bmp|All Files (*.*)|*.*"
   End
   Begin PicClip.PictureClip clpTiles 
      Left            =   9360
      Top             =   840
      _ExtentX        =   1905
      _ExtentY        =   397
      _Version        =   393216
      Cols            =   58
   End
   Begin VB.ComboBox cmbSpecial 
      Height          =   315
      Left            =   9360
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "java"
      DialogTitle     =   "Export Java Fragment"
      Filter          =   "Java Files (*.java;*.jav)|*.java;*.jav|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "rtm"
      Filter          =   "Racing Track Maps (*.rtm)|*.rtm|All Files (*.*)|*.*"
   End
   Begin VB.VScrollBar vsbTracks 
      Height          =   1935
      LargeChange     =   4
      Left            =   10200
      Max             =   0
      TabIndex        =   21
      Top             =   4200
      Width           =   255
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      ClipControls    =   0   'False
      Height          =   1980
      Left            =   8160
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   128
      TabIndex        =   20
      Top             =   4200
      Width           =   1980
   End
   Begin VB.VScrollBar vsbTiles 
      Height          =   2895
      LargeChange     =   192
      Left            =   10200
      SmallChange     =   64
      TabIndex        =   19
      Top             =   1200
      Width           =   255
   End
   Begin VB.PictureBox picTilePort 
      Height          =   2895
      Left            =   8160
      ScaleHeight     =   189
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   16
      Top             =   1200
      Width           =   2055
      Begin VB.PictureBox picTiles 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   0
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   137
         TabIndex        =   17
         Top             =   0
         Width           =   2055
         Begin VB.OptionButton optTile 
            Alignment       =   1  'Right Justify
            Height          =   975
            Index           =   0
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   0
            Width           =   975
         End
      End
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   11
      Left            =   10080
      Picture         =   "TileEditor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Checkpoint"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   10
      Left            =   9720
      Picture         =   "TileEditor.frx":0497
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Starting Point"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   9
      Left            =   9360
      Picture         =   "TileEditor.frx":04F7
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Powerup"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   2
      Left            =   8880
      Picture         =   "TileEditor.frx":0552
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "AI Guide"
      Top             =   720
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   3
      Left            =   8520
      Picture         =   "TileEditor.frx":05AA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "AI Guide"
      Top             =   720
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   4
      Left            =   8160
      Picture         =   "TileEditor.frx":0604
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "AI Guide"
      Top             =   720
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   1
      Left            =   8880
      Picture         =   "TileEditor.frx":065C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "AI Guide"
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   0
      Left            =   8520
      Picture         =   "TileEditor.frx":06B1
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Delete AI Guide"
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   5
      Left            =   8160
      Picture         =   "TileEditor.frx":070A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "AI Guide"
      Top             =   360
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   8
      Left            =   8880
      Picture         =   "TileEditor.frx":075F
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "AI Guide"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   7
      Left            =   8520
      Picture         =   "TileEditor.frx":07B7
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "AI Guide"
      Top             =   0
      Width           =   375
   End
   Begin VB.OptionButton optTool 
      Height          =   375
      Index           =   6
      Left            =   8160
      Picture         =   "TileEditor.frx":0812
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "AI Guide"
      Top             =   0
      Width           =   375
   End
   Begin VB.VScrollBar vsbMap 
      Height          =   5775
      LargeChange     =   64
      Left            =   7800
      Max             =   2048
      SmallChange     =   8
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar hsbMap 
      Height          =   255
      LargeChange     =   64
      Left            =   120
      Max             =   2048
      SmallChange     =   8
      TabIndex        =   2
      Top             =   5880
      Width           =   7695
   End
   Begin VB.PictureBox picViewport 
      Height          =   5760
      Left            =   120
      ScaleHeight     =   380
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   0
      Top             =   120
      Width           =   7680
      Begin VB.PictureBox picMap 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   30720
         Left            =   0
         ScaleHeight     =   2048
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   2048
         TabIndex        =   1
         Top             =   0
         Width           =   30720
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuTileSet 
         Caption         =   "&Load Tile Set..."
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppend 
         Caption         =   "&Append Tracks..."
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export Java Fragment..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTrack 
      Caption         =   "&Track"
      Begin VB.Menu mnuGrid 
         Caption         =   "&Grid"
         Checked         =   -1  'True
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInsertBefore 
         Caption         =   "Insert Track &Before Current"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuInsertAfter 
         Caption         =   "Insert Track &After Current"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete Current Track"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveUp 
         Caption         =   "Move Track &Up"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuMoveDown 
         Caption         =   "Move Track D&own"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWidth 
         Caption         =   "Change &Width..."
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuHeight 
         Caption         =   "Change &Height..."
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpInstructions 
         Caption         =   "&Instructions..."
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About TileEditor..."
      End
   End
End
Attribute VB_Name = "frmTileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const intMinWidth = 8
Const intMinHeight = 8
Const intMaxWidth = 128
Const intMaxHeight = 128

Dim NumTiles As Integer
Dim Tile As Integer
Dim Sel As Integer
Dim SelTool As Integer
Dim intMap() As Integer
Dim intArrow() As Integer
Dim intSpecial() As Integer
Dim blnCancelMove As Boolean
Dim blnCancelRepaint As Boolean
Dim blnPreventEdit As Boolean
Dim intMoveX As Integer
Dim intMoveY As Integer
Dim intWidth() As Integer
Dim intHeight() As Integer
Dim intTrack As Integer
Dim intNumTracks As Integer
Dim sngScaledSize As Single
Dim intLayer As Integer
Dim spcArrow(7) As StdPicture

Private Sub Form_Load()
    Dim i As Integer
    
    dlgFile.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    dlgExport.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    dlgTileSet.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    
    ReDim intMap(intMaxWidth - 1, intMaxHeight - 1, 0)
    ReDim intArrow(intMaxWidth - 1, intMaxHeight - 1, 0)
    ReDim intSpecial(intMaxWidth - 1, intMaxHeight - 1, 0)
    ReDim intWidth(0)
    ReDim intHeight(0)
    intWidth(0) = intMinWidth
    intHeight(0) = intMinHeight
    intTrack = 0
    intNumTracks = 1
    
    CalculateScaledSize
    
    LoadTileSet App.Path & "\default.bmp"
    
    For i = 0 To 7
        Set spcArrow(i) = LoadPicture(App.Path & "\arrow" & i + 1 & ".bmp")
    Next i
    
    PaintPreview
End Sub

Private Sub hsbMap_Change()
    picMap.Left = -hsbMap.Value
End Sub

Private Sub hsbMap_Scroll()
    picMap.Left = -hsbMap.Value
End Sub

Private Sub mnuAppend_Click()
    Dim intMap2() As Integer
    Dim intArrow2() As Integer
    Dim intSpecial2() As Integer
    Dim intWidth2() As Integer
    Dim intHeight2() As Integer
    Dim intNumTracks2 As Integer
    Dim i As Integer, j As Integer, k As Integer
    
    On Error Resume Next
    blnPreventEdit = True
    dlgFile.ShowOpen
    tmrPreventEdit.Enabled = True
    If Err.Number Then Exit Sub
    Open dlgFile.Filename For Binary Access Read As 1
    Get #1, , intNumTracks2
    ReDim intMap2(intMaxWidth - 1, intMaxHeight - 1, intNumTracks2 - 1)
    ReDim intArrow2(intMaxWidth - 1, intMaxHeight - 1, intNumTracks2 - 1)
    ReDim intSpecial2(intMaxWidth - 1, intMaxHeight - 1, intNumTracks2 - 1)
    ReDim intWidth2(intNumTracks2 - 1)
    ReDim intHeight2(intNumTracks2 - 1)
    Get #1, , intWidth2
    Get #1, , intHeight2
    Get #1, , intMap2
    Get #1, , intArrow2
    Get #1, , intSpecial2
    Close 1
    
    ReDim Preserve intMap(intMaxWidth - 1, intMaxHeight - 1, intNumTracks + intNumTracks2 - 1)
    ReDim Preserve intArrow(intMaxWidth - 1, intMaxHeight - 1, intNumTracks + intNumTracks2 - 1)
    ReDim Preserve intSpecial(intMaxWidth - 1, intMaxHeight - 1, intNumTracks + intNumTracks2 - 1)
    ReDim Preserve intWidth(intNumTracks + intNumTracks2 - 1)
    ReDim Preserve intHeight(intNumTracks + intNumTracks2 - 1)
    
    For i = 0 To intNumTracks2 - 1
        For j = 0 To intMaxWidth - 1
            For k = 0 To intMaxHeight - 1
                intMap(j, k, i + intNumTracks) = intMap2(j, k, i)
                intArrow(j, k, i + intNumTracks) = intArrow2(j, k, i)
                intSpecial(j, k, i + intNumTracks) = intSpecial2(j, k, i)
            Next k
        Next j
        intWidth(i + intNumTracks) = intWidth2(i)
        intHeight(i + intNumTracks) = intHeight2(i)
    Next i
    
    intNumTracks = intNumTracks + intNumTracks2
    vsbTracks.Max = intNumTracks - 1
End Sub

Private Sub mnuDelete_Click()
    Dim i As Integer, j As Integer, k As Integer
    
    If intNumTracks = 1 Then
        MsgBox "There cannot be less that one track.", vbCritical
    ElseIf MsgBox("Are you sure you want to delete this track?", vbYesNo + vbQuestion) = vbYes Then
        For i = intTrack To intNumTracks - 2
            intWidth(i) = intWidth(i + 1)
            intHeight(i) = intHeight(i + 1)
            For j = 0 To intMaxHeight - 1
                For k = 0 To intMaxWidth - 1
                    intMap(k, j, i) = intMap(k, j, i + 1)
                    intArrow(k, j, i) = intArrow(k, j, i + 1)
                    intSpecial(k, j, i) = intSpecial(k, j, i + 1)
                Next k
            Next j
        Next i
    
        intNumTracks = intNumTracks - 1
        ReDim Preserve intWidth(intNumTracks - 1)
        ReDim Preserve intHeight(intNumTracks - 1)
        ReDim Preserve intMap(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
        ReDim Preserve intArrow(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
        ReDim Preserve intSpecial(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
        If intTrack = intNumTracks Then intTrack = intTrack - 1
        
        blnCancelRepaint = True
        vsbTracks.Max = intNumTracks - 1
        CalculateScaledSize
        PaintPreview
        PaintMap
    End If
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuExport_Click()
    Dim i As Integer, j As Integer, k As Integer
    Dim intAmount As Integer
    Dim intMax As Integer
    
    On Error Resume Next
    dlgExport.ShowSave
    If Err.Number Then Exit Sub
    Open dlgExport.Filename For Output Access Write As 1
    
    Print #1, Chr(9); "static final byte tracks[][][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            Print #1, "{";
            For k = 0 To intHeight(i) - 2
                Print #1, Format(intMap(j, k, i)); ", ";
            Next k
            Print #1, Format(intMap(j, intHeight(i) - 1, i)); "}";
            If j < intWidth(i) - 1 Then
                Print #1, ","
                Print #1, Chr(9); Chr(9);
            End If
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte arrows[][][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            Print #1, "{";
            For k = 0 To intHeight(i) - 2
                Print #1, Format(intArrow(j, k, i)); ", ";
            Next k
            Print #1, Format(intArrow(j, intHeight(i) - 1, i)); "}";
            If j < intWidth(i) - 1 Then
                Print #1, ","
                Print #1, Chr(9); Chr(9);
            End If
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
'    Print #1, Chr(9); "static final byte special[][][] ="
'    Print #1, Chr(9); Chr(9); "{";
'    For i = 0 To intNumTracks - 1
'        Print #1, "{";
'        For j = 0 To intWidth(i) - 1
'            Print #1, "{";
'            For k = 0 To intHeight(i) - 2
'                Print #1, Format(intSpecial(j, k, i)); ", ";
'            Next k
'            Print #1, Format(intSpecial(j, intHeight(i) - 1, i)); "}";
'            If j < intWidth(i) - 1 Then
'                Print #1, ","
'                Print #1, Chr(9); Chr(9);
'            End If
'        Next j
'        Print #1, "}";
'        If i < intNumTracks - 1 Then
'            Print #1, ","
'            Print #1, Chr(9); Chr(9);
'        End If
'    Next i
'    Print #1, "};"
    
    Print #1, Chr(9); "static final byte powerupsX[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intAmount = 0
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) = 1 Then
                    If intAmount > 0 Then Print #1, ", ";
                    Print #1, Format(j);
                    intAmount = intAmount + 1
                End If
            Next k
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte powerupsY[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intAmount = 0
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) = 1 Then
                    If intAmount > 0 Then Print #1, ", ";
                    Print #1, Format(k);
                    intAmount = intAmount + 1
                End If
            Next k
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte startPosX[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        Print #1, "{"; Format(FindStartPosX(0, i)); ", ";
        Print #1, Format(FindStartPosX(1, i)); ", ";
        Print #1, Format(FindStartPosX(2, i)); ", ";
        Print #1, Format(FindStartPosX(3, i)); "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte startPosY[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        Print #1, "{"; Format(FindStartPosY(0, i)); ", ";
        Print #1, Format(FindStartPosY(1, i)); ", ";
        Print #1, Format(FindStartPosY(2, i)); ", ";
        Print #1, Format(FindStartPosY(3, i)); "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte checkpointsX[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intAmount = 0
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) >= 6 Then
                    If intAmount > 0 Then Print #1, ", ";
                    Print #1, Format(j);
                    intAmount = intAmount + 1
                End If
            Next k
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte checkpointsY[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intAmount = 0
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) >= 6 Then
                    If intAmount > 0 Then Print #1, ", ";
                    Print #1, Format(k);
                    intAmount = intAmount + 1
                End If
            Next k
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte checkpointsZ[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intAmount = 0
        Print #1, "{";
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) > 5 Then
                    If intAmount > 0 Then Print #1, ", ";
                    Print #1, Format(intSpecial(j, k, i) - 6);
                    intAmount = intAmount + 1
                End If
            Next k
        Next j
        Print #1, "}";
        If i < intNumTracks - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "static final byte numCheckpoints[] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To intNumTracks - 1
        intMax = 0
        For j = 0 To intWidth(i) - 1
            For k = 0 To intHeight(i) - 1
                If intSpecial(j, k, i) > 5 And intSpecial(j, k, i) - 5 > intMax Then intMax = intSpecial(j, k, i) - 5
            Next k
        Next j
        Print #1, Format(intMax);
        If i < intNumTracks - 1 Then Print #1, ", ";
    Next i
    Print #1, "};"
    
    Close 1
End Sub

Private Sub mnuGrid_Click()
    mnuGrid.Checked = Not mnuGrid.Checked
    PaintMap
End Sub

Private Sub mnuHeight_Click()
    Dim strResponse As String
    
    strResponse = InputBox("Enter new track height (in tiles):", , intHeight(intTrack))
    
    If strResponse <> "" Then
        If Val(strResponse) < intMinHeight Or Val(strResponse) > intMaxHeight Then
            MsgBox "The height must range from " & intMinHeight & " to " & intMaxHeight & ".", vbExclamation
        Else
            intHeight(intTrack) = strResponse
            CalculateScaledSize
            PaintPreview
        End If
    End If
End Sub

Private Sub mnuHelpAbout_Click()
    MsgBox "TileEditor version 1.11" & vbNewLine & "for Handy Games" & vbNewLine & vbNewLine & "Copyright (C) 2000 Adrian Grucza"
End Sub

Private Sub mnuHelpInstructions_Click()
    MsgBox "The terrain tiles on the right let you create the landscape (trees, snow, etc...). The arrow buttons above this let you tell the computer-controlled racers where to go. These are necessary. The other three buttons to the right of these are for inserting powerups, starting locations, and checkpoints. These buttons work in conjunction with the dropdown combo box below them. See the sample track for an idea of how to use the arrows and checkpoints." & vbNewLine & vbNewLine & "Use the left mouse button to draw and the right mouse button to scroll. Drawing and scrolling can be done in both the large main window and the small preview window at the bottom-right. The scrollbar next to this preview window lets you switch between tracks. You can insert, delete, and change the size of tracks via the Track menu."
End Sub

Private Sub mnuInsertAfter_Click()
    Dim i As Integer, j As Integer, k As Integer
    
    ReDim Preserve intWidth(intNumTracks)
    ReDim Preserve intHeight(intNumTracks)
    ReDim Preserve intMap(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)
    ReDim Preserve intArrow(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)
    ReDim Preserve intSpecial(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)

    For i = intNumTracks To intTrack + 2 Step -1
        intWidth(i) = intWidth(i - 1)
        intHeight(i) = intHeight(i - 1)
        For j = 0 To intMaxHeight - 1
            For k = 0 To intMaxWidth - 1
                intMap(k, j, i) = intMap(k, j, i - 1)
                intArrow(k, j, i) = intArrow(k, j, i - 1)
                intSpecial(k, j, i) = intSpecial(k, j, i - 1)
            Next k
        Next j
    Next i

    intWidth(intTrack + 1) = intMinWidth
    intHeight(intTrack + 1) = intMinHeight
    For j = 0 To intMaxHeight - 1
        For k = 0 To intMaxWidth - 1
            intMap(k, j, intTrack + 1) = 0
            intArrow(k, j, intTrack + 1) = 0
            intSpecial(k, j, intTrack + 1) = 0
        Next k
    Next j
    
    vsbTracks.Max = intNumTracks
    intNumTracks = intNumTracks + 1
End Sub

Private Sub mnuInsertBefore_Click()
    Dim i As Integer, j As Integer, k As Integer
    
    ReDim Preserve intWidth(intNumTracks)
    ReDim Preserve intHeight(intNumTracks)
    ReDim Preserve intMap(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)
    ReDim Preserve intArrow(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)
    ReDim Preserve intSpecial(intMaxWidth - 1, intMaxHeight - 1, intNumTracks)

    For i = intNumTracks To intTrack + 1 Step -1
        intWidth(i) = intWidth(i - 1)
        intHeight(i) = intHeight(i - 1)
        For j = 0 To intMaxHeight - 1
            For k = 0 To intMaxWidth - 1
                intMap(k, j, i) = intMap(k, j, i - 1)
                intArrow(k, j, i) = intArrow(k, j, i - 1)
                intSpecial(k, j, i) = intSpecial(k, j, i - 1)
            Next k
        Next j
    Next i

    intWidth(intTrack) = intMinWidth
    intHeight(intTrack) = intMinHeight
    For j = 0 To intMaxHeight - 1
        For k = 0 To intMaxWidth - 1
            intMap(k, j, intTrack) = 0
            intArrow(k, j, intTrack) = 0
            intSpecial(k, j, intTrack) = 0
        Next k
    Next j
    
    blnCancelRepaint = True
    vsbTracks.Max = intNumTracks
    intNumTracks = intNumTracks + 1
    intTrack = intTrack + 1
    vsbTracks.Value = intTrack
End Sub

Private Sub mnuMoveDown_Click()
    If intTrack = intNumTracks - 1 Then
        MsgBox "The current track is already at the bottom.", vbExclamation
    Else
        SwapTracks intTrack, intTrack + 1
        blnCancelRepaint = True
        intTrack = intTrack + 1
        vsbTracks.Value = intTrack
    End If
End Sub

Private Sub mnuMoveUp_Click()
    If intTrack = 0 Then
        MsgBox "The current track is already at the top.", vbExclamation
    Else
        SwapTracks intTrack - 1, intTrack
        blnCancelRepaint = True
        intTrack = intTrack - 1
        vsbTracks.Value = intTrack
    End If
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next
    blnPreventEdit = True
    dlgFile.ShowOpen
    tmrPreventEdit.Enabled = True
    If Err.Number Then Exit Sub
    Open dlgFile.Filename For Binary Access Read As 1
    Get #1, , intNumTracks
    ReDim intMap(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
    ReDim intArrow(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
    ReDim intSpecial(intMaxWidth - 1, intMaxHeight - 1, intNumTracks - 1)
    ReDim intWidth(intNumTracks - 1)
    ReDim intHeight(intNumTracks - 1)
    Get #1, , intWidth
    Get #1, , intHeight
    Get #1, , intMap
    Get #1, , intArrow
    Get #1, , intSpecial
    Close 1
    intTrack = 0
    blnCancelRepaint = True
    vsbTracks.Max = intNumTracks - 1
    vsbTracks.Value = 0
    blnCancelRepaint = False
    CalculateScaledSize
    PaintPreview
    PaintMap
End Sub

Private Sub mnuSave_Click()
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number Then Exit Sub
    Open dlgFile.Filename For Binary Access Write As 1
    Put #1, , intNumTracks
    Put #1, , intWidth
    Put #1, , intHeight
    Put #1, , intMap
    Put #1, , intArrow
    Put #1, , intSpecial
    Close 1
End Sub

Private Sub mnuTileSet_Click()
    On Error Resume Next
    blnPreventEdit = True
    dlgTileSet.ShowOpen
    tmrPreventEdit.Enabled = True
    If Err.Number Then Exit Sub
    LoadTileSet dlgTileSet.Filename
    PaintPreview
    PaintMap
End Sub

Private Sub mnuWidth_Click()
    Dim strResponse As String
    
    strResponse = InputBox("Enter new track width (in tiles):", , intWidth(intTrack))
    
    If strResponse <> "" Then
        If Val(strResponse) < intMinWidth Or Val(strResponse) > intMaxWidth Then
            MsgBox "The width must range from " & intMinWidth & " to " & intMaxWidth & ".", vbExclamation
        Else
            intWidth(intTrack) = strResponse
            CalculateScaledSize
            PaintPreview
        End If
    End If
End Sub

Private Sub optTile_Click(Index As Integer)
    optTool(SelTool).Value = False
    Sel = Index
    If intLayer > 0 Then
        intLayer = 0
        PaintMap
    End If
    cmbSpecial.Visible = False
End Sub

Private Sub optTool_Click(Index As Integer)
    Dim intOldLayer As Integer
    Dim i As Integer
    
    optTile(Sel).Value = False
    SelTool = Index
    intOldLayer = intLayer
    With cmbSpecial
        If SelTool < 9 Then
            intLayer = 1
            .Visible = False
        Else
            intLayer = 2
            Select Case SelTool
                Case 9
                    .Visible = False
                Case 10
                    .Clear
                    For i = Asc("A") To Asc("D")
                        .AddItem "Player " & Chr(i)
                    Next i
                    .ListIndex = 0
                    .Visible = True
                Case 11
                    .Clear
                    For i = 0 To 9
                        .AddItem "Point " & i
                    Next i
                    .ListIndex = 0
                    .Visible = True
            End Select
        End If
    End With
    If intLayer <> intOldLayer Then PaintMap
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blnPreventEdit Then Exit Sub

    If Button And vbLeftButton Then DrawTile x \ 64, y \ 64
    If Button And vbRightButton Then
        intMoveX = x
        intMoveY = y
    End If
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If blnPreventEdit Then Exit Sub

    If blnCancelMove Then
        blnCancelMove = False
    Else
        If Button And vbLeftButton Then DrawTile x \ 64, y \ 64
        If Button And vbRightButton Then
            blnCancelMove = True
            MapMove x, y
        End If
    End If
End Sub

Private Sub picMap_Paint()
    PaintMap
End Sub

Private Sub picPreview_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbLeftButton Then DrawTile x / sngScaledSize, y / sngScaledSize
    If Button And vbRightButton Then PreviewMove x, y
End Sub

Private Sub picPreview_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button And vbLeftButton Then DrawTile x / sngScaledSize, y / sngScaledSize
    If Button And vbRightButton Then PreviewMove x, y
End Sub

Private Sub tmrPreventEdit_Timer()
    tmrPreventEdit.Enabled = False
    blnPreventEdit = False
End Sub

Private Sub vsbMap_Change()
    picMap.Top = -vsbMap.Value
End Sub

Private Sub vsbMap_Scroll()
    picMap.Top = -vsbMap.Value
End Sub

Private Sub PaintMap()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    
    For i = -picMap.Top \ 64 To (picViewport.ScaleHeight - picMap.Top - 1) \ 64
        For j = -picMap.Left \ 64 To (picViewport.ScaleWidth - picMap.Left - 1) \ 64
            PaintTile optTile(intMap(j, i, intTrack)).picture, j * 64, i * 64
            Select Case intLayer
                Case 1
                    If intArrow(j, i, intTrack) > 0 Then picMap.PaintPicture spcArrow(intArrow(j, i, intTrack) - 1), j * 64, i * 64, , , , , , , vbSrcInvert
                Case 2
                    picMap.CurrentX = j * 64 + 8
                    picMap.CurrentY = i * 64 - 4
                    Select Case intSpecial(j, i, intTrack)
                        Case 1
                            picMap.Print "P"
                        Case 2 To 5
                            picMap.Print Chr(Asc("A") + intSpecial(j, i, intTrack) - 2)
                        Case 6 To 15
                            picMap.CurrentX = picMap.CurrentX - 12
                            picMap.Print intSpecial(j, i, intTrack) - 6
                    End Select
            End Select
        Next j
    Next i
End Sub

Private Sub PaintPreview()
    Dim i As Integer, j As Integer
    
    On Error Resume Next
    Screen.MousePointer = vbHourglass
    picPreview.Cls
    
    For i = 0 To intHeight(intTrack) - 1
        For j = 0 To intWidth(intTrack) - 1
            picPreview.PaintPicture optTile(intMap(j, i, intTrack)).picture, j * sngScaledSize, i * sngScaledSize, sngScaledSize, sngScaledSize
        Next j
    Next i

    Screen.MousePointer = vbDefault
End Sub

Private Sub vsbTiles_Change()
    picTiles.Top = -vsbTiles.Value
End Sub

Private Sub vsbTiles_Scroll()
    picTiles.Top = -vsbTiles.Value
End Sub

Private Sub MapMove(ByVal x As Integer, ByVal y As Integer)
    Dim intX As Integer, intY As Integer
    
    intX = intMoveX - x - picMap.Left
    intY = intMoveY - y - picMap.Top
    
    If intX < 0 Then intX = 0
    If intY < 0 Then intY = 0
    If intX > hsbMap.Max Then intX = hsbMap.Max
    If intY > vsbMap.Max Then intY = vsbMap.Max
    
    hsbMap.Value = intX
    vsbMap.Value = intY
End Sub

Private Sub PreviewMove(ByVal x As Integer, ByVal y As Integer)
    Dim intX As Integer, intY As Integer
    
    intX = x * 64 / sngScaledSize - picViewport.ScaleWidth / 2
    intY = y * 64 / sngScaledSize - picViewport.ScaleHeight / 2
    
    If intX < 0 Then intX = 0
    If intY < 0 Then intY = 0
    If intX > hsbMap.Max Then intX = hsbMap.Max
    If intY > vsbMap.Max Then intY = vsbMap.Max
    
    hsbMap.Value = intX
    vsbMap.Value = intY
End Sub

Private Sub DrawTile(intX As Integer, intY As Integer)
    If intX < 0 Then intX = 0
    If intY < 0 Then intY = 0
    If intX >= intWidth(intTrack) Then intX = intWidth(intTrack) - 1
    If intY >= intHeight(intTrack) Then intY = intHeight(intTrack) - 1
    
    Select Case intLayer
        Case 0
            intMap(intX, intY, intTrack) = Sel
            PaintTile optTile(Sel).picture, intX * 64, intY * 64
            picPreview.PaintPicture optTile(Sel).picture, intX * sngScaledSize, intY * sngScaledSize, sngScaledSize, sngScaledSize
        Case 1
            If intArrow(intX, intY, intTrack) <> SelTool Then
                intArrow(intX, intY, intTrack) = SelTool
                PaintTile optTile(intMap(intX, intY, intTrack)).picture, intX * 64, intY * 64
                If SelTool > 0 Then picMap.PaintPicture spcArrow(SelTool - 1), intX * 64, intY * 64, , , , , , , vbSrcInvert
            End If
        Case 2
            PaintTile optTile(intMap(intX, intY, intTrack)).picture, intX * 64, intY * 64
            picMap.CurrentX = intX * 64 + 8
            picMap.CurrentY = intY * 64 - 4
            Select Case SelTool
                Case 9
                    If intSpecial(intX, intY, intTrack) = 1 Then
                        intSpecial(intX, intY, intTrack) = 0
                    Else
                        intSpecial(intX, intY, intTrack) = 1
                        picMap.Print "P"
                    End If
                Case 10
                    If intSpecial(intX, intY, intTrack) = cmbSpecial.ListIndex + 2 Then
                        intSpecial(intX, intY, intTrack) = 0
                    Else
                        intSpecial(intX, intY, intTrack) = cmbSpecial.ListIndex + 2
                        picMap.Print Chr(Asc("A") + cmbSpecial.ListIndex)
                    End If
                Case 11
                    If intSpecial(intX, intY, intTrack) = cmbSpecial.ListIndex + 6 Then
                        intSpecial(intX, intY, intTrack) = 0
                    Else
                        intSpecial(intX, intY, intTrack) = cmbSpecial.ListIndex + 6
                        picMap.CurrentX = picMap.CurrentX - 12
                        picMap.Print cmbSpecial.ListIndex
                    End If
            End Select
    End Select
End Sub

Private Sub CalculateScaledSize()
    Dim sngScaledSizeX As Single, sngScaledSizeY As Single
    
    picMap.Width = intWidth(intTrack) * 64
    hsbMap.Max = intWidth(intTrack) * 64 - picViewport.ScaleWidth
    picMap.Height = intHeight(intTrack) * 64
    vsbMap.Max = intHeight(intTrack) * 64 - picViewport.ScaleHeight
    
    sngScaledSizeX = picPreview.ScaleWidth / intWidth(intTrack)
    sngScaledSizeY = picPreview.ScaleHeight / intHeight(intTrack)
    If sngScaledSizeX <= sngScaledSizeY Then sngScaledSize = sngScaledSizeX Else sngScaledSize = sngScaledSizeY
End Sub

Private Sub vsbTracks_Change()
    If blnCancelRepaint Then
        blnCancelRepaint = False
    Else
        intTrack = vsbTracks.Value
        CalculateScaledSize
        PaintPreview
        PaintMap
    End If
End Sub

Private Function FindStartPosX(Player As Integer, Track As Integer)
    Dim i As Integer, j As Integer
    
    For i = 0 To intHeight(Track) - 1
        For j = 0 To intWidth(Track) - 1
            If intSpecial(j, i, Track) = Player + 2 Then
                FindStartPosX = j
            End If
        Next j
    Next i
End Function

Private Function FindStartPosY(Player As Integer, Track As Integer)
    Dim i As Integer, j As Integer
    
    For i = 0 To intHeight(Track) - 1
        For j = 0 To intWidth(Track) - 1
            If intSpecial(j, i, Track) = Player + 2 Then
                FindStartPosY = i
            End If
        Next j
    Next i
End Function

Private Function SwapTracks(Track1 As Integer, Track2 As Integer)
    Dim i As Integer, j As Integer, k As Integer
    Dim intTemp As Integer
    
    For j = 0 To intMaxWidth - 1
        For k = 0 To intMaxHeight - 1
            intTemp = intMap(j, k, Track1)
            intMap(j, k, Track1) = intMap(j, k, Track2)
            intMap(j, k, Track2) = intTemp
            intTemp = intArrow(j, k, Track1)
            intArrow(j, k, Track1) = intArrow(j, k, Track2)
            intArrow(j, k, Track2) = intTemp
            intTemp = intSpecial(j, k, Track1)
            intSpecial(j, k, Track1) = intSpecial(j, k, Track2)
            intSpecial(j, k, Track2) = intTemp
        Next k
    Next j
    intTemp = intWidth(Track1)
    intWidth(Track1) = intWidth(Track2)
    intWidth(Track2) = intTemp
    intTemp = intHeight(Track1)
    intHeight(Track1) = intHeight(Track2)
    intHeight(Track2) = intTemp
End Function

Private Sub LoadTileSet(Filename As String)
    Dim i As Integer
    Dim spcTiles As StdPicture

    For i = 1 To NumTiles - 1
        Unload optTile(i)
    Next i
    
    Set spcTiles = LoadPicture(Filename)
    clpTiles.picture = spcTiles
    NumTiles = spcTiles.Width / spcTiles.Height
    clpTiles.Cols = NumTiles
    optTile(0).picture = clpTiles.GraphicCell(0)
    
    For i = 1 To NumTiles - 1
        Load optTile(i)
        optTile(i).Move (i Mod 2) * 64, (i \ 2) * 64
        optTile(i).picture = clpTiles.GraphicCell(i)
        optTile(i).Visible = True
    Next i
    
    vsbTiles.Max = (((NumTiles + 1) \ 2) * 64) - picTilePort.ScaleHeight
    picTiles.Height = ((NumTiles + 1) \ 2) * 64
    Sel = 0
End Sub

Private Sub PaintTile(picture As StdPicture, x As Integer, y As Integer)
    picMap.PaintPicture picture, x, y
    
    If mnuGrid.Checked Then
        picMap.Line (x, y)-(x + 64, y)
        picMap.Line (x, y)-(x, y + 64)
    End If
End Sub
