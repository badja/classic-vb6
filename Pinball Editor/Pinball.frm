VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPinball 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pinball Table Editor"
   ClientHeight    =   6060
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9600
   Icon            =   "Pinball.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   404
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgExport 
      Left            =   8880
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "java"
      DialogTitle     =   "Export Java Fragment"
      Filter          =   "Java Files (*.java;*.jav)|*.java;*.jav|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgPicture 
      Left            =   8280
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "bmp"
      DialogTitle     =   "Save Picture"
      Filter          =   "Bitmap Files (*.bmp)|*.bmp|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgImage 
      Left            =   7680
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Load Tracing Image"
      Filter          =   "All Images (*.bmp;*.rle;*.wmf;*.emf;*.gif;*.jpg)|*.bmp;*.rle;*.wmf;*.emf;*.gif;*.jpg|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   7080
      Top             =   3840
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "ptd"
      Filter          =   "Pinball Table Definition Files (*.ptd)|*.ptd|All Files (*.*)|*.*"
   End
   Begin VB.Frame fraLightProperties 
      Caption         =   "Light Properties"
      Height          =   855
      Left            =   7080
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.TextBox txtLightGroup 
         Height          =   285
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblLightGroup 
         Caption         =   "&Group #:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraCoordinates 
      Caption         =   "Coordinates"
      Height          =   1815
      Left            =   7080
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton cmdDeletePoint 
         Caption         =   "&Delete Point"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdAddPoint 
         Caption         =   "&Add Point"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtX 
         Height          =   285
         Left            =   960
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtY 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblY 
         Caption         =   "&Y:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   135
      End
      Begin VB.Label lblX 
         Caption         =   "&X:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.Frame fraWallProperties 
      Caption         =   "Wall Properties"
      Height          =   855
      Left            =   7080
      TabIndex        =   19
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.ComboBox cmbWallAction 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblWallAction 
         Caption         =   "A&ction:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame fraTriggerProperties 
      Caption         =   "Trigger Properties"
      Height          =   1815
      Left            =   7080
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
      Begin VB.ComboBox cmbTriggerAction 
         Height          =   315
         ItemData        =   "Pinball.frx":0442
         Left            =   960
         List            =   "Pinball.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox txtTriggerScore 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox cmbTriggerValue 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblTriggerAction 
         Caption         =   "A&ction:"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblTriggerScore 
         Caption         =   "&Score:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lblTriggerValue 
         Caption         =   "&Value:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1320
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lvwObjects 
      Height          =   5835
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   10292
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.PictureBox picTable 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5820
      Left            =   3240
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   3660
   End
   Begin VB.Label lblCopyright 
      Alignment       =   1  'Right Justify
      Caption         =   "Copyright (c) 2000 Adrian Grucza"
      Height          =   255
      Left            =   7080
      TabIndex        =   22
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version 1.01"
      Height          =   255
      Left            =   8400
      TabIndex        =   23
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Label lblCoordinates 
      Caption         =   "(0, 0)"
      Height          =   255
      Left            =   7080
      TabIndex        =   24
      Top             =   5700
      Width           =   2415
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "&Export Java Fragment..."
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTable 
      Caption         =   "&Table"
      Begin VB.Menu mnuTableLoadTracing 
         Caption         =   "&Load Tracing Image..."
      End
      Begin VB.Menu mnuTableHideTracing 
         Caption         =   "&Hide Tracing Image"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuTableSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTableSavePicture 
         Caption         =   "&Save Picture..."
      End
      Begin VB.Menu mnuTableCopyPicture 
         Caption         =   "&Copy Picture to Clipboard"
      End
      Begin VB.Menu mnuTableSeparator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTableClosePolygons 
         Caption         =   "Close &Polygons"
         Checked         =   -1  'True
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuObject 
      Caption         =   "&Object"
      Begin VB.Menu mnuObjectWall 
         Caption         =   "Add &Wall"
      End
      Begin VB.Menu mnuObjectTrigger 
         Caption         =   "Add &Trigger"
      End
      Begin VB.Menu mnuObjectLight 
         Caption         =   "Add &Light"
      End
      Begin VB.Menu mnuObjectMulti 
         Caption         =   "Add &Bonus Multiplier Light"
      End
      Begin VB.Menu mnuObjectArrow 
         Caption         =   "Add &Arrow"
      End
      Begin VB.Menu mnuObjectLetter 
         Caption         =   "Add L&etter"
      End
      Begin VB.Menu mnuObjectSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObjectDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmPinball"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFileName As String
Private blnDirty As Boolean
Private blnHover As Boolean
Private strHover As String
Private intHover As Integer
Private intHoverX As Integer
Private intHoverY As Integer
Private intHoverIndex As Integer
Private strEdit As String
Private intEdit As Integer
Private intEditIndex As Integer
Private blnAdjust As Boolean
Private blnEdit As Boolean

Private wallX() As Integer
Private wallY() As Integer
Private wallAction() As Integer
Private numPointsWall() As Integer
Private numWalls As Integer
Private flipperX(1023, 3, 4) As Integer
Private flipperY(1023, 3, 4) As Integer
Private numPointsFlipper(3, 4) As Integer
Private triggerX() As Integer
Private triggerY() As Integer
Private triggerAction() As Integer
Private triggerScore() As Long
Private triggerValue() As Integer
Private numTriggers As Integer
Private lightX() As Integer
Private lightY() As Integer
Private lightGroup() As Integer
Private numLights As Integer
Private arrowX() As Integer
Private arrowY() As Integer
Private numArrows As Integer
Private multiX() As Integer
Private multiY() As Integer
Private numMultis As Integer
Private letterX() As Integer
Private letterY() As Integer
Private numLetters As Integer

Private Sub cmbTriggerAction_Click()
    blnDirty = True
    triggerAction(intEditIndex) = cmbTriggerAction.ListIndex
    PopulateTriggerValue
End Sub

Private Sub cmbTriggerValue_Click()
    blnDirty = True
    If triggerAction(intEditIndex) = 0 Then
        triggerValue(intEditIndex) = cmbTriggerValue.ListIndex
    Else
        triggerValue(intEditIndex) = cmbTriggerValue.ListIndex - 1
    End If
End Sub

Private Sub cmbWallAction_Click()
    blnDirty = True
    wallAction(intEdit, intEditIndex) = cmbWallAction.ListIndex - 2
End Sub

Private Sub cmdAddPoint_Click()
    Dim i As Integer, j As Integer
    Dim intFrame As Integer
    Dim nextIndex As Integer
    
    If IsWall(strEdit) Then
        For i = numPointsWall(intEditIndex) To intEdit + 2 Step -1
            wallX(i, intEditIndex) = wallX(i - 1, intEditIndex)
            wallY(i, intEditIndex) = wallY(i - 1, intEditIndex)
            wallAction(i, intEditIndex) = wallAction(i - 1, intEditIndex)
        Next i
        numPointsWall(intEditIndex) = numPointsWall(intEditIndex) + 1
        nextIndex = i + 1
        If nextIndex = numPointsWall(intEditIndex) Then nextIndex = 0
        wallX(i, intEditIndex) = (wallX(i - 1, intEditIndex) + wallX(nextIndex, intEditIndex)) / 2
        wallY(i, intEditIndex) = (wallY(i - 1, intEditIndex) + wallY(nextIndex, intEditIndex)) / 2
        wallAction(i, intEditIndex) = -2
    ElseIf IsFlipper(strEdit) Then
        intFrame = Val(Right(strEdit, 1)) - 1
        For i = numPointsFlipper(intEditIndex, intFrame) + 2 To intEdit + 2 Step -1
            flipperX(i, intEditIndex, intFrame) = flipperX(i - 1, intEditIndex, intFrame)
            flipperY(i, intEditIndex, intFrame) = flipperY(i - 1, intEditIndex, intFrame)
        Next i
        numPointsFlipper(intEditIndex, intFrame) = numPointsFlipper(intEditIndex, intFrame) + 1
        nextIndex = i + 1
        If nextIndex = numPointsFlipper(intEditIndex, intFrame) Then nextIndex = 0
        flipperX(i, intEditIndex, intFrame) = (flipperX(i - 1, intEditIndex, intFrame) + flipperX(nextIndex, intEditIndex, intFrame)) / 2
        flipperY(i, intEditIndex, intFrame) = (flipperY(i - 1, intEditIndex, intFrame) + flipperY(nextIndex, intEditIndex, intFrame)) / 2
    End If
    
    lvwObjects_ItemClick lvwObjects.SelectedItem
End Sub

Private Sub cmdDeletePoint_Click()
    Dim i As Integer, j As Integer
    Dim intFrame As Integer
    
    If IsWall(strEdit) Then
        For i = intEdit To numPointsWall(intEditIndex) - 2
            wallX(i, intEditIndex) = wallX(i + 1, intEditIndex)
            wallY(i, intEditIndex) = wallY(i + 1, intEditIndex)
            wallAction(i, intEditIndex) = wallAction(i + 1, intEditIndex)
        Next i
        numPointsWall(intEditIndex) = numPointsWall(intEditIndex) - 1
    ElseIf IsFlipper(strEdit) Then
        intFrame = Val(Right(strEdit, 1)) - 1
        For i = intEdit To numPointsFlipper(intEditIndex, intFrame) - 2
            flipperX(i, intEditIndex, intFrame) = flipperX(i + 1, intEditIndex, intFrame)
            flipperY(i, intEditIndex, intFrame) = flipperY(i + 1, intEditIndex, intFrame)
        Next i
        numPointsFlipper(intEditIndex, intFrame) = numPointsFlipper(intEditIndex, intFrame) - 1
    End If
    
    lvwObjects_ItemClick lvwObjects.SelectedItem
End Sub

Private Sub Form_Load()
    dlgFile.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    dlgImage.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    dlgPicture.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    dlgExport.Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    UpdateTitleBar
    NewFile
End Sub

Private Sub Form_Paint()
    DrawObjects
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim intResponse As Integer
    
    If blnDirty Then
        intResponse = PromptToSave
        If intResponse = vbYes Then
            mnuFileSave_Click
        ElseIf intResponse = vbCancel Then
            Cancel = True
            Exit Sub
        End If
    End If
End Sub

Private Sub lvwObjects_BeforeLabelEdit(Cancel As Integer)
    If IsLight(lvwObjects.SelectedItem.Key) Or IsArrow(lvwObjects.SelectedItem.Key) Or IsMulti(lvwObjects.SelectedItem.Key) Or IsLetter(lvwObjects.SelectedItem.Key) Then
        Cancel = True
        MsgBox "Cannot rename this kind of object.", vbExclamation
    End If
End Sub

Private Sub lvwObjects_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    blnDirty = True
    blnEdit = False
    fraCoordinates.Visible = False
    fraWallProperties.Visible = False
    fraTriggerProperties.Visible = False
    fraLightProperties.Visible = False
    If Item.Selected And Item.Checked Then
        If IsTrigger(Item.Key) Then
            ShowTriggerProperties
        ElseIf IsLight(Item.Key) Then
            ShowLightProperties
        ElseIf Not IsWall(Item.Key) And Not IsFlipper(Item.Key) Then
            ShowCoordinates
        End If
    End If
    DrawObjects
End Sub

Private Sub lvwObjects_ItemClick(ByVal Item As MSComctlLib.ListItem)
    blnEdit = False
    fraCoordinates.Visible = False
    fraWallProperties.Visible = False
    fraTriggerProperties.Visible = False
    fraLightProperties.Visible = False
    If Item.Checked Then
        If IsTrigger(Item.Key) Then
            ShowTriggerProperties
        ElseIf IsLight(Item.Key) Then
            ShowLightProperties
        ElseIf Not IsWall(Item.Key) And Not IsFlipper(Item.Key) Then
            ShowCoordinates
        End If
    End If
    DrawObjects
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileExport_Click()
    On Error Resume Next
    dlgExport.ShowSave
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    ExportFile
End Sub

Private Sub mnuFileNew_Click()
    Dim intResponse As Integer
    
    If blnDirty Then
        intResponse = PromptToSave
        If intResponse = vbYes Then
            mnuFileSave_Click
        ElseIf intResponse = vbCancel Then
            Exit Sub
        End If
    End If
    
    strFileName = ""
    UpdateTitleBar
    NewFile
End Sub

Private Sub mnuFileOpen_Click()
    Dim intResponse As Integer
    
    If blnDirty Then
        intResponse = PromptToSave
        If intResponse = vbYes Then
            mnuFileSaveAs_Click
            If Err.Number > 0 Then Exit Sub
        ElseIf intResponse = vbCancel Then
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    dlgFile.ShowOpen
    If Err.Number > 0 Then Exit Sub
    On Error GoTo 0
    strFileName = dlgFile.FileName
    UpdateTitleBar
    OpenFile
End Sub

Private Sub mnuFileSave_Click()
    If strFileName = "" Then
        mnuFileSaveAs_Click
    Else
        SaveFile
    End If
End Sub

Private Sub mnuFileSaveAs_Click()
    On Error Resume Next
    dlgFile.ShowSave
    If Err.Number > 0 Then Exit Sub
    strFileName = dlgFile.FileName
    UpdateTitleBar
    SaveFile
End Sub

Private Sub mnuObjectArrow_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve arrowX(numArrows)
    ReDim Preserve arrowY(numArrows)
    arrowX(numArrows) = 120
    arrowY(numArrows) = 192
    numArrows = numArrows + 1
    strName = "Arrow " & numArrows
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuObjectDelete_Click()
    Dim intIndex As Integer
    Dim intFrame As Integer
    Dim i As Integer, j As Integer
    
    intIndex = GetIndex(lvwObjects.SelectedItem.Key)
    
    If IsWall(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numWalls - 2
            For j = 0 To 1023
                wallX(j, i) = wallX(j, i + 1)
                wallY(j, i) = wallY(j, i + 1)
                wallAction(j, i) = wallAction(j, i + 1)
            Next j
            numPointsWall(i) = numPointsWall(i + 1)
            lvwObjects.ListItems("Wall " & i + 2).Key = "Wall " & i + 1
        Next i
        numWalls = numWalls - 1
        If numWalls > 0 Then
            ReDim Preserve wallX(1023, numWalls - 1)
            ReDim Preserve wallY(1023, numWalls - 1)
            ReDim Preserve wallAction(1023, numWalls - 1)
            ReDim Preserve numPointsWall(numWalls - 1)
        Else
            Erase wallX, wallY, wallAction, numPointsWall
        End If
    ElseIf IsFlipper(lvwObjects.SelectedItem.Key) Then
        intIndex = Asc(Mid(lvwObjects.SelectedItem.Key, 8, 1)) - Asc("A")
        intFrame = Val(Right(lvwObjects.SelectedItem.Key, 1)) - 1
        numPointsFlipper(intIndex, intFrame) = 0
    ElseIf IsTrigger(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numTriggers - 2
            triggerX(i) = triggerX(i + 1)
            triggerY(i) = triggerY(i + 1)
            triggerAction(i) = triggerAction(i + 1)
            triggerScore(i) = triggerScore(i + 1)
            triggerValue(i) = triggerValue(i + 1)
            lvwObjects.ListItems("Trigger " & i + 2).Key = "Trigger " & i + 1
        Next i
        numTriggers = numTriggers - 1
        If numTriggers > 0 Then
            ReDim Preserve triggerX(numTriggers - 1)
            ReDim Preserve triggerY(numTriggers - 1)
            ReDim Preserve triggerAction(numTriggers - 1)
            ReDim Preserve triggerScore(numTriggers - 1)
            ReDim Preserve triggerValue(numTriggers - 1)
        Else
            Erase triggerX, triggerY, triggerAction, triggerScore, triggerValue
        End If
    ElseIf IsLight(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numLights - 2
            lightX(i) = lightX(i + 1)
            lightY(i) = lightY(i + 1)
            lightGroup(i) = lightGroup(i + 1)
            lvwObjects.ListItems("Light " & i + 2).Key = "Light " & i + 1
            lvwObjects.ListItems("Light " & i + 1).Text = lvwObjects.ListItems("Light " & i + 1).Key
        Next i
        numLights = numLights - 1
        If numLights > 0 Then
            ReDim Preserve lightX(numLights - 1)
            ReDim Preserve lightY(numLights - 1)
            ReDim Preserve lightGroup(numLights - 1)
        Else
            Erase lightX, lightY, lightGroup
        End If
    ElseIf IsArrow(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numArrows - 2
            arrowX(i) = arrowX(i + 1)
            arrowY(i) = arrowY(i + 1)
            lvwObjects.ListItems("Arrow " & i + 2).Key = "Arrow " & i + 1
            lvwObjects.ListItems("Arrow " & i + 1).Text = lvwObjects.ListItems("Arrow " & i + 1).Key
        Next i
        numArrows = numArrows - 1
        If numArrows > 0 Then
            ReDim Preserve arrowX(numArrows - 1)
            ReDim Preserve arrowY(numArrows - 1)
        Else
            Erase arrowX, arrowY
        End If
    ElseIf IsMulti(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numMultis - 2
            multiX(i) = multiX(i + 1)
            multiY(i) = multiY(i + 1)
            lvwObjects.ListItems("Multi " & i + 2).Key = "Multi " & i + 1
            lvwObjects.ListItems("Multi " & i + 1).Text = lvwObjects.ListItems("Multi " & i + 1).Key
        Next i
        numMultis = numMultis - 1
        If numMultis > 0 Then
            ReDim Preserve multiX(numMultis - 1)
            ReDim Preserve multiY(numMultis - 1)
        Else
            Erase multiX, multiY
        End If
    ElseIf IsLetter(lvwObjects.SelectedItem.Key) Then
        lvwObjects.ListItems.Remove lvwObjects.SelectedItem.Index
        For i = intIndex To numLetters - 2
            letterX(i) = letterX(i + 1)
            letterY(i) = letterY(i + 1)
            lvwObjects.ListItems("Letter " & i + 2).Key = "Letter " & i + 1
            lvwObjects.ListItems("Letter " & i + 1).Text = lvwObjects.ListItems("Letter " & i + 1).Key
        Next i
        numLetters = numLetters - 1
        If numLetters > 0 Then
            ReDim Preserve letterX(numLetters - 1)
            ReDim Preserve letterY(numLetters - 1)
        Else
            Erase letterX, letterY
        End If
    End If
    
    lvwObjects.SelectedItem.Selected = True
    lvwObjects_ItemClick lvwObjects.SelectedItem
End Sub

Private Sub mnuObjectLetter_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve letterX(numLetters)
    ReDim Preserve letterY(numLetters)
    letterX(numLetters) = 120
    letterY(numLetters) = 192
    numLetters = numLetters + 1
    strName = "Letter " & numLetters
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuObjectLight_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve lightX(numLights)
    ReDim Preserve lightY(numLights)
    ReDim Preserve lightGroup(numLights)
    lightX(numLights) = 120
    lightY(numLights) = 192
    lightGroup(numLights) = 1
    numLights = numLights + 1
    strName = "Light " & numLights
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuObjectMulti_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve multiX(numMultis)
    ReDim Preserve multiY(numMultis)
    multiX(numMultis) = 120
    multiY(numMultis) = 192
    numMultis = numMultis + 1
    strName = "Multi " & numMultis
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuObjectTrigger_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve triggerX(numTriggers)
    ReDim Preserve triggerY(numTriggers)
    ReDim Preserve triggerAction(numTriggers)
    ReDim Preserve triggerScore(numTriggers)
    ReDim Preserve triggerValue(numTriggers)
    triggerX(numTriggers) = 120
    triggerY(numTriggers) = 192
    triggerAction(numTriggers) = 1
    triggerScore(numTriggers) = 0
    triggerValue(numTriggers) = -1
    numTriggers = numTriggers + 1
    strName = "Trigger " & numTriggers
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuObjectWall_Click()
    Dim objObject As ListItem
    Dim strName As String

    ReDim Preserve wallX(1023, numWalls)
    ReDim Preserve wallY(1023, numWalls)
    ReDim Preserve wallAction(1023, numWalls)
    ReDim Preserve numPointsWall(numWalls)
    numWalls = numWalls + 1
    strName = "Wall " & numWalls
    Set objObject = lvwObjects.ListItems.Add(, strName, strName)
    objObject.Checked = True
    objObject.Selected = True
    lvwObjects_ItemClick objObject
End Sub

Private Sub mnuTableClosePolygons_Click()
    mnuTableClosePolygons.Checked = Not mnuTableClosePolygons.Checked
    DrawObjects
End Sub

Private Sub mnuTableCopyPicture_Click()
    Clipboard.Clear
    Clipboard.SetData picTable.Image
End Sub

Private Sub mnuTableHideTracing_Click()
    On Error Resume Next
    
    mnuTableHideTracing.Checked = Not mnuTableHideTracing.Checked
    
    If mnuTableHideTracing.Checked Then
        picTable.Picture = LoadPicture()
    Else
        picTable.Picture = LoadPicture(dlgImage.FileName)
    End If
    
    DrawObjects
End Sub

Private Sub mnuTableLoadTracing_Click()
    On Error Resume Next
    dlgImage.ShowOpen
    If Err.Number > 0 Then Exit Sub
    mnuTableHideTracing.Checked = False
    picTable.Picture = LoadPicture(dlgImage.FileName)
    DrawObjects
End Sub

Private Sub mnuTableSavePicture_Click()
    On Error Resume Next
    dlgPicture.ShowSave
    If Err.Number > 0 Then Exit Sub
    SavePicture picTable.Image, dlgPicture.FileName
End Sub

Private Sub picTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intIndex As Integer
    Dim intFrame As Integer

    blnDirty = True
    
    If blnHover Then
        blnAdjust = True
        txtX.Text = intHoverX
        txtY.Text = intHoverY
        fraCoordinates.Visible = True
        If IsWall(strHover) Then
            cmdAddPoint.Enabled = True
            cmdDeletePoint.Enabled = True
            blnEdit = True
            strEdit = strHover
            intEdit = intHover
            intEditIndex = intHoverIndex
            PopulateWallAction
            On Error Resume Next
            cmbWallAction.ListIndex = wallAction(intEdit, intEditIndex) + 2
            fraWallProperties.Visible = True
        ElseIf IsFlipper(strHover) Then
            cmdAddPoint.Enabled = True
            cmdDeletePoint.Enabled = True
            blnEdit = True
            strEdit = strHover
            intEdit = intHover
            intEditIndex = intHoverIndex
        End If
        
        DrawObjects
    Else
        lvwObjects.SelectedItem.Checked = True
        intIndex = GetIndex(lvwObjects.SelectedItem.Key)
        
        If IsWall(lvwObjects.SelectedItem.Key) Then
            wallX(numPointsWall(intIndex), intIndex) = X
            wallY(numPointsWall(intIndex), intIndex) = Y
            wallAction(numPointsWall(intIndex), intIndex) = -2
            numPointsWall(intIndex) = numPointsWall(intIndex) + 1
        ElseIf IsFlipper(lvwObjects.SelectedItem.Key) Then
            intIndex = Asc(Mid(lvwObjects.SelectedItem.Key, 8, 1)) - Asc("A")
            intFrame = Val(Right(lvwObjects.SelectedItem.Key, 1)) - 1
            flipperX(numPointsFlipper(intIndex, intFrame), intIndex, intFrame) = X
            flipperY(numPointsFlipper(intIndex, intFrame), intIndex, intFrame) = Y
            numPointsFlipper(intIndex, intFrame) = numPointsFlipper(intIndex, intFrame) + 1
        ElseIf IsTrigger(lvwObjects.SelectedItem.Key) Then
            triggerX(intIndex) = X
            triggerY(intIndex) = Y
        ElseIf IsLight(lvwObjects.SelectedItem.Key) Then
            lightX(intIndex) = X
            lightY(intIndex) = Y
        ElseIf IsArrow(lvwObjects.SelectedItem.Key) Then
            arrowX(intIndex) = X
            arrowY(intIndex) = Y
        ElseIf IsMulti(lvwObjects.SelectedItem.Key) Then
            multiX(intIndex) = X
            multiY(intIndex) = Y
        ElseIf IsLetter(lvwObjects.SelectedItem.Key) Then
            letterX(intIndex) = X
            letterY(intIndex) = Y
        End If
        
        If Not IsWall(lvwObjects.SelectedItem.Key) And Not IsFlipper(lvwObjects.SelectedItem.Key) Then
            txtX.Text = X
            txtY.Text = Y
            cmdAddPoint.Enabled = False
            cmdDeletePoint.Enabled = False
            fraCoordinates.Visible = True
        End If
        
        DrawObjects
    End If
End Sub

Private Sub DrawObjects()
    Dim i As Integer
    Dim intIndex As Integer
    Dim intFrame As Integer
    Dim objObject As ListItem

    picTable.Cls
    
    For Each objObject In lvwObjects.ListItems
        If objObject.Checked Then
            intIndex = GetIndex(objObject.Key)
            If IsWall(objObject.Key) Then
                If numPointsWall(intIndex) > 0 Then
                    If blnEdit And objObject.Key = strEdit And 0 = intEdit Then
                        picTable.ForeColor = QBColor(12)
                        picTable.FillColor = QBColor(12)
                        picTable.FillStyle = vbSolid
                    Else
                        picTable.ForeColor = QBColor(0)
                        picTable.FillStyle = 1
                    End If
                    If objObject.Selected Then picTable.Circle (wallX(0, intIndex), wallY(0, intIndex)), 2
                    For i = 1 To numPointsWall(intIndex) - 1
                        If blnEdit And objObject.Key = strEdit And i = intEdit Then
                            picTable.ForeColor = QBColor(12)
                            picTable.FillColor = QBColor(12)
                            picTable.FillStyle = vbSolid
                        Else
                            picTable.ForeColor = QBColor(0)
                            picTable.FillStyle = 1
                        End If
                        If objObject.Selected Then picTable.Circle (wallX(i, intIndex), wallY(i, intIndex)), 2
                        
                        If blnEdit And objObject.Key = strEdit And i - 1 = intEdit Then
                            picTable.ForeColor = QBColor(12)
                        ElseIf wallAction(i - 1, intIndex) = -1 Then
                            picTable.ForeColor = QBColor(4)
                        ElseIf wallAction(i - 1, intIndex) > -1 Then
                            picTable.ForeColor = QBColor(3)
                        Else
                            picTable.ForeColor = QBColor(1)
                        End If
                        picTable.Line (wallX(i - 1, intIndex), wallY(i - 1, intIndex))-(wallX(i, intIndex), wallY(i, intIndex))
                    Next i
                    
                    If mnuTableClosePolygons.Checked Then
                        If blnEdit And objObject.Key = strEdit And i - 1 = intEdit Then
                            picTable.ForeColor = QBColor(12)
                        ElseIf wallAction(i - 1, intIndex) = -1 Then
                            picTable.ForeColor = QBColor(4)
                        ElseIf wallAction(i - 1, intIndex) > -1 Then
                            picTable.ForeColor = QBColor(3)
                        Else
                            picTable.ForeColor = QBColor(1)
                        End If
                        picTable.Line (wallX(i - 1, intIndex), wallY(i - 1, intIndex))-(wallX(0, intIndex), wallY(0, intIndex))
                    End If
                End If
            ElseIf IsFlipper(objObject.Key) Then
                intIndex = Asc(Mid(objObject.Key, 8, 1)) - Asc("A")
                intFrame = Val(Right(objObject.Key, 1)) - 1
                If numPointsFlipper(intIndex, intFrame) > 0 Then
                    If blnEdit And objObject.Key = strEdit And 0 = intEdit Then
                        picTable.ForeColor = QBColor(12)
                        picTable.FillColor = QBColor(12)
                        picTable.FillStyle = vbSolid
                    Else
                        picTable.ForeColor = QBColor(0)
                        picTable.FillStyle = 1
                    End If
                    If objObject.Selected Then picTable.Circle (flipperX(0, intIndex, intFrame), flipperY(0, intIndex, intFrame)), 2
                    For i = 1 To numPointsFlipper(intIndex, intFrame) - 1
                        If blnEdit And objObject.Key = strEdit And i = intEdit Then
                            picTable.ForeColor = QBColor(12)
                            picTable.FillColor = QBColor(12)
                            picTable.FillStyle = vbSolid
                        Else
                            picTable.ForeColor = QBColor(0)
                            picTable.FillStyle = 1
                        End If
                        If objObject.Selected Then picTable.Circle (flipperX(i, intIndex, intFrame), flipperY(i, intIndex, intFrame)), 2
                        picTable.Line (flipperX(i - 1, intIndex, intFrame), flipperY(i - 1, intIndex, intFrame))-(flipperX(i, intIndex, intFrame), flipperY(i, intIndex, intFrame)), QBColor(2)
                    Next i
                    If mnuTableClosePolygons.Checked Then picTable.Line (flipperX(i - 1, intIndex, intFrame), flipperY(i - 1, intIndex, intFrame))-(flipperX(0, intIndex, intFrame), flipperY(0, intIndex, intFrame)), QBColor(2)
                End If
            ElseIf IsTrigger(objObject.Key) Then
                picTable.ForeColor = QBColor(3)
                picTable.FillStyle = 1
                If objObject.Selected Then picTable.Circle (triggerX(intIndex), triggerY(intIndex)), 2, QBColor(0)
                picTable.Circle (triggerX(intIndex), triggerY(intIndex)), 4
            ElseIf IsLight(objObject.Key) Then
                picTable.ForeColor = QBColor(7)
                picTable.FillStyle = 1
                If objObject.Selected Then picTable.Circle (lightX(intIndex), lightY(intIndex)), 2, QBColor(0)
                picTable.Line (lightX(intIndex) + 8, lightY(intIndex))-(lightX(intIndex), lightY(intIndex))
                picTable.Line -(lightX(intIndex), lightY(intIndex) + 8)
            ElseIf IsArrow(objObject.Key) Then
                picTable.ForeColor = QBColor(7)
                picTable.FillStyle = 1
                If objObject.Selected Then picTable.Circle (arrowX(intIndex), arrowY(intIndex)), 2, QBColor(0)
                picTable.Line (arrowX(intIndex) + 8, arrowY(intIndex))-(arrowX(intIndex), arrowY(intIndex))
                picTable.Line -(arrowX(intIndex), arrowY(intIndex) + 8)
            ElseIf IsMulti(objObject.Key) Then
                picTable.ForeColor = QBColor(7)
                picTable.FillStyle = 1
                If objObject.Selected Then picTable.Circle (multiX(intIndex), multiY(intIndex)), 2, QBColor(0)
                picTable.Line (multiX(intIndex) + 8, multiY(intIndex))-(multiX(intIndex), multiY(intIndex))
                picTable.Line -(multiX(intIndex), multiY(intIndex) + 8)
            ElseIf IsLetter(objObject.Key) Then
                picTable.ForeColor = QBColor(7)
                picTable.FillStyle = 1
                If objObject.Selected Then picTable.Circle (letterX(intIndex), letterY(intIndex)), 2, QBColor(0)
                picTable.Line (letterX(intIndex) + 8, letterY(intIndex))-(letterX(intIndex), letterY(intIndex))
                picTable.Line -(letterX(intIndex), letterY(intIndex) + 8)
            End If
        End If
    Next objObject
End Sub

Private Sub picTable_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim intIndex As Integer
    Dim intFrame As Integer
    
    lblCoordinates = "(" & X & ", " & Y & ")"
    
    If blnAdjust Then
        intIndex = GetIndex(strHover)
        If IsWall(strHover) Then
            wallX(intHover, intIndex) = X
            wallY(intHover, intIndex) = Y
        ElseIf IsFlipper(strHover) Then
            intIndex = Asc(Mid(strHover, 8, 1)) - Asc("A")
            intFrame = Val(Right(strHover, 1)) - 1
            flipperX(intHover, intIndex, intFrame) = X
            flipperY(intHover, intIndex, intFrame) = Y
        ElseIf IsTrigger(strHover) Then
            triggerX(intIndex) = X
            triggerY(intIndex) = Y
        End If
        
        txtX.Text = X
        txtY.Text = Y
        fraCoordinates.Visible = True
        DrawObjects
    Else
        blnHover = False
    
        If lvwObjects.SelectedItem.Checked Then
            intIndex = GetIndex(lvwObjects.SelectedItem.Key)
            If IsWall(lvwObjects.SelectedItem.Key) Then
                For i = 0 To numPointsWall(intIndex) - 1
                    If X > wallX(i, intIndex) - 2 And X < wallX(i, intIndex) + 2 And Y > wallY(i, intIndex) - 2 And Y < wallY(i, intIndex) + 2 Then
                        blnHover = True
                        strHover = lvwObjects.SelectedItem.Key
                        intHover = i
                        intHoverIndex = intIndex
                        intHoverX = wallX(i, intIndex)
                        intHoverY = wallY(i, intIndex)
                    End If
                Next i
            ElseIf IsFlipper(lvwObjects.SelectedItem.Key) Then
                intIndex = Asc(Mid(lvwObjects.SelectedItem.Key, 8, 1)) - Asc("A")
                intFrame = Val(Right(lvwObjects.SelectedItem.Key, 1)) - 1
                For i = 0 To numPointsFlipper(intIndex, intFrame) - 1
                    If X > flipperX(i, intIndex, intFrame) - 2 And X < flipperX(i, intIndex, intFrame) + 2 And Y > flipperY(i, intIndex, intFrame) - 2 And Y < flipperY(i, intIndex, intFrame) + 2 Then
                        blnHover = True
                        strHover = lvwObjects.SelectedItem.Key
                        intHover = i
                        intHoverIndex = intIndex
                        intHoverX = flipperX(i, intIndex, intFrame)
                        intHoverY = flipperY(i, intIndex, intFrame)
                    End If
                Next i
            End If
        End If
        
        If blnHover Then picTable.MousePointer = vbArrow Else picTable.MousePointer = vbCrosshair
    End If
End Sub

Private Sub picTable_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    blnAdjust = False
End Sub

Private Sub Highlight(strObject As String)
    Dim i As Integer
    Dim objObject As ListItem
    
    For Each objObject In lvwObjects.ListItems
        If objObject.Key = strObject Then
            objObject.Checked = True
            Exit Sub
        End If
    Next objObject
End Sub

Private Sub PopulateWallAction()
    Dim i As Integer
    
    cmbWallAction.Clear
    cmbWallAction.AddItem "Normal"
    cmbWallAction.AddItem "Bumper"

    For i = 1 To numLights
        cmbWallAction.AddItem lvwObjects.ListItems("Light " & i).Text
    Next i
End Sub

Private Sub PopulateTriggerValue()
    Dim i As Integer
    On Error Resume Next
    
    cmbTriggerValue.Clear
    
    If triggerAction(intEditIndex) = 0 Then
        For i = 1 To numLights
            cmbTriggerValue.AddItem lvwObjects.ListItems("Light " & i).Text
        Next i
        cmbTriggerValue.ListIndex = triggerValue(intEditIndex)
    Else
        cmbTriggerValue.AddItem "No arrow"
        For i = 1 To numArrows
            cmbTriggerValue.AddItem lvwObjects.ListItems("Arrow " & i).Text
        Next i
        cmbTriggerValue.ListIndex = triggerValue(intEditIndex) + 1
    End If
End Sub

Private Sub txtLightGroup_Change()
    blnDirty = True
    lightGroup(intEditIndex) = Val(txtLightGroup.Text)
End Sub

Private Sub txtTriggerScore_Change()
    On Error Resume Next
    blnDirty = True
    triggerScore(intEditIndex) = Val(txtTriggerScore.Text)
End Sub

Private Sub ShowTriggerProperties()
    strEdit = lvwObjects.SelectedItem.Key
    intEditIndex = GetIndex(strEdit)
    blnEdit = True
    cmbTriggerAction.ListIndex = triggerAction(intEditIndex)
    txtTriggerScore.Text = triggerScore(intEditIndex)
    PopulateTriggerValue
    ShowCoordinates
    fraTriggerProperties.Visible = True
End Sub

Private Sub ShowLightProperties()
    strEdit = lvwObjects.SelectedItem.Key
    intEditIndex = GetIndex(strEdit)
    blnEdit = True
    txtLightGroup.Text = lightGroup(intEditIndex)
    ShowCoordinates
    fraLightProperties.Visible = True
End Sub

Private Sub ShowCoordinates()
    strEdit = lvwObjects.SelectedItem.Key
    intEditIndex = GetIndex(strEdit)
    If IsTrigger(strEdit) Then
        txtX.Text = triggerX(intEditIndex)
        txtY.Text = triggerY(intEditIndex)
    ElseIf IsLight(strEdit) Then
        txtX.Text = lightX(intEditIndex)
        txtY.Text = lightY(intEditIndex)
    ElseIf IsArrow(strEdit) Then
        txtX.Text = arrowX(intEditIndex)
        txtY.Text = arrowY(intEditIndex)
    ElseIf IsMulti(strEdit) Then
        txtX.Text = multiX(intEditIndex)
        txtY.Text = multiY(intEditIndex)
    ElseIf IsLetter(strEdit) Then
        txtX.Text = letterX(intEditIndex)
        txtY.Text = letterY(intEditIndex)
    End If
    blnEdit = True
    cmdAddPoint.Enabled = False
    cmdDeletePoint.Enabled = False
    fraCoordinates.Visible = True
End Sub

Private Function IsWall(strObject)
    If Left(strObject, 4) = "Wall" Then IsWall = True
End Function

Private Function IsFlipper(strObject)
    If Left(strObject, 7) = "Flipper" Then IsFlipper = True
End Function

Private Function IsTrigger(strObject)
    If Left(strObject, 7) = "Trigger" Then IsTrigger = True
End Function

Private Function IsLight(strObject)
    If Left(strObject, 5) = "Light" Then IsLight = True
End Function

Private Function IsArrow(strObject)
    If Left(strObject, 5) = "Arrow" Then IsArrow = True
End Function

Private Function IsMulti(strObject)
    If Left(strObject, 5) = "Multi" Then IsMulti = True
End Function

Private Function IsLetter(strObject)
    If Left(strObject, 6) = "Letter" Then IsLetter = True
End Function

Private Function GetIndex(strObject)
    GetIndex = Val(Mid(strObject, InStr(strObject, " ") + 1)) - 1
End Function

Private Sub txtX_KeyPress(KeyAscii As Integer)
    Dim intFrame As Integer

    If KeyAscii = 13 Then
        KeyAscii = 0
        blnDirty = True
        If IsWall(strEdit) Then
            wallX(intEdit, intEditIndex) = Val(txtX.Text)
        ElseIf IsFlipper(strEdit) Then
            intFrame = Val(Right(strEdit, 1)) - 1
            flipperX(intEdit, intEditIndex, intFrame) = Val(txtX.Text)
        ElseIf IsTrigger(strEdit) Then
            triggerX(intEditIndex) = Val(txtX.Text)
        ElseIf IsLight(strEdit) Then
            lightX(intEditIndex) = Val(txtX.Text)
        ElseIf IsArrow(strEdit) Then
            arrowX(intEditIndex) = Val(txtX.Text)
        ElseIf IsMulti(strEdit) Then
            multiX(intEditIndex) = Val(txtX.Text)
        ElseIf IsLetter(strEdit) Then
            letterX(intEditIndex) = Val(txtX.Text)
        End If
        DrawObjects
    End If
End Sub

Private Sub txtY_KeyPress(KeyAscii As Integer)
    Dim intFrame As Integer
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        blnDirty = True
        If IsWall(strEdit) Then
            wallY(intEdit, intEditIndex) = Val(txtY.Text)
        ElseIf IsFlipper(strEdit) Then
            intFrame = Val(Right(strEdit, 1)) - 1
            flipperY(intEdit, intEditIndex, intFrame) = Val(txtY.Text)
        ElseIf IsTrigger(strEdit) Then
            triggerY(intEditIndex) = Val(txtY.Text)
        ElseIf IsLight(strEdit) Then
            lightY(intEditIndex) = Val(txtY.Text)
        ElseIf IsArrow(strEdit) Then
            arrowY(intEditIndex) = Val(txtY.Text)
        ElseIf IsMulti(strEdit) Then
            multiY(intEditIndex) = Val(txtY.Text)
        ElseIf IsLetter(strEdit) Then
            letterY(intEditIndex) = Val(txtY.Text)
        End If
        DrawObjects
    End If
End Sub

Private Sub SaveFile()
    Dim i As Integer, j As Integer

    Dim wallNames() As String
    Dim flipperNames(3, 4) As String
    Dim triggerNames() As String
    Dim lightNames() As String
    Dim multiNames() As String
    Dim arrowNames() As String
    Dim letterNames() As String
    
    Dim wallChecked() As Boolean
    Dim flipperChecked(3, 4) As Boolean
    Dim triggerChecked() As Boolean
    Dim lightChecked() As Boolean
    Dim multiChecked() As Boolean
    Dim arrowChecked() As Boolean
    Dim letterChecked() As Boolean
    
    On Error Resume Next
    Kill strFileName
    On Error GoTo 0
    
    If numWalls > 0 Then
        ReDim wallNames(numWalls - 1)
        ReDim wallChecked(numWalls - 1)
    End If
    If numTriggers > 0 Then
        ReDim triggerNames(numTriggers - 1)
        ReDim triggerChecked(numTriggers - 1)
    End If
    If numLights > 0 Then
        ReDim lightNames(numLights - 1)
        ReDim lightChecked(numLights - 1)
    End If
    If numMultis > 0 Then
        ReDim multiNames(numMultis - 1)
        ReDim multiChecked(numMultis - 1)
    End If
    If numArrows > 0 Then
        ReDim arrowNames(numArrows - 1)
        ReDim arrowChecked(numArrows - 1)
    End If
    If numLetters > 0 Then
        ReDim letterNames(numLetters - 1)
        ReDim letterChecked(numLetters - 1)
    End If
    
    For i = 0 To numWalls - 1
        wallNames(i) = lvwObjects.ListItems("Wall " & i + 1).Text
        wallChecked(i) = lvwObjects.ListItems("Wall " & i + 1).Checked
    Next i
    For i = 0 To 3
        For j = 0 To 4
            flipperNames(i, j) = lvwObjects.ListItems("Flipper" & Chr(i + Asc("A")) & " " & j + 1).Text
            flipperChecked(i, j) = lvwObjects.ListItems("Flipper" & Chr(i + Asc("A")) & " " & j + 1).Checked
        Next j
    Next i
    For i = 0 To numTriggers - 1
        triggerNames(i) = lvwObjects.ListItems("Trigger " & i + 1).Text
        triggerChecked(i) = lvwObjects.ListItems("Trigger " & i + 1).Checked
    Next i
    For i = 0 To numLights - 1
        lightNames(i) = lvwObjects.ListItems("Light " & i + 1).Text
        lightChecked(i) = lvwObjects.ListItems("Light " & i + 1).Checked
    Next i
    For i = 0 To numArrows - 1
        arrowNames(i) = lvwObjects.ListItems("Arrow " & i + 1).Text
        arrowChecked(i) = lvwObjects.ListItems("Arrow " & i + 1).Checked
    Next i
    For i = 0 To numMultis - 1
        multiNames(i) = lvwObjects.ListItems("Multi " & i + 1).Text
        multiChecked(i) = lvwObjects.ListItems("Multi " & i + 1).Checked
    Next i
    For i = 0 To numLetters - 1
        letterNames(i) = lvwObjects.ListItems("Letter " & i + 1).Text
        letterChecked(i) = lvwObjects.ListItems("Letter " & i + 1).Checked
    Next i

    Open strFileName For Binary Access Write As #1
    Put #1, , numWalls
    Put #1, , numPointsWall()
    Put #1, , wallX()
    Put #1, , wallY()
    Put #1, , wallAction()
    Put #1, , wallNames()
    Put #1, , wallChecked()
    Put #1, , numPointsFlipper()
    Put #1, , flipperX()
    Put #1, , flipperY()
    Put #1, , flipperNames()
    Put #1, , flipperChecked()
    Put #1, , numTriggers
    Put #1, , triggerX()
    Put #1, , triggerY()
    Put #1, , triggerAction()
    Put #1, , triggerScore()
    Put #1, , triggerValue()
    Put #1, , triggerNames()
    Put #1, , triggerChecked()
    Put #1, , numLights
    Put #1, , lightX()
    Put #1, , lightY()
    Put #1, , lightGroup()
    Put #1, , lightNames()
    Put #1, , lightChecked()
    Put #1, , numArrows
    Put #1, , arrowX()
    Put #1, , arrowY()
    Put #1, , arrowNames()
    Put #1, , arrowChecked()
    Put #1, , numMultis
    Put #1, , multiX()
    Put #1, , multiY()
    Put #1, , multiNames()
    Put #1, , multiChecked()
    Put #1, , numLetters
    Put #1, , letterX()
    Put #1, , letterY()
    Put #1, , letterNames()
    Put #1, , letterChecked()
    Close #1
    
    blnDirty = False
End Sub

Private Sub OpenFile()
    Dim i As Integer, j As Integer
    Dim intDim As Integer
    Dim objObject As ListItem

    Dim wallNames() As String
    Dim flipperNames(3, 4) As String
    Dim triggerNames() As String
    Dim lightNames() As String
    Dim multiNames() As String
    Dim arrowNames() As String
    Dim letterNames() As String
    
    Dim wallChecked() As Boolean
    Dim flipperChecked(3, 4) As Boolean
    Dim triggerChecked() As Boolean
    Dim lightChecked() As Boolean
    Dim multiChecked() As Boolean
    Dim arrowChecked() As Boolean
    Dim letterChecked() As Boolean
    
    Erase wallX, wallY, wallAction, numPointsWall, flipperX, flipperY, numPointsFlipper, triggerX, triggerY, triggerAction, triggerScore, triggerValue, lightX, lightY, lightGroup, arrowX, arrowY, multiX, multiY, letterX, letterY
    
    Open strFileName For Binary Access Read As #1
    Get #1, , numWalls
    If numWalls > 0 Then
        ReDim numPointsWall(numWalls - 1)
        ReDim wallX(1023, numWalls - 1)
        ReDim wallY(1023, numWalls - 1)
        ReDim wallAction(1023, numWalls - 1)
        ReDim wallNames(numWalls - 1)
        ReDim wallChecked(numWalls - 1)
    End If
    Get #1, , numPointsWall()
    Get #1, , wallX()
    Get #1, , wallY()
    Get #1, , wallAction()
    Get #1, , wallNames()
    Get #1, , wallChecked()
    
    Get #1, , numPointsFlipper()
    Get #1, , flipperX()
    Get #1, , flipperY()
    Get #1, , flipperNames()
    Get #1, , flipperChecked()
    
    Get #1, , numTriggers
    If numTriggers > 0 Then
        ReDim triggerX(numTriggers - 1)
        ReDim triggerY(numTriggers - 1)
        ReDim triggerAction(numTriggers - 1)
        ReDim triggerScore(numTriggers - 1)
        ReDim triggerValue(numTriggers - 1)
        ReDim triggerNames(numTriggers - 1)
        ReDim triggerChecked(numTriggers - 1)
    End If
    Get #1, , triggerX()
    Get #1, , triggerY()
    Get #1, , triggerAction()
    Get #1, , triggerScore()
    Get #1, , triggerValue()
    Get #1, , triggerNames()
    Get #1, , triggerChecked()
    
    Get #1, , numLights
    If numLights > 0 Then
        ReDim lightX(numLights - 1)
        ReDim lightY(numLights - 1)
        ReDim lightGroup(numLights - 1)
        ReDim lightNames(numLights - 1)
        ReDim lightChecked(numLights - 1)
    End If
    Get #1, , lightX()
    Get #1, , lightY()
    Get #1, , lightGroup()
    Get #1, , lightNames()
    Get #1, , lightChecked()
    
    Get #1, , numArrows
    If numArrows > 0 Then
        ReDim arrowX(numArrows - 1)
        ReDim arrowY(numArrows - 1)
        ReDim arrowNames(numArrows - 1)
        ReDim arrowChecked(numArrows - 1)
    End If
    Get #1, , arrowX()
    Get #1, , arrowY()
    Get #1, , arrowNames()
    Get #1, , arrowChecked()
    
    Get #1, , numMultis
    If numMultis > 0 Then
        ReDim multiX(numMultis - 1)
        ReDim multiY(numMultis - 1)
        ReDim multiNames(numMultis - 1)
        ReDim multiChecked(numMultis - 1)
    End If
    Get #1, , multiX()
    Get #1, , multiY()
    Get #1, , multiNames()
    Get #1, , multiChecked()
    
    Get #1, , numLetters
    If numLetters > 0 Then
        ReDim letterX(numLetters - 1)
        ReDim letterY(numLetters - 1)
        ReDim letterNames(numLetters - 1)
        ReDim letterChecked(numLetters - 1)
    End If
    Get #1, , letterX()
    Get #1, , letterY()
    Get #1, , letterNames()
    Get #1, , letterChecked()
    Close #1
    
    lvwObjects.ListItems.Clear
    
    For i = 0 To numWalls - 1
        Set objObject = lvwObjects.ListItems.Add(, "Wall " & i + 1, wallNames(i))
        objObject.Checked = wallChecked(i)
    Next i
    For i = 0 To 3
        For j = 0 To 4
            Set objObject = lvwObjects.ListItems.Add(, "Flipper" & Chr(i + Asc("A")) & " " & j + 1, flipperNames(i, j))
            objObject.Checked = flipperChecked(i, j)
        Next j
    Next i
    For i = 0 To numTriggers - 1
        Set objObject = lvwObjects.ListItems.Add(, "Trigger " & i + 1, triggerNames(i))
        objObject.Checked = triggerChecked(i)
    Next i
    For i = 0 To numLights - 1
        Set objObject = lvwObjects.ListItems.Add(, "Light " & i + 1, lightNames(i))
        objObject.Checked = lightChecked(i)
    Next i
    For i = 0 To numArrows - 1
        Set objObject = lvwObjects.ListItems.Add(, "Arrow " & i + 1, arrowNames(i))
        objObject.Checked = arrowChecked(i)
    Next i
    For i = 0 To numMultis - 1
        Set objObject = lvwObjects.ListItems.Add(, "Multi " & i + 1, multiNames(i))
        objObject.Checked = multiChecked(i)
    Next i
    For i = 0 To numLetters - 1
        Set objObject = lvwObjects.ListItems.Add(, "Letter " & i + 1, letterNames(i))
        objObject.Checked = letterChecked(i)
    Next i
    
    lvwObjects.ListItems(1).Selected = True
    lvwObjects_ItemClick lvwObjects.SelectedItem
    DrawObjects
    blnDirty = False
End Sub

Private Sub NewFile()
    Dim i As Integer, j As Integer
    Dim strName As String
    
    numWalls = 0
    numTriggers = 0
    numLights = 0
    numArrows = 0
    numMultis = 0
    numLetters = 0
    
    Erase wallX, wallY, wallAction, numPointsWall, flipperX, flipperY, numPointsFlipper, triggerX, triggerY, triggerAction, triggerScore, triggerValue, lightX, lightY, lightGroup, arrowX, arrowY, multiX, multiY, letterX, letterY
    lvwObjects.ListItems.Clear
    
    For i = 0 To 3
        For j = 0 To 4
            strName = "Flipper" & Chr(i + Asc("A")) & " " & j + 1
            lvwObjects.ListItems.Add , strName, strName
        Next j
    Next i
    
    blnDirty = False
    lvwObjects_ItemClick lvwObjects.SelectedItem
End Sub

Private Function PromptToSave()
    Dim intResponse As Integer
    
    PromptToSave = MsgBox("Do you want to save the changes you made to " & GetFileName & "?", vbYesNoCancel + vbQuestion)
End Function

Private Function GetFileName()
    If strFileName = "" Then
        GetFileName = "Untitled"
    Else
        GetFileName = dlgFile.FileTitle
    End If
End Function

Private Sub UpdateTitleBar()
    Caption = "Pinball Table Editor - " & GetFileName
End Sub

Private Sub ExportFile()
    Dim i As Integer, j As Integer, k As Integer
    Dim numFlippers As Integer
    Dim numGroups As Integer
    Dim blnFoundLight As Boolean
    Dim flipperTop(3) As Integer
    Dim flipperBottom(3) As Integer
    Dim intLowest As Integer, intHighest As Integer
    
    For i = 0 To 3
        If numPointsFlipper(i, 0) > 0 Then numFlippers = i + 1
    Next i
    
    For i = 0 To numFlippers - 1
        intLowest = 0
        For j = 0 To numPointsFlipper(i, 0) - 1
            If flipperY(j, i, 0) > intLowest Then intLowest = flipperY(j, i, 0)
        Next j
        flipperBottom(i) = intLowest + 1
        intHighest = 383
        For j = 0 To numPointsFlipper(i, 4) - 1
            If flipperY(j, i, 4) < intHighest Then intHighest = flipperY(j, i, 4)
        Next j
        flipperTop(i) = intHighest - 1
    Next i
    
    numGroups = 1
    For i = 0 To numLights - 1
        If lightGroup(i) > numGroups Then numGroups = lightGroup(i)
    Next i
    
    Open dlgExport.FileName For Output Access Write As #1
    Print #1, Chr(9); "final int NUM_WALLS = "; Format(numWalls); ";"
    Print #1, Chr(9); "final int NUM_FLIPPERS = "; Format(numFlippers); ";"
    Print #1, Chr(9); "final int NUM_FLIPPER_SIDES = "; Format(numPointsFlipper(0, 0)); ";"
    Print #1, Chr(9); "final int NUM_TRIGGERS = "; Format(numTriggers); ";"
    Print #1, Chr(9); "final int NUM_LIGHTS = "; Format(numLights); ";"
    Print #1, Chr(9); "final int NUM_GROUPS = "; Format(numGroups); ";"
    Print #1, Chr(9); "final int NUM_ARROWS = "; Format(numArrows); ";"
    Print #1, Chr(9); "final int NUM_MULTIS = "; Format(numMultis); ";"
    Print #1, Chr(9); "final int NUM_LETTERS = "; Format(numLetters); ";"
    
    Print #1, Chr(9); "final int wallX[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To numWalls - 1
        Print #1, "{";
        For j = 0 To numPointsWall(i) - 1
            Print #1, Format(wallX(j, i)); ", ";
        Next j
        Print #1, Format(wallX(0, i)); "}";
        If i < numWalls - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int wallY[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To numWalls - 1
        Print #1, "{";
        For j = 0 To numPointsWall(i) - 1
            Print #1, Format(wallY(j, i)); ", ";
        Next j
        Print #1, Format(wallY(0, i)); "}";
        If i < numWalls - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int bump[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To numWalls - 1
        Print #1, "{";
        For j = 0 To numPointsWall(i) - 2
            Print #1, Format(wallAction(j, i)); ", ";
        Next j
        Print #1, Format(wallAction(j, i)); "}";
        If i < numWalls - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int flipperX[][][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To numFlippers - 1
        Print #1, "{";
        For j = 0 To 4
            Print #1, "{";
            For k = 0 To numPointsFlipper(i, j) - 1
                Print #1, Format(flipperX(k, i, j)); ", ";
            Next k
            Print #1, Format(flipperX(0, i, j)); "}";
            If j < 4 Then
                Print #1, ","
                Print #1, Chr(9); Chr(9);
            End If
        Next j
        Print #1, "}";
        If i < numFlippers - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int flipperY[][][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 0 To numFlippers - 1
        Print #1, "{";
        For j = 0 To 4
            Print #1, "{";
            For k = 0 To numPointsFlipper(i, j) - 1
                Print #1, Format(flipperY(k, i, j)); ", ";
            Next k
            Print #1, Format(flipperY(0, i, j)); "}";
            If j < 4 Then
                Print #1, ","
                Print #1, Chr(9); Chr(9);
            End If
        Next j
        Print #1, "}";
        If i < numFlippers - 1 Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int flipperTop[] = {";
    For i = 0 To numFlippers - 2
        Print #1, Format(flipperTop(i)); ", ";
    Next i
    If numFlippers > 0 Then Print #1, Format(flipperTop(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int flipperBottom[] = {";
    For i = 0 To numFlippers - 2
        Print #1, Format(flipperBottom(i)); ", ";
    Next i
    If numFlippers > 0 Then Print #1, Format(flipperBottom(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int triggerX[] = {";
    For i = 0 To numTriggers - 2
        Print #1, Format(triggerX(i)); ", ";
    Next i
    If numTriggers > 0 Then Print #1, Format(triggerX(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int triggerY[] = {";
    For i = 0 To numTriggers - 2
        Print #1, Format(triggerY(i)); ", ";
    Next i
    If numTriggers > 0 Then Print #1, Format(triggerY(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int triggerAct[] = {";
    For i = 0 To numTriggers - 2
        Print #1, Format(triggerAction(i)); ", ";
    Next i
    If numTriggers > 0 Then Print #1, Format(triggerAction(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int triggerScore[] = {";
    For i = 0 To numTriggers - 2
        Print #1, Format(triggerScore(i)); ", ";
    Next i
    If numTriggers > 0 Then Print #1, Format(triggerScore(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int triggerVal[] = {";
    For i = 0 To numTriggers - 2
        Print #1, Format(triggerValue(i)); ", ";
    Next i
    If numTriggers > 0 Then Print #1, Format(triggerValue(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int lightX[] = {";
    For i = 0 To numLights - 2
        Print #1, Format(lightX(i)); ", ";
    Next i
    If numLights > 0 Then Print #1, Format(lightX(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int lightY[] = {";
    For i = 0 To numLights - 2
        Print #1, Format(lightY(i)); ", ";
    Next i
    If numLights > 0 Then Print #1, Format(lightY(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int group[][] ="
    Print #1, Chr(9); Chr(9); "{";
    For i = 1 To numGroups
        Print #1, "{";
        blnFoundLight = False
        For j = 0 To numLights - 1
            If lightGroup(j) = i Then
                If blnFoundLight Then Print #1, ", "; Else blnFoundLight = True
                Print #1, Format(j);
            End If
        Next j
        Print #1, "}";
        If i < numGroups Then
            Print #1, ","
            Print #1, Chr(9); Chr(9);
        End If
    Next i
    Print #1, "};"
    
    Print #1, Chr(9); "final int arrowX[] = {";
    For i = 0 To numArrows - 2
        Print #1, Format(arrowX(i)); ", ";
    Next i
    If numArrows > 0 Then Print #1, Format(arrowX(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int arrowY[] = {";
    For i = 0 To numArrows - 2
        Print #1, Format(arrowY(i)); ", ";
    Next i
    If numArrows > 0 Then Print #1, Format(arrowY(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int multiX[] = {";
    For i = 0 To numMultis - 2
        Print #1, Format(multiX(i)); ", ";
    Next i
    If numMultis > 0 Then Print #1, Format(multiX(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int multiY[] = {";
    For i = 0 To numMultis - 2
        Print #1, Format(multiY(i)); ", ";
    Next i
    If numMultis > 0 Then Print #1, Format(multiY(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int letterX[] = {";
    For i = 0 To numLetters - 2
        Print #1, Format(letterX(i)); ", ";
    Next i
    If numLetters > 0 Then Print #1, Format(letterX(i));
    Print #1, "};"
    
    Print #1, Chr(9); "final int letterY[] = {";
    For i = 0 To numLetters - 2
        Print #1, Format(letterY(i)); ", ";
    Next i
    If numLetters > 0 Then Print #1, Format(letterY(i));
    Print #1, "};"
    
    Close #1
End Sub
