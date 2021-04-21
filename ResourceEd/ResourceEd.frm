VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmResourceEd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ResourceEd"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9510
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   5520
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Browse"
      Filter          =   "All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgPreset 
      Left            =   3960
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "red"
      Filter          =   "ResourceEd Files (*.red)|*.red|All Files (*.*)|*.*"
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clea&r"
      Height          =   615
      Left            =   2760
      Picture         =   "ResourceEd.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save Preset..."
      Height          =   615
      Left            =   1440
      Picture         =   "ResourceEd.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Preset..."
      Height          =   615
      Left            =   120
      Picture         =   "ResourceEd.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraContents 
      Caption         =   "File Con&tents"
      Height          =   3375
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   5655
      Begin VB.CommandButton cmdOpenWith 
         Caption         =   "Open &With..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   1560
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdExport 
         Caption         =   "&Export..."
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   15
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Enabled         =   0   'False
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2880
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvwFiles 
         Height          =   2535
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Ext'n"
            Object.Width           =   1191
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Size"
            Object.Width           =   2328
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "#"
            Object.Width           =   794
         EndProperty
      End
   End
   Begin VB.Frame fraFiles 
      Caption         =   "Files"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   5655
      Begin VB.TextBox txtIndexFile 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton cmdBrowseIndex 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   4440
         TabIndex        =   7
         Top             =   345
         Width           =   975
      End
      Begin VB.TextBox txtDataFile 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.CommandButton cmdBrowseData 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   4440
         TabIndex        =   10
         Top             =   705
         Width           =   975
      End
      Begin VB.Label lblIndexFile 
         Caption         =   "I&ndex File:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   405
         Width           =   735
      End
      Begin VB.Label lblDataFile 
         Caption         =   "Data F&ile:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   765
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdCreateIndex 
      Caption         =   "&Create Index"
      Default         =   -1  'True
      Height          =   615
      Left            =   4560
      Picture         =   "ResourceEd.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraIndex 
      Caption         =   "Index structure"
      Height          =   5415
      Left            =   6000
      TabIndex        =   16
      Top             =   120
      Width           =   3375
      Begin VB.TextBox txtExtension 
         Height          =   285
         Left            =   1800
         TabIndex        =   35
         Top             =   4920
         Width           =   1335
      End
      Begin VB.TextBox txtFooter 
         Height          =   285
         Left            =   1800
         TabIndex        =   32
         Text            =   "0"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add Field"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1920
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete Field"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Frame fraProperties 
         Caption         =   "Field properties"
         Height          =   1695
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   2895
         Begin MSComCtl2.UpDown updSize 
            Height          =   285
            Left            =   2400
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   1200
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtSize"
            BuddyDispid     =   196632
            OrigLeft        =   1440
            OrigTop         =   1320
            OrigRight       =   1680
            OrigBottom      =   1605
            Max             =   65535
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.ComboBox cmbType 
            Height          =   315
            ItemData        =   "ResourceEd.frx":0408
            Left            =   240
            List            =   "ResourceEd.frx":0421
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   600
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.ComboBox cmbSize 
            Height          =   315
            ItemData        =   "ResourceEd.frx":0474
            Left            =   240
            List            =   "ResourceEd.frx":0481
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1200
            Visible         =   0   'False
            Width           =   2415
         End
         Begin VB.TextBox txtSize 
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   1200
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label lblType 
            Caption         =   "T&ype:"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblSize 
            Caption         =   "Si&ze:"
            Height          =   255
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.TextBox txtHeader 
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Text            =   "0"
         Top             =   4440
         Width           =   1095
      End
      Begin MSComctlLib.ListView lvwRecord 
         Height          =   1215
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Field type"
            Object.Width           =   2249
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Size"
            Object.Width           =   2275
         EndProperty
      End
      Begin MSComCtl2.UpDown updHeader 
         Height          =   285
         Left            =   1320
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   4440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtHeader"
         BuddyDispid     =   196635
         OrigLeft        =   1800
         OrigTop         =   600
         OrigRight       =   2040
         OrigBottom      =   885
         Max             =   65535
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown updFooter 
         Height          =   285
         Left            =   2880
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   4440
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         BuddyControl    =   "txtFooter"
         BuddyDispid     =   196626
         OrigLeft        =   1800
         OrigTop         =   600
         OrigRight       =   2040
         OrigBottom      =   885
         Max             =   65535
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblExtension 
         Caption         =   "Default E&xtension:"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   4965
         Width           =   1335
      End
      Begin VB.Label lblFooter 
         Caption         =   "&Footer size:"
         Height          =   255
         Left            =   1800
         TabIndex        =   31
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblRecordStruct 
         Caption         =   "Record str&ucture:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblHeader 
         Caption         =   "&Header size:"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   4200
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmResourceEd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type ArchiveFile
    FileName As String
    Address As Long
    Size As Long
    FileNameLen As Integer
    Data() As Variant
End Type

Private Type Preset
    IndexFile As String
    DataFile As String
    RecordDef() As Integer
    HeaderSize As Integer
    FooterSize As Integer
    DefaultExt As String
End Type


Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type


Private Const SEE_MASK_INVOKEIDLIST = &HC
Private Const SEE_MASK_NOCLOSEPROCESS = &H40
Private Const SEE_MASK_FLAG_NO_UI = &H400

Private Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type


Private blnAddressField As Boolean
Private blnSizeField As Boolean
Private blnFileNameField As Boolean
Private blnFNLenField As Boolean
Private blnStopComboClick As Boolean
Private colTempFiles As New Collection
Private intDataField As Integer
Private intLastSort As Integer
Private intNumDataFields As Integer
Private intNumFields As Integer
Private intNumFiles As Integer
Private intRecordDef() As Integer
Private lngArchiveLen As Long
Private lngTotalDataLen As Long
Private udtFileArray() As ArchiveFile

Private Sub cmbSize_Click()
    Dim intIndex As Integer
    
    intIndex = lvwRecord.SelectedItem.Index - 1
    intRecordDef(1, intIndex) = cmbSize.ListIndex
    UpdateField intIndex
End Sub

Private Sub cmbType_Click()
    Dim intIndex As Integer
    Dim i As Integer
    Dim blnVariable As Boolean
    
    If blnStopComboClick Then
        blnStopComboClick = False
        Exit Sub
    End If
    intIndex = lvwRecord.SelectedItem.Index - 1
    intRecordDef(0, intIndex) = cmbType.ListIndex
    intRecordDef(1, intIndex) = Choose(intRecordDef(0, intIndex) + 1, 12, 2, 2, 0, 2, 0, 0)
    
    If intRecordDef(0, intIndex) = 0 Then
        For i = 0 To intIndex - 1
            If intRecordDef(0, i) = 3 Then
                blnVariable = True
                Exit For
            End If
        Next i
    End If
    
    If blnVariable Then
        intRecordDef(1, intIndex) = -1
        lblSize.Visible = False
        cmbSize.Visible = False
        txtSize.Visible = False
        updSize.Visible = False
    Else
        ShowSize intIndex
        ShowSizeControl intIndex
    End If
    
    UpdateField intIndex
End Sub

Private Sub cmdAdd_Click()
    Dim itmX As ListItem
    
    Set itmX = lvwRecord.ListItems.Add
    ReDim Preserve intRecordDef(1, intNumFields)
    intRecordDef(0, itmX.Index - 1) = cmbType.ListCount - 1
    intNumFields = intNumFields + 1
    lvwRecord.SelectedItem = itmX
    lvwRecord_ItemClick itmX
    cmbType_Click
    cmdDelete.Enabled = True
    cmbType.SetFocus
End Sub

Private Sub cmdBrowseData_Click()
    On Error Resume Next
    dlgBrowse.FileName = txtDataFile.Text
    dlgBrowse.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    txtDataFile.Text = dlgBrowse.FileName
End Sub

Private Sub cmdBrowseIndex_Click()
    On Error Resume Next
    dlgBrowse.FileName = txtIndexFile.Text
    dlgBrowse.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    txtIndexFile.Text = dlgBrowse.FileName
End Sub

Private Sub cmdClear_Click()
    txtIndexFile.Text = ""
    txtDataFile.Text = ""
    ClearContents
    lvwRecord.ListItems.Clear
    Erase intRecordDef
    intNumFields = 0
    txtHeader.Text = "0"
    txtFooter.Text = "0"
    txtExtension.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Dim i As Integer
    Dim intIndex As Integer
    Dim blnVariable As Boolean
    
    intIndex = lvwRecord.SelectedItem.Index - 1
    
    For i = intIndex To intNumFields - 2
        intRecordDef(0, i) = intRecordDef(0, i + 1)
        intRecordDef(1, i) = intRecordDef(1, i + 1)
    Next i
    
    intNumFields = intNumFields - 1
    If intNumFields = 0 Then
        cmdDelete.Enabled = False
    Else
        ReDim Preserve intRecordDef(1, intNumFields - 1)
    End If
    lvwRecord.ListItems.Remove intIndex + 1
    lblType.Visible = False
    cmbType.Visible = False
    lblSize.Visible = False
    cmbSize.Visible = False
    txtSize.Visible = False
    updSize.Visible = False
    
    For i = 0 To intNumFields - 1
        If intRecordDef(0, i) = 3 Then blnVariable = True
        If intRecordDef(1, i) = -1 And Not blnVariable Then
            intRecordDef(1, i) = 12
            UpdateField i
        End If
    Next i
End Sub

Private Sub cmdExport_Click()
    'Opens a Browse Folders Dialog Box that displays the
    'directories in your computer
    Dim lpIDList As Long ' Declare Varibles
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    
    Dim itmX As ListItem
    Dim fso As New FileSystemObject
    
    szTitle = "Please select a folder to export the selected files into."
    ' Text to appear in the the gray area under the
    ' title bar telling you what to do
    With tBrowseInfo
       .hWndOwner = Me.hwnd ' Owner Form
       .lpszTitle = lstrcat(szTitle, "")
       .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)
    If lpIDList = 0 Then Exit Sub
    sBuffer = Space(MAX_PATH)
    SHGetPathFromIDList lpIDList, sBuffer
    sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
    
    For Each itmX In lvwFiles.ListItems
        If itmX.Selected Then ExportFile itmX.Tag, fso.BuildPath(sBuffer, fso.GetFileName(itmX.Text)), True
        If Err.Number Then Exit Sub
    Next itmX
    
    MsgBox "All selected files were successfully exported", vbInformation
End Sub

Private Sub cmdLoad_Click()
    Dim udtPreset As Preset
    Dim i As Integer
    Dim itmX As ListItem
    
    On Error Resume Next
    dlgPreset.DialogTitle = "Load Preset"
    dlgPreset.ShowOpen
    If Err.Number = cdlCancel Then Exit Sub
    dlgPreset.InitDir = ""
    
    Open dlgPreset.FileName For Binary Access Read As #1
    If Err.Number Then
        MsgBox "Cannot open preset file '" & dlgPreset.FileName & "'", vbExclamation
        Exit Sub
    End If
    Get #1, , udtPreset
    Close #1
    
    With udtPreset
        txtIndexFile.Text = .IndexFile
        txtDataFile.Text = .DataFile
        intRecordDef() = .RecordDef
        txtHeader.Text = .HeaderSize
        txtFooter.Text = .FooterSize
        txtExtension.Text = .DefaultExt
    End With
    
    intNumFields = UBound(intRecordDef, 2) + 1
    lvwRecord.ListItems.Clear
    cmdOpen.Enabled = False
    cmdOpenWith.Enabled = False
    cmdExport.Enabled = False
    
    For i = 0 To intNumFields - 1
        itmX = lvwRecord.ListItems.Add(, , cmbType.List(intRecordDef(0, i)))
        UpdateField i
    Next i
    
    If intNumFields > 0 Then cmdDelete.Enabled = True
    ClearContents
End Sub

Private Sub cmdCreateIndex_Click()
    Dim strHeader As String
    Dim i As Integer
    Dim itmX As ListItem
    Dim blnLoopWhile As Boolean
    Dim varData As Variant
    Dim fso As New FileSystemObject
    
    On Error Resume Next
    If Not CheckRecordStruct Then Exit Sub
    Screen.MousePointer = vbHourglass
    intNumFiles = 0
    lngTotalDataLen = 0
    Open txtIndexFile.Text For Binary Access Read As #1
    If Err.Number Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot open index file '" & txtIndexFile.Text & "'", vbExclamation
        GoTo AbortIndex
    End If
    strHeader = String(txtHeader.Text, 0)
    Get #1, , strHeader
    
    Do
        ReDim Preserve udtFileArray(intNumFiles)
        If intNumDataFields > 0 Then ReDim udtFileArray(intNumFiles).Data(intNumDataFields - 1)
        intDataField = 0
        
        For i = 0 To intNumFields - 1
            LoadValue intNumFiles, i
            If Err.Number Then GoTo AbortIndex
        Next i
        
        If Not blnFileNameField Then udtFileArray(intNumFiles).FileName = "File" & Format(intNumFiles + 1, "0000") & IIf(txtExtension = "", "", "." & txtExtension.Text)
        Set itmX = lvwFiles.ListItems.Add(intNumFiles + 1, , udtFileArray(intNumFiles).FileName)
        itmX.Tag = intNumFiles
        itmX.SubItems(1) = fso.GetExtensionName(udtFileArray(intNumFiles).FileName)
        If blnSizeField Then itmX.SubItems(2) = Format(udtFileArray(intNumFiles).Size, "#,0") & " bytes"
        itmX.SubItems(3) = intNumFiles + 1
        For i = 0 To intNumDataFields - 1
            varData = udtFileArray(intNumFiles).Data(i)
            itmX.SubItems(4 + i) = IIf(IsNumeric(varData), Format(varData, "#,0"), varData)
        Next i
        
        If LCase(txtIndexFile.Text) <> LCase(txtDataFile) Then
            blnLoopWhile = Loc(1) + txtFooter.Text < LOF(1)
        ElseIf Not blnAddressField Then
            blnLoopWhile = Loc(1) + lngTotalDataLen < LOF(1)
        Else
            blnLoopWhile = Loc(1) + txtFooter.Text < udtFileArray(0).Address
        End If
            
        intNumFiles = intNumFiles + 1
    Loop While blnLoopWhile
    
    If blnAddressField And Not blnSizeField Then FillSize
    If Err.Number Then GoTo AbortIndex
    If Not blnAddressField And blnSizeField Then FillAddress
    cmdOpen.Enabled = True
    cmdOpenWith.Enabled = True
    cmdExport.Enabled = True
    Screen.MousePointer = vbDefault

AbortIndex:
    Close #1
End Sub

Private Sub cmdOpen_Click()
    Dim intFileIndex As Integer
    Dim lngTempPathLen As Long
    Dim strTempPath As String
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long

    intFileIndex = lvwFiles.SelectedItem.Tag
    strTempPath = String(MAX_PATH, 0)
    lngTempPathLen = GetTempPath(MAX_PATH, strTempPath)
    strTempPath = Left(strTempPath, lngTempPathLen)
    strFileName = fso.BuildPath(strTempPath, fso.GetFileName(udtFileArray(intFileIndex).FileName))
    colTempFiles.Add strFileName
    ExportFile intFileIndex, strFileName, False
    If Err.Number Then Exit Sub
    
    With SEI
      .cbSize = Len(SEI)
      .fMask = SEE_MASK_NOCLOSEPROCESS Or _
      SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
      .hwnd = Me.hwnd
      .lpVerb = "open"
      .lpFile = strFileName
      .lpParameters = vbNullChar
      .lpDirectory = vbNullChar
      .nShow = 0
      .hInstApp = 0
      .lpIDList = 0
    End With
         
    r = ShellExecuteEX(SEI)
End Sub

Private Sub cmdOpenWith_Click()
    Dim intFileIndex As Integer
    Dim lngTempPathLen As Long
    Dim strTempPath As String
    Dim strFileName As String
    Dim fso As New FileSystemObject
    Dim r As Long

    intFileIndex = lvwFiles.SelectedItem.Tag
    strTempPath = String(MAX_PATH, 0)
    lngTempPathLen = GetTempPath(MAX_PATH, strTempPath)
    strTempPath = Left(strTempPath, lngTempPathLen)
    strFileName = fso.BuildPath(strTempPath, fso.GetFileName(udtFileArray(intFileIndex).FileName))
    colTempFiles.Add strFileName
    ExportFile intFileIndex, strFileName, False
    If Err.Number Then Exit Sub
    r = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & strFileName)
End Sub

Private Sub cmdSave_Click()
    Dim udtPreset As Preset
    
    On Error Resume Next
    dlgPreset.DialogTitle = "Save Preset"
    dlgPreset.ShowSave
    If Err.Number = cdlCancel Then Exit Sub
    dlgPreset.InitDir = ""
    
    With udtPreset
        .IndexFile = txtIndexFile.Text
        .DataFile = txtDataFile.Text
        .RecordDef = intRecordDef()
        .HeaderSize = txtHeader.Text
        .FooterSize = txtFooter.Text
        .DefaultExt = txtExtension.Text
    End With
    
    Open dlgPreset.FileName For Binary Access Write As #1
    If Err.Number Then
        MsgBox "Cannot open preset file '" & dlgPreset.FileName & "'", vbExclamation
        Exit Sub
    End If
    Put #1, , udtPreset
    Close #1
End Sub

Private Sub Form_Load()
    dlgPreset.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    dlgBrowse.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNPathMustExist
    dlgPreset.InitDir = CurDir
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim varTempFile As Variant
    
    On Error Resume Next
    For Each varTempFile In colTempFiles
        Kill varTempFile
    Next varTempFile
End Sub

Private Sub lvwFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim itmX As ListItem
    Dim varData
    Dim lngEndPos As Long

    For Each itmX In lvwFiles.ListItems
        Select Case ColumnHeader.Index
            Case 3
                itmX.SubItems(2) = Format(udtFileArray(itmX.Tag).Size, "0000000000")
            Case 4
                itmX.SubItems(3) = Format(itmX.Tag + 1, "0000000000")
            Case Is >= 5
                varData = udtFileArray(itmX.Tag).Data(ColumnHeader.Index - 5)
                If IsNumeric(varData) Then itmX.SubItems(ColumnHeader.Index - 1) = Format(varData, "0000000000")
        End Select
    Next itmX

    lvwFiles.SortKey = ColumnHeader.Index - 1
    If intLastSort = ColumnHeader.Index Then
        lvwFiles.SortOrder = -Not -lvwFiles.SortOrder
    Else
        lvwFiles.SortOrder = lvwAscending
    End If
    lvwFiles.Sorted = True

    For Each itmX In lvwFiles.ListItems

        Select Case ColumnHeader.Index
            Case 3
                itmX.SubItems(2) = Format(udtFileArray(itmX.Tag).Size, "0,#") & " bytes"
            Case 4
                itmX.SubItems(3) = itmX.Tag + 1
            Case Is >= 5
                varData = udtFileArray(itmX.Tag).Data(ColumnHeader.Index - 5)
                itmX.SubItems(ColumnHeader.Index - 1) = IIf(IsNumeric(varData), Format(varData, "#,0"), varData)
        End Select
    Next itmX

    lvwFiles.Sorted = False
    intLastSort = ColumnHeader.Index
End Sub

Private Sub lvwFiles_DblClick()
    If cmdOpen.Enabled Then cmdOpen.Value = True
End Sub

Private Sub lvwRecord_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim intIndex As Integer
    
    intIndex = Item.Index - 1
    blnStopComboClick = True
    cmbType.ListIndex = intRecordDef(0, intIndex)
    lblType.Visible = True
    cmbType.Visible = True
    lblSize.Visible = True
    ShowSize intIndex
    ShowSizeControl intIndex
End Sub

Private Sub txtDataFile_GotFocus()
    txtDataFile.SelStart = 0
    txtDataFile.SelLength = Len(txtDataFile.Text)
End Sub

Private Sub txtExtension_GotFocus()
    txtExtension.SelStart = 0
    txtExtension.SelLength = Len(txtExtension.Text)
End Sub

Private Sub txtFooter_GotFocus()
    txtFooter.SelStart = 0
    txtFooter.SelLength = Len(txtFooter.Text)
End Sub

Private Sub txtHeader_GotFocus()
    txtHeader.SelStart = 0
    txtHeader.SelLength = Len(txtHeader.Text)
End Sub

Private Sub txtIndexFile_GotFocus()
    txtIndexFile.SelStart = 0
    txtIndexFile.SelLength = Len(txtIndexFile.Text)
End Sub

Private Sub txtSize_Change()
    Dim intIndex As Integer
    
    If Val(txtSize.Text) = 0 Then txtSize.Text = "zero-terminated"
    intIndex = lvwRecord.SelectedItem.Index - 1
    intRecordDef(1, intIndex) = Val(txtSize.Text)
    UpdateField intIndex
End Sub

Private Sub UpdateField(intIndex As Integer)
    lvwRecord.ListItems(intIndex + 1).Text = cmbType.List(intRecordDef(0, intIndex))
    
    If intRecordDef(0, intIndex) >= 1 And intRecordDef(0, intIndex) <= 4 Then
        lvwRecord.ListItems(intIndex + 1).SubItems(1) = cmbSize.List(intRecordDef(1, intIndex))
    Else
        Select Case intRecordDef(1, intIndex)
            Case -1
                lvwRecord.ListItems(intIndex + 1).SubItems(1) = "variable"
            Case 0
                lvwRecord.ListItems(intIndex + 1).SubItems(1) = "zero-terminated"
            Case Else
                lvwRecord.ListItems(intIndex + 1).SubItems(1) = intRecordDef(1, intIndex) & " byte" & IIf(intRecordDef(1, intIndex) = 1, "", "s")
        End Select
    End If
End Sub

Sub ShowSizeControl(intIndex As Integer)
    lblSize.Visible = True
    If intRecordDef(0, intIndex) >= 1 And intRecordDef(0, intIndex) <= 4 Then
        cmbSize.Visible = True
        txtSize.Visible = False
        updSize.Visible = False
    Else
        txtSize.Visible = True
        updSize.Visible = True
        cmbSize.Visible = False
    End If
End Sub

Sub ShowSize(intIndex As Integer)
    If intRecordDef(0, intIndex) >= 1 And intRecordDef(0, intIndex) <= 4 Then
        cmbSize.ListIndex = intRecordDef(1, intIndex)
    Else
        txtSize.Text = IIf(intRecordDef(1, intIndex) = 0, "zero-terminated", intRecordDef(1, intIndex))
    End If
End Sub

Private Sub LoadValue(intIndex As Integer, intField As Integer)
    Dim bytByte As Byte
    Dim intInteger As Integer
    Dim lngLong As Long
    Dim strString As String
    Dim varValue As Variant
    Dim intStringLen As Integer
    Dim lngDataFileLen As Long
    Dim lngZeroPos As Long
    
    On Error Resume Next
    lngDataFileLen = FileLen(txtDataFile.Text)
    If Err.Number Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot find data file '" & txtDataFile.Text & "'", vbExclamation
        Exit Sub
    End If
    
    If intRecordDef(0, intField) >= 1 And intRecordDef(0, intField) <= 4 Then
        Select Case intRecordDef(1, intField)
            Case 0
                Get #1, , bytByte
                varValue = bytByte
            Case 1
                Get #1, , intInteger
                varValue = intInteger
            Case 2
                Get #1, , lngLong
                varValue = lngLong
        End Select
    Else
        intStringLen = intRecordDef(1, intField)
        Select Case intStringLen
            Case -1
                strString = String(udtFileArray(intIndex).FileNameLen, 0)
                Get #1, , strString
            Case 0
                strString = GetString
            Case Else
                strString = String(intStringLen, 0)
                Get #1, , strString
        End Select
        varValue = strString
    End If
    
    Select Case intRecordDef(0, intField)
        Case 0
            lngZeroPos = InStr(varValue, vbNullChar)
            If lngZeroPos > 0 Then
                udtFileArray(intIndex).FileName = Left(varValue, lngZeroPos - 1)
            Else
                udtFileArray(intIndex).FileName = varValue
            End If
        Case 1
            udtFileArray(intIndex).Address = varValue
            If varValue < 0 Or varValue >= lngDataFileLen Then
                Screen.MousePointer = vbDefault
                MsgBox "An invalid address was read from the index. Please check that the index structure is correct.", vbCritical
                Err.Raise 513
                Exit Sub
            End If
        Case 2
            udtFileArray(intIndex).Size = varValue
            If varValue < 0 Or varValue >= lngDataFileLen Then
                Screen.MousePointer = vbDefault
                MsgBox "An invalid size was read from the index. Please check that the index structure is correct.", vbCritical
                Err.Raise 513
                Exit Sub
            End If
            lngTotalDataLen = lngTotalDataLen + varValue
            If lngTotalDataLen > lngDataFileLen Then
                Screen.MousePointer = vbDefault
                MsgBox "The amount of data specified in the index has exceeded the size of the data file. Please check that the index structure is correct.", vbCritical
                Err.Raise 513
                Exit Sub
            End If
        Case 3
            udtFileArray(intIndex).FileNameLen = varValue
        Case 4, 5
            udtFileArray(intIndex).Data(intDataField) = varValue
            intDataField = intDataField + 1
    End Select
End Sub

Private Function CheckRecordStruct() As Boolean
    Dim i As Integer
    Dim itmX As ColumnHeader

    If intNumFields = 0 Then
        MsgBox "Record structure not defined", vbExclamation
        Exit Function
    End If
    blnFileNameField = False
    blnAddressField = False
    blnSizeField = False
    blnFNLenField = False
    intNumDataFields = 0
    ClearContents
    lvwFiles.Visible = False

    For i = 0 To intNumFields - 1
        Select Case intRecordDef(0, i)
            Case 0
                If blnFileNameField Then
                    lvwFiles.Visible = True
                    MsgBox "Only one file name field is allowed.", vbExclamation
                    Exit Function
                Else
                    blnFileNameField = True
                End If
            Case 1
                If blnAddressField Then
                    lvwFiles.Visible = True
                    MsgBox "Only one address field is allowed.", vbExclamation
                    Exit Function
                Else
                    blnAddressField = True
                End If
            Case 2
                If blnSizeField Then
                    lvwFiles.Visible = True
                    MsgBox "Only one size field is allowed.", vbExclamation
                    Exit Function
                Else
                    blnSizeField = True
                End If
            Case 3
                If blnFNLenField Then
                    lvwFiles.Visible = True
                    MsgBox "Only one file name length field is allowed.", vbExclamation
                    Exit Function
                Else
                    blnFNLenField = True
                End If
            Case 4
                intNumDataFields = intNumDataFields + 1
                Set itmX = lvwFiles.ColumnHeaders.Add(, , "Data " & intNumDataFields, 675, lvwColumnRight)
            Case 5
                intNumDataFields = intNumDataFields + 1
                Set itmX = lvwFiles.ColumnHeaders.Add(, , "Data " & intNumDataFields, 675, lvwColumnLeft)
        End Select
    Next i
    
    If Not blnAddressField And Not blnSizeField Then
        MsgBox "An address field and/or a size field is required", vbExclamation
        Exit Function
    End If
    intLastSort = 0
    lvwFiles.Visible = True
    CheckRecordStruct = True
End Function

Private Sub FillSize()
    Dim i As Integer
    Dim lngDataFileLen As Long
    
    On Error Resume Next
    lngDataFileLen = FileLen(txtDataFile.Text)
    If Err.Number Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot find data file '" & txtDataFile.Text & "'", vbExclamation
        Exit Sub
    End If
    
    For i = 1 To intNumFiles - 1
        udtFileArray(i - 1).Size = udtFileArray(i).Address - udtFileArray(i - 1).Address
        
        If udtFileArray(i - 1).Size < 0 Or udtFileArray(i - 1).Size >= lngDataFileLen Then
            Screen.MousePointer = vbDefault
            MsgBox "An invalid size was calculated from the index. Please check that the index structure is correct.", vbCritical
            Err.Raise 513
            Exit Sub
        End If
        
        lvwFiles.ListItems(i).SubItems(2) = Format(udtFileArray(i - 1).Size, "#,0") & " bytes"
    Next i
    udtFileArray(i - 1).Size = lngDataFileLen - udtFileArray(i - 1).Address
    
    If udtFileArray(i - 1).Size < 0 Or udtFileArray(i - 1).Size >= lngDataFileLen Then
        Screen.MousePointer = vbDefault
        MsgBox "An invalid size was calculated from the index. Please check that the index structure is correct.", vbCritical
        Err.Raise 513
        Exit Sub
    End If
    
    lvwFiles.ListItems(i).SubItems(2) = Format(udtFileArray(i - 1).Size, "#,0") & " bytes"
End Sub

Private Sub FillAddress()
    Dim i As Integer
    Dim lngDataFileLen As Long
    
    On Error Resume Next
    lngDataFileLen = FileLen(txtDataFile.Text)
    If Err.Number Then
        Screen.MousePointer = vbDefault
        MsgBox "Cannot find data file '" & txtDataFile.Text & "'", vbExclamation
        Exit Sub
    End If
    udtFileArray(0).Address = Loc(1) + txtFooter.Text
    
    If udtFileArray(0).Address < 0 Or udtFileArray(0).Address >= lngDataFileLen Then
        Screen.MousePointer = vbDefault
        MsgBox "An invalid address was calculated from the index. Please check that the index structure is correct.", vbCritical
        Err.Raise 513
        Exit Sub
    End If
    
    For i = 1 To intNumFiles - 1
        udtFileArray(i).Address = udtFileArray(i - 1).Address + udtFileArray(i - 1).Size
    
        If udtFileArray(i).Address < 0 Or udtFileArray(i).Address >= lngDataFileLen Then
            Screen.MousePointer = vbDefault
            MsgBox "An invalid address was calculated from the index. Please check that the index structure is correct.", vbCritical
            Err.Raise 513
            Exit Sub
        End If
    Next i
End Sub

Private Function GetString() As String
    Dim bytChar As Byte
    
    Get #1, , bytChar
    
    Do While bytChar > 0 And Not EOF(1)
        GetString = GetString + Chr(bytChar)
        Get #1, , bytChar
    Loop
End Function

Private Sub txtSize_GotFocus()
    txtSize.SelStart = 0
    txtSize.SelLength = Len(txtSize.Text)
End Sub

Private Sub ClearContents()
    Dim i As Integer
    
    lvwFiles.Visible = False
    lvwFiles.ListItems.Clear
    cmdOpen.Enabled = False
    cmdOpenWith.Enabled = False
    cmdExport.Enabled = False
    For i = lvwFiles.ColumnHeaders.Count To 5 Step -1
        lvwFiles.ColumnHeaders.Remove i
    Next i
    lvwFiles.Visible = True
End Sub

Private Sub ExportFile(intIndex As Integer, strFileName As String, blnReplacePrompt As Boolean)
    Dim strData As String
    Dim intResponse As Integer

    On Error Resume Next
    strData = String(udtFileArray(intIndex).Size, 0)
    
    Open txtDataFile.Text For Binary Access Read As 1
    If Err.Number Then
        MsgBox "Cannot open data file '" & txtDataFile.Text & "'", vbExclamation
        Exit Sub
    End If
    Get #1, udtFileArray(intIndex).Address + 1, strData
    Close #1
    
    If blnReplacePrompt And Dir(strFileName) <> "" Then
        intResponse = MsgBox("The file '" & strFileName & "' already exists. Do you want to replace it?", vbQuestion + vbYesNoCancel)
        If intResponse = vbNo Then Exit Sub
        If intResponse = vbCancel Then
            Err.Raise 513
            Exit Sub
        End If
    End If
    
    Open strFileName For Binary Access Write As #1
    If Err.Number Then
        MsgBox "Cannot create file '" & strFileName & "'", vbCritical
        Exit Sub
    End If
    Put #1, , strData
    If Err.Number Then
        MsgBox "Cannot create file '" & strFileName & "'", vbCritical
        Close #1
        Exit Sub
    End If
    Close #1
End Sub
