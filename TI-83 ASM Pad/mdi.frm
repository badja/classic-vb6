VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "TI-83 ASM Pad"
   ClientHeight    =   3495
   ClientLeft      =   915
   ClientTop       =   2205
   ClientWidth     =   5520
   Icon            =   "mdi.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList imlToolbar 
      Left            =   0
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0556
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":066A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":077E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0E0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":0F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":1036
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdi.frx":114A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbToolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Assemble"
            Object.ToolTipText     =   "Assemble"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Emulate"
            Object.ToolTipText     =   "Launch VTI"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Send"
            Object.ToolTipText     =   "Send to VTI"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Find"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Font"
            Object.ToolTipText     =   "Font"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgAssemble 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "83p"
      DialogTitle     =   "Assemble As"
      Filter          =   "TI-83 Programs (*.83p)|*.83p|All Files (*.*)|*.*"
   End
   Begin MSComDlg.CommonDialog dlgFont 
      Left            =   960
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   0
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "TXT"
      Filter          =   "Assembly Code (*.z80;*.asm)|*.z80;*.asm|Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
      FilterIndex     =   557
      FontSize        =   1.27584e-37
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
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile1"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile2"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile3"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile4"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   "RecentFile5"
         Index           =   5
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuOSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFixed 
         Caption         =   "&Show Fixed-Pitch Fonts Only"
      End
      Begin VB.Menu mnuDefaultFont 
         Caption         =   "&Default Font..."
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Main MDI form for MDI Notepad sample       ***
'*** application.                               ***
'**************************************************
Option Explicit

Private intMainLeft As Integer           ' Holds the main window's left pos
Private intMainTop As Integer            ' Holds the main window's top pos
Private intMainWidth As Integer          ' Holds the main window's width
Private intMainHeight As Integer         ' Holds the main window's height

Private Sub MDIForm_Load()
    Dim lngLength As Long
    
    intMainLeft = GetSetting(ThisApp, SetKey, "MainLeft", 1000)
    intMainTop = GetSetting(ThisApp, SetKey, "MainTop", 1000)
    intMainWidth = GetSetting(ThisApp, SetKey, "MainWidth", 6500)
    intMainHeight = GetSetting(ThisApp, SetKey, "MainHeight", 6500)
    Me.Move intMainLeft, intMainTop, intMainWidth, intMainHeight
    Me.WindowState = -vbMaximized * GetSetting(ThisApp, SetKey, "MainMax", 0)
    blnChildMax = GetSetting(ThisApp, SetKey, "ChildMax", False)
    tlbToolbar.Visible = GetSetting(ThisApp, SetKey, "Toolbar", True)
    blnFixed = GetSetting(ThisApp, SetKey, "FixedPitch", True)
    
    strTempPath = Space(256)
    lngLength = GetTempPath(256, strTempPath)
    strTempPath = Left(strTempPath, lngLength)
    If Right(strTempPath, 1) <> "\" Then strTempPath = strTempPath & "\"
    CMDialog1.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    dlgAssemble.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    dlgFont.Flags = cdlCFForceFontExist Or cdlCFScreenFonts
    
    ' Application starts here (Load event of Startup form).
    Show
    ' Always set the working directory to the directory containing the application.
    'ChDir App.Path
    ' Initialize the document form array, and show the first document.
    ReDim Document(1)
    ReDim FState(1)
    Document(1).Tag = 1
    FState(1).Dirty = False
    
    If Command <> "" Then
        ' Call the file open procedure, passing a
        ' reference to the selected file name
        OpenFile (Command)
        ' Show the toolbar if they aren't already visible.
        If gToolsHidden Then
            frmMDI.tlbToolbar.Buttons("Save").Enabled = True
            frmMDI.tlbToolbar.Buttons("Assemble").Enabled = True
            frmMDI.tlbToolbar.Buttons("Emulate").Enabled = True
            frmMDI.tlbToolbar.Buttons("Send").Enabled = True
            frmMDI.tlbToolbar.Buttons("Cut").Enabled = True
            frmMDI.tlbToolbar.Buttons("Copy").Enabled = True
            frmMDI.tlbToolbar.Buttons("Paste").Enabled = True
            frmMDI.tlbToolbar.Buttons("Find").Enabled = True
            frmMDI.tlbToolbar.Buttons("Font").Enabled = True
            gToolsHidden = False
        End If
    End If
    
    ' Read System registry and set the recent menu file list control array appropriately.
    GetRecentFiles
    ' Set public variable gFindDirection which determines which direction
    ' the FindIt function will search in.
    gFindDirection = 1
End Sub

Private Sub MDIForm_Resize()
    If Me.WindowState = vbNormal Then
        intMainLeft = Me.Left
        intMainTop = Me.Top
        intMainWidth = Me.Width
        intMainHeight = Me.Height
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    ' If the Unload event was not cancelled (in the QueryUnload events for the Notepad forms),
    ' there will be no document window left, so go ahead and end the application.
    If Not AnyPadsLeft() Then
        
        If Me.WindowState = vbNormal Then
            intMainLeft = Me.Left
            intMainTop = Me.Top
            intMainWidth = Me.Width
            intMainHeight = Me.Height
        End If
        
        SaveSetting ThisApp, SetKey, "MainLeft", intMainLeft
        SaveSetting ThisApp, SetKey, "MainTop", intMainTop
        SaveSetting ThisApp, SetKey, "MainWidth", intMainWidth
        SaveSetting ThisApp, SetKey, "MainHeight", intMainHeight
        SaveSetting ThisApp, SetKey, "MainMax", CInt(Me.WindowState = vbMaximized)
        SaveSetting ThisApp, SetKey, "ChildMax", CInt(blnChildMax)
        SaveSetting ThisApp, SetKey, "Toolbar", CInt(tlbToolbar.Visible)
        SaveSetting ThisApp, SetKey, "FixedPitch", CInt(blnFixed)
        
        End
    End If
End Sub





Private Sub mnuDefaultFont_Click()
    DefaultFont
End Sub

Private Sub mnuFileExit_Click()
    ' End the application.
    End
End Sub

Private Sub mnuFileNew_Click()
    ' Call the new file procedure
    FileNew
End Sub

Private Sub mnuFileOpen_Click()
    ' Call the file open procedure.
    FileOpenProc
End Sub

Private Sub mnuOptions_Click()
    ' Toggle the visibility of the toolbar.
    mnuOptionsToolbar.Checked = frmMDI.tlbToolbar.Visible
    mnuOptionsFixed.Checked = blnFixed
End Sub


Private Sub mnuOptionsFixed_Click()
    ' Call the fixed-pitch procedure, passing a reference
    ' to this form.
    OptionsFixedProc Me
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Call the toolbar procedure, passing a reference
    ' to this form.
    OptionsToolbarProc Me
End Sub


Private Sub mnuRecentFile_Click(index As Integer)
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile (mnuRecentFile(index).Caption)
    ' Update the list of the most recently opened files.
    GetRecentFiles
    ' Show the toolbar if they aren't already visible.
    If gToolsHidden Then
        frmMDI.tlbToolbar.Buttons("Save").Enabled = True
        frmMDI.tlbToolbar.Buttons("Assemble").Enabled = True
        frmMDI.tlbToolbar.Buttons("Emulate").Enabled = True
        frmMDI.tlbToolbar.Buttons("Send").Enabled = True
        frmMDI.tlbToolbar.Buttons("Cut").Enabled = True
        frmMDI.tlbToolbar.Buttons("Copy").Enabled = True
        frmMDI.tlbToolbar.Buttons("Paste").Enabled = True
        frmMDI.tlbToolbar.Buttons("Find").Enabled = True
        frmMDI.tlbToolbar.Buttons("Font").Enabled = True
        gToolsHidden = False
    End If
End Sub

Private Sub tlbToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.key
        Case "New"
            FileNew
        Case "Open"
            FileOpenProc
        Case "Save"
            FileSaveProc frmMDI.ActiveForm
        Case "Assemble"
            FileAssembleProc frmMDI.ActiveForm
        Case "Emulate"
            FileEmulateProc frmMDI.ActiveForm
        Case "Send"
            FileSendProc frmMDI.ActiveForm
        Case "Cut"
            EditCutProc
        Case "Copy"
            EditCopyProc
        Case "Paste"
            EditPasteProc
        Case "Find"
            SearchFindProc frmMDI.ActiveForm
        Case "Font"
            FontProc frmMDI.ActiveForm
    End Select
End Sub
