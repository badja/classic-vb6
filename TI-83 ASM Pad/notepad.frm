VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmNotePad 
   Caption         =   "Untitled"
   ClientHeight    =   3990
   ClientLeft      =   1515
   ClientTop       =   3315
   ClientWidth     =   5670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "notepad.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3990
   ScaleMode       =   0  'User
   ScaleWidth      =   101.07
   Begin MSComctlLib.StatusBar staStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3735
      Width           =   5670
      _ExtentX        =   10001
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   6588
      _Version        =   393217
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"notepad.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
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
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "De&lete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuESep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditTime 
         Caption         =   "Time/&Date"
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu mnuProgram 
      Caption         =   "&Program"
      Begin VB.Menu mnuProgramAssemble 
         Caption         =   "Asse&mble"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuProgramAssembleAs 
         Caption         =   "Assem&ble As..."
      End
      Begin VB.Menu mnuProgramSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramEmulate 
         Caption         =   "&Launch VTI Emulator"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuProgramSend 
         Caption         =   "S&end to VTI Emulator"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuProgramSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProgramMode 
         Caption         =   "TI-83 &Mode"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuProgramMode 
         Caption         =   "TI-83 &Plus Mode"
         Index           =   1
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuLocations 
         Caption         =   "File &Locations..."
      End
      Begin VB.Menu mnuOSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFixed 
         Caption         =   "&Show Fixed-Pitch Fonts Only"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuDefaultFont 
         Caption         =   "&Default Font..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTile 
         Caption         =   "&Tile Horizontally"
      End
      Begin VB.Menu mnuWindowTileV 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowArrange 
         Caption         =   "&Arrange Icons"
      End
   End
End
Attribute VB_Name = "frmNotePad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*** Child form for the MDI Notepad sample application  ***
'**********************************************************
Option Explicit

Private Sub Form_Load()
    Me.WindowState = -vbMaximized * blnChildMax
    Text1.Font.Bold = GetSetting(ThisApp, SetKey, "FontBold", False)
    Text1.Font.Italic = GetSetting(ThisApp, SetKey, "FontItalic", False)
    Text1.Font.Name = GetSetting(ThisApp, SetKey, "FontName", "Fixedsys")
    Text1.Font.Size = GetSetting(ThisApp, SetKey, "FontSize", 9)
    staStatus.SimpleText = "Line 2"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim strMsg As String
    Dim strFilename As String
    Dim intResponse As Integer

    ' Check to see if the text has been changed.
    If FState(Me.Tag).Dirty Then
        strFilename = Me.Caption
        strMsg = "The text in [" & strFilename & "] has changed."
        strMsg = strMsg & vbCrLf
        strMsg = strMsg & "Do you want to save the changes?"
        intResponse = MsgBox(strMsg, 51, frmMDI.Caption)
        Select Case intResponse
            Case 6      ' User chose Yes.
                If Left(Me.Caption, 8) = "Untitled" Then
                    ' The file hasn't been saved yet.
                    strFilename = "untitled.txt"
                    ' Get the strFilename, and then call the save procedure, GetstrFilename.
                    strFilename = GetFileName(strFilename)
                Else
                    ' The form's Caption contains the name of the open file.
                    strFilename = Me.Caption
                End If
                ' Call the save procedure. If strFilename = Empty, then
                ' the user chose Cancel in the Save As dialog box; otherwise,
                ' save the file.
                If strFilename <> "" Then
                    SaveFileAs strFilename, False, ""
                End If
            Case 7      ' User chose No. Unload the file.
                Cancel = False
            Case 2      ' User chose Cancel. Cancel the unload.
                Cancel = True
        End Select
    End If
End Sub

Private Sub Form_Resize()
    Dim intTextHeight As Integer
    
    ' Expand text box to fill the current child form's internal area.
    intTextHeight = ScaleHeight - staStatus.Height
    If intTextHeight < 0 Then intTextHeight = 0
    Text1.Height = intTextHeight
    Text1.Width = ScaleWidth
    If Me.WindowState = vbMaximized Then blnChildMax = True Else blnChildMax = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Show the current form instance as deleted
    FState(Me.Tag).Deleted = True
    
    ' Hide the toolbar edit buttons if no notepad windows exist.
    If Not AnyPadsLeft() And blnFirstDocDirty Or Not blnKeepEdit Then
        frmMDI.tlbToolbar.Buttons("Save").Enabled = False
        frmMDI.tlbToolbar.Buttons("Assemble").Enabled = False
        frmMDI.tlbToolbar.Buttons("Emulate").Enabled = False
        frmMDI.tlbToolbar.Buttons("Send").Enabled = False
        frmMDI.tlbToolbar.Buttons("Cut").Enabled = False
        frmMDI.tlbToolbar.Buttons("Copy").Enabled = False
        frmMDI.tlbToolbar.Buttons("Paste").Enabled = False
        frmMDI.tlbToolbar.Buttons("Find").Enabled = False
        frmMDI.tlbToolbar.Buttons("Font").Enabled = False
        ' Toggle the public tool state variable
        gToolsHidden = True
        ' Call the recent file list procedure
        GetRecentFiles
    End If
    
    blnKeepEdit = True
End Sub

Private Sub mnuDefaultFont_Click()
    DefaultFont
End Sub

Private Sub mnuEditCopy_Click()
    ' Call the copy procedure
    EditCopyProc
End Sub

Private Sub mnuEditCut_Click()
    ' Call the cut procedure
    EditCutProc
End Sub

Private Sub mnuEditDelete_Click()
    ' If the mouse pointer is not at the end of the notepad...
    If Screen.ActiveControl.SelStart <> Len(Screen.ActiveControl.Text) Then
        ' If nothing is selected, extend the selection by one.
        If Screen.ActiveControl.SelLength = 0 Then
            Screen.ActiveControl.SelLength = 1
            ' If the mouse pointer is on a blank line, extend the selection by two.
            If Asc(Screen.ActiveControl.SelText) = 13 Then
                Screen.ActiveControl.SelLength = 2
            End If
        End If
        ' Delete the selected text.
        Screen.ActiveControl.SelText = ""
    End If
End Sub

Private Sub mnuEditPaste_Click()
    ' Call the paste procedure.
    EditPasteProc
End Sub

Private Sub mnuEditSelectAll_Click()
    ' Use SelStart & SelLength to select the text.
    frmMDI.ActiveForm.Text1.SelStart = 0
    frmMDI.ActiveForm.Text1.SelLength = Len(frmMDI.ActiveForm.Text1.Text)
End Sub

Private Sub mnuEditTime_Click()
    ' Insert the current time and date.
    Text1.SelText = Now
End Sub

Private Sub mnuFileClose_Click()
    ' Unload this form.
    Unload Me
End Sub

Private Sub mnuFileExit_Click()
    ' Unloading the MDI form invokes the QueryUnload event
    ' for each child form, and then the MDI form.
    ' Setting the Cancel argument to True in any of the
    ' QueryUnload events cancels the unload.
    Unload frmMDI
End Sub

Private Sub mnuFileNew_Click()
    ' Call the new form procedure
    FileNew
End Sub

Private Sub mnuFileOpen_Click()
    ' Call the file open procedure.
    FileOpenProc
End Sub

Private Sub mnuFileSave_Click()
    FileSaveProc Me
End Sub

Private Sub mnuFileSaveAs_Click()
    Dim strSaveFileName As String
    Dim strDefaultName As String
    
    ' Assign the form caption to the variable.
    strDefaultName = Me.Caption
    If Left(Me.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, strSaveFileName.
        
        strSaveFileName = GetFileName("Untitled.txt")
        If strSaveFileName <> "" Then
            FState(Me.Tag).ProgName(0) = ""
            FState(Me.Tag).ProgName(1) = ""
            SaveFileAs (strSaveFileName), False, ""
        End If
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    Else
        ' The form's Caption contains the name of the open file.
        
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then
            FState(Me.Tag).ProgName(0) = ""
            FState(Me.Tag).ProgName(1) = ""
            SaveFileAs (strSaveFileName), False, ""
        End If
        ' Update the list of recently opened files in the File menu control array.
        UpdateFileMenu (strSaveFileName)
    End If

    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Private Sub mnuFont_Click()
    FontProc Me
End Sub

Private Sub mnuLocations_Click()
    frmFileLocations.Show 1
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Private Sub mnuOptions_Click()
    ' Toggle the Checked property to match the .Visible property.
    mnuOptionsToolbar.Checked = frmMDI.tlbToolbar.Visible
    mnuOptionsFixed.Checked = blnFixed
End Sub

Private Sub mnuOptionsFixed_Click()
    ' Call the toolbar procedure, passing a reference
    ' to this form instance.
    OptionsFixedProc Me
End Sub

Private Sub mnuOptionsToolbar_Click()
    ' Call the toolbar procedure, passing a reference
    ' to this form instance.
    OptionsToolbarProc Me
End Sub

Private Sub mnuProgramAssemble_Click()
    FileAssembleProc Me
End Sub

Private Sub mnuProgramAssembleAs_Click()
    FileAssembleAsProc Me
End Sub

Private Sub mnuProgramEmulate_Click()
    FileEmulateProc Me
End Sub

Private Sub mnuProgramMode_Click(index As Integer)
    mnuProgramMode(index).Checked = True
    mnuProgramMode(-Not -index).Checked = False
End Sub

Private Sub mnuProgramSend_Click()
    FileSendProc Me
End Sub

Private Sub mnuRecentFile_Click(index As Integer)
    ' Call the file open procedure, passing a
    ' reference to the selected file name
    OpenFile (mnuRecentFile(index).Caption)
    ' Update the list of recently opened files in the File menu control array.
    GetRecentFiles
End Sub

Private Sub mnuSearchFind_Click()
    SearchFindProc Me
End Sub

Private Sub mnuSearchFindNext_Click()
    ' If the public variable isn't empty, call the
    ' find procedure, otherwise call the find menu
    If Len(gFindString) > 0 Then
        FindIt
    Else
        mnuSearchFind_Click
    End If
End Sub

Private Sub mnuWindowArrange_Click()
    ' Arrange the icons for any minimzied child forms.
    frmMDI.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowCascade_Click()
    ' Cascade the child forms.
    frmMDI.Arrange vbCascade
End Sub

Private Sub mnuWindowTile_Click()
    ' Tile the child forms.
    frmMDI.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowTileV_Click()
    ' Tile the child forms.
    frmMDI.Arrange vbTileVertical
End Sub

Private Sub Text1_Change()
    ' Set the public variable to show that text has changed.
    FState(Me.Tag).Dirty = True
    blnFirstDocDirty = True
End Sub

Private Sub Text1_SelChange()
    staStatus.SimpleText = "Line " & Text1.GetLineFromChar(Text1.SelStart) + 2
End Sub
