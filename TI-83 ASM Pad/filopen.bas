Attribute VB_Name = "Module2"
'*** Standard module with procedures for working with   ***
'*** files. Part of the MDI Notepad sample application. ***
'**********************************************************
Option Explicit

Sub FileOpenProc()
    Dim intRetVal
    On Error Resume Next
    Dim strOpenFileName As String
    frmMDI.CMDialog1.Filename = ""
    frmMDI.CMDialog1.ShowOpen
    If Err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = frmMDI.CMDialog1.Filename
        
        OpenFile (strOpenFileName)
        UpdateFileMenu (strOpenFileName)
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
    
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Function GetFileName(Filename As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    frmMDI.CMDialog1.Filename = Filename
    frmMDI.CMDialog1.ShowSave
    If Err <> 32755 Then    ' User chose Cancel.
        GetFileName = frmMDI.CMDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function

Function OnRecentFilesList(Filename) As Integer
  Dim I         ' Counter variable.

  For I = 1 To 4
    If frmMDI.mnuRecentFile(I).Caption = Filename Then
      OnRecentFilesList = True
      Exit Function
    End If
  Next I
    OnRecentFilesList = False
End Function

Sub OpenFile(Filename)
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Open the selected file.
    Open Filename For Input As #1
    If Err Then
        MsgBox "Can't open file: " + Filename, vbExclamation
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
    ' Change the form's caption and display the new text.
    If Not blnFirstDocDirty Then
        blnKeepEdit = True
        Unload Document(1)
        blnFirstDocDirty = True
    End If
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = UCase(Filename)
    Document(fIndex).Text1.Text = Input(LOF(1), 1)
    FState(fIndex).Dirty = False
    FState(fIndex).ProgName(0) = ""
    FState(fIndex).ProgName(1) = ""
    Document(fIndex).Show
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(Filename, blnAssemble As Boolean, strHeader As String)
    On Error Resume Next
    Dim strContents As String

    ' Open the file.
    Open Filename For Output As #1
    ' Place the contents of the notepad into a variable.
    strContents = frmMDI.ActiveForm.Text1.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    'Write the optional header
    If strHeader <> "" Then Print #1, strHeader
    ' Write the variable contents to a saved file.
    Print #1, strContents;
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    ' Set the form's caption.
    If Err Then
        MsgBox Error, 48, App.Title
    ElseIf Not blnAssemble Then
        frmMDI.ActiveForm.Caption = UCase(Filename)
        ' Reset the dirty flag.
        FState(frmMDI.ActiveForm.Tag).Dirty = False
    End If
End Sub

Sub UpdateFileMenu(Filename)
        Dim intRetVal As Integer
        ' Check if the open filename is already in the File menu control array.
        intRetVal = OnRecentFilesList(Filename)
        If Not intRetVal Then
            ' Write open filename to the registry.
            WriteRecentFiles (Filename)
        End If
        ' Update the list of the most recently opened files in the File menu control array.
        GetRecentFiles
End Sub

Function GetTitle(strFilename As String, blnKeepDir As Boolean) As String
    Dim intI As Integer, intStart As Integer
    
    For intI = Len(strFilename) To 1 Step -1
        If Mid(strFilename, intI, 1) = "\" Then
            intStart = intI
            Exit For
        End If
    Next intI
    
    intStart = intStart + 1
    
    For intI = intStart To Len(strFilename)
        If Mid(strFilename, intI, 1) = "." Then Exit For
    Next intI
    
    If blnKeepDir Then
        GetTitle = Left(strFilename, intI - 1)
    Else
        GetTitle = Mid(strFilename, intStart, intI - intStart)
    End If
End Function

Sub FileSaveProc(CurrentForm As Form)
    Dim strFilename As String

    If Left(CurrentForm.Caption, 8) = "Untitled" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = CurrentForm.Caption
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename, False, ""
    End If
    
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Sub FileAssembleProc(CurrentForm As Form)
    Dim dblDummy1 As Double, dblDummy2 As Double
    Dim strCurDir As String
    Dim strProgName As String, strLine As String
    Dim strTASMPath As String, strDevpacPath As String
    
    strProgName = GetTitle(FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked), False)
    
    If strProgName = "" Then
        FileAssembleAsProc CurrentForm
        Exit Sub
    End If
    
    On Error GoTo Error1
    strCurDir = CurDir
    SaveFileAs strTempPath & strProgName & ".z80", True, Choose(-CurrentForm.mnuProgramMode(1).Checked + 1, "#define TI83", "#define TI83P")
    
    strTASMPath = GetSetting(ThisApp, SetKey, "TASMDir")
    strDevpacPath = GetSetting(ThisApp, SetKey, "Devpac83Dir")
    
    ChDir strTASMPath
    dblDummy1 = Shell("tasm -t80 -b -i " & strTempPath & strProgName & ".z80 " & strTempPath & strProgName & ".bin", vbNormalFocus)
    On Error Resume Next
    
    Do Until FileLen(strTempPath & strProgName & ".bin")
        DoEvents
    Loop
    
    Do Until FileLen(strTempPath & strProgName & ".lst")
        DoEvents
    Loop
    
    Open strTempPath & strProgName & ".lst" For Input As #1
    
    Do Until EOF(1)
        Line Input #1, strLine
    Loop
    
    Close #1
    
    If Right(strLine, 2) = " 0" Then
        On Error GoTo Error2
        ChDir strDevpacPath
        FileCopy strTempPath & strProgName & ".bin", strProgName & ".bin"
        dblDummy2 = Shell("devpac83.com " & strProgName, vbMinimizedNoFocus)
        On Error Resume Next
        Do Until FileLen(strProgName & ".83p")
            DoEvents
        Loop
        AppActivate dblDummy1
        SendKeys "%({F4})"
        AppActivate dblDummy2
        SendKeys "%({F4})"
        If UCase(strDevpacPath & strProgName & ".83p") <> UCase(frmMDI.dlgAssemble.Filename) Then
            FileCopy strProgName & ".83p", FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked)
            Kill strDevpacPath & strProgName & ".83p"
        End If
    End If
    
    On Error Resume Next
    Kill strTempPath & strProgName & ".z80"
    Kill strTempPath & strProgName & ".bin"
    Kill strTempPath & strProgName & ".lst"
    Kill strDevpacPath & strProgName & ".bin"
    GoTo Quit

Error1:
    Select Case Err.Number
        Case 76
            MsgBox "TASM path not found. Please check File Locations under the Options menu.", vbExclamation
        Case 53
            MsgBox "TASM program file not found. Please check File Locations under the Options menu.", vbExclamation
    End Select
    
    GoTo Quit
    
Error2:
    MsgBox "Devpac83 program file not found.  Please check File Locations under the Options menu.", vbExclamation
   
Quit:
    ChDir strCurDir
    Exit Sub
End Sub

Sub FileAssembleAsProc(CurrentForm As Form)
    Dim strProgName As String
    
    On Error GoTo Cancel
    frmMDI.dlgAssemble.Filename = FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked)
    If FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked) = "" Then frmMDI.dlgAssemble.Filename = GetTitle(CurrentForm.Caption, True)
    frmMDI.dlgAssemble.Filter = Choose(-CurrentForm.mnuProgramMode(1).Checked + 1, "TI-83 Programs (*.83p)|*.83p|All Files (*.*)|*.*", "TI-83 Plus Programs (*.8xp)|*.8xp|All Files (*.*)|*.*")
    frmMDI.dlgAssemble.ShowSave
    strProgName = GetTitle(frmMDI.dlgAssemble.FileTitle, False)
    
    If Len(strProgName) > 8 Then
        MsgBox "Program name must be 8 characters or less.", vbExclamation
        frmMDI.ActiveForm.Text1.SetFocus
    Else
        FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked) = frmMDI.dlgAssemble.Filename
        FileAssembleProc CurrentForm
    End If
    
    Exit Sub
    
Cancel:
    frmMDI.ActiveForm.Text1.SetFocus
    Exit Sub
End Sub

Sub FileEmulateProc(CurrentForm As Form)
    Dim strVTI As String, strAutoFile As String, strBaseFile As String
    Dim strSavedDir As String
    
    On Error GoTo ErrorHandler
    
    strVTI = GetSetting(ThisApp, SetKey, "VTIDir")
    strAutoFile = GetSetting(ThisApp, SetKey, "AutoSaveFile")
    strBaseFile = GetSetting(ThisApp, SetKey, "BaseSaveFile")
    FileCopy strBaseFile, strAutoFile
    strSavedDir = CurDir
    ChDir strVTI
    dblTaskID = Shell(strVTI & "VTI.exe", vbNormalFocus)
    ChDir strSavedDir
    Exit Sub
    
ErrorHandler:
    If Err.Number = 53 Then MsgBox "Virtual TI program file or automatic save state file not found. Please check File Locations under the Options menu.", vbExclamation
    frmMDI.ActiveForm.Text1.SetFocus
    Exit Sub
End Sub

Sub FileSendProc(CurrentForm As Form)
    Dim strResponse As String
    Dim intErrorCount As Integer
    
    On Error GoTo NotRunning
    
    If FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked) = "" Or Dir(FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked)) = "" Then
        strResponse = InputBox("If this program has already been assembled, enter the filename of the program file below. Otherwise, press Cancel and assemble the program first.", "Send to VTI", GetTitle(CurrentForm.Caption, True) & Choose(-CurrentForm.mnuProgramMode(1).Checked + 1, ".83P", ".8XP"))
        
        If strResponse = "" Then
            Exit Sub
        Else
            If Dir(strResponse) = "" Then
                MsgBox "The program file could not be found. Make sure you have assembled the program first.", vbExclamation
                Exit Sub
            Else
                FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked) = strResponse
            End If
        End If
    End If
        
    intErrorCount = 0
    AppActivate dblTaskID
    AppActivate "Virtual TI-83"
    SendKeys "{SCROLLLOCK}{SCROLLLOCK}{PGDN}{PGDN}{ESC}{ESC}{PGDN}{PGDN}{F10}" & Chr(34) & FState(CurrentForm.Tag).ProgName(-CurrentForm.mnuProgramMode(1).Checked) & Chr(34) & "{ENTER}", True
    Exit Sub

NotRunning:
    intErrorCount = intErrorCount + 1
    If intErrorCount = 1 Then Resume Next
    MsgBox "Please Launch Virtual TI first.", vbInformation
End Sub

Sub FontProc(CurrentForm As Form)
    On Error GoTo Cancel
    
    If blnFixed Then
        frmMDI.dlgFont.Flags = frmMDI.dlgFont.Flags Or cdlCFFixedPitchOnly
    Else
        frmMDI.dlgFont.Flags = frmMDI.dlgFont.Flags And Not cdlCFFixedPitchOnly
    End If
    
    With frmMDI.dlgFont
        .FontBold = CurrentForm.Text1.Font.Bold
        .FontItalic = CurrentForm.Text1.Font.Italic
        .FontName = CurrentForm.Text1.Font.Name
        .FontSize = CurrentForm.Text1.Font.Size
    End With
    
    frmMDI.dlgFont.ShowFont
    
    With CurrentForm.Text1.Font
        .Bold = frmMDI.dlgFont.FontBold
        .Italic = frmMDI.dlgFont.FontItalic
        .Name = frmMDI.dlgFont.FontName
        .Size = frmMDI.dlgFont.FontSize
    End With
    
Cancel:
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Sub DefaultFont()
    On Error GoTo Cancel
    
    If blnFixed Then
        frmMDI.dlgFont.Flags = frmMDI.dlgFont.Flags Or cdlCFFixedPitchOnly
    Else
        frmMDI.dlgFont.Flags = frmMDI.dlgFont.Flags And Not cdlCFFixedPitchOnly
    End If
    
    frmMDI.dlgFont.FontBold = GetSetting(ThisApp, SetKey, "FontBold", False)
    frmMDI.dlgFont.FontItalic = GetSetting(ThisApp, SetKey, "FontItalic", False)
    frmMDI.dlgFont.FontName = GetSetting(ThisApp, SetKey, "FontName", "Fixedsys")
    frmMDI.dlgFont.FontSize = GetSetting(ThisApp, SetKey, "FontSize", 9)
    frmMDI.dlgFont.ShowFont
    SaveSetting ThisApp, SetKey, "FontBold", CInt(frmMDI.dlgFont.FontBold)
    SaveSetting ThisApp, SetKey, "FontItalic", CInt(frmMDI.dlgFont.FontItalic)
    SaveSetting ThisApp, SetKey, "FontName", frmMDI.dlgFont.FontName
    SaveSetting ThisApp, SetKey, "FontSize", frmMDI.dlgFont.FontSize
    
Cancel:
    frmMDI.ActiveForm.Text1.SetFocus
End Sub

Sub SearchFindProc(CurrentForm As Form)
    ' If there is text in the textbox, assign it to
    ' the textbox on the Find form, otherwise assign
    ' the last findtext value.
    If CurrentForm.Text1.SelText <> "" Then
        frmFind.Text1.Text = CurrentForm.Text1.SelText
    Else
        frmFind.Text1.Text = gFindString
    End If
    
    frmFind.Text1.SelLength = Len(frmFind.Text1.Text)

    ' Set the public variable to start at the beginning.
    gFirstTime = True
    ' Set the case checkbox to match the public variable
    If (gFindCase) Then
        frmFind.chkCase = 1
    End If
    ' Display the Find form.
    frmFind.Show vbModal
    frmMDI.ActiveForm.Text1.SetFocus
End Sub
