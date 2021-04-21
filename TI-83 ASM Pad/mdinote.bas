Attribute VB_Name = "Module1"
'*** Global module for MDI Notepad sample application.  ***
'**********************************************************
Option Explicit

' User-defined type to store information about child forms
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
    ProgName(1) As String
End Type

Public FState()  As FormState           ' Array of user-defined types
Public Document() As New frmNotePad     ' Array of child form objects
Public gFindString As String            ' Holds the search text.
Public gFindCase As Integer             ' Key for case sensitive search
Public gFindDirection As Integer        ' Key for search direction.
Public gCurPos As Long                  ' Holds the cursor location.
Public gFirstTime As Long               ' Key for start position.
Public gToolsHidden As Boolean          ' Holds toolbar state.
Public Const ThisApp = "TI-83 ASM Pad"  ' Registry App constant.
Public Const ThisKey = "Recent Files"   ' Registry Key constant.
Public Const SetKey = "Settings"        ' Registry Key constant.
Public blnFixed As Boolean              ' Show only fixed-pitch fonts?
Public strTempPath As String            ' Holds the temporary directory
Public blnChildMax As Boolean           ' Are child windows maximized?
Public blnFirstDocDirty As Boolean      ' Has the first document been changed?
Public blnKeepEdit As Boolean           ' Set before default window is auto-closed
Public dblTaskID As Double

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long


Function AnyPadsLeft() As Integer
    Dim I As Integer        ' Counter variable

    ' Cycle through the document array.
    ' Return true if there is at least one open document.
    For I = 1 To UBound(Document)
        If Not FState(I).Deleted Then
            AnyPadsLeft = True
            Exit Function
        End If
    Next
End Function


Sub EditCopyProc()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
End Sub

Sub EditCutProc()
    ' Copy the selected text onto the Clipboard.
    Clipboard.SetText frmMDI.ActiveForm.ActiveControl.SelText
    ' Delete the selected text.
    frmMDI.ActiveForm.ActiveControl.SelText = ""
End Sub

Sub EditPasteProc()
    ' Place the text from the Clipboard into the active control.
    frmMDI.ActiveForm.ActiveControl.SelText = Clipboard.GetText()
End Sub

Sub FileNew()
    Dim fIndex As Integer

    ' Find the next available index and show the child form.
    fIndex = FindFreeIndex()
    Document(fIndex).Tag = fIndex
    Document(fIndex).Caption = "Untitled:" & fIndex
    FState(fIndex).ProgName(0) = ""
    FState(fIndex).ProgName(1) = ""
    blnFirstDocDirty = True
    Document(fIndex).Show

    ' Make sure the toolbar edit buttons are visible.
    frmMDI.tlbToolbar.Buttons("Save").Enabled = True
    frmMDI.tlbToolbar.Buttons("Assemble").Enabled = True
    frmMDI.tlbToolbar.Buttons("Emulate").Enabled = True
    frmMDI.tlbToolbar.Buttons("Send").Enabled = True
    frmMDI.tlbToolbar.Buttons("Cut").Enabled = True
    frmMDI.tlbToolbar.Buttons("Copy").Enabled = True
    frmMDI.tlbToolbar.Buttons("Paste").Enabled = True
    frmMDI.tlbToolbar.Buttons("Find").Enabled = True
    frmMDI.tlbToolbar.Buttons("Font").Enabled = True
    blnKeepEdit = True
End Sub

Function FindFreeIndex() As Integer
    Dim I As Integer
    Dim ArrayCount As Integer

    ArrayCount = UBound(Document)

    ' Cycle through the document array. If one of the
    ' documents has been deleted, then return that index.
    For I = 1 To ArrayCount
        If FState(I).Deleted Then
            FindFreeIndex = I
            FState(I).Deleted = False
            Exit Function
        End If
    Next

    ' If none of the elements in the document array have
    ' been deleted, then increment the document and the
    ' state arrays by one and return the index to the
    ' new element.
    ReDim Preserve Document(ArrayCount + 1)
    ReDim Preserve FState(ArrayCount + 1)
    FindFreeIndex = UBound(Document)
End Function

Sub FindIt()
    Dim lngStart As Long
    Dim lngPos As Long
    Dim strFindString As String
    Dim strSourceString As String
    Dim strMsg As String
    Dim intResponse As Integer
    Dim intOffset As Integer
    
    ' Set offset variable based on cursor position.
    If (gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart) Then
        intOffset = 1
    Else
        intOffset = 0
    End If

    ' Read the public variable for start position.
    If gFirstTime Then intOffset = 0
    ' Assign a value to the start value.
    lngStart = frmMDI.ActiveForm.ActiveControl.SelStart + intOffset
        
    ' If not case sensitive, convert the string to upper case
    If gFindCase Then
        strFindString = gFindString
        strSourceString = frmMDI.ActiveForm.ActiveControl.Text
    Else
        strFindString = UCase(gFindString)
        strSourceString = UCase(frmMDI.ActiveForm.ActiveControl.Text)
    End If
            
    ' Search for the string.
    If gFindDirection = 1 Then
        lngPos = InStr(lngStart + 1, strSourceString, strFindString)
    Else
        For lngPos = lngStart - 1 To 0 Step -1
            If lngPos = 0 Then Exit For
            If Mid(strSourceString, lngPos, Len(strFindString)) = strFindString Then Exit For
        Next
    End If

    ' If the string is found...
    If lngPos Then
        frmMDI.ActiveForm.ActiveControl.SelStart = lngPos - 1
        frmMDI.ActiveForm.ActiveControl.SelLength = Len(strFindString)
    Else
        strMsg = "Cannot find " & Chr(34) & gFindString & Chr(34)
        intResponse = MsgBox(strMsg, vbInformation, App.Title)
    End If
    
    ' Reset the public variables
    gCurPos = frmMDI.ActiveForm.ActiveControl.SelStart
    gFirstTime = False
End Sub

Sub GetRecentFiles()
    ' This procedure demonstrates the use of the GetAllSettings function,
    ' which returns an array of values from the Windows registry. In this
    ' case, the registry contains the files most recently opened.  Use the
    ' SaveSetting statement to write the names of the most recent files.
    ' That statement is used in the WriteRecentFiles procedure.
    Dim I, j As Integer
    Dim varFiles As Variant ' Varible to store the returned array.
    
    ' Get recent files from the registry using the GetAllSettings statement.
    ' ThisApp and ThisKey are constants defined in this module.
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
    For I = 0 To UBound(varFiles, 1)
        
        frmMDI.mnuRecentFile(0).Visible = True
        frmMDI.mnuRecentFile(I).Caption = varFiles(I, 1)
        frmMDI.mnuRecentFile(I).Visible = True
            ' Iterate through all the documents and update each menu.
            For j = 1 To UBound(Document)
                If Not FState(j).Deleted Then
                    Document(j).mnuRecentFile(0).Visible = True
                    Document(j).mnuRecentFile(I + 1).Caption = varFiles(I, 1)
                    Document(j).mnuRecentFile(I + 1).Visible = True
                End If
            Next j
    Next I

End Sub

Sub OptionsToolbarProc(CurrentForm As Form)
    ' Toggle the check
    CurrentForm.mnuOptionsToolbar.Checked = Not CurrentForm.mnuOptionsToolbar.Checked
    ' If not the MDI form, set the MDI form's check.
    If Not TypeOf CurrentForm Is MDIForm Then
        frmMDI.mnuOptionsToolbar.Checked = CurrentForm.mnuOptionsToolbar.Checked
    End If
    ' Toggle the toolbar based on the value.
    If CurrentForm.mnuOptionsToolbar.Checked Then
        frmMDI.tlbToolbar.Visible = True
    Else
        frmMDI.tlbToolbar.Visible = False
    End If
End Sub

Sub OptionsFixedProc(CurrentForm As Form)
    ' Toggle the check
    CurrentForm.mnuOptionsFixed.Checked = Not CurrentForm.mnuOptionsFixed.Checked
    ' If not the MDI form, set the MDI form's check.
    If Not TypeOf CurrentForm Is MDIForm Then
        frmMDI.mnuOptionsFixed.Checked = CurrentForm.mnuOptionsFixed.Checked
    End If
    ' Toggle the flag based on the value.
    blnFixed = CurrentForm.mnuOptionsFixed.Checked
End Sub

Sub WriteRecentFiles(OpenFileName)
    ' This procedure uses the SaveSettings statement to write the names of
    ' recently opened files to the System registry. The SaveSetting
    ' statement requires three parameters. Two of the parameters are
    ' stored as constants and are defined in this module.  The GetAllSettings
    ' function is used in the GetRecentFiles procedure to retrieve the
    ' file names stored in this procedure.
    
    Dim I, j As Integer
    Dim strFile, key As String

    ' Copy RecentFile1 to RecentFile2, and so on.
    For I = 3 To 1 Step -1
        key = "RecentFile" & I
        strFile = GetSetting(ThisApp, ThisKey, key)
        If strFile <> "" Then
            key = "RecentFile" & (I + 1)
            SaveSetting ThisApp, ThisKey, key, strFile
        End If
    Next I
  
    ' Write the open file to first recent file.
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

