VERSION 5.00
Begin VB.Form frmBackupWizard 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Backup Wizard"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFolder 
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   2520
      Width           =   3615
   End
   Begin VB.OptionButton optOption2 
      Height          =   255
      Left            =   2520
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.OptionButton optOption1 
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   2280
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.PictureBox picPicture 
      BackColor       =   &H00808000&
      Enabled         =   0   'False
      Height          =   4425
      Left            =   120
      Picture         =   "BackupWizard.frx":0000
      ScaleHeight     =   4365
      ScaleWidth      =   2190
      TabIndex        =   0
      Top             =   120
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4920
      TabIndex        =   10
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   8
      Top             =   5040
      Width           =   1215
   End
   Begin VB.DirListBox dirFolder 
      Height          =   1890
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.DriveListBox drvDrive 
      Height          =   315
      Left            =   2520
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Line linLine 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   6120
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Shape shpLine 
      BorderColor     =   &H00FFFFFF&
      Height          =   30
      Left            =   120
      Top             =   4800
      Width           =   6015
   End
   Begin VB.Label lblMessage 
      Height          =   1815
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblFolder 
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   1920
      Visible         =   0   'False
      Width           =   3615
   End
End
Attribute VB_Name = "frmBackupWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private blnMade As Boolean, blnReadOnly As Boolean
Private intBackups As Integer
Private intStep As Integer
Private strBackupList As String
Private strBackups() As String
Private strInitDrive As String

Private Sub cmdBack_Click()
    Select Case intStep
        Case 1
            Step0
        Case 2
            Step1
        Case 3
            Step0
        Case 4
            Step3
    End Select
End Sub

Private Sub cmdCancel_Click()
    If intWizardVersion = 1 Then
        If GetSetting("GTA Wave", "Options", "RunWizard", 1) = 1 Then MsgBox "You can complete this Wizard later by selecting GTA Backup Wizard from the File menu.", vbInformation
    Else
        If GetSetting("GTA Wave", "Options", "RunWizard2", 1) = 1 Then MsgBox "You can complete this Wizard later by selecting GTA2 Backup Wizard from the File menu.", vbInformation
    End If
    
    Unload Me
End Sub

Private Sub cmdNext_Click()
    Dim intI As Integer
    Dim strSource As String, strDest As String
    
    On Error GoTo CannotMake
    
    Select Case intStep
        Case 0
            If blnMade Then Step1 Else Step3
        
        Case 1
            If intWizardVersion = 1 Then
                If LCase(Right(lblFolder.Caption, 14)) = "\gtadata\audio" Then
                    If MsgBox("This folder appears to be the one which contains the sound files actually used by GTA. The backup folder must be different from the GTA audio folder. Are you sure this is the folder which contains your backups?" & strNewLine & strNewLine & "To have this Wizard backup the files for you, click No now, then click Back, and answer 'No' to the question.", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
                End If
            Else
                If LCase(Right(lblFolder.Caption, 11)) = "\data\audio" Then
                    If MsgBox("This folder appears to be the one which contains the sound files actually used by GTA2. The backup folder must be different from the GTA2 audio folder. Are you sure this is the folder which contains your backups?" & strNewLine & strNewLine & "To have this Wizard backup the files for you, click No now, then click Back, and answer 'No' to the question.", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Exit Sub
                End If
            End If
            
            FindFiles
            
            If intBackups = 0 Then
                MsgBox "No backups were found in this folder. Make sure you have backed up both the SDT and the RAW files, and that they have not been renamed." & strNewLine & strNewLine & "To have this Wizard backup the files for you, click OK now, then click Back, and answer 'No' to the question.", vbExclamation
            Else
                strBackupList = ""
                
                For intI = 0 To intBackups - 2
                    strBackupList = strBackupList & strBackups(intI) & ", "
                Next intI
                
                strBackupList = strBackupList & strBackups(intBackups - 1)
                Step2
            End If
        
        Case 2
            If blnReadOnly Then MakeReadOnly
            Step5
        
        Case 3
            FindFiles
            
            If intBackups = 0 Then
                If intWizardVersion = 1 Then
                    MsgBox "The GTA sound files were not found in this folder. Make sure you select the AUDIO folder inside the GTADATA folder.", vbExclamation
                Else
                    MsgBox "The GTA2 sound files were not found in this folder. Make sure you select the AUDIO folder inside the DATA folder.", vbExclamation
                End If
            Else
                strBackupList = ""
                
                For intI = 0 To intBackups - 2
                    strBackupList = strBackupList & strBackups(intI) & ", "
                Next intI
                
                strBackupList = strBackupList & strBackups(intBackups - 1)
                Step4
            End If
        
        Case 4
            strSource = lblFolder.Caption
            strDest = txtFolder.Text
            If Right(strSource, 1) <> "\" Then strSource = strSource & "\"
            If Right(strDest, 1) <> "\" Then strDest = strDest & "\"
            
            If strSource = strDest Then
                If intWizardVersion = 1 Then
                    MsgBox "The backup folder must be different to your GTA audio folder. If you accept the default and click Next, the folder will automatically be created for you.", vbExclamation
                Else
                    MsgBox "The backup folder must be different to your GTA2 audio folder. If you accept the default and click Next, the folder will automatically be created for you.", vbExclamation
                End If
                
                Step4
            Else
                MkDir strDest
                BackupFiles strSource, strDest
                Step5
            End If
            
        Case 5
            If blnMade Then
                If intWizardVersion = 1 Then
                    SaveSetting "GTA Wave", "Options", "BackupDir", lblFolder.Caption
                Else
                    SaveSetting "GTA Wave", "Options", "BackupDir2", lblFolder.Caption
                End If
            Else
                If intWizardVersion = 1 Then
                    SaveSetting "GTA Wave", "Options", "BackupDir", txtFolder.Text
                Else
                    SaveSetting "GTA Wave", "Options", "BackupDir2", txtFolder.Text
                End If
            End If
            
            Unload Me
    End Select

    Exit Sub
    
CannotMake:
    If Err = 76 Then
        MsgBox "Cannot create more than one folder at once. Please create this folder yourself or accept the default.", vbExclamation
        Step4
    Else
        Resume Next
    End If
End Sub

Private Sub dirFolder_Change()
    lblFolder.Caption = dirFolder.Path
End Sub

Private Sub drvDrive_Change()
    On Error GoTo Unavailable
    dirFolder.Path = drvDrive.Drive
    Exit Sub
    
Unavailable:
    MsgBox "Drive is unavailable", vbExclamation
    drvDrive.Drive = strInitDrive
    Resume Next
End Sub

Private Sub Form_Load()
    If intWizardVersion = 1 Then
        frmBackupWizard.Caption = "GTA Backup Wizard"
    Else
        frmBackupWizard.Caption = "GTA2 Backup Wizard"
    End If
    blnMade = True
    blnReadOnly = True
    strInitDrive = drvDrive.Drive
    Step0
End Sub

Private Sub Step0()
    intStep = 0
    If intWizardVersion = 1 Then
        lblMessage.Caption = "This Wizard will enable you to use the Restore commands in GTA Wave. You will be able to restore the original GTA sounds individually or all at once." & strNewLine & strNewLine & "Have you already made backups of the original GTA sound files?"
    Else
        lblMessage.Caption = "This Wizard will enable you to use the Restore commands in GTA Wave. You will be able to restore the original GTA2 sounds individually or all at once." & strNewLine & strNewLine & "Have you already made backups of the original GTA2 sound files?"
    End If
    optOption1.Caption = "Yes, I have"
    optOption2.Caption = "No, I haven't"
    If blnMade Then optOption1.Value = True Else optOption2.Value = True
    optOption1.Visible = True
    optOption2.Visible = True
    lblFolder.Visible = False
    dirFolder.Visible = False
    drvDrive.Visible = False
    txtFolder.Visible = False
    cmdBack.Enabled = False
    cmdNext.Enabled = True
End Sub

Private Sub Step1()
    intStep = 1
    lblMessage.Caption = "Select the folder in which these backups are stored."
    If intWizardVersion = 1 Then
        dirFolder.Path = GetSetting("GTA Wave", "Options", "BackupDir")
    Else
        dirFolder.Path = GetSetting("GTA Wave", "Options", "BackupDir2")
    End If
    lblFolder.Caption = dirFolder.Path
    optOption1.Visible = False
    optOption2.Visible = False
    lblFolder.Visible = True
    dirFolder.Visible = True
    drvDrive.Visible = True
    txtFolder.Visible = False
    cmdBack.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub Step2()
    intStep = 2
    lblMessage.Caption = "SDT and RAW files were found for the following:" & strNewLine & strNewLine & strBackupList & strNewLine & strNewLine & "These backups can be made read-only to prevent them from being modified. Do you want to make them read-only?"
    optOption1.Caption = "Yes, make them read-only"
    optOption2.Caption = "No, leave them alone"
    If blnReadOnly Then optOption1.Value = True Else optOption2.Value = True
    optOption1.Visible = True
    optOption2.Visible = True
    lblFolder.Visible = False
    dirFolder.Visible = False
    drvDrive.Visible = False
    txtFolder.Visible = False
    cmdBack.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub FindFiles()
    Dim strPath As String, strDir As String
    Dim intSDTs As Integer
    Dim strSDTs() As String
    Dim intI As Integer
    
    strPath = lblFolder.Caption
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strDir = Dir(strPath & "*.sdt")
    
    Do While strDir <> ""
        ReDim Preserve strSDTs(intSDTs)
        strSDTs(intSDTs) = Left(strDir, Len(strDir) - 4)
        intSDTs = intSDTs + 1
        strDir = Dir()
    Loop

    intBackups = 0
    
    For intI = 0 To intSDTs - 1
        strDir = Dir(strPath & strSDTs(intI) & ".raw")
        
        If strDir <> "" Then
            ReDim Preserve strBackups(intBackups)
            strBackups(intBackups) = Left(strDir, Len(strDir) - 4)
            intBackups = intBackups + 1
        End If
    Next intI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If intWizardVersion = 1 Then
        SaveSetting "GTA Wave", "Options", "RunWizard", 0
    Else
        SaveSetting "GTA Wave", "Options", "RunWizard2", 0
    End If
End Sub

Private Sub optOption1_Click()
    If intStep = 0 Then blnMade = True Else blnReadOnly = True
End Sub

Private Sub optOption2_Click()
    If intStep = 2 Then blnReadOnly = False Else blnMade = False
End Sub

Private Sub Step5()
    intStep = 5
    If intWizardVersion = 1 Then
        If Not blnMade Then lblMessage.Caption = "The GTA sound files have been backed up and made read-only. " Else lblMessage.Caption = ""
    Else
        If Not blnMade Then lblMessage.Caption = "The GTA2 sound files have been backed up and made read-only. " Else lblMessage.Caption = ""
    End If
    
    If intWizardVersion = 1 Then
        lblMessage.Caption = lblMessage.Caption & "The Restore and Restore All commands can now be used to restore the original GTA sounds."
    Else
        lblMessage.Caption = lblMessage.Caption & "The Restore and Restore All commands can now be used to restore the original GTA2 sounds."
    End If
    
    cmdNext.Caption = "&Finish"
    optOption1.Visible = False
    optOption2.Visible = False
    lblFolder.Visible = False
    dirFolder.Visible = False
    drvDrive.Visible = False
    txtFolder.Visible = False
    cmdBack.Enabled = False
    cmdNext.Enabled = True
    cmdCancel.Enabled = False
End Sub

Private Sub MakeReadOnly()
    Dim strPath As String
    Dim intI As Integer
    Dim strFileName As String
    
    On Error GoTo CannotSet
    strPath = lblFolder.Caption
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    
    For intI = 0 To intBackups - 1
        strFileName = strPath & strBackups(intI) & ".sdt"
        SetAttr strFileName, GetAttr(strFileName) Or vbReadOnly
        strFileName = strPath & strBackups(intI) & ".raw"
        SetAttr strFileName, GetAttr(strFileName) Or vbReadOnly
    Next intI
    
    Exit Sub
    
CannotSet:
    MsgBox "Cannot make '" & strFileName & "' read-only.", vbExclamation
    Resume Next
End Sub

Private Sub Step3()
    intStep = 3
    If intWizardVersion = 1 Then
        lblMessage.Caption = "This Wizard will now backup the original GTA sound files. To do this, you need to specify the folder in which they are stored. They will be inside GTADATA\AUDIO in your GTA folder." & strNewLine & strNewLine & "Select the folder in which the original, unmodified GTA sound files are stored."
    Else
        lblMessage.Caption = "This Wizard will now backup the original GTA2 sound files. To do this, you need to specify the folder in which they are stored. They will be inside DATA\AUDIO in your GTA2 folder." & strNewLine & strNewLine & "Select the folder in which the original, unmodified GTA2 sound files are stored."
    End If
    lblFolder.Caption = dirFolder.Path
    optOption1.Visible = False
    optOption2.Visible = False
    lblFolder.Visible = True
    dirFolder.Visible = True
    drvDrive.Visible = True
    txtFolder.Visible = False
    cmdBack.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub Step4()
    Dim strPath As String
    
    intStep = 4
    lblMessage.Caption = "SDT and RAW files were found for the following:" & strNewLine & strNewLine & strBackupList & strNewLine & strNewLine & "Enter a folder to backup these files to. If you don't know where to backup to, just accept the default below and click Next."
    strPath = lblFolder.Caption
    If Right(strPath, 1) <> "\" Then strPath = strPath & "\"
    strPath = strPath & "Backup"
    txtFolder.Text = strPath
    optOption1.Visible = False
    optOption2.Visible = False
    lblFolder.Visible = False
    dirFolder.Visible = False
    drvDrive.Visible = False
    txtFolder.Visible = True
    cmdBack.Enabled = True
    cmdNext.Enabled = True
End Sub

Private Sub BackupFiles(strSource As String, strDest As String)
    Dim intI As Integer, intJ As Integer
    Dim strExt As String
    Dim strSourceFile As String, strDestFile As String
    Dim intResponse As Integer
    
    On Error GoTo ErrorHandler
    
    For intI = 0 To intBackups - 1
        
        For intJ = 0 To 1
            If intJ = 0 Then strExt = "sdt" Else strExt = "raw"
            strSourceFile = strSource & strBackups(intI) & "." & strExt
            strDestFile = strDest & strBackups(intI) & "." & strExt
            intResponse = vbYes
            If Dir(strDestFile) <> "" Then intResponse = MsgBox("The file '" & strDestFile & "' already exists. Do you want to replace it?", vbYesNo + vbDefaultButton2 + vbQuestion)
            
            If intResponse = vbYes Then
                FileCopy strSourceFile, strDestFile
                SetAttr strDestFile, GetAttr(strDestFile) Or vbReadOnly
NextFile:
            End If
            
        Next intJ
        
    Next intI
    
    Exit Sub
    
ErrorHandler:
    If Err = 75 Then MsgBox "The file '" & strDestFile & "' is read-only.", vbExclamation Else MsgBox "Cannot create the file '" & strDestFile & "'.", vbExclamation
    Resume NextFile
End Sub
