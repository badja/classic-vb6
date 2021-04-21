Attribute VB_Name = "modPublic"
Option Explicit

Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public blnCancel As Boolean
Public blnEditing As Boolean
Public blnLastConst As Boolean
Public blnLastScaleVar As Boolean
Public blnLooping As Boolean
Public blnScale As Boolean
Public blnScaleVar As Boolean
Public dblExternalTaskID As Double
Public intBrowseVersion As Integer
Public intGTAVersion As Integer
Public intPlayIndex As Integer
Public intWizardVersion As Integer
Public lngCurRate As Long, lngPlayRate As Long
Public lngCurVariation As Long, lngPlayVariation As Long
Public lngCurSize As Long, lngPlayLoopStart As Long
Public lngPlayLoopEnd As Long
Public sngScale As Single
Public strBackupDir As String
Public strInfo() As String
Public strNewLine As String
Public strTempFile As String
Public varExternalDate As Variant

Public Sub CreateFile(strFileName As String, intIndex As Integer, strInfoFile As String, blnLoop As Boolean, blnRandom As Boolean)
    Dim lngChannels As Long, lngBits As Long
    Dim lngBegin As Long, lngLength As Long
    Dim strData As String
    Dim strRawFile As String
    Dim strSampleRate As String
    Dim lngLoopEnd As Long
    
    On Error GoTo ErrorHandler
    lngChannels = 1
    
    If intGTAVersion = 1 Then
        If LCase(Right(strInfoFile, 12)) = "level000.sdt" Then
            lngBits = 16
            If intIndex <= 2 Then lngChannels = 2
        Else
            lngBits = 8
        End If
    Else
        If intIndex >= 69 And intIndex <= 136 Then lngBits = 8 Else lngBits = 16
        lngChannels = 1
    End If
    
    lngBegin = AddHex(Left(strInfo(intIndex), 4))
    lngLength = AddHex(Mid(strInfo(intIndex), 5, 4))
    
    If blnLoop And intGTAVersion = 2 Then
        If lngPlayLoopStart >= 0 Then
            lngBegin = lngBegin + lngPlayLoopStart
            
            If lngPlayLoopEnd = -1 Then
                lngLength = lngLength - lngPlayLoopStart
            Else
                lngLength = lngPlayLoopEnd - lngPlayLoopStart
            End If
        Else
            lngBegin = lngBegin + AddHex(Mid(strInfo(intIndex), 17, 4))
            lngLoopEnd = AddHex(Mid(strInfo(intIndex), 21, 4))
            
            If lngLoopEnd = -1 Then
                lngLength = lngLength - AddHex(Mid(strInfo(intIndex), 17, 4))
            Else
                lngLength = lngLoopEnd - AddHex(Mid(strInfo(intIndex), 17, 4))
            End If
        End If
    End If
    
    strData = Space(lngLength)
    strRawFile = Left(strInfoFile, Len(strInfoFile) - 4) & ".raw"
    
    If Dir(strRawFile) = "" Then
        MsgBox "Cannot find '" & strRawFile & "'", vbCritical
    Else
        Open strRawFile For Binary Access Read As #1
        Get #1, lngBegin + 1, strData
        Close #1

        SafeKill strFileName
        Open strFileName For Binary Access Write As #1
        
        If lngPlayRate > 0 Then
            strSampleRate = MakeHexString(lngPlayRate, 4)
        ElseIf lngPlayVariation <> 0 Then
            strSampleRate = MakeHexString(AddHex(Mid(strInfo(intIndex), 9, 4)) + lngPlayVariation, 4)
        Else
            strSampleRate = Mid(strInfo(intIndex), 9, 4)
            If blnRandom Then strSampleRate = MakeHexString(Int((2 * AddHex(Mid(strInfo(intIndex), 13, 4)) + 1) * Rnd + AddHex(strSampleRate) - AddHex(Mid(strInfo(intIndex), 13, 4))), 4)
        End If
        
        Put #1, , "RIFF" & MakeHexString(lngLength + 36, 4) & "WAVE"
        Put #1, , "fmt " & MakeHexString(16, 4) & MakeHexString(1, 2) & MakeHexString(lngChannels, 2) & strSampleRate & MakeHexString(AddHex(strSampleRate) * lngChannels * lngBits / 8, 4) & MakeHexString(lngChannels * lngBits / 8, 2) & MakeHexString(lngBits, 2)
        Put #1, , "data" & MakeHexString(lngLength, 4)
        Put #1, , strData
        Close #1
    End If
    
    Exit Sub

ErrorHandler:
    MsgBox "File error", vbCritical
    Close #1
End Sub

Public Function AddHex(strData As String) As Long
    Dim intI As Integer
        
    If strData = String(Len(strData), Chr(255)) Then
        AddHex = -1
    Else
        For intI = 1 To Len(strData)
            AddHex = AddHex + 256 ^ (intI - 1) * Asc(Mid(strData, intI, 1))
        Next intI
    End If
    
End Function

Public Sub SafeKill(strPathName As String)
    On Error Resume Next
    
    If Dir(strPathName) <> "" Then Kill strPathName
End Sub

Public Function MakeHexString(lngDecimal As Long, intBytes As Integer) As String
    Dim intI As Integer
    Dim lngRest As Long
    Dim strData As String
    
    If lngDecimal = -1 Then
        MakeHexString = String(intBytes, Chr(255))
    Else
        lngRest = lngDecimal
        
        For intI = 1 To intBytes
            strData = strData & Chr(lngRest Mod 256)
            lngRest = lngRest \ 256
        Next intI
        
        MakeHexString = strData
    End If
End Function

Public Sub StopPlaying()
    Dim lngSuccess As Long
    
    lngSuccess = sndPlaySound(vbNullString, 0)
    blnLooping = False
End Sub

Public Function GetPath(strPathName As String) As String
    Dim intI As Integer
    
    If InStr(strPathName, "\") Then
        
        intI = Len(strPathName)
        
        Do While Mid(strPathName, intI, 1) <> "\" And intI > 0
            intI = intI - 1
        Loop
    
    End If
    
    GetPath = Left(strPathName, intI)
End Function
