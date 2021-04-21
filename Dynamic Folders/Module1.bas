Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Dim strRealShortcuts() As String
    Dim strBackupShortcuts() As String
    Dim strRealFolder As String
    Dim strBackupFolder As String
    Dim intReal As Integer
    Dim intBackup As Integer
    Dim blnQuotes As Boolean
    Dim intI As Integer, intJ As Integer
    Dim strCommand As String
    Dim strTemp As String
    Dim strPath As String
    Dim dblDummy As Double
    Dim intStart As Integer
    
    strCommand = Trim(Command)
    For intI = 1 To Len(strCommand)
        If Mid(strCommand, intI, 1) = Chr(34) Then
            If blnQuotes Then
                Exit For
            Else
                blnQuotes = True
            End If
        ElseIf Mid(strCommand, intI, 1) = " " And Not blnQuotes Then
            intI = intI - 1
            Exit For
        End If
    Next intI
    
    strRealFolder = Left(strCommand, intI)
    strBackupFolder = Trim(Right(strCommand, Len(strCommand) - intI))
    If Left(strRealFolder, 1) = Chr(34) And Right(strRealFolder, 1) = Chr(34) Then strRealFolder = Trim(Mid(strRealFolder, 2, Len(strRealFolder) - 2))
    If Left(strBackupFolder, 1) = Chr(34) And Right(strBackupFolder, 1) = Chr(34) Then strBackupFolder = Trim(Mid(strBackupFolder, 2, Len(strBackupFolder) - 2))
    If Right(strRealFolder, 1) <> "\" Then strRealFolder = strRealFolder & "\"
    If Right(strBackupFolder, 1) <> "\" Then strBackupFolder = strBackupFolder & "\"
    
    strTemp = Dir(strRealFolder & "*.lnk")
    Do While strTemp <> ""
        ReDim Preserve strRealShortcuts(intReal)
        strRealShortcuts(intReal) = strTemp
        intReal = intReal + 1
        strTemp = Dir()
    Loop
    
    strTemp = Dir(strRealFolder & "*.pif")
    Do While strTemp <> ""
        ReDim Preserve strRealShortcuts(intReal)
        strRealShortcuts(intReal) = strTemp
        intReal = intReal + 1
        strTemp = Dir()
    Loop

    strTemp = Dir(strBackupFolder & "*.lnk")
    Do While strTemp <> ""
        ReDim Preserve strBackupShortcuts(intBackup)
        strBackupShortcuts(intBackup) = strTemp
        intBackup = intBackup + 1
        strTemp = Dir()
    Loop
    
    strTemp = Dir(strBackupFolder & "*.pif")
    Do While strTemp <> ""
        ReDim Preserve strBackupShortcuts(intBackup)
        strBackupShortcuts(intBackup) = strTemp
        intBackup = intBackup + 1
        strTemp = Dir()
    Loop
    
    For intI = 0 To intReal - 1
        Open strRealFolder & strRealShortcuts(intI) For Binary Access Read As 1
        strTemp = Space(LOF(1))
        Get #1, , strTemp
        Close 1
        
        If LCase(Right(strRealShortcuts(intI), 1)) = "f" Then
            For intJ = 37 To 98
                If Mid(strTemp, intJ, 1) = Chr(0) Then Exit For
            Next intJ
            strPath = Mid(strTemp, 37, intJ - 37)
            If Dir(strPath) = "" Then
                Name strRealFolder & strRealShortcuts(intI) As strBackupFolder & strRealShortcuts(intI)
            End If
        Else
            intStart = InStr(strTemp, ":\")
            Do
                intStart = InStr(intStart + 1, strTemp, ":\")
            Loop Until Mid(strTemp, intStart - 1, 1) >= "A" And Mid(strTemp, intStart - 1, 1) <= "Z"
            For intJ = intStart To Len(strTemp)
                If Mid(strTemp, intJ, 1) = Chr(0) Then Exit For
            Next intJ
            strPath = Mid(strTemp, intStart - 1, intJ - intStart + 1)
            If LCase(Right(strPath, 4)) = ".exe" And Dir(strPath) = "" Then
                Name strRealFolder & strRealShortcuts(intI) As strBackupFolder & strRealShortcuts(intI)
            End If
        End If
    Next intI

    For intI = 0 To intBackup - 1
        Open strBackupFolder & strBackupShortcuts(intI) For Binary Access Read As 1
        strTemp = Space(LOF(1))
        Get #1, , strTemp
        Close 1
        
        If LCase(Right(strBackupShortcuts(intI), 1)) = "f" Then
            For intJ = 37 To 98
                If Mid(strTemp, intJ, 1) = Chr(0) Then Exit For
            Next intJ
            strPath = Mid(strTemp, 37, intJ - 37)
            If Dir(strPath) <> "" Then
                Name strBackupFolder & strBackupShortcuts(intI) As strRealFolder & strBackupShortcuts(intI)
            End If
        Else
            intStart = InStr(strTemp, ":\")
            Do
                intStart = InStr(intStart + 1, strTemp, ":\")
            Loop Until Mid(strTemp, intStart - 1, 1) >= "A" And Mid(strTemp, intStart - 1, 1) <= "Z"
            For intJ = intStart To Len(strTemp)
                If Mid(strTemp, intJ, 1) = Chr(0) Then Exit For
            Next intJ
            strPath = Mid(strTemp, intStart - 1, intJ - intStart + 1)
            If LCase(Right(strPath, 4)) = ".exe" And Dir(strPath) <> "" Then
                Name strBackupFolder & strBackupShortcuts(intI) As strRealFolder & strBackupShortcuts(intI)
            End If
        End If
    Next intI
    
    dblDummy = Shell("explorer.exe " & strRealFolder, vbNormalFocus)
End Sub
