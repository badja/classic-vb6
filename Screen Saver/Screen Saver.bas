Attribute VB_Name = "Module1"
Option Explicit
Public Const WS_CHILD = &H40000000
Public Const GWL_STYLE = (-16)
Public Const GWL_HWNDPARENT = (-8)
Public Const HWND_TOP = 0&
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SPI_SCREENSAVERRUNNING = 97&

Declare Function ShowCursor Lib "user32" _
(ByVal bShow As Long) As Long

Declare Function GetClientRect Lib "user32" _
(ByVal hwnd As Long, lpRect As RECT) As Long

Declare Function SetParent Lib "user32" _
(ByVal hWndChild As Long, _
ByVal hWndNewParent As Long) As Long

Declare Function IsWindowVisible Lib "user32" _
(ByVal hwnd As Long) As Long

Declare Sub SetWindowPos Lib "user32" (ByVal _
hwnd As Long, ByVal hWndInsertAfter As Long, _
ByVal X As Long, ByVal Y As Long, ByVal _
cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Declare Function SetWindowLong Lib "user32" _
Alias "SetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Declare Function GetWindowLong Lib "user32" _
Alias "GetWindowLongA" (ByVal hwnd As Long, _
ByVal nIndex As Long) As Long

Declare Function PwdChangePassword& Lib "mpr" _
Alias "PwdChangePasswordA" (ByVal lpcRegkeyname$, _
ByVal hwnd&, ByVal uiReserved1&, ByVal uiReserved2&)

Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As _
Long, ByVal uParam As Long, lpvParam As Any, ByVal _
fuWinIni As Long) As Long

Declare Function VerifyScreenSavePwd Lib _
"password.cpl" (ByVal hwnd&) As Boolean

' Registry API functions
Private Const HKEY_CURRENT_USER = &H80000001
Public Const REG_DWORD As Long = 4

Declare Function RegOpenKey Lib "advapi32.dll" _
Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal _
lpSubKey As String, phkResult As Long) As Long

Declare Function RegQueryValueEx Lib "advapi32.dll" _
Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal _
lpValueName As String, ByVal lpReserved As Long, _
lpType As Long, ByVal lpData As String, lpcbData _
As Long) As Long

Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Public MouseMoveCount As Long
Public DispRec As RECT
Private dispHWND, style As Long

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Public Function ReadRegistry(ByVal Group _
As Long, ByVal Section As String, ByVal Key _
As String) As String

Dim lResult As Long, lKeyValue As Long, _
lDataTypeValue As Long, lValueLength As Long, _
sValue As String, td As Double

On Error Resume Next
lResult = RegOpenKey(Group, Section, lKeyValue)
sValue = Space$(2048)
lValueLength = Len(sValue)
lResult = RegQueryValueEx(lKeyValue, Key, 0&, _
lDataTypeValue, sValue, lValueLength)
If (lResult = 0) And (Err.Number = 0) Then
   If lDataTypeValue = REG_DWORD Then
      td = Asc(Mid$(sValue, 1, 1)) + &H100& * _
      Asc(Mid$(sValue, 2, 1)) + &H10000 * _
      Asc(Mid$(sValue, 3, 1)) + &H1000000 * _
      CDbl(Asc(Mid$(sValue, 4, 1)))
      sValue = Format$(td, "000")
   End If
   sValue = Left$(sValue, lValueLength - 1)
Else
   sValue = "Not Found"
End If
lResult = RegCloseKey(lKeyValue)
ReadRegistry = sValue
End Function

Public Sub EnableCtrlAltDelete(bEnabled As Boolean)
    If bEnabled = False Then
      'disable ctrl+alt+delete
      SystemParametersInfo SPI_SCREENSAVERRUNNING, 1&, 0&, 0&
    Else
      'enable ctrl+alt+delete
      SystemParametersInfo SPI_SCREENSAVERRUNNING, 0&, 0&, 0&
    End If
End Sub

Function IsPasswordEnabled()
IsPasswordEnabled = ReadRegistry(HKEY_CURRENT_USER, _
"Control Panel\Desktop", "ScreenSaveUsePassword")
End Function

Public Sub DoPreviewMode()
dispHWND = CLng(Right$(Command$, Len(Command$) - 3))
Load frmPreview
GetClientRect dispHWND, DispRec
style = GetWindowLong(frmPreview.hwnd, GWL_STYLE)
style = style Or WS_CHILD ' Append "WS_CHILD"
SetWindowLong frmPreview.hwnd, GWL_STYLE, style
SetParent frmPreview.hwnd, dispHWND
SetWindowLong frmPreview.hwnd, GWL_HWNDPARENT, dispHWND
SetWindowPos frmPreview.hwnd, HWND_TOP, 0&, 0&, _
DispRec.Right, DispRec.Bottom, _
SWP_NOZORDER Or SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Public Sub Main()
    If Left(Command$, 2) = "/p" Then
      ' Display the mini preview form
      DoPreviewMode
    ElseIf Left(Command$, 2) = "/c" Then
       MsgBox "Settings Dialog Box", vbInformation
    ElseIf Left(Command$, 2) = "/s" Then
      ' Screensaver Normal Mode - Disable Ctrl+Alt+Delete
      Call EnableCtrlAltDelete(False)
      Call ShowCursor(False)
      frmMain.Show
    ElseIf Left(Command$, 2) = "/a" Then
      ' show the change password box
      ' retrieve the handle of the display
      ' dialog box from the command line
      dispHWND = CLng(Right$(Command$, Len(Command$) - 3))
      ' show the password change box
      Call PwdChangePassword("SCRSAVE", dispHWND, 0, 0)
    End If
End Sub

