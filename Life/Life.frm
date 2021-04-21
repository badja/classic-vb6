VERSION 5.00
Begin VB.Form frmLife 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Life"
   ClientHeight    =   1395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2280
   Icon            =   "Life.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   93
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmLife"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Cell() As Boolean
Dim XSize As Integer, YSize As Integer
Private Sub Form_Click()
    Initialize
    Life
End Sub
Private Sub Form_Load()
    Randomize
End Sub
Sub Initialize()
    Dim X As Integer, Y As Integer
    
    XSize = ScaleWidth
    YSize = ScaleHeight
    ReDim Cell(-1 To 0, XSize - 1, YSize - 1) As Boolean
    
    For Y = 0 To YSize - 1
        
        For X = 0 To XSize - 1
            If Rnd < 0.5 Then
                Cell(0, X, Y) = True
                PSet (X, Y)
            Else
                Cell(0, X, Y) = False
                PSet (X, Y), &HFFFFFF
            End If
        Next X
    
    Next Y

End Sub
Sub Life()
    Dim X As Integer, Y As Integer
    Dim Count1 As Integer, Count2 As Integer, Count3 As Integer
    Dim Page As Boolean, Total As Integer
    Dim XCount As Integer, YCount As Integer
    Dim XReal As Integer, YReal As Integer
    
    Do
        For Y = 0 To YSize - 1
            Count1 = 0
            Count2 = 0
            
            For YCount = Y - 1 To Y + 1
                YReal = YCount
                If YReal = -1 Then
                    YReal = YSize - 1
                ElseIf YReal = YSize Then
                    YReal = 0
                End If
                If Cell(Page, XSize - 1, YReal) Then Count1 = Count1 + 1
                If Cell(Page, 0, YReal) Then Count2 = Count2 + 1
            Next YCount

            For X = 0 To XSize - 1
                
                Select Case X Mod 3
                    Case 0
                        Count3 = 0
                        XReal = X + 1
                        If XReal = XSize Then XReal = 0
                        
                        For YCount = Y - 1 To Y + 1
                            YReal = YCount
                            If YReal = -1 Then
                                YReal = YSize - 1
                            ElseIf YReal = YSize Then
                                YReal = 0
                            End If
                            If Cell(Page, XReal, YReal) Then Count3 = Count3 + 1
                        Next YCount
                    
                    Case 1
                        Count1 = 0
                        XReal = X + 1
                        If XReal = XSize Then XReal = 0
                        
                        For YCount = Y - 1 To Y + 1
                            YReal = YCount
                            If YReal = -1 Then
                                YReal = YSize - 1
                            ElseIf YReal = YSize Then
                                YReal = 0
                            End If
                            If Cell(Page, XReal, YReal) Then Count1 = Count1 + 1
                        Next YCount
                    
                    Case 2
                        Count2 = 0
                        XReal = X + 1
                        If XReal = XSize Then XReal = 0
                        
                        For YCount = Y - 1 To Y + 1
                            YReal = YCount
                            If YReal = -1 Then
                                YReal = YSize - 1
                            ElseIf YReal = YSize Then
                                YReal = 0
                            End If
                            If Cell(Page, XReal, YReal) Then Count2 = Count2 + 1
                        Next YCount
                End Select
                
                Total = Count1 + Count2 + Count3
                Cell(Not Page, X, Y) = Cell(Page, X, Y)
                
                If Cell(Page, X, Y) Then
                    If Total < 3 Or Total > 4 Then
                        Cell(Not Page, X, Y) = False
                        PSet (X, Y), &HFFFFFF
                    End If
                ElseIf Total = 3 Then
                    Cell(Not Page, X, Y) = True
                    PSet (X, Y)
                End If
            
            Next X
        
        Next Y
        
        Page = Not Page
        Refresh
        DoEvents
    Loop

End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub
