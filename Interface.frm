VERSION 5.00
Begin VB.Form InterfaceWindow 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   ClientHeight    =   4755
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ForeColor       =   &H0000FF00&
   KeyPreview      =   -1  'True
   ScaleHeight     =   317
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "InterfaceWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main inteface.
Option Explicit




'This procedure gives the command to draw the triangle.
Private Sub Form_Activate()
On Error GoTo ErrorTrap

   DrawTriangle Me, CurrentAngle, CurrentZoom
   DisplayHelp Me
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure handles the user's keystrokes.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

   If (Shift And vbCtrlMask) = vbCtrlMask Then
      Select Case KeyCode
         Case vbKeyAdd
            CurrentZoom = CurrentZoom + 1
         Case vbKeySubtract
            If CurrentZoom > 1 Then CurrentZoom = CurrentZoom - 1
      End Select
   Else
      Select Case KeyCode
         Case vbKeyAdd
            If CurrentAngle = 359 Then CurrentAngle = 0 Else CurrentAngle = CurrentAngle + 1
         Case vbKeySubtract
            If CurrentAngle = 0 Then CurrentAngle = 359 Else CurrentAngle = CurrentAngle - 1
      End Select
   End If
   
   DrawTriangle Me, CurrentAngle, CurrentZoom
   DisplayHelp Me
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
On Error GoTo ErrorTrap

   With Me
      .Width = Screen.Width / 1.5
      .Height = Screen.Height / 1.5
      .Caption = ProgramInformation()
   End With
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbRetry Then Resume
End Sub


'This procedure adjusts this window's contents to its new size.
Private Sub Form_Resize()
On Error Resume Next

   DrawTriangle Me, CurrentAngle, CurrentZoom
   DisplayHelp Me
End Sub


