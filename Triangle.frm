VERSION 5.00
Begin VB.Form TriangleWindow 
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
Attribute VB_Name = "TriangleWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This module contains this program's main inteface.
Option Explicit

Private Const PI As Double = 3.14159265358979           'Defines the value of PI.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI   'Defines the number of degrees per radian.

'This structure defines a triangle.
Private Type TriangleStr
   x(0 To 2) As Double       'Contains the triangle's horizontal coordinates.
   y(0 To 2) As Double       'Contains the triangle's vertical coordinates.
   Sides(0 To 2) As Double   'Contains the lengths of triangle's sides.
End Type

Private CurrentAngle As Long   'The current angle used to draw the triangle.
Private CurrentZoom As Long    'The current zoom used to display the triangle.

'This procedure displays the help.
Private Sub DisplayHelp()
   Me.CurrentX = 0
   Me.CurrentY = Me.ScaleHeight - 16
   Me.Print "Use the +/- keys to adjust the triangle's angles and CTRL + +/- to zoom in and out.";
End Sub

'This procedure displays status information.
Private Sub DisplayStatus(Angle As Long, Zoom As Long, Triangle As TriangleStr)
Dim Side As Long

   Me.CurrentX = 0
   Me.CurrentY = 16
   Me.Print " Zoom: "; CStr(Zoom)
   Me.Print " Angle: " & CStr(Angle)
   Me.Print
   Me.Print " SINe: " & Format$(Sin(Angle / DEGREES_PER_RADIAN), "0.00")
   Me.Print " COSine: " & Format$(Cos(Angle / DEGREES_PER_RADIAN), "0.00")
   Me.Print " TANgent: ";
   
   If Angle = 90 Or Angle = 270 Then
      Me.Print "-"
   Else
      Me.Print Format$(Tan(Angle / DEGREES_PER_RADIAN), "00.00")
   End If
   
   With Triangle
      Me.Print
      Me.Print " 'My'-SINe: " & Format$(MySine(.Sides()), "0.00")
      Me.Print " 'My'-COSine: " & Format$(MyCosine(.Sides()), "0.00")
      Me.Print " 'My'-TANgent: " & Format$(MyTangent(.Sides()), "00.00")

      Me.Print
      For Side = LBound(.Sides()) To UBound(.Sides())
         Me.Print " Side "; CStr(Side); ": "; Format$(.Sides(Side), "0.00")
      Next Side
   End With
End Sub

'This procedure draws a triangle using the specified angle.
Private Sub DrawTriangle(Angle As Long, Zoom As Long)
Dim Side As Long
Dim Triangle As TriangleStr

   With Triangle
      .x(0) = (Me.ScaleWidth / 2) + (Cos(Angle / DEGREES_PER_RADIAN) * Zoom)
      .y(0) = (Me.ScaleHeight / 2) - (Sin(Angle / DEGREES_PER_RADIAN) * Zoom)
      .x(1) = Me.ScaleWidth / 2
      .y(1) = Me.ScaleHeight / 2
      .x(2) = .x(0)
      .y(2) = .y(1)
   
      .Sides(0) = .y(2) - .y(0)
      .Sides(1) = .x(2) - .x(1)
      .Sides(2) = Sqr(((.x(1) - .x(0)) ^ 2) + ((.y(1) - .y(0)) ^ 2))

      Me.Cls
      Me.PSet (.x(LBound(.Sides())), .y(LBound(.Sides())))
      For Side = LBound(.Sides()) To UBound(.Sides())
         Me.Line -(.x(Side), .y(Side))
      Next Side
      Me.Line -(.x(LBound(.Sides())), .y(LBound(.Sides())))
      Me.Circle (.x(1), .y(1)), Zoom

      For Side = LBound(.Sides()) To UBound(.Sides())
         Me.CurrentX = .x(Side)
         Me.CurrentY = .y(Side)
         Me.Print CStr(Side);
      Next Side
   End With

   DisplayStatus Angle, Zoom, Triangle
End Sub

'This procedure returns the cosine based on the specified triangle side lengths.
Private Function MyCosine(Sides() As Double) As Double
Dim Cosine As Double

   Cosine = Sides(1) / Sides(2)
   
   MyCosine = Cosine
End Function

'This procedure returns the sine based on the specified triangle side lengths.
Private Function MySine(Sides() As Double) As Double
Dim Sine As Double

   Sine = Sides(0) / Sides(2)
   
   MySine = Sine
End Function


'This procedure returns the tangent based on the specified triangle side lengths.
Private Function MyTangent(Sides() As Double) As Double
Dim Tangent As Double

   Tangent = 0
   If Not Sides(1) = 0 Then Tangent = Sides(0) / Sides(1)
   If Tangent < -60 Or Tangent > 60 Then Tangent = 0
   
   MyTangent = Tangent
End Function



'This procedure gives the command to draw the triangle.
Private Sub Form_Activate()
   DrawTriangle CurrentAngle, CurrentZoom
   DisplayHelp
End Sub

'This procedure handles the user's keystrokes.
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
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
   
   DrawTriangle CurrentAngle, CurrentZoom
   DisplayHelp
End Sub

'This procedure initializes this window.
Private Sub Form_Load()
   With App
      ChDrive Left$(.Path, InStr(.Path, "\"))
      ChDir .Path

      Me.Caption = App.Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With
   
   With Me
      .Width = Screen.Width / 1.5
      .Height = Screen.Height / 1.5
   End With
   
   CurrentAngle = 45
   CurrentZoom = 200
End Sub


'This procedure adjusts this window's contents to its new size.
Private Sub Form_Resize()
   DisplayHelp
End Sub


