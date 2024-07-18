Attribute VB_Name = "CoreModule"
'This module contains this program's core procedures.
Option Explicit

Private Const PI As Double = 3.14159265358979           'Defines the value of PI.
Private Const DEGREES_PER_RADIAN As Double = 180 / PI   'Defines the number of degrees per radian.

'This structure defines a triangle.
Private Type TriangleStr
   x(0 To 2) As Double       'Defines the triangle's horizontal coordinates.
   y(0 To 2) As Double       'Defines the triangle's vertical coordinates.
   Sides(0 To 2) As Double   'Defines the lengths of triangle's sides.
End Type

Public CurrentAngle As Long   'Contains the current angle used to draw the triangle.
Public CurrentZoom As Long    'Contains the current zoom used to display the triangle.

'This procedure displays the help.
Public Sub DisplayHelp(Canvas As Object)
On Error GoTo ErrorTrap

   Canvas.CurrentX = 0
   Canvas.CurrentY = Canvas.ScaleHeight - 16
   Canvas.Print "Use the +/- keys to adjust the triangle's angles and CTRL + +/- to zoom in and out.";
   
EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure displays status information.
Private Sub DisplayStatus(Canvas As Object, Angle As Long, Zoom As Long, Triangle As TriangleStr)
On Error GoTo ErrorTrap
Dim Side As Long

   Canvas.CurrentX = 0
   Canvas.CurrentY = 16
   Canvas.Print " Zoom: "; CStr(Zoom)
   Canvas.Print " Angle: " & CStr(Angle)
   Canvas.Print
   Canvas.Print " SINe: " & Format$(Sin(Angle / DEGREES_PER_RADIAN), "0.00")
   Canvas.Print " COSine: " & Format$(Cos(Angle / DEGREES_PER_RADIAN), "0.00")
   Canvas.Print " TANgent: ";
   
   If Angle = 90 Or Angle = 270 Then
      Canvas.Print "-"
   Else
      Canvas.Print Format$(Tan(Angle / DEGREES_PER_RADIAN), "00.00")
   End If
   
   With Triangle
      Canvas.Print
      Canvas.Print " 'My'-SINe: " & Format$(MySine(.Sides()), "0.00")
      Canvas.Print " 'My'-COSine: " & Format$(MyCosine(.Sides()), "0.00")
      Canvas.Print " 'My'-TANgent: " & Format$(MyTangent(.Sides()), "00.00")

      Canvas.Print
      For Side = LBound(.Sides()) To UBound(.Sides())
         Canvas.Print " Side "; CStr(Side); ": "; Format$(.Sides(Side), "0.00")
      Next Side
   End With

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure draws a triangle using the specified angle.
Public Sub DrawTriangle(Canvas As Object, Angle As Long, Zoom As Long)
On Error GoTo ErrorTrap
Dim Side As Long
Dim Triangle As TriangleStr

   With Triangle
      .x(0) = (Canvas.ScaleWidth / 2) + (Cos(Angle / DEGREES_PER_RADIAN) * Zoom)
      .y(0) = (Canvas.ScaleHeight / 2) - (Sin(Angle / DEGREES_PER_RADIAN) * Zoom)
      .x(1) = Canvas.ScaleWidth / 2
      .y(1) = Canvas.ScaleHeight / 2
      .x(2) = .x(0)
      .y(2) = .y(1)
   
      .Sides(0) = .y(2) - .y(0)
      .Sides(1) = .x(2) - .x(1)
      .Sides(2) = Sqr(((.x(1) - .x(0)) ^ 2) + ((.y(1) - .y(0)) ^ 2))

      Canvas.Cls
      Canvas.PSet (.x(LBound(.Sides())), .y(LBound(.Sides())))
      For Side = LBound(.Sides()) To UBound(.Sides())
         Canvas.Line -(.x(Side), .y(Side))
      Next Side
      Canvas.Line -(.x(LBound(.Sides())), .y(LBound(.Sides())))
      Canvas.Circle (.x(1), .y(1)), Zoom

      For Side = LBound(.Sides()) To UBound(.Sides())
         Canvas.CurrentX = .x(Side)
         Canvas.CurrentY = .y(Side)
         Canvas.Print CStr(Side);
      Next Side
   End With

   DisplayStatus Canvas, Angle, Zoom, Triangle

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure handles any errors that occur.
Public Function HandleError(Optional ReturnPreviousChoice As Boolean = False) As Long
Dim Description As String
Dim ErrorCode As Long
Static Choice As Long

   Description = Err.Description
   ErrorCode = Err.Number
   On Error Resume Next
   If Not ReturnPreviousChoice Then
      Choice = MsgBox(Description & "." & vbCr & "Error code: " & CStr(ErrorCode), vbAbortRetryIgnore Or vbDefaultButton2 Or vbExclamation)
   End If
   
   If Choice = vbAbort Then End
   
   HandleError = Choice
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   CurrentAngle = 45
   CurrentZoom = 200

   InterfaceWindow.Show

EndProcedure:
   Exit Sub
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Sub

'This procedure returns the cosine based on the specified triangle side lengths.
Private Function MyCosine(Sides() As Double) As Double
On Error GoTo ErrorTrap
Dim Cosine As Double

   Cosine = Sides(1) / Sides(2)
  
EndProcedure:
   MyCosine = Cosine
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure returns the sine based on the specified triangle side lengths.
Private Function MySine(Sides() As Double) As Double
On Error GoTo ErrorTrap
Dim Sine As Double

   Sine = Sides(0) / Sides(2)
   
EndProcedure:
   MySine = Sine
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure returns the tangent based on the specified triangle side lengths.
Private Function MyTangent(Sides() As Double) As Double
On Error GoTo ErrorTrap
Dim Tangent As Double

   Tangent = 0
   If Not Sides(1) = 0 Then Tangent = Sides(0) / Sides(1)
   If Tangent < -60 Or Tangent > 60 Then Tangent = 0
   
EndProcedure:
   MyTangent = Tangent
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function

'This procedure returns information about this program.
Public Function ProgramInformation() As String
On Error GoTo ErrorTrap
Dim Information As String

   With App
      Information = .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName
   End With

EndProcedure:
   ProgramInformation = Information
   Exit Function
   
ErrorTrap:
   If HandleError() = vbIgnore Then Resume EndProcedure
   If HandleError(ReturnPreviousChoice:=True) = vbIgnore Then Resume
End Function



