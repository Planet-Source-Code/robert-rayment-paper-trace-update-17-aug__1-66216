Attribute VB_Name = "ModPublics"
Option Explicit

Public ImageData() As Byte
Public DrawData() As Byte
Public PaperData() As Byte
Public saveDRAWData() As Byte
Public savePAPERData() As Byte

Public Tool As Long

' On Form1
Public Enum EToolType
   None = 0
   FreeDraw
   ALine
   Rectangle
   Cirlipse
   PolyLine
   Curvyline
   SmallRubber
   LargeRubber
End Enum

' Cirlipse
Public zrad As Single, zratio As Single


Public Sub EvalCirlipse(xs As Single, ys As Single, xe As Single, ye As Single)
' Cirlipse
' xs,ys start, xe,ye end
Dim zradx As Single, zrady As Single
   zradx = Abs(xe - xs)
   zrady = Abs(ye - ys)
   If zradx = 0 Then
      zrad = zrady
      zratio = 10
   ElseIf zradx >= zrady Then
      zrad = zradx
      zratio = zrady / zradx
   Else  'zradx<zrady
      zrad = zrady
      zratio = zrady / zradx
   End If
End Sub

Public Sub GenerateCurvyPoints(xsto() As Integer, ysto() As Integer, PointCount As Long)
Dim xfrac As Single
Dim i As Long
Dim S As Long
Dim SUP As Long
Dim oldpts As Long
Dim newpts As Long
Dim xaa() As Single, yaa() As Single
Dim xdx As Single, ydy As Single

   xfrac = 0.25
   SUP = 3
   oldpts = PointCount
   For S = 1 To SUP
       ReDim xaa(1 To oldpts), yaa(1 To oldpts)
       For i = 1 To oldpts
           xaa(i) = xsto(i): yaa(i) = ysto(i)
       Next i
       newpts = 2 * oldpts - 2
       ReDim xsto(1 To newpts), ysto(1 To newpts)
       xsto(1) = xaa(1): ysto(1) = yaa(1)
       For i = 2 To oldpts - 1
           xdx = xaa(i) - xaa(i - 1)
           xsto(2 * i - 2) = xaa(i) - xfrac * xdx
           ydy = yaa(i) - yaa(i - 1)
           ysto(2 * i - 2) = yaa(i) - xfrac * ydy
           xdx = xaa(i + 1) - xaa(i)
           xsto(2 * i - 1) = xaa(i) + xfrac * xdx
           ydy = yaa(i + 1) - yaa(i)
           ysto(2 * i - 1) = yaa(i) + xfrac * ydy
       Next i
       xsto(newpts) = xaa(oldpts): ysto(newpts) = yaa(oldpts)
       oldpts = newpts
   Next S
   PointCount = newpts
End Sub
