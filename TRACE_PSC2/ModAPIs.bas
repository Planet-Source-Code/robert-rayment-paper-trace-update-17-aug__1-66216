Attribute VB_Name = "ModAPIs"
' ModAPIs.bas

Option Explicit

Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
  
Public Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

' Getting image to data array
Public Declare Function GetDIBits Lib "gdi32.dll" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, _
   ByVal nStartScan As Long, ByVal nNumScans As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long) As Long
   
Public Declare Function GetObject Lib "gdi32.dll" Alias "GetObjectA" _
(ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
   
' For display & zooming
Public Declare Function SetStretchBltMode Lib "gdi32.dll" _
(ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Const HALFTONE As Long = 4
Public Const COLORONCOLOR As Long = 3

Public Declare Function StretchDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, _
   ByVal x As Long, ByVal y As Long, _
   ByVal dx As Long, ByVal dy As Long, _
   ByVal SrcX As Long, ByVal SrcY As Long, _
   ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long, _
   ByVal dwRop As Long) As Long
   
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
(ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

   

' wUsage
'Public Const DIB_RGB_COLORS As Long = 0
'Public Const DIB_PAL_COLORS As Long = 1
' eg dwRop
' vbSrcCopy = &H00CC0020

Public BHI As BITMAPINFOHEADER
Public tBitmap As BITMAP

