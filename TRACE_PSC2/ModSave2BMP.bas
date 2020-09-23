Attribute VB_Name = "ModSave2BMP"
Option Explicit

Private Type BITMAPFILEHEADER   ' For 1bpp
    bType             As Integer  ' BM        2
    bSize             As Long     ' 54+8 + (W\8 + Mod(4))*H  FileSize B
    bReserved1        As Integer  ' 0         2
    bReserved2        As Integer  ' 0         2
    bOffBits          As Long     ' 54+8      4
    bHeaderSize       As Long     ' 40        4
    bWidth            As Long     ' W         4
    bHeight           As Long     ' H         4
    bNumPlanes        As Integer  ' 1         2
    bBPP              As Integer  ' 1         2
    bCompress         As Long     ' 0         4
    bBytesInImage     As Long     ' (W\8 + Mod(4))*H  4  Image size B
    bHRES             As Long     ' 0 ignore  4
    bVRES             As Long     ' 0 ignore  4
    bUsedIndexes      As Long     ' 0 ignore  4
    bImportantIndexes As Long     ' 0 ignore  4  Total = 62
End Type

'Private Const BI_RLE4 As Long = 2&
'Private Const BI_RLE8 As Long = 1&

Private bARRIndexes() As Byte
Private Pal() As Long

Public Function SaveBMP2(FSpec$, bARR() As Byte, bWidth As Long, bHeight As Long) As Boolean
' Save 2 color BMP
' bArr() = Public PaperData(0 To 3, 0 To W - 1, 0 To H - 1)

Dim BFH As BITMAPFILEHEADER ' 54 bytes
Dim fnum As Long
Dim BytesPerScanLine As Long
Dim ix As Long, iy As Long
Dim k As Long

Dim B As Long, n As Long

On Error GoTo SaveBMPError

   BytesPerScanLine = (bWidth + 7) \ 8
   ' Expand to 4B boundary
   BytesPerScanLine = ((BytesPerScanLine + 3) \ 4) * 4
   
   With BFH
      .bType = &H4D42    ' BM
      .bWidth = bWidth
      .bHeight = bHeight
      .bSize = 54 + 8 + BytesPerScanLine * Abs(bHeight)
      .bOffBits = 54 + 8
      .bHeaderSize = 40
      .bNumPlanes = 1
      .bBPP = 1
      .bCompress = 0
      .bBytesInImage = BytesPerScanLine * Abs(bHeight)
   End With
   ' 2bpp image map.  Fill from 32bpp image map.
   ReDim bARRIndexes(0 To BytesPerScanLine - 1, 0 To Abs(bHeight) - 1)
   ReDim Pal(0 To 1)
   Pal(1) = &HFFFFFFFF
'   ' 0 0 0 0 255 255 255 255  Index 0(Black) & 1(White)
'   For k = 4 To 7
'      Pal(k) = 255
'   Next k
   
   For iy = 0 To bHeight - 1
   n = 0
   B = 0
   For ix = 0 To bWidth - 1
      If bARR(0, ix, iy) > 250 Then
         bARRIndexes(n, iy) = bARRIndexes(n, iy) Or 1
         If B < 7 Then bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2
      Else
         If B < 7 Then bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2
      End If
      B = B + 1
      If B = 8 Then
         n = n + 1
         B = 0
      End If
   Next ix
   If B <> 0 Then
      bARRIndexes(n, iy) = bARRIndexes(n, iy) * 2 ^ (7 - B)
   End If
   Next iy
   
   '-- Kill previous
   On Error Resume Next
   Kill FSpec$
   On Error GoTo 0
   
   ' bARRIndexes() & PAL() could be Input to GIF save
   ' width = BytesPerScanLine, height = bHeight +/- ?
   
   fnum = FreeFile
   Open FSpec$ For Binary As fnum
   Put #fnum, , BFH
   Put #fnum, , Pal()
   Put #fnum, , bARRIndexes()
   Close #fnum
   Erase bARRIndexes()
   SaveBMP2 = True
   On Error GoTo 0
   Exit Function
'=======
SaveBMPError:
   Close
   SaveBMP2 = False
End Function


