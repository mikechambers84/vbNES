Attribute VB_Name = "modRender"
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Type BITMAP
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long

Private Type RGBQUAD
   rgbBlue As Byte
   rgbGreen As Byte
   rgbRed As Byte
   rgbAlpha As Byte
End Type

Private Type BITMAPINFOHEADER
   bmSize As Long
   bmWidth As Long
   bmHeight As Long
   bmPlanes As Integer
   bmBitCount As Integer
   bmCompression As Long
   bmSizeImage As Long
   bmXPelsPerMeter As Long
   bmYPelsPerMeter As Long
   bmClrUsed As Long
   bmClrImportant As Long
End Type

Private Type BITMAPINFO
   bmHeader As BITMAPINFOHEADER
   bmColors(0 To 255) As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal hDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long

'Public drawscreen(0 To 2, 0 To 511, 0 To 479) As Byte
Public drawscreen() As Byte

'Routine to set an image's pixel information from an array dimensioned (rgb, x, y)
Public Sub SetImageData(ByRef DstPictureBox As PictureBox, ByRef ImageData() As Byte)

'Variables for the necessary bitmap types
Dim bm As BITMAP
Dim bmi As BITMAPINFO

'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data - this is identical to what you've already seen
bmi.bmHeader.bmSize = 40
bmi.bmHeader.bmPlanes = 1
bmi.bmHeader.bmBitCount = 24
bmi.bmHeader.bmCompression = 0

'Calculate the size of the bitmap type (in bytes)
Dim bmLen As Long
bmLen = Len(bm)

'Get the picture box information from DstPictureBox and put it into our 'bm' variable
GetObject DstPictureBox.Image, bmLen, bm

'Now that we know the object's size, finish building the temporary header to pass to the StretchDIBits call (continuing to use the 'bmi' we used above)
bmi.bmHeader.bmWidth = bm.bmWidth
bmi.bmHeader.bmHeight = bm.bmHeight

'Now that we've built the temporary header, we use StretchDIBits to take the data from the ImageData() array and put it into SrcPictureBox using the settings specified in 'bmi' (the StretchDIBits call should be on one continuous line)
StretchDIBits DstPictureBox.hDC, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, bm.bmWidth, bm.bmHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy

'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
'Note: always keep AutoRedraw as 'True' when using DIB sections. Otherwise, you WILL get unpredictable results.
If DstPictureBox.AutoRedraw = True Then
   DstPictureBox.Picture = DstPictureBox.Image
   DstPictureBox.Refresh
End If
End Sub

