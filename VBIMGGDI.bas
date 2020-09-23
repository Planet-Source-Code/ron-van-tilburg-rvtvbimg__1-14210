Attribute VB_Name = "VBIMGGDI"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'VBIMGGDI.bas
' API Functions for capturing/setting memory dc and bitmaps

Public Type BITMAP '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Integer
  bmBitsPixel As Integer
  bmBits As Long
End Type

Public Type BITMAPFILEHEADER
  bfType As Integer
  bfSize As Long
  bfReserved1 As Integer
  bfReserved2 As Integer
  bfOffBits As Long
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type RGBQUAD
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Public Type BITMAPINFO
  bmiHeader As BITMAPINFOHEADER        '40&
  bmiColors(1 To 256) As RGBQUAD       '256&*4&=1024
End Type

Public Type PALETTEENTRY
  peRed As Byte
  peGreen As Byte
  peBlue As Byte
  peFlags As Byte
End Type

Public Type LOGPALETTE
  palVersion As Integer
  palNumEntries As Integer
  palPalEntry(1 To 256) As PALETTEENTRY
End Type

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hdc As Long, ByVal uObjectType As Long) As Long
Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function GetObjectX Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As POINTAPI) As Long
Public Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Public Const DIB_RGB_COLORS = 0         '  color table in RGBs
Public Const BI_RGB = 0&

Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)

Public Const HALFTONE = 4
Public Const GDI_ERROR = &HFFFF

Public Type POINTAPI
  x As Long
  y As Long
End Type

'Allocate a DC and Bitmap area compatible with the current display - they are NOT CONNECTED YET
Private Function AllocateDCandBitmap(ByVal Width As Long, ByVal Height As Long, _
                                     ByRef NewMemDC As Long, ByRef NewBitMap As Long) As Long
  Dim rc As Long, hdc As Long
  
  rc = 0: NewMemDC = 0: NewBitMap = 0
  hdc = CreateDC("DISPLAY", "", "", 0)     'get a DisplayDC
  If hdc <> 0 Then
    NewMemDC = CreateCompatibleDC(hdc)
    If NewMemDC <> 0 Then
      NewBitMap = CreateCompatibleBitmap(hdc, Width, Height)
      If NewBitMap <> 0 Then               'we have a suitable bitmap ready for populating
        rc = 1
      End If
    End If
  End If
  If hdc <> 0 Then Call DeleteDC(hdc)                          'get rid of this we dont need it anymore
  If rc = 0 Then Call FreeDCandBitmap(NewMemDC, NewBitMap)
  AllocateDCandBitmap = rc
End Function

'Clone a DC and Bitmap area compatible with the given hDC - they are NOT CONNECTED YET
Private Function CloneDCandBitmap(ByVal hdc As Long, ByVal Width As Long, ByVal Height As Long, _
                                  ByRef NewMemDC As Long, ByRef NewBitMap As Long) As Long
  Dim rc As Long
  
  rc = 0: NewMemDC = 0: NewBitMap = 0
  NewMemDC = CreateCompatibleDC(hdc)
  If NewMemDC <> 0 Then
    NewBitMap = CreateCompatibleBitmap(hdc, Width, Height)
    If NewBitMap <> 0 Then               'we have a suitable bitmap ready for populating
      rc = 1
    End If
  End If
  If rc = 0 Then Call FreeDCandBitmap(NewMemDC, NewBitMap)
  CloneDCandBitmap = rc
End Function
  
  'Free the memory allocated in DC and BitMap - they should not be connected
  'WARNING: Dont try this with anything NOT allocated with AllocateDCandBitmap(), or CloneDCandBitmap
  
Private Sub FreeDCandBitmap(ByRef hdc As Long, ByRef hBitmap As Long)
  If hBitmap <> 0 Then Call DeleteObject(hBitmap): hBitmap = 0
  If hdc <> 0 Then Call DeleteDC(hdc): hdc = 0
End Sub

  '--------------------------------------------------------------------------------------------------------
  '================= This fills a GDI hDC and Bitmap from the data in PixBits() =======
  '--------------------------------------------------------------------------------------------------------

Private Function DCBitMapFromImage(ByVal Width As Long, ByVal Height As Long, _
                                   ByRef PixBits() As Byte, ByVal BitsPerPixel As Long, _
                                   ByRef CMap() As Byte, ByVal NMapColors As Long, _
                                   ByRef DestDC As Long, ByRef DestBitMap As Long, _
                                   Optional OpCodes As Long = 0, _
                                   Optional OpParm1 As Long = 0, _
                                   Optional OpParm2 As Long = 0) As Long
  Dim BMI  As BITMAPINFO
  Dim rc As Long, i As Long, p As Long, RasterOp As Long, lppt As POINTAPI
  Dim NewWidth As Long, NewHeight As Long
  
  rc = 0
  On Error GoTo ErrorFound
  
  NewWidth = Width: If OpParm1 <> 0 Then NewWidth = OpParm1
  NewHeight = Height: If OpParm2 <> 0 Then NewHeight = OpParm2
  
  If (OpCodes And PIC_INVERT_COLOR) = PIC_INVERT_COLOR Then RasterOp = NOTSRCCOPY Else RasterOp = SRCCOPY
  
  '--------------------------------------------------------------------------------------------------------
  '------------------------------------ BitMap CREATION FROM STORAGE --------------------------------------
  '--------------------------------------------------------------------------------------------------------
  
  rc = AllocateDCandBitmap(Abs(NewWidth), Abs(NewHeight), DestDC, DestBitMap)
  If rc <> 0 Then
    With BMI.bmiHeader
      .biSize = 40                     'sizeof(BITMAPINFOHEADER
      .biWidth = Width                 '{width of the bitmapclip}
      If (OpCodes And PIC_FLIP_VERT) = PIC_FLIP_VERT Then
        .biHeight = -Height            '{height of the bitmapclip} make sure its top to bottom
      Else
        .biHeight = Height             '{height of the bitmapclip} make sure its bottom to top
      End If
      .biPlanes = 1
      .biBitCount = BitsPerPixel       '{desired color resolution (1, 4, 8, or 16,24)}
      .biClrUsed = NMapColors
      .biCompression = BI_RGB
      .biSizeImage = BMPRowModulo(Width, BitsPerPixel) * Height
    End With
        
    If BitsPerPixel < PIC_16BPP Then   'we have a Colormap to use
      p = 0
      For i = 1 To NMapColors
        With BMI.bmiColors(i)
          .rgbRed = CMap(p): p = p + 1
          .rgbGreen = CMap(p): p = p + 1
          .rgbBlue = CMap(p): p = p + 1
          .rgbReserved = 0
        End With
      Next
    End If
    
    If rc <> 0 Then
      DestBitMap = SelectObject(DestDC, DestBitMap)   'select it in
      Call SetStretchBltMode(DestDC, HALFTONE)
      Call SetBrushOrgEx(DestDC, 0, 0, lppt)
      'for some reason I cant get x and y mirroring to work as advertised'
      'FLIP_VERT is kludged by flipping the BMI.biHeight parameter above. The other one just doesnt work???
      'the rest of the code is setup for it, the Abs() need to be removed below to test it
      'Can someone get this to work????
      rc = StretchDIBits(DestDC, 0, 0, Abs(NewWidth), Abs(NewHeight), _
                                 0, 0, Width, Height, PixBits(0), BMI, DIB_RGB_COLORS, RasterOp)
      DestBitMap = SelectObject(DestDC, DestBitMap)   'unselect it again
      If rc = GDI_ERROR Then rc = 0 Else rc = 1
    End If
  End If
 
  DCBitMapFromImage = rc
  On Error GoTo 0
  Exit Function
  
ErrorFound:
  DCBitMapFromImage = 0
  On Error GoTo 0
End Function                          'the values NewDC and NewBitMap will contain the Image

    
'This function manipulates the image using the API functions (quite quick)
'Supports: PIC_UNMAP_COLOR, PIC_FLIP_VERT, PIC_FLIP_HORZ, PIC_IMAGE_RESIZE or PIC_IMAGE_ZOOM, PIC_INVERT_COLOR
'Also OR combinations of the above, i.e. All could be combined in one call

Public Function APIOperations(ByRef Width As Long, ByRef Height As Long, _
                              ByRef PixBits() As Byte, ByRef BPP As Long, _
                              ByRef CMap() As Byte, ByRef NMapColors As Long, _
                              ByVal OpCodes As Long, _
                              Optional ByVal OpParm1 As Long = 0, _
                              Optional ByVal OpParm2 As Long = 0) As Long
  
  Dim hdc As Long, hBitmap As Long, rc As Long
  Dim NewWidth As Long, NewHeight As Long, NewBPP As Long
  
  NewWidth = Width: NewHeight = Height: NewBPP = BPP    'assume nothing changes
  
  'ZOOM and RESIZE are mutually exclusive as they are just different ways of defining the same thing
  If (OpCodes And PIC_IMAGE_ZOOM) = PIC_IMAGE_ZOOM Then
    If OpParm1 < 0 Then
      NewWidth = (Width * 10000&) \ -OpParm1      'these are scaled parameters
    ElseIf OpParm1 > 0 Then
      NewWidth = (Width * 10000&) * OpParm1
    End If
    If OpParm2 < 0 Then
      NewHeight = (Height * 10000&) \ -OpParm2
    ElseIf OpParm2 > 0 Then
      NewHeight = (Height * 10000&) * OpParm2
    End If
    If NewWidth < 1 Then NewWidth = Width     'if zero revert
    If NewHeight < 1 Then NewHeight = Height  'if zero revert
  ElseIf (OpCodes And PIC_IMAGE_RESIZE) = PIC_IMAGE_RESIZE Then
    If OpParm1 > 0 Then NewWidth = OpParm1
    If OpParm2 > 0 Then NewHeight = OpParm2
  End If
  If NewWidth < 1 Then NewWidth = Width     'if <=zero revert
  If NewHeight < 1 Then NewHeight = Height  'if <=zero revert
  
  If (OpCodes And PIC_FLIP_HORZ) = PIC_FLIP_HORZ Then NewWidth = -NewWidth
  If (OpCodes And PIC_FLIP_VERT) = PIC_FLIP_VERT Then NewHeight = -NewHeight
  
  If (OpCodes And PIC_UNMAP_COLOR) = PIC_UNMAP_COLOR Then
    NewBPP = PIC_24BPP
  ElseIf (OpCodes And PIC_MSMAP_COLOR) = PIC_MSMAP_COLOR Then
    If OpParm1 <= PIC_1BPP Then
      OpParm1 = PIC_1BPP
    ElseIf OpParm1 <= PIC_4BPP Then
      OpParm1 = PIC_4BPP
    ElseIf OpParm1 <= PIC_8BPP Then
      OpParm1 = PIC_8BPP
    ElseIf OpParm1 <> PIC_16BPP And OpParm1 <> PIC_24BPP Then
      OpParm1 = PIC_24BPP
    End If
    NewBPP = OpParm1                       'BE careful here only 1,4,8,16,24 OK
  End If
  
  'a new DC and Bitmap are allocated in this call
  rc = DCBitMapFromImage(Width, Height, PixBits(), BPP, CMap(), NMapColors, _
                         hdc, hBitmap, OpCodes, NewWidth, NewHeight)
   
  If rc <> 0 Then
    rc = ImageFromDCBitMap(hdc, hBitmap, Abs(NewWidth), Abs(NewHeight), PixBits(), NewBPP, CMap(), NMapColors)
    If rc <> 0 Then
      Width = Abs(NewWidth)
      Height = Abs(NewHeight)
      BPP = NewBPP
    End If
  End If

  Call FreeDCandBitmap(hdc, hBitmap)
  APIOperations = rc
End Function


Public Function ImageFromDCBitMap(ByVal SrcDC As Long, ByVal SrcBitmap As Long, _
                                  ByVal Width As Long, ByVal Height As Long, _
                                  ByRef PixBits() As Byte, ByVal BitsPerPixel As Long, _
                                  ByRef CMap() As Byte, ByRef NMapColors As Long) As Long
  Dim BMI As BITMAPINFO
  Dim rc  As Long, i As Long, p As Long

  rc = 0
  On Error GoTo ErrorFound
  
  '--------------------------------------------------------------------------------------------------------
  '-----------------------------------  BitMap CAPTURE INTO STORAGE ---------------------------------------
  '--------------------------------------------------------------------------------------------------------
  NMapColors = 2& ^ BitsPerPixel: If BitsPerPixel >= PIC_16BPP Then NMapColors = 0
 
  With BMI.bmiHeader
    .biSize = 40                              'sizeof(BITMAPINFOHEADER
    .biWidth = Width                          '{width of the bitmapclip}
    .biHeight = Height                        '{height of the bitmapclip} make sure its top to bottom
    .biPlanes = 1
    .biBitCount = BitsPerPixel                '{desired color resolution (1, 4, 8, or 24)}
    .biClrUsed = NMapColors
    .biCompression = BI_RGB
    .biSizeImage = BMPRowModulo(Width, BitsPerPixel) * Abs(Height)
  End With
  ReDim PixBits(0 To BMI.bmiHeader.biSizeImage - 1) As Byte       'the real image size
      
  rc = GetDIBits(SrcDC, SrcBitmap, 0, Abs(Height), PixBits(0), BMI, DIB_RGB_COLORS)
  
  If rc <> 0 Then                             'we now have the whole thing captured
    If BitsPerPixel < PIC_16BPP Then          'we have a Colormap to use
      p = 0
      ReDim CMap(0 To 3 * NMapColors - 1) As Byte
      For i = 1 To NMapColors
        With BMI.bmiColors(i)
          CMap(p) = .rgbRed:   p = p + 1
          CMap(p) = .rgbGreen: p = p + 1
          CMap(p) = .rgbBlue:  p = p + 1
        End With
      Next
    End If
  End If
  
  ImageFromDCBitMap = rc
  On Error GoTo 0
  Exit Function
  
ErrorFound:
  ImageFromDCBitMap = 0
  On Error GoTo 0
End Function
  
  '--------------------------------------------------------------------------------------------------------
  '================= This takes a GDI hDC, makes a copy of the bitmap and saves it to DIBits() =======
  '--------------------------------------------------------------------------------------------------------
  'any gDI Device Context should work here from eg. Form, PictureBox, Image, Control...
  'Will not work with Printer.object (someone explain why to me please)
  'It can fail for lack of memory   0=failure,1=success, NO error checking on other inputs (shielded by Class)
  '--------------------------------------------------------------------------------------------------------

Public Function ImageFromDCClip(hdc As Long, _
                                ByVal PicType As Long, ByRef PicState As Long, _
                                ByRef Width As Long, ByRef Height As Long, _
                                ByRef PixBits() As Byte, ByVal IntBPP As Long, _
                                ByRef CMap() As Byte, ByRef NMapColors As Long, _
                                ByVal ClipX0 As Long, ByVal ClipY0 As Long, _
                                ByVal ClipX1 As Long, ByVal ClipY1 As Long) As Long
                                  
  Dim hwBitMap As Long, hwMemDC As Long         'BITMAP,hDC2
  Dim rc As Long, i As Long, p As Long
  Dim ClipWidth As Long, ClipHeight As Long
  
  rc = 0
  On Error GoTo ErrorFound
  '--------------------------------------  PARAMETER VALIDATION  ------------------------------------------
  
  PicState = PicState And IS_VALID_PIPELINE     'clear all other bits
  
  'NOTE: no validation of clipping rectangle in this backend routine - MAKE SURE ITS RIGHT
  ClipWidth = ClipX1 - ClipX0 + 1
  ClipHeight = ClipY1 - ClipY0 + 1
  If ClipWidth < 1 Or ClipHeight < 1 Then GoTo ErrorFound     'Que??? its a bit small or overlapped
  
  PicState = PicState Or IS_VALID_CLIP
  
  '--------------------------------------------------------------------------------------------------------
  '--------------------------------------  DIB CAPTURE INTO STORAGE ---------------------------------------
  '--------------------------------------------------------------------------------------------------------
  rc = CloneDCandBitmap(hdc, ClipWidth, ClipHeight, hwMemDC, hwBitMap)
  If rc <> 0 Then
  
    hwBitMap = SelectObject(hwMemDC, hwBitMap)    'we get a clip of the right size to fill the cloned map
    Call BitBlt(hwMemDC, 0, 0, ClipWidth, ClipHeight, hdc, ClipX0, ClipY0, SRCCOPY)
    hwBitMap = SelectObject(hwMemDC, hwBitMap)    'hwBitMap is now a clipped copy, and unselected
    
    If PicType <> PIC_BMP Then
      rc = ImageFromDCBitMap(hwMemDC, hwBitMap, ClipWidth, -ClipHeight, PixBits(), IntBPP, CMap(), NMapColors)
      If rc <> 0 Then PicState = PicState Or IS_TOP_TO_BOTTOM   'mark it right way up
    Else
      rc = ImageFromDCBitMap(hwMemDC, hwBitMap, ClipWidth, ClipHeight, PixBits(), IntBPP, CMap(), NMapColors)
    End If
    
    If rc <> 0 Then
      Width = ClipWidth
      Height = ClipHeight
      If NMapColors <> 0 Then PicState = PicState Or IS_CMAPPED
    End If
  End If
  
  Call FreeDCandBitmap(hwMemDC, hwBitMap)
  ImageFromDCClip = rc
  On Error GoTo 0
  Exit Function
  
ErrorFound:
  Call FreeDCandBitmap(hwMemDC, hwBitMap)
  ImageFromDCClip = 0
  On Error GoTo 0
End Function

