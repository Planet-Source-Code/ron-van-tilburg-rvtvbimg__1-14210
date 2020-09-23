Attribute VB_Name = "BMPSave"

Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'BMPSave.bas
' for capturing a memory dc and saving pictures as bitmaps

'==================================== SAVING A BMP ==============================================================

Public Function SaveBMP(Path As String, _
                        ByVal Width As Long, ByVal Height As Long, _
                        ByRef PixBits() As Byte, ByVal BitsPerPixel As Long, _
                        ByRef CMap() As Byte, ByVal NMapColors As Long)
                        
  Dim BMFH As BITMAPFILEHEADER
  Dim BMIH As BITMAPINFOHEADER
  Dim BMCM() As RGBQUAD
  Dim i As Long, p As Long
  
  On Error GoTo WriteError
      
  With BMIH
    .biSize = 40                              'sizeof(BITMAPINFOHEADER
    .biWidth = Width                          '{width of the bitmapclip}
    .biHeight = Height                        '{height of the bitmapclip} make sure its bottom to top
    .biPlanes = 1
    .biBitCount = BitsPerPixel                '{desired color resolution (1, 4, 8, or 24)}
    .biClrUsed = NMapColors
    .biCompression = BI_RGB
    .biSizeImage = BMPRowModulo(Width, BitsPerPixel) * Height
  End With
     
  With BMFH
    .bfType = &H4D42                                                          'as integer                   '
    .bfSize = 14& + 40& + 4& * NMapColors + BMIH.biSizeImage                  'sizeof file  As Long
    .bfReserved1 = 0                                                          'As Integer
    .bfReserved2 = 0                                                          'As Integer
    .bfOffBits = 14& + 40& + 4& * NMapColors                                  'As Long
  End With
  
  Open Path For Binary Access Write As #99
  Put #99, , BMFH
  Put #99, , BMIH
  If BitsPerPixel <= PIC_8BPP Then    'write out the colormap
    ReDim BMCM(1 To NMapColors)
    p = 0
    For i = 1 To NMapColors
      With BMCM(i)
        .rgbRed = CMap(p): p = p + 1
        .rgbGreen = CMap(p): p = p + 1
        .rgbBlue = CMap(p): p = p + 1
        .rgbReserved = 0
      End With
    Next
    Put #99, , BMCM()
  End If
  Put #99, , PixBits()
  Close #99
  SaveBMP = 1
  On Error GoTo 0
  Exit Function
  
WriteError:
  SaveBMP = 0
  On Error GoTo 0
End Function

'=============================================================================================================
'This is the correct RowModulo for a BMP file

Public Function BMPRowModulo(ByVal Width As Long, ByVal BitsPerPixel As Long) As Long
  BMPRowModulo = (((Width * BitsPerPixel) + 31&) And Not 31&) \ 8&
End Function

'==================================== LOADING A BMP ==========================================================

Public Function LoadBMP(Path As String, _
                        ByRef Width As Long, ByRef Height As Long, _
                        ByRef PixBits() As Byte, ByRef BitsPerPixel As Long, _
                        ByRef CMap() As Byte, ByRef NMapColors As Long)
                        
  Dim BMFH As BITMAPFILEHEADER
  Dim BMIH As BITMAPINFOHEADER
  Dim BMCM() As RGBQUAD
  Dim i As Long, p As Long
  
  On Error GoTo ReadError
      
  'I assume that we only have a 'Modern' bitmap ie BMIH.bisize=40, so that we have QUADRGB colors
  
  Open Path For Binary Access Read As #99
  Get #99, , BMFH
  If BMFH.bfType <> &H4D42 Then GoTo ReadError 'not a BMP
  Get #99, , BMIH
  
  With BMIH
    Width = .biWidth
    Height = .biHeight
    BitsPerPixel = .biBitCount
    If .biCompression <> BI_RGB Then GoTo ReadError   'I havent a clue how to deal with other ones
  End With
  
  Select Case BitsPerPixel:
    Case 16, 24:
      NMapColors = 0  'do nothing, we assume this is not mapped
    
    Case 8, 4, 1:
      NMapColors = 2 ^ BitsPerPixel
      ReDim BMCM(1 To NMapColors)
      ReDim CMap(0 To 3 * NMapColors - 1)
      Get #99, , BMCM()
      p = 0
      For i = 1 To NMapColors
        With BMCM(i)
          CMap(p) = .rgbRed: p = p + 1
          CMap(p) = .rgbGreen: p = p + 1
          CMap(p) = .rgbBlue: p = p + 1
        End With
      Next
  End Select
  
  'get the bitdata
  ReDim PixBits(0 To BMPRowModulo(Width, BitsPerPixel) * Height - 1)
  Seek #99, 1 + BMFH.bfOffBits
  Get #99, , PixBits()
  Close #99
  LoadBMP = 1
  On Error GoTo 0
  Exit Function       'if height<0 then the bitmap is top-down
  
ReadError:
  On Error Resume Next
  Erase PixBits(), CMap()
  NMapColors = 0
  Width = 0
  Height = 0
  LoadBMP = 0
  BitsPerPixel = 0
  On Error GoTo 0
End Function

'This routine takes an image in 8BPP format which needs to be packed into 1BPP or 4BPP format
Public Function PackBMPImage(ByVal Width As Long, ByVal Height As Long, _
                             ByRef PixBits() As Byte, _
                             ByRef IntBPP As Long, _
                             ByVal ReqBPP As Long) As Long  'deals with all MS formats

  Dim x As Long, y As Long, z As Long, w As Long, maskx As Long
  Dim RowMod As Long, NewRowMod As Long, p As Long, i As Integer, c As Integer
  
  If IntBPP <> PIC_8BPP _
  Or (ReqBPP <> PIC_1BPP And ReqBPP <> PIC_4BPP) Then PackBMPImage = 0: Exit Function
   
  ' Scan the PixBits() and pack all pixels in each row of the matrix
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)  'the byte width of a given row
  NewRowMod = BMPRowModulo(Width, ReqBPP)                      'the byte width of a row

  For y = 0 To Height - 1                                      'assume bmap is right way up
    x = y * RowMod                                             'this is the byte where we will put the result
    p = y * NewRowMod
    i = 0
    Select Case ReqBPP
      Case PIC_1BPP:   '8 pixels per byte
        maskx = 128: c = 0
        For i = 0 To Width - 1
          If PixBits(x) <> 0 Then c = c Or maskx
          maskx = maskx \ 2
          If maskx = 0 Then
            PixBits(p) = c
            p = p + 1
            maskx = 128
            c = 0
          End If
          x = x + 1
        Next
        If maskx <> 0 Then PixBits(p) = c         'the remaining bits
        
      Case PIC_4BPP:   '2 pixels per byte
        For i = 0 To Width - 1
          If (x And 1) = 0 Then
            PixBits(p) = PixBits(x) * 16          'this guarantees it being fill in on oddlength rows
          Else
            PixBits(p) = PixBits(p) Or PixBits(x)
            p = p + 1
          End If
          x = x + 1
        Next
    End Select
  Next y
 
  IntBPP = ReqBPP
  
'  MsgBox "OK PackBMPImage"
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
  PackBMPImage = 1
End Function


