Attribute VB_Name = "CRemap"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'CRemap - Remapping colors by various methods, Simple, Dithered, FS Dithers, also ColorMap Mutations

'these routines make extensive use of HistCmap.bas for three functions,
'  InitColorMappingHistogram()
'  FreeColorMappingHistogram()
'  MatchColorbyHistogram()

'/*================== GENERAL DITHERING ====================================*/

Dim DMat(0 To 63) As Integer        'a dither matrix against which values are compared
Dim delv As Long                    'a small measure by which to vary pixels in colour dithering

Private Sub SetDitherMatrix(ByVal WhichDither As Long)
  Dim zdmap(), i As Integer
  
  Select Case WhichDither
    Case PIC_DITHER_BIN:                                '/* Binary */
      zdmap = Array(0, 252, 0, 252, 0, 252, 0, 252, _
                    252, 0, 252, 0, 252, 0, 252, 0, _
                    0, 252, 0, 252, 0, 252, 0, 252, _
                    252, 0, 252, 0, 252, 0, 252, 0, _
                    0, 252, 0, 252, 0, 252, 0, 252, _
                    252, 0, 252, 0, 252, 0, 252, 0, _
                    0, 252, 0, 252, 0, 252, 0, 252, _
                    252, 0, 252, 0, 252, 0, 252, 0)

      '/*------------------------------------------------------------------------*/
    Case PIC_DITHER_HTC:                                '/* Printer HalfTone */
      zdmap = Array(80, 192, 224, 48, 88, 200, 232, 56, _
                   128, 0, 96, 160, 136, 8, 104, 168, _
                   208, 32, 64, 240, 216, 40, 72, 248, _
                   112, 176, 144, 16, 120, 184, 152, 24, _
                    92, 204, 236, 60, 84, 196, 228, 52, _
                   140, 12, 108, 172, 132, 4, 100, 164, _
                   220, 44, 76, 252, 212, 36, 68, 244, _
                   124, 188, 156, 28, 116, 180, 148, 20)

      '/*-----------------------------------------------------------------------*/
    Case PIC_DITHER_FDIAG:                                     'forward diagonal
      zdmap = Array(0, 112, 192, 240, 8, 120, 200, 248, _
                   96, 16, 128, 208, 104, 24, 136, 216, _
                  176, 80, 32, 144, 184, 88, 40, 152, _
                  224, 160, 64, 48, 232, 168, 72, 56, _
                   12, 124, 204, 252, 4, 116, 196, 244, _
                  108, 28, 140, 220, 100, 20, 132, 212, _
                  188, 92, 44, 156, 180, 84, 36, 148, _
                  236, 172, 76, 60, 228, 164, 68, 52)

      '/*-----------------------------------------------------------------------*/
    Case PIC_DITHER_BDIAG:                                     'backward diagonal
      zdmap = Array(240, 208, 144, 48, 248, 216, 152, 56, _
                    192, 128, 32, 64, 200, 136, 40, 72, _
                    112, 16, 80, 160, 120, 24, 88, 168, _
                      0, 96, 176, 224, 8, 104, 184, 232, _
                    252, 220, 156, 60, 244, 212, 148, 52, _
                    204, 140, 44, 76, 196, 132, 36, 68, _
                    124, 28, 92, 172, 116, 20, 84, 164, _
                     12, 108, 188, 236, 4, 100, 180, 228)

      '/*-----------------------------------------------------------------------*/
    Case PIC_DITHER_HORZ:                                     'horizontal
      zdmap = Array(0, 32, 64, 96, 8, 40, 72, 104, _
                  128, 160, 192, 224, 136, 168, 200, 232, _
                   80, 112, 16, 48, 88, 120, 24, 56, _
                  208, 240, 144, 176, 216, 248, 152, 184, _
                   12, 44, 76, 108, 4, 36, 68, 100, _
                  140, 172, 204, 236, 132, 164, 196, 228, _
                   92, 124, 28, 60, 84, 116, 20, 52, _
                  220, 252, 156, 188, 212, 244, 148, 180)

      '/*-----------------------------------------------------------------------*/
    Case PIC_DITHER_VERT:                                     'vertical
      zdmap = Array(0, 128, 80, 208, 8, 136, 88, 216, _
                   32, 160, 112, 240, 40, 168, 120, 248, _
                   64, 192, 16, 144, 72, 200, 24, 152, _
                   96, 224, 48, 176, 104, 232, 56, 184, _
                   12, 140, 92, 220, 4, 132, 84, 212, _
                   44, 172, 124, 252, 36, 164, 116, 244, _
                   76, 204, 28, 156, 68, 196, 20, 148, _
                  108, 236, 60, 188, 100, 228, 52, 180)

      '/*-----------------------------------------------------------------------*/
    Case Else                                       'PIC_DITHER_ORD:  '/* Printer Ordered  */
      zdmap = Array(0, 128, 32, 160, 8, 136, 40, 168, _
                    192, 64, 224, 96, 200, 72, 232, 104, _
                     48, 176, 16, 144, 56, 184, 24, 152, _
                    240, 112, 208, 80, 248, 120, 216, 88, _
                     12, 140, 44, 172, 4, 132, 36, 164, _
                    204, 76, 236, 108, 196, 68, 228, 100, _
                     60, 188, 28, 156, 52, 180, 20, 148, _
                    252, 124, 220, 92, 244, 116, 212, 84)
  End Select
  For i = 0 To 63
    DMat(i) = zdmap(i)
  Next
End Sub

Private Sub set_dv(ByRef CMap() As Byte, ByVal NMapColors As Long)
  Dim i As Long, low As Long, high As Long, z As Long, p As Long
  
  low = 256: high = 0:  p = 0
  For i = 0 To NMapColors - 1
    z = RGBtoGrey(CMap(p), CMap(p + 1), CMap(p + 2))
    p = p + 3
    If z > high Then high = z
    If z < low Then low = z
  Next i
  delv = (high - low + 1) \ NMapColors
  If delv = 0 Then delv = 1
  
'  if (distrmeth==DISTRMETH_FCMAP) delv=12;
'  msgbox "delv = " & delv
End Sub
'/*=======================================================================*/

'dither into printer coordinates  Implicit FIXED_CMAP 8 color palette KRGBWCMY
Private Function dither_pixelc8(ByVal r As Integer, _
                                ByVal g As Integer, _
                                ByVal b As Integer, _
                                ByVal dval As Integer) As Integer
  
  Dim k As Integer, c As Integer, m As Integer, y As Integer, cz As Integer
  
  cz = 7
  c = 255 - r:  k = c
  m = 255 - g:  If m < k Then k = m
  y = 255 - b:  If y < k Then k = y
  
  If k > dval Then
    cz = 0
  Else
    If y > dval Then cz = cz - 4  ''/* Y */
    If m > dval Then cz = cz - 2  ''/* M */
    If c > dval Then cz = cz - 1  ''/* C */
  End If

  dither_pixelc8 = cz
End Function

'dither into Black and white    'Implicit FIXED CMAP 2 color CMap required B,W
Private Function dither_pixelbw(ByVal r As Integer, _
                                ByVal g As Integer, _
                                ByVal b As Integer, _
                                ByVal dval As Integer) As Integer
  Dim k As Integer

'  k = r
'  If k < g Then k = g
'  If k < b Then k = b
'  If (255 - k) > dval Then k = 0 Else k = 1
  If RGBtoGrey(r, g, b) > dval Then k = 1 Else k = 0
  dither_pixelbw = k
End Function

'dither into greys  (implicit FIXED CMAP n)
Private Function dither_pixelgrey(ByVal r As Integer, _
                                  ByVal g As Integer, _
                                  ByVal b As Integer, _
                                  ByVal dval As Integer) As Integer
  Dim v As Integer

  v = RGBtoGrey(r, g, b)
  If v > dval Then
    v = v - delv: If v < 0 Then v = 0
  Else
    v = v + delv: If v > 255 Then v = 255
  End If
  
  dither_pixelgrey = MatchColorbyHistogram(v, v, v)
End Function

'generalised dither function
Private Function dither_pixel(ByVal r As Integer, _
                              ByVal g As Integer, _
                              ByVal b As Integer, _
                              ByVal dval As Integer) As Integer
  Dim v As Integer

  v = RGBtoGrey(r, g, b)
  If v > dval Then
    r = r - delv: If r < 0 Then r = 0
    g = g - delv: If g < 0 Then g = 0
    b = b - delv: If b < 0 Then b = 0
  Else
    r = r + delv: If r > 255 Then r = 255
    g = g + delv: If g > 255 Then g = 255
    b = b + delv: If b > 255 Then b = 255
  End If
  
  dither_pixel = MatchColorbyHistogram(r, g, b)
End Function

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by Dithering against a a dither matrix. This minimises the error propagation of
'approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map

Public Sub DitherMapColors(ByVal Width As Long, ByVal Height As Long, _
                           ByRef PixBits() As Byte, _
                           ByVal BitsPerPixel As Long, _
                           ByRef CMap() As Byte, _
                           ByVal NMapColors As Long, _
                           ByVal CMAPMode As Long, _
                           ByVal DitherMode As Long)
                                 
  Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, NewRowMod As Long, i As Integer
  Dim r As Long, g As Long, b As Long, wColor As Long, p As Long, wmap As Long

   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Sub           'will only work for unmapped DIBs
   
   ' Initialize Remapping variables
  Call InitColorMappingHistogram(CMap(), NMapColors)
  Call SetDitherMatrix(DitherMode)
  Call set_dv(CMap(), NMapColors)
  wmap = CMAPMode
  
   ' Scan the PixBits() and build the octree
  Skip = (BitsPerPixel \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  NewRowMod = BMPRowModulo(Width, PIC_8BPP)
  For y = 0 To Height - 1                                               'assume bmap is right way up
    z = y * RowMod
    p = y * NewRowMod                                                   'this is where we put the new byte
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row
    
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          b = (wColor And &H1F&) * 8&
          g = (wColor And &H3E0&) \ 4&
          r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          r = PixBits(x + 1)
          g = PixBits(x + 2)
          b = PixBits(x + 3)
      End Select
      
      'find the nearest the color to the given color
      i = DMat(8 * (y And 7) + (x And 7))
      Select Case wmap
        Case PIC_FIXED_CMAP_BW:   i = dither_pixelbw(r, g, b, i)
        Case PIC_FIXED_CMAP_C8:   i = dither_pixelc8(r, g, b, i)
        Case PIC_FIXED_CMAP_GREY: i = dither_pixel(r, g, b, i)      'dither_pixelgrey(r, g, b, i)
        Case Else:                i = dither_pixel(r, g, b, i)
      End Select
      
      PixBits(p) = i: p = p + 1
    Next x
  Next y
    
  Call FreeColorMappingHistogram
  
  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
'  MsgBox "OK DitheredRemap"
End Sub

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by direct substitution. This should really be dithered to minimise the
'error propagation of approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map

Public Sub SimpleMapColors(ByVal Width As Long, ByVal Height As Long, _
                           ByRef PixBits() As Byte, _
                           ByVal BitsPerPixel As Long, _
                           ByRef CMap() As Byte, _
                           ByVal NMapColors As Long)
                                 
  Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, NewRowMod As Long, i As Long
  Dim r As Long, g As Long, b As Long, wColor As Long, p As Long

   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Sub           'will only work for unmapped DIBs
   
   ' Initialize Remapping variables
  Call InitColorMappingHistogram(CMap(), NMapColors)
  
   ' Scan the PixBits() and build the octree
  Skip = (BitsPerPixel \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  NewRowMod = BMPRowModulo(Width, PIC_8BPP)
  For y = 0 To Height - 1                                               'assume bmap is right way up
    z = y * RowMod
    p = y * NewRowMod                                                   'this is where we put the new byte
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row
    
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          b = (wColor And &H1F&) * 8&
          g = (wColor And &H3E0&) \ 4&
          r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          r = PixBits(x + 1)
          g = PixBits(x + 2)
          b = PixBits(x + 3)
      End Select
      
      'find the nearest the color to the given color
      PixBits(p) = MatchColorbyHistogram(r, g, b): p = p + 1
    Next x
  Next y
    
  Call FreeColorMappingHistogram
  
  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
'  MsgBox "OK SimpleRemap"
End Sub

'/*================== FLOYD-STEINBERG DITHERING ==========================*/

'This routine takes a new colormap, and the original PixBit array and remaps the array a pixel at a time
'to the best fitting color by Floyd-Steinberg Dithering. This minimises the error propagation of
'approximate colormaps. There may be unused (unmappable) colors in the colormap
'the original Pixel array will be resized to the equivalent of a 8BPP array, correctly size without padding
'every entry will be the index into the map - this routine is fairly involved

Public Sub FSDitherMapColors(ByVal Width As Long, ByVal Height As Long, _
                             ByRef PixBits() As Byte, _
                             ByVal BitsPerPixel As Long, _
                             ByRef CMap() As Byte, _
                             ByVal NMapColors As Long, _
                             ByVal CMAPMode As Long, _
                             ByVal DitherMode As Long)
  
  Dim x As Long, Skip As Long, RowMod As Long, NewRowMod As Long
  Dim r As Long, g As Long, b As Long, wColor As Long, p As Long, q As Long
        
  Const FS_SCALE As Integer = 128       'errors are scaled by this much to limit roundoff error
  
  Dim n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer
  Dim curr_rerr() As Integer, curr_gerr() As Integer, curr_berr() As Integer
  Dim next_rerr() As Integer, next_gerr() As Integer, next_berr() As Integer
  Dim sr As Long, sg As Long, sb As Long, errx As Long, aw As Long
  Dim col As Integer, limitcol As Integer, row As Integer, ind As Integer, fs_LtoR As Boolean

   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Sub           'will only work for unmapped DIBs
   
  ' error matrix
  ' FS1  original FS  7,3,5,1
  ' FS2  binary   FS  6,4,3,2   'slightly lossy but sometimes better looking
  
  Select Case DitherMode
    Case PIC_DITHER_FS1:                    'the weights for redistributing errors in colors
      n1 = 7: n2 = 3: n3 = 5: n4 = 1
    Case PIC_DITHER_FS2:
      n1 = 6: n2 = 4: n3 = 3: n4 = 2
    Case PIC_DITHER_FS3:
      n1 = 6: n2 = 2: n3 = 6: n4 = 2
  End Select
  
   ' Initialize Remapping variables
  Call InitColorMappingHistogram(CMap(), NMapColors)
  
   ' Establish the amounts needed to scan the bitmap
  Skip = (BitsPerPixel \ 8)                                               'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)             'the byte width of a row
  NewRowMod = BMPRowModulo(Width, PIC_8BPP)                               'the RowMod after Mapping
  
  '/* Initialize Floyd-Steinberg error vectors. */

  aw = (Width + 2)                                    ' sizeof(short);
  
  'Establish current error arrays
  
  ReDim curr_rerr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);
  ReDim curr_gerr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);
  ReDim curr_berr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);

  fs_LtoR = True             'ie moving Left to Right, FALSE is moving right to left

  For row = 0 To Height - 1
  
    'Clear Next Error Arrays
    ReDim next_rerr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);
    ReDim next_gerr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);
    ReDim next_berr(0 To aw - 1) As Integer             ' = (short *)Calloc(sb);

    If fs_LtoR Then
      col = 0: limitcol = Width
    Else
      col = Width - 1: limitcol = -1
    End If
   
    'we need to be cunning about this overwriting trick here. We have a serpentine movement LtoR then RtoL
    'on the first row we will at worst (16bit case) use half the bytes of the first original pixel row
    'on the second row we will therefore start writing 1 byte before the last pixel to be used on the first
    'RtoL pass - hence we havent clobbered anything we need later - I just thought Id share that :-)
    
    p = row * NewRowMod + col       'where we will start putting the row of mapped bytes
    x = row * RowMod + col * Skip   'where the first pixel will be - assume bmap is right way up
    
    Do      'get a pixel
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&
          b = (wColor And &H1F&) * 8&
          g = (wColor And &H3E0&) \ 4&
          r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          r = PixBits(x + 1)
          g = PixBits(x + 2)
          b = PixBits(x + 3)
      End Select

      '/* Use Floyd-Steinberg errors to adjust actual color. */

      sr = r + curr_rerr(col + 1) \ FS_SCALE
      sg = g + curr_gerr(col + 1) \ FS_SCALE
      sb = b + curr_berr(col + 1) \ FS_SCALE
      
      If sr < 0 Then sr = 0 Else If sr > 255 Then sr = 255
      If sg < 0 Then sg = 0 Else If sg > 255 Then sg = 255
      If sb < 0 Then sb = 0 Else If sb > 255 Then sb = 255
      
      '/* find best match for this color */

      ind = MatchColorbyHistogram(sr, sg, sb)
      q = 3 * ind
      
      '/* Propagate Floyd-Steinberg error terms. */

      If fs_LtoR Then
        errx = (sr - CMap(q)) * FS_SCALE: q = q + 1
        curr_rerr(col + 2) = curr_rerr(col + 2) + (errx * n1) \ 16
        next_rerr(col) = next_rerr(col) + (errx * n2) \ 16
        next_rerr(col + 1) = next_rerr(col + 1) + (errx * n3) \ 16
        next_rerr(col + 2) = next_rerr(col + 2) + (errx * n4) \ 16

        errx = (sg - CMap(q)) * FS_SCALE: q = q + 1
        curr_gerr(col + 2) = curr_gerr(col + 2) + (errx * n1) \ 16
        next_gerr(col) = next_gerr(col) + (errx * n2) \ 16
        next_gerr(col + 1) = next_gerr(col + 1) + (errx * n3) \ 16
        next_gerr(col + 2) = next_gerr(col + 2) + (errx * n4) \ 16

        errx = (sb - CMap(q)) * FS_SCALE: q = q + 1
        curr_berr(col + 2) = curr_berr(col + 2) + (errx * n1) \ 16
        next_berr(col) = next_berr(col) + (errx * n2) \ 16
        next_berr(col + 1) = next_berr(col + 1) + (errx * n3) \ 16
        next_berr(col + 2) = next_berr(col + 2) + (errx * n4) \ 16
      Else
        errx = (sr - CMap(q)) * FS_SCALE: q = q + 1
        curr_rerr(col) = curr_rerr(col) + (errx * n1) \ 16
        next_rerr(col + 2) = next_rerr(col + 2) + (errx * n2) \ 16
        next_rerr(col + 1) = next_rerr(col + 1) + (errx * n3) \ 16
        next_rerr(col) = next_rerr(col) + (errx * n4) \ 16

        errx = (sg - CMap(q)) * FS_SCALE: q = q + 1
        curr_gerr(col) = curr_gerr(col) + (errx * n1) \ 16
        next_gerr(col + 2) = next_gerr(col + 2) + (errx * n2) \ 16
        next_gerr(col + 1) = next_gerr(col + 1) + (errx * n3) \ 16
        next_gerr(col) = next_gerr(col) + (errx * n4) \ 16

        errx = (sb - CMap(q)) * FS_SCALE: q = q + 1
        curr_berr(col) = curr_berr(col) + (errx * n1) \ 16
        next_berr(col + 2) = next_berr(col + 2) + (errx * n2) \ 16
        next_berr(col + 1) = next_berr(col + 1) + (errx * n3) \ 16
        next_berr(col) = next_berr(col) + (errx * n4) \ 16
      End If

      PixBits(p) = ind  'yes we can finally store something, the rest is housekeeping

      If (fs_LtoR) Then
        col = col + 1: p = p + 1: x = x + Skip
      Else
        col = col - 1: p = p - 1: x = x - Skip
      End If
    Loop While col <> limitcol
       
    'copy the error arrays up for next line
    curr_rerr() = next_rerr()
    curr_gerr() = next_gerr()
    curr_berr() = next_berr()

    fs_LtoR = Not fs_LtoR     'and reverse direction
  Next row
  
  'all dynamically assigned arrays die on exit
  Call FreeColorMappingHistogram
  
  'OK everything is now mapped so lets resize the PixBits array
  ReDim Preserve PixBits(0 To NewRowMod * Height - 1) As Byte
'  MsgBox "OK FSDitherRemap"
End Sub

'a generic color to grey function
Public Function RGBtoGrey(ByVal r As Long, ByVal g As Long, ByVal b As Long) As Long
  RGBtoGrey = (18940& * r + 39125 * g + 7471& * b) \ 65536
End Function

Public Sub CMaptoGrey(ByRef CMap() As Byte)  'change the RGB palette to CIY greys (it better be a multiple of 3 long)
  Dim i As Long, z As Long
  
  For i = LBound(CMap) To UBound(CMap) Step 3
    z = (18940& * CMap(i) + 39125 * CMap(i + 1) + 7471& * CMap(i + 2)) \ 65536
    CMap(i) = z: CMap(i + 1) = z: CMap(i + 2) = z
  Next
End Sub

Public Sub GenFixedMap(ByRef CMap() As Byte, ByRef NMapColors As Long, ByRef BitsPerPixel As Long, CMAPMode As Long)
  Dim i As Long, fcm() As Variant
  
  Select Case CMAPMode
    Case PIC_FIXED_CMAP_BW:
      BitsPerPixel = 1: NMapColors = 2
      ReDim CMap(0 To 5) As Byte
      CMap(3) = &HFF: CMap(4) = &HFF: CMap(5) = &HFF
    
    Case PIC_FIXED_CMAP_C4:
      BitsPerPixel = 2: NMapColors = 4
      ReDim CMap(0 To 11) As Byte
      fcm() = Array(0, 0, 0, 0, 255, 255, 255, 0, 255, 255, 255, 0) 'KCMY
      For i = 0 To 11
        CMap(i) = fcm(i)
      Next
    Case PIC_FIXED_CMAP_VGA:
      BitsPerPixel = 4: NMapColors = 16
      ReDim CMap(0 To 47) As Byte
      fcm() = Array(0, 0, 0, 128, 0, 0, 0, 128, 0, 128, 128, 0, _
                    0, 0, 128, 128, 0, 128, 0, 128, 128, 192, 192, 192, _
                    128, 128, 128, 255, 0, 0, 0, 255, 0, 255, 255, 0, _
                    0, 0, 255, 255, 0, 255, 0, 255, 255, 255, 255, 255) 'KRGYBMC/2,3W/4,W/2,RGYBMCW
      For i = 0 To 47
        CMap(i) = fcm(i)
      Next
    Case PIC_FIXED_CMAP_INET:
      BitsPerPixel = 8: NMapColors = 256: Call GenCMap(CMap(), NMapColors, 5, 5, 5)   '216
   
    Case PIC_FIXED_CMAP_C8:
      BitsPerPixel = 3: NMapColors = 8:  Call GenCMap(CMap(), NMapColors, 1, 1, 1)  '8
    
    Case PIC_FIXED_CMAP_C16:
      BitsPerPixel = 4: NMapColors = 16: Call GenCMap(CMap(), NMapColors, 1, 2, 1, 64, 128, 192, 224) '12 +4
    
    Case PIC_FIXED_CMAP_C32:
      BitsPerPixel = 5: NMapColors = 32:  Call GenCMap(CMap(), NMapColors, 2, 2, 2, 64, 96, 160, 192, 224) '27+5
    
    Case PIC_FIXED_CMAP_C64:
      BitsPerPixel = 6: NMapColors = 64:  Call GenCMap(CMap(), NMapColors, 3, 3, 3)    '64
    
    Case PIC_FIXED_CMAP_C128:
      BitsPerPixel = 7: NMapColors = 128: Call GenCMap(CMap(), NMapColors, 4, 4, 4, 96, 160, 224) '125+3
    
    Case PIC_FIXED_CMAP_C256:
      BitsPerPixel = 8: NMapColors = 256: Call GenCMap(CMap(), NMapColors, 5, 6, 5)   '252
    
    Case PIC_FIXED_CMAP_MS256:
      BitsPerPixel = 8: NMapColors = 256: Call GenCMap(CMap(), NMapColors, 7, 7, 3)   '256
  End Select
End Sub
        
Public Sub GenGreyMap(ByRef CMap() As Byte, ByRef NMapColors As Long, ByRef BitsPerPixel As Long)
  Dim i As Long, z As Long
  
  NMapColors = 2 ^ BitsPerPixel
  ReDim CMap(0 To 3 * NMapColors - 1) As Byte
  For i = 0 To NMapColors - 1
    z = (256 * i) \ (NMapColors - 1): If z > 255 Then z = 255
    CMap(3 * i) = z: CMap(3 * i + 1) = z: CMap(3 * i + 2) = z
  Next
End Sub
        
Public Sub GenUserCMap(ByRef CMap() As Byte, ByRef NMapColors As Long, ByRef UserCMap() As Byte)
  Dim i As Long
  
  ReDim CMap(0 To 3 * NMapColors - 1) As Byte
  For i = 0 To 3 * NMapColors - 1
    CMap(i) = UserCMap(i)
  Next
End Sub

Private Sub GenCMap(ByRef CMap() As Byte, ByVal NMapColors As Long, _
                    ByVal nr As Integer, ByVal ng As Integer, ByVal nb As Integer, _
                    ParamArray Greys())
  
  Dim p As Integer, r As Integer, g As Integer, b As Integer, sr As Integer, sg As Integer, sb As Integer
  
  ReDim CMap(0 To 3 * NMapColors - 1) As Byte
  p = 0
  For b = 0 To nb
    sb = (b * 256) \ nb: If sb > 255 Then sb = 255
    For g = 0 To ng
      sg = (g * 256) \ ng: If sg > 255 Then sg = 255
      For r = 0 To nr
        sr = (r * 256) \ nr: If sr > 255 Then sr = 255
        CMap(p) = sr: p = p + 1
        CMap(p) = sg: p = p + 1
        CMap(p) = sb: p = p + 1
      Next
    Next
  Next
  For g = LBound(Greys) To UBound(Greys)
    CMap(p) = Greys(g): p = p + 1
    CMap(p) = Greys(g): p = p + 1
    CMap(p) = Greys(g): p = p + 1
  Next
End Sub

Public Function ShrinkCMap(PixBits() As Byte, ByVal PixelWidth As Long, _
                              CMap() As Byte, ByVal BitsPerPixel As Long) As Long
  Dim i As Long, j As Long, k As Long, nc As Long, nnc As Long
  Dim Idx() As Integer, cc() As Long
  
  ShrinkCMap = BitsPerPixel
  nc = (UBound(CMap) - LBound(CMap) + 1) \ 3      'the current number of colours in CMAP
  ReDim Idx(0 To nc - 1) As Integer               'the indices of old and new pixels
  ReDim cc(0 To nc - 1) As Long                   'the colour count of colours used
  
  'count the number of each sort of pixel
  For i = LBound(PixBits) To UBound(PixBits)  'PixelWidth is 4 or 8 only
    If PixelWidth = 8 Then
      j = PixBits(i):                        cc(j) = cc(j) + 1
    Else
      j = (PixBits(i) And &HF0&) \ 16&:      cc(j) = cc(j) + 1
      j = (PixBits(i) And &HF&):             cc(j) = cc(j) + 1
    End If
  Next

  'now set up the idx array, newcolor=idx(oldcolor)
  j = 0
  For i = 0 To nc - 1
    If cc(i) > 0 Then     'its been used
      Idx(i) = j          'and is kept
      j = j + 1
    Else
      Idx(i) = -1
    End If
  Next
  
  nnc = j
  If nnc <= (nc \ 2) Then     'we did reduce the number of colors enough to gain something
    For i = 0 To nc - 1       'so adjust the cmap
      j = 3 * Idx(i)
      k = 3 * i
      If j >= 0 And j < k Then       'move it up
        CMap(j) = CMap(k)
        CMap(j + 1) = CMap(k + 1)
        CMap(j + 2) = CMap(k + 2)
      End If
    Next
    i = 1: j = 0              'find the new palette size
    Do
      j = j + 1
      i = i + i
    Loop Until i > nnc
    ReDim Preserve CMap(0 To 3 * i - 1) As Byte   '2^j colours
    ShrinkCMap = j                                'New Bits Per Pixel
    
    'now the jolly task of remapping the Pixels
    For i = LBound(PixBits) To UBound(PixBits)  'PixelWidth is 4 or 8 only
      If PixelWidth = 8 Then
        PixBits(i) = Idx(PixBits(i))
      Else
        j = Idx((PixBits(i) And &HF0&) \ 16&)
        k = Idx((PixBits(i) And &HF&))
        PixBits(i) = j * 16 + k           'the new value
      End If
    Next
  End If
End Function

'This routine takes a BMP and turns all its pixels into grey shades
Public Sub MakePixelsGrey(ByVal Width As Long, ByVal Height As Long, _
                          ByRef PixBits() As Byte, ByVal BitsPerPixel As Long)
                                 
  Dim x As Long, y As Long, z As Long, w As Long, Skip As Long, RowMod As Long, i As Long
  Dim r As Long, g As Long, b As Long, wColor As Long, p As Long, v As Long

   If BitsPerPixel <> PIC_16BPP _
  And BitsPerPixel <> PIC_24BPP _
  And BitsPerPixel <> PIC_32BPP Then Exit Sub           'will only work for unmapped DIBs
  
  Skip = (BitsPerPixel \ 8)                                             'size of a pixel in bytes
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  
  For y = 0 To Height - 1                                               'assume bmap is right way up
    p = y * RowMod                                                      'this is where we put the new byte
    z = y * RowMod
    w = z + Skip * (Width - 1)
    For x = z To w Step Skip        'pixel 0,1,2,3 in a row
    
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          wColor = PixBits(x) + PixBits(x + 1) * 256&     'LOW-HIGH
          b = (wColor And &H1F&) * 8&
          g = (wColor And &H3E0&) \ 4&
          r = (wColor And &H7C00&) \ 128&
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          b = PixBits(x)
          g = PixBits(x + 1)
          r = PixBits(x + 2)
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          r = PixBits(x + 1)
          g = PixBits(x + 2)
          b = PixBits(x + 3)
      End Select
      
      v = RGBtoGrey(r, g, b)
      
      Select Case BitsPerPixel
        Case PIC_16BPP:  ' One case for 16-bit DIBs
          v = v \ 8&
          wColor = (v * 32& + v) * 32& + v
          PixBits(p) = wColor And &HFF: p = p + 1
          PixBits(p) = (wColor And &HFF00) \ 256&: p = p + 1
    
        Case PIC_24BPP:  ' Another for 24-bit DIBs
          PixBits(p) = v: p = p + 1
          PixBits(p) = v: p = p + 1
          PixBits(p) = v: p = p + 1
    
        Case PIC_32BPP:  ' And another for 32-bit DIBs
          p = p + 1
          PixBits(p) = v: p = p + 1
          PixBits(p) = v: p = p + 1
          PixBits(p) = v: p = p + 1
      End Select
    
    Next x
  Next y
  
'  MsgBox "OK MakePixelsGrey"
End Sub


