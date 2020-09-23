Attribute VB_Name = "ImageProcs"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please
'ImageProcs.bas

' New processing options should be added in this file

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'This routine takes an image and flips it about the Y axis ie. Mirrors in the X axis
Public Function FlipImageHorz(ByVal Width As Long, ByVal Height As Long, _
                              ByRef PixBits() As Byte, _
                              ByVal BitsPerPixel As Long) As Long  'deals with all MS formats

  Dim wk() As Byte, maskw As Integer, maskx As Integer, c As Integer
  Dim x As Long, y As Long, z As Long, w As Long, RowMod As Long, i As Integer
  
   ' Scan the PixBits() and reverse all pixels in each row of the matrix
  RowMod = ((UBound(PixBits) - LBound(PixBits) + 1) \ Height)           'the byte width of a row
  ReDim wk(0 To RowMod - 1)
  For y = 0 To Height - 1                                      'assume bmap is right way up
    w = ((Width - 1) * BitsPerPixel) \ 8&                      'this is byte where the rightmost pixel is
    x = y * RowMod                                             'this is the byte where we will put the result
    Call CopyMemory(wk(0), PixBits(x), RowMod)                 'our working copy
    
    Select Case BitsPerPixel
      Case PIC_1BPP:   '8 pixels per byte
        maskw = 2 ^ (7 - (Width - 1 - 8 * w)): maskx = 128
        Do
          c = 0
          Do
            If (wk(w) And maskw) <> 0 Then c = c Or maskx
            maskw = maskw + maskw
            If maskw > 128 Then maskw = 1: w = w - 1
            maskx = maskx \ 2
          Loop Until maskx = 0
          PixBits(x) = c
          x = x + 1
          maskx = 128
        Loop Until w < 1
        c = 0
        Do      'there are still some bits in byte0 left
          If (wk(w) And maskw) <> 0 Then c = c Or maskx
          maskw = maskw + maskw
          maskx = maskx \ 2
        Loop Until maskw > 128
        PixBits(x) = c
        
      Case PIC_4BPP:   '2 pixels per byte
        If (Width And 1) = 1 Then                              ' ab,cd,ef,g0  -> gf,ed,cb,a0
          Do
            PixBits(x) = (wk(w) And &HF0) Or (wk(w - 1) And &HF)
            x = x + 1: w = w - 1
          Loop Until w < 1
          PixBits(x) = wk(w) And &HF0
        Else                                                   ' ab,cd,ef,gh  -> hg,fe,dc,ba
          Do
            PixBits(x) = (16 * (wk(w) And &HF)) Or ((wk(w) And &HF0) \ 16)
            x = x + 1: w = w - 1
          Loop Until w < 0
        End If
        
      Case PIC_8BPP:   '1 pixel per byte    'The easiest one
        Do
          PixBits(x) = wk(w): x = x + 1: w = w - 1
        Loop Until w < 0
        
      Case PIC_16BPP:  ' One case for 16-bit DIBs
        Do
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w - 3
        Loop Until w < 0
    
      Case PIC_24BPP:  ' Another for 24-bit DIBs
        Do
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w - 5
        Loop Until w < 0
    
      Case PIC_32BPP:  ' And another for 32-bit DIBs
        Do
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w + 1
          PixBits(x) = wk(w): x = x + 1: w = w - 7
        Loop Until w < 0
    End Select
  Next y
    
'  MsgBox "OK FlipImageHorz"
  FlipImageHorz = 1
End Function
