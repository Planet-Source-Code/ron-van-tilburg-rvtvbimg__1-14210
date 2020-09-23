Attribute VB_Name = "GIFSave"
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

' GIFSave.bas  -  master file for writing GIF files
' from the C copyright ©1997 Ron van Tilburg 25.12.1997
' VB copyright ©2000 Ron van Tilburg 24.12.2000     'what xmas holidays are good for <:-)
' and copyrights of the original C code from which this is derived are given in the body
' Documentation of GIF structures is from the GIF standard as attached as html documents
' All copyrights applying there continue to apply

' Unisys Corp believes it has the Copyright on all LZW algorithms for GIF files. If it worries you then
' dont use this code. Read the HTML standards for the owner of the copyright of GIFs and its usability

' Start at the bottom of this file at the SaveGIF function and work upwards

' General Disclaimer: I think this all works ok (but it needs some exercis to prove it) but you use and
' adapt it at your own risk. However reference to my authorship would be appreciated if you use it in
' the public domain.  Ron van Tilburg Xmas 2000.

' GIF structures    (not actually Used)
Private Type GifScreenHdr
  Width As Integer
  Height As Integer
  MCR0Pix As Byte     'see below
  BC As Byte
  Aspect As Byte
End Type

Private Type GifImageHdr
  Left As Integer
  Top As Integer
  Width  As Integer
  Height As Integer
  MIPixBits As Byte
  CodeSize As Byte
End Type

'GLOBAL VRIABLES for the Encoding Routines ============================================================

'***************************************************************************'
'*  FROM GIFCOMPR.C       - GIF Image compression routines
'*
'*  Lempel-Ziv compression based on 'compress'.  GIF modifications by
'*  David Rowley (mgardi@watdcsu.waterloo.edu)
'*
'***************************************************************************/
'* an Integer must be able to hold 2**BITS values of type int, and also -1 */

Const MAXBITS      As Integer = 12              ' user settable max - bits/code
Const MAXBITSHIFT  As Integer = 2 ^ MAXBITS
Const MAXMAXCODE   As Integer = 2 ^ MAXBITS     ' should NEVER generate this code
Const HASHTABSIZE  As Integer = 5003            ' 80% occupancy

' * GIF Image compression - modified 'compress'
' *
' * Based on: compress.c - File compression ala IEEE Computer, June 1984.
' *
' * By Authors:  Spencer W. Thomas       (decvax!harpo!utah-cs!utah-gr!thomas)
' *              Jim McKie               (decvax!mcvax!jim)
' *              Steve Davies            (decvax!vax135!petsd!peora!srd)
' *              Ken Turkowski           (decvax!decwrl!turtlevax!ken)
' *              James A. Woods          (decvax!ihnp4!ames!jaw)
' *              Joe Orost               (decvax!vax135!petsd!joe)
' * VB code by   Ron van Tilburg          rivit@f1.net.au

Dim n_Bits As Integer                    ' number of bits/code
Dim MaxCode As Integer                   ' maximum code, given n_bits

'-define MAXCODE(n_bits) (((Integer)1 << (n_bits)) - 1)    '=masks(n_bits)

Dim HashTab(0 To HASHTABSIZE - 1) As Long
Dim CodeTab(0 To HASHTABSIZE - 1) As Integer

' To save much memory, we overlay the table used by compress() with those used by decompress().
' The tab_prefix table is the same size and type as the codetab.  The tab_suffix table needs
' 2**MAXBITS characters.  We get this from the beginning of HashTab.  The output stack uses the rest
' of HashTab, and contains characters.  There is plenty of room for any possible stack
' (stack used to be 8000 characters).

'-define tab_prefixof(i) CodeTabOf(i)
'-define tab_suffixof(i) ((byte*)(HashTab))[i]
'-define de_stack        ((byte*)&tab_suffixof((Integer)1<<MAXBITS))

Dim Free_Ent As Integer                      ' first unused entry

' block compression parameters -- after all codes are used up, and compression rate changes, start over.
Dim Clear_Flg As Boolean
Dim Offset    As Integer
Dim In_Count  As Long                        ' length of input
Dim Out_Count As Long                        ' - of codes output (for debugging)

' Algorithm:  use open addressing double hashing (no chaining) on the prefix code / next character
' combination.  We do a variant of Knuth's algorithm D (vol. 3, sec. 6.4) along with G. Knott's
' relatively-prime secondary probe.  Here, the modular division first probe is gives way to a faster
' exclusive-or manipulation.  Also do block compression with an adaptive reset, whereby the code table
' is cleared when the compression ratio decreases, but after the table fills.  The variable-length output
' codes are re-sized at this point, and a special CLEAR code is generated for the decompressor.  Late
' addition:  construct the table according to file size for noticeable speed improvement on small files.
' Please direct questions about this implementation to ames!jaw.

Dim g_Init_Bits As Integer
Dim ClearCode   As Integer
Dim EOFCode     As Integer

'variables for positioning and control

Dim CurX    As Integer         'current xpos
Dim CurY    As Integer         'current ypos
Dim PWidth  As Long            'the Width
Dim PHeight As Long            'the Height
Dim RowMod  As Long            'the rowModulo - BMPs can have extension bits for padding <=PWidth
Dim PixSize As Long            'the nr of bits for a given pixel

Dim Countdown As Long          'pixels left to do
Dim Pass      As Integer       'which pass in interlaced mode
Dim Interlace As Boolean       'use interlace mode
Dim FileCount As Long          'bytes output so far

Const EOF As Integer = -1     'END of input

'variables for the code accumulator (OutputCode)

Dim cur_Accum As Long
Dim cur_Bits As Integer
Dim Masks(0 To 16) As Long         'powers of 2 -1

'variables for the outputbyte accumulator

Dim A_Count As Integer      'Number of characters so far in this 'packet'
Dim Accum() As Byte         'will be max 256 bytes long, first byte is length

'======================================================================================================
'======================================================================================================
Private Sub CompressAndWriteBits(init_bits As Integer, ByRef PixBits() As Byte)

  Dim fcode As Long
  Dim i As Long, c As Long, ent As Long, disp As Long
  Dim hshift As Long, zm() As Variant

  'set up where we are starting
  i = 0
  FileCount = 0
  Pass = 0
  CurX = 0
  CurY = 0
  Countdown = PWidth * PHeight

  'set up the code accumulator
  cur_Accum = 0
  cur_Bits = 0
  zm = Array(&H0&, &H1&, &H3&, &H7&, &HF&, _
                   &H1F&, &H3F&, &H7F&, &HFF&, _
                   &H1FF&, &H3FF&, &H7FF&, &HFFF&, _
                   &H1FFF&, &H3FFF&, &H7FFF&, &HFFFF&)   'Array values of 2^N-1  N=0,1,2,,..16
  For i = 0 To 16
    Masks(i) = CLng(zm(i))
  Next
  
  '  Set up the globals:  g_init_bits - initial number of bits
  g_Init_Bits = init_bits

  '  Set up the necessary values

  Offset = 0
  Out_Count = 0
  Clear_Flg = False

  n_Bits = g_Init_Bits
  MaxCode = Masks(n_Bits)               'MAXCODE(n_bits);

  ClearCode = 2 ^ (init_bits - 1)
  EOFCode = ClearCode + 1
  Free_Ent = ClearCode + 2

  Call char_init                        'set up output buffers

  hshift = 0
  fcode = HASHTABSIZE
  Do While fcode < 65536
    hshift = hshift + 1
    fcode = fcode + fcode
  Loop
  hshift = 1 + Masks(8 - hshift)        'set hash code range bound for shifting
  
  Call cl_hash                          'clear hash table
  Call OutputCode(ClearCode)            'get ready to go
  Out_Count = 1
  
  ent = GIFNextPixel(PixBits)
  In_Count = 1
  
  c = GIFNextPixel(PixBits)
  Do While c <> EOF
    In_Count = In_Count + 1

    fcode = c * MAXBITSHIFT + ent
    i = (c * hshift) Xor ent          '/* xor hashing */

    If HashTab(i) = fcode Then
      ent = CodeTab(i)
      GoTo NextPixel
    ElseIf HashTab(i) < 0 Then       '/* empty slot */
      GoTo NoMatch
    End If
      
    disp = HASHTABSIZE - i          '/* secondary hash (after G. Knott) */
    If i = 0 Then disp = 1

Probe:
    i = i - disp
    If i < 0 Then i = i + HASHTABSIZE

    If HashTab(i) = fcode Then
      ent = CodeTab(i)
      GoTo NextPixel
    End If
    
    If HashTab(i) > 0 Then GoTo Probe

NoMatch:
    Call OutputCode(ent)
    Out_Count = Out_Count + 1
    ent = c

    If Free_Ent < MAXMAXCODE Then
      CodeTab(i) = Free_Ent
      Free_Ent = Free_Ent + 1             '/* code -> hashtable */
      HashTab(i) = fcode
    Else
      Call cl_block
    End If
NextPixel:
    c = GIFNextPixel(PixBits)
  Loop

  '  Put out the final code.

  Call OutputCode(ent)
  Out_Count = Out_Count + 1
  Call OutputCode(EOFCode)
  Out_Count = Out_Count + 1
End Sub

'Return the next pixel from the image and increment positions
Private Function GIFNextPixel(ByRef PixBits() As Byte) As Integer
  Dim RowOffset As Long, Mask As Long
  
  If (Countdown = 0) Then
    GIFNextPixel = EOF
  Else
    Countdown = Countdown - 1
    RowOffset = LBound(PixBits) + RowMod * CurY
    
    Select Case PixSize         '1,4,8 from a bitmap
      Case 8:                                                            'every byte is a pixel
        GIFNextPixel = PixBits(RowOffset + CurX)
      
      Case 4:                                                            'every nibble is a pixel
        If (CurX And 1) = 1 Then
          GIFNextPixel = CLng(PixBits(RowOffset + CurX \ 2)) And &HF&    'odd
        Else
          GIFNextPixel = (CLng(PixBits(RowOffset + CurX \ 2)) And &HF0&) \ 16 'even
        End If
        
      Case 1:                                                            'every bit is a pixel
        Mask = 2& ^ (7 - CurX Mod 8)
        GIFNextPixel = (CLng(PixBits(RowOffset + CurX \ 8)) And Mask) \ Mask
    End Select
    
  '   Bump the current X position
    
    CurX = CurX + 1
  
  '   If we are at the end of a scan line, set curx back to the beginning
  '   If we are interlaced, bump the cury to the appropriate spot, otherwise, just increment it.
    
    If CurX = PWidth Then
      CurX = 0
      If Interlace = False Then
        CurY = CurY + 1
      Else
        Select Case Pass
          Case 0:
            CurY = CurY + 8
            If CurY >= PHeight Then
              Pass = Pass + 1
              CurY = 4
            End If
                 
          Case 1:
            CurY = CurY + 8
            If CurY >= PHeight Then
              Pass = Pass + 1
              CurY = 2
            End If
  
          Case 2:
            CurY = CurY + 4
            If CurY >= PHeight Then
              Pass = Pass + 1
              CurY = 1
            End If
               
          Case 3:
            CurY = CurY + 2
        End Select
      End If
    End If
  End If
End Function

' TAG( OutputCode )
' Output the given code.
'  Inputs:
'    code: A n_bits-bit integer.  If == -1, then EOF.  This assumes that n_bits =< (long)wordsize - 1.
'  Outputs:
'    Outputs code to the file.
'  Assumptions:
'    Chars are 8 bits long.
'  Algorithm:
'    Maintain a MAXBITS character long buffer (so that 8 codes will fit in it exactly).
'    When the buffer fills up empty it and start over.

Private Sub OutputCode(ByVal code As Long)

  cur_Accum = cur_Accum And Masks(cur_Bits)

  If (cur_Bits > 0) Then
    cur_Accum = cur_Accum Or (code * (1 + Masks(cur_Bits)))
  Else
    cur_Accum = code
  End If
  
  cur_Bits = cur_Bits + n_Bits

  Do While (cur_Bits >= 8)
    Call char_out(cur_Accum And &HFF&)
    cur_Accum = cur_Accum \ 256&
    cur_Bits = cur_Bits - 8
  Loop

  ' If the next entry is going to be too big for the code size, then increase it, if possible.

  If (Free_Ent > MaxCode Or Clear_Flg = True) Then
    If (Clear_Flg = True) Then
      n_Bits = g_Init_Bits
      MaxCode = Masks(n_Bits)       'MAXCODE(n_bits);
      Clear_Flg = False
    Else
      n_Bits = n_Bits + 1
      If (n_Bits = MAXBITS) Then
        MaxCode = MAXMAXCODE
      Else
       MaxCode = Masks(n_Bits)     'MAXCODE(n_bits);
      End If
    End If
  End If

  If (code = EOFCode) Then    'At EOF, write the rest of the buffer.
    While (cur_Bits > 0)
      Call char_out(cur_Accum And &HFF&)
      cur_Accum = cur_Accum \ 256&
      cur_Bits = cur_Bits - 8
    Wend
    Call flush_char
  End If
End Sub

' Clear out the hash table

Private Sub cl_block()                          '/* table clear for block compress */
  Call cl_hash
  Free_Ent = ClearCode + 2
  Clear_Flg = True
  Call OutputCode(ClearCode)
  Out_Count = Out_Count + 1
End Sub

Private Sub cl_hash()                           '/* reset code table */
  Dim i As Long
  
  For i = 0 To HASHTABSIZE - 1
    HashTab(i) = -1
  Next
End Sub

' Set up the 'byte output' routine and Define the storage for the packet accumulator
Private Sub char_init()
  A_Count = 0
  ReDim Accum(0 To 255) As Byte
End Sub

' Add a character to the end of the current packet, and if it is 254 characters, flush the packet to disk.
Private Sub char_out(ByVal c As Integer)
  Accum(A_Count + 1) = c              '0,1,2,3 ....mapped to 1,2,3,4...255
  A_Count = A_Count + 1
  If A_Count >= 255 Then Call flush_char      'in the original this was >=254, the std allows 255
End Sub                                       '(most art programs Ive got seem to use this 254 code)

' Flush the current packet to disk, and reset the accumulator
Private Sub flush_char()
  If A_Count > 0 Then
    Accum(0) = A_Count                                'set block length
    ReDim Preserve Accum(0 To A_Count) As Byte        'and redimension to this length
    Put #98, , Accum                                  'write it to disk
    FileCount = FileCount + A_Count + 1               'track bytes written
    Call char_init
  End If
End Sub

'============================ THE REAL ROUTINES PUBLICLY VISIBLE =========================================

Public Function SaveGIF(ByVal Path As String, _
                        ByVal Width As Long, ByVal Height As Long, ByVal BitsPerPixel As Long, _
                        ByRef PixBits() As Byte, ByVal PixelWidth As Long, ByRef CMap() As Byte, _
                        Optional ByVal Interlaced As Boolean = False) As Long     '<=0=failure,1=success
  
  'Path: Where you will store the file should end .gif
  'Width,Height:  Pic size in pixels
  'BitsPerPixel: in planes 1=BW, 4=16 colours, 8=256 colours
  'PixBits: the bits of the picture from top-left in BitsPerPixel=1 1 pixel=1 bit, BitsPerPixel=4 1 pixel=4bits, 8 1pixel=8bits
  'PixelWidth how wide is a pixel packed in bits should be 1,4,8 for MS bitmaps
  'When calling this routine independently make sure the image is the right way up, and that a colour map exists
  'CMap: the three byte tuples r,g,b for each colour in the image. (should be 3*2^n bytes n=1,,8) NOT CHECKED
  'Interlaced: make the GIF interlaced (see doco)
  
  Dim ID As String
  Dim GSH As GifScreenHdr
  Dim GIH As GifImageHdr
  
    ' attempt to save a gif file of the bitmap data with colormap cmap
    ' CMAP contains 2^BitsPerPixel colours of 3 bytes each r,g,b
    ' Bits contains the colour mapped data as 1 byte per pixel as mapped by colormap
  
  On Error GoTo BadPath
  Open Path For Binary Access Write As #98
  On Error GoTo GIFSaveFailed
  
  'File identifier
  ID = "GIF87a"         'We are making the lowest common file type here
  Put #98, , ID
  
  'ScreenDescriptor
'              Bits
'         7 6 5 4 3 2 1 0  Byte -
'        +---------------+
'        |               |  1
'        +-Screen Width -+      Raster width in pixels (LSB first)
'        |               |  2
'        +---------------+
'        |               |  3
'        +-Screen Height-+      Raster height in pixels (LSB first)
'        |               |  4
'        +-+-----+-+-----+      M = 1, Global color map follows Descriptor
'        |M|  cr |0|pixel|  5   cr+1 = - bits of color resolution
'        +-+-----+-+-----+      pixel+1 = - bits/pixel in image
'        |   background  |  6   background=Color index of screen background
'        +---------------+          (color is defined from the Global color
'        |0 0 0 0 0 0 0 0|  7        map or default map if none specified)
'        +---------------+
        
'        The logical screen width and height can both  be  larger  than  the
'   physical  display.   How  images  larger  than  the physical display are
'   handled is implementation dependent and can take advantage  of  hardware
'   characteristics  (e.g.   Macintosh scrolling windows).  Otherwise images
'   can be clipped to the edges of the display.

'        The value of 'pixel' also defines  the  maximum  number  of  colors
'   within  an  image.   The  range  of  values  for 'pixel' is 0 to 7 which
'   represents 1 to 8 bits.  This translates to a range of 2 (B & W) to  256
'   colors.   Bit  3 of word 5 is reserved for future definition and must be
'   zero.
  
  With GSH
    .Width = Width
    .Height = Height
    .MCR0Pix = &HF0 Or (BitsPerPixel - 1)
    .BC = 0
    .Aspect = 0
  End With
  
  'done this way to make sure LoHi storage in file
  Put #98, , GSH.Width
  Put #98, , GSH.Height
  Put #98, , GSH.MCR0Pix
  Put #98, , GSH.BC
  Put #98, , GSH.Aspect
    
  'Global ColorMap
'        The Global Color Map is optional but recommended for  images  where
'   accurate color rendition is desired.  The existence of this color map is
'   indicated in the 'M' field of byte 5 of the Screen Descriptor.  A  color
'   map  can  also  be associated with each image in a GIF file as described
'   later.  However this  global  map  will  normally  be  used  because  of
'   hardware  restrictions  in equipment available today.  In the individual
'   Image Descriptors the 'M' flag will normally be  zero.   If  the  Global
'   color Map Is present, it    's definition immediately follows the Screen
'   Descriptor.   The  number  of  color  map  entries  following  a  Screen
'   Descriptor  is equal to 2**(- bits per pixel), where each entry consists
'   of three byte values representing the relative intensities of red, green
'   and blue respectively.  The structure of the Color Map block is:

'              Bits
'         7 6 5 4 3 2 1 0  Byte -
'        +---------------+
'        | red intensity |  1    Red value for color index 0
'        +---------------+
'        |green intensity|  2    Green value for color index 0
'        +---------------+
'        | blue intensity|  3    Blue value for color index 0
'        +---------------+
'        | red intensity |  4    Red value for color index 1
'        +---------------+
'        |green intensity|  5    Green value for color index 1
'        +---------------+
'        | blue intensity|  6    Blue value for color index 1
'        +---------------+
'        :               :       (Continues for remaining colors)

'        Each image pixel value received will be displayed according to  its
'   closest match with an available color of the display based on this color
'   map.  The color components represent a fractional intensity  value  from
'   none  (0)  to  full (255).  White would be represented as (255,255,255),
'   black as (0,0,0) and medium yellow as (180,180,0).  For display, if  the
'   device  supports fewer than 8 bits per color component, the higher order
'   bits of each component are used.  In the creation of  a  GIF  color  map
'   entry  with  hardware  supporting  fewer  than 8 bits per component, the
'   component values for the hardware  should  be  converted  to  the  8-bit
'   format with the following calculation:

'        map_value> = component_value>*255/(2**nbits> -1)

'        This assures accurate translation of colors for all  displays.   In
'   the  cases  of  creating  GIF images from hardware without color palette
'   capability, a fixed palette should be created  based  on  the  available
'   display  colors for that hardware.  If no Global Color Map is indicated,
'   a default color map is generated internally  which  maps  each  possible
'   incoming  color  index to the same hardware color index modulo where
'   is the number of available hardware colors.

  Put #98, , CMap     'NOTE THIS IS NOT CHECKED should be of size 3*2^k, k=1..8
  
  'ImageDescriptor
'        The Image Descriptor defines the actual placement  and  extents  of
'   the  following  image within the space defined in the Screen Descriptor.
'   Also defined are flags to indicate the presence of a local color  lookup
'   map, and to define the pixel display sequence.  Each Image Descriptor is
'   introduced by an image separator  character.   The  role  of  the  Image
'   Separator  is simply to provide a synchronization character to introduce
'   an Image Descriptor.  This is desirable if a GIF file happens to contain
'   more  than  one  image.   This  character  is defined as 0x2C hex or ','
'   (comma).  When this character is encountered between images,  the  Image
'   Descriptor will follow immediately.
        
'        Any characters encountered between the end of a previous image  and
'   the image separator character are to be ignored.  This allows future GIF
'   enhancements to be present in newer image formats and yet ignored safely
'   by older software decoders.

'              Bits
'         7 6 5 4 3 2 1 0  Byte -
'        +---------------+
'        |0 0 1 0 1 1 0 0|  1    ',' - Image separator character &H2C
'        +---------------+
'        |               |  2    Start of image in pixels from the
'        +-  Image Left -+       left side of the screen (LSB first)
'        |               |  3
'        +---------------+
'        |               |  4
'        +-  Image Top  -+       Start of image in pixels from the
'        |               |  5    top of the screen (LSB first)
'        +---------------+
'        |               |  6
'        +- Image Width -+       Width of the image in pixels (LSB first)
'        |               |  7
'        +---------------+
'        |               |  8
'        +- Image Height-+       Height of the image in pixels (LSB first)
'        |               |  9
'        +-+-+-+-+-+-----+       M=0 - Use global color map, ignore 'pixel'
'        |M|I|0|0|0|pixel| 10    M=1 - Local color map follows, use 'pixel'
'        +-+-+-+-+-+-----+       I=0 - Image formatted in Sequential order
'                                I=1 - Image formatted in Interlaced order
'                                pixel+1 - - bits per pixel for this image
'
'        The specifications for the image position and size must be confined
'   to  the  dimensions defined by the Screen Descriptor.  On the other hand
'   it is not necessary that the image fill the entire screen defined.
  
  With GIH
    .Left = 0
    .Top = 0
    .Width = Width
    .Height = Height
    .MIPixBits = BitsPerPixel - 1
    If Interlaced = True Then .MIPixBits = .MIPixBits Or &H40
    'code size is part of the raster stream but for convenience Ive added it to the ImageHeader
    If BitsPerPixel = 1 Then .CodeSize = 2 Else .CodeSize = BitsPerPixel    'see below
  End With
    
  'done this way to make sure LoHi storage in file
  Put #98, , Chr$(&H2C)
  Put #98, , GIH.Left
  Put #98, , GIH.Top
  Put #98, , GIH.Width
  Put #98, , GIH.Height
  Put #98, , GIH.MIPixBits
  Put #98, , GIH.CodeSize
    
  'The Compressed Bits
'        The Raster Data stream that represents the actual output image  can
'   be represented as:

'         7 6 5 4 3 2 1 0
'        +---------------+
'        |   code size   |
'        +---------------+     ---+
'        |blok byte count|        |
'        +---------------+        |
'        :               :        +-- Repeated as many times as necessary
'        |  data bytes   |        |
'        :               :        |
'        +---------------+     ---+
'        . . .       . . .
'        +---------------+
'        |0 0 0 0 0 0 0 0|       zero byte count (terminates data stream)
'        +---------------+

'        The conversion of the image from a series  of  pixel  values  to  a
'   transmitted or stored character stream involves several steps.  In brief
'   these steps are:

'    Establish the Code Size -
'       Define  the  number  of  bits  needed  to
'       represent the actual data.

'   Compress the Data -
'       Compress the series of image pixels to a  series
'       of compression codes.

'   Build a Series of Bytes -
'       Take the  set  of  compression  codes  and
'       convert to a string of 8-bit bytes.

'   Package the Bytes -
'       Package sets of bytes into blocks  preceeded  by
'       character counts and output.

'   Establish Code Size
'        The first byte of the GIF Raster Data stream is a value  indicating
'   the minimum number of bits required to represent the set of actual pixel
'   values.  Normally this will be the same as the  number  of  color  bits.
'   Because  of  some  algorithmic constraints however, black & white images
'   which have one color bit must be indicated as having a code size  of  2.
'   This  code size value also implies that the compression codes must start
'   out one bit longer.
  
  'set some external globals
  PWidth = Width
  PHeight = Height
  PixSize = PixelWidth
  RowMod = (UBound(PixBits) - LBound(PixBits) + 1) \ PHeight   'BMPS can have hanging bits at the end of rows
  Interlace = Interlaced
  Call CompressAndWriteBits(GIH.CodeSize + 1, PixBits)
  
  'write the trailer, terminator
  Put #98, , Chr$(&H3B)
  
  Close #98
  SaveGIF = 1
  Exit Function

BadPath:
  SaveGIF = 0
  On Error GoTo 0
  Exit Function
  
GIFSaveFailed:
  Close #98
  Call Kill(Path)     'no idea if file is any good so kill it (could fail here if open failed)
  SaveGIF = 0
  On Error GoTo 0
End Function
