VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRVTVBIMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'cRVTVBIMG.cls    - Image saving Pipeline Processing  - saves GIF and BMP

Public Enum RVTVBIMG_PICTYPE 'PIC Save Type
  PIC_BMP = 0               'save as BMP
  PIC_GIF = 1               'save as GIF
  PIC_GIF_LACED = 2         'save as interlaced gif
                            'others should be added here
End Enum

Public Enum RVTVBIMG_COLORMODES
  PIC_COLOR = 0           'Color style  default
  PIC_BW = 1
  PIC_GREY = 2              'colours are converted into greys (not necessarily unique)
End Enum

'NOTE: 32BIT Options do not work - I dont know why - but 24bits is equivalent

Public Enum RVTVBIMG_NCOLORMODES      'Depth in Bits and the synonymous Number of colors
  PIC_1BPP = 1              'GIF & BMP implies Black and White
  PIC_2BPP = 2              'GIF 4 colors
  PIC_3BPP = 3              'GIF 8 colors
  PIC_4BPP = 4              'GIF 16 or BMP implies VGA 16 color palette
  PIC_5BPP = 5              'GIF 32 colors
  PIC_6BPP = 6              'GIF 64 colors
  PIC_7BPP = 7              'GIF 128 colors
  PIC_8BPP = 8              'GIF 256 or BMP implies Internet palette default
  PIC_16BPP = 16            'BMP          colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
  PIC_24BPP = 24            'BMP default  colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
  PIC_32BPP = 32            'BMP          colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
                          
                          'NCOlors in Bits
  PIC_C2 = 1                'GIF & BMP implies Black and White
  PIC_C4 = 2                'GIF 4 colors
  PIC_C8 = 3                'GIF 8 colors
  PIC_C16 = 4               'GIF 16 or BMP implies VGA 16 color palette
  PIC_C32 = 5               'GIF 32 colors
  PIC_C64 = 6               'GIF 64 colors
  PIC_C128 = 7              'GIF 128 colors
  PIC_C256 = 8              'GIF 256 or BMP implies Internet palette default
  PIC_C65536 = 16           'BMP          colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
  PIC_C16M = 24             'BMP default  colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
  PIC_C16M32 = 32           'BMP          colour implicit in pixeldata  (no Grey Mode), => 8BPP GIF
End Enum

Public Enum RVTVBIMG_CMAPMODES 'CMAP Actions
  PIC_USE_MS_CMAP = &H1000&       'Just use the Cmap passed on from GDI or REMAP using MS API
  PIC_SMALLEST_CMAP = &H2000&     'Use the smallest palette possible (of the colors given by MS)
  PIC_OPTIMAL_CMAP = &H4000&      'Make an optimal palette to correspond with BPP (use 24bits to n bits)
  PIC_FIXED_CMAP = &H8000&        'A fixed CMAP  if supplied a 256 colormap will be used ie. &H8080
  PIC_FIXED_CMAP_BW = &H8002&     'Force the two tone dithering Black and White (Fixed CMAP)
  PIC_FIXED_CMAP_C4 = &H8004&     'Force the 4 tone dither WCMY (fixed CMAP)
  PIC_FIXED_CMAP_C8 = &H8008&     'Force the 8 tone dither KRGBWCMY (fixed CMAP)
  PIC_FIXED_CMAP_C16 = &H8010&    'Force the VGA Colormap
  PIC_FIXED_CMAP_VGA = &H8210&    'Force the VGA Colormap
  PIC_FIXED_CMAP_C32 = &H8020&    'Force the Exteded VGA Map
  PIC_FIXED_CMAP_C64 = &H8040&    'Force the 64 ColorMAP
  PIC_FIXED_CMAP_C128 = &H8080&   'Force the 128 ColorMAP
  PIC_FIXED_CMAP_C256 = &H8100&   'Force the RVT 256 colormap
  PIC_FIXED_CMAP_MS256 = &H8300&  'Force the MS  256 colormap
  PIC_FIXED_CMAP_INET = &H84D8&   'Force the 216 InternetSafe colormap
  PIC_FIXED_CMAP_GREY = &H88FF&   'Use a nColor Grey map
  PIC_FIXED_CMAP_USER = &H80FF&   'Use a user supplied nColorMap   (see SetCustomCmap())
 End Enum
  
Public Enum RVTVBIMG_DITHERMODES 'DITHER Actions
  PIC_DITHER_NONE = 0
  PIC_DITHER_BIN = &H1&          'Dither in a binary method
  PIC_DITHER_ORD = &H2&          'Dither to an ordered dithermatrix (default)
  PIC_DITHER_HTC = &H4&          'Dither to a printer Halftone
  PIC_DITHER_FDIAG = &H8&        'Dither to an forward diagonal dithermatrix
  PIC_DITHER_BDIAG = &H10&       'Dither to an backward diagonal dithermatrix
  PIC_DITHER_HORZ = &H20&        'Dither to an horizontal dithermatrix
  PIC_DITHER_VERT = &H40&        'Dither to an vertical dithermatrix
  PIC_DITHER_FS1 = &H100&        'Floyd-Steinberg dithering  Original Matrix
  PIC_DITHER_FS2 = &H200&        'Floyd-Steinberg dithering  Even Matrix
  PIC_DITHER_FS3 = &H400&        'Floyd-Steinberg dithering  Equal Matrix
End Enum

Public Enum RVTVBIMG_ACTIONS_INTERNAL
  PIC_INTERLACED = &H1000&      'for GIF produce an interlaced GIF
End Enum

Public Enum RVTVBIMG_PICPROC_COMMANDS
  'The constants here have the form
  'BYTE 3 special control codes, BYTE 2 API OpCodes, BYTE 1 Command Nr, BYTE 0 The number of parameters it uses
  '00=Ignored                  '00 - Ignored
  '01=APIAssisted              '01 - FLIP_VERT
  '02=Requires Unmap           '02 - FLIP_HORZ
  '04=                         '04 - IMAGE_RESIZE )
  '08=                         '08 - IMAGE_ZOOM   ) can only define one of these
  '10=                         '10 - INVERT_COLOR
  '20=                         '20 - MSMAP_COLOR  not with Resize or Zoom, but UNMAP=REMAPto24BPP
  '40=
  '80 DO NOT USE
  
  PIC_UNMAP_COLOR = &H3000000    'We must unmap (to 24BPP) if mapped, (and remap afterwards)
  PIC_FLIP_VERT = &H1010000      'flip image vertically
  PIC_FLIP_HORZ = &H1020000      'flip image horizontally
  PIC_IMAGE_RESIZE = &H3040002   'Resize to parm1=NewWidth, parm2=NewHeight
  PIC_IMAGE_ZOOM = &H3080002     'Zoom image by parm1=XZoom times,parm2=YZoom Times
                                 'if xzoom or yzoom<0 then shrink
                                 'these paramaters should be scaled by 10000 and passed in as Long
  PIC_INVERT_COLOR = &H3100000   'Invert all colors 00=FF etc
  PIC_MSMAP_COLOR = &H1200001    'Parm1 = Color BPP required only 1,4,8,16,24 allowed
  'new commands (and their VB code)
End Enum

Public Enum RVTVBIMG_PICSTATES 'additive bitstates
  IS_VALID_PIPELINE = &H1&
  IS_VALID_CLIP = &H2&         'the clipping rectangle is inside the image
  IS_CMAPPED = &H4&            'Colormapping has taken place (all pixels are 1,4 or 8bit indices to CMap)
  IS_TOP_TO_BOTTOM = &H8&      'the image is topdown oriented, not bottom up as a BMP usually is
End Enum

Public Type ZRGB    'all values 0..255
  Grey  As Byte     'the grey value of the pixel
  Red   As Byte     'the red value
  Green As Byte     'the green value
  Blue  As Byte     'the blue value
End Type

'============================= THE BASIC ELEMENTS OF THE IMAGE ================================================

Private zWidth        As Long   'in pixels
Private zHeight       As Long   'in pixels
Private zPixBits()    As Byte   'the array of pixel bits, may be indices into the CMap
Private zIntBPP       As Long   'The Internal BitsPerPixel (1,4,8,16,24 for BMPS, =24 for GIFS)
Private zCMap()       As Byte   'the colormap in R,G,B tuples of 0..255  (MUST BE 3*NColors)
Private zNMapColors   As Long   'The number of Colors in the CMap, 2^BitsPerPixel

Private zPicState     As Long   'see enum above
Private zPicType      As Long   'BMP or GIF
Private zColorMode    As Long   'COLOR, BW, GREY
Private zReqBPP       As Long   'The requested format in Bits Per Pixel
Private zCMAPMode     As Long   'See Enum above
Private zDitherMode   As Long   'See Enum Above

Private zUserCMAP()   As Byte   'will be used if CMAPMode=PIC_FIXED_CMAP_USER
                                'your problem to make it the right size

'for DemoPurposes Only
Public etCMAP    As Single
Attribute etCMAP.VB_VarDescription = "Elapsed time for ColorMapping Routines"
Public etRemap   As Single
Attribute etRemap.VB_VarDescription = "Elapsed time (secs) for the color pixel remapping routines"
Public etSave    As Single
Attribute etSave.VB_VarDescription = "Elapsed time (secs) for the I/O step of saving an image"
Public etElapsed As Single
Attribute etElapsed.VB_VarDescription = "Elapsed time (secs) for the entire saving process"

Private Sub Class_Terminate()
  Erase zCMap(), zPixBits()
End Sub

'==================================================================================================================
'This is the AutoPilot User function, everything is done for you, The whole or part of an Object hDC will be
'saved according to your passed in preferences
'The simplest possible call is
'    SaveObjDCCLip(Obj,Path) which saves the whole Obj.hDC as a 256Color Bitmap to Path
'  a return code =0 indicates something went wrong.
'
'I havent been particularly industrial strength in error handling (sorry), so why may be hard to fathom (trace it)
'==================================================================================================================

Public Function SaveObjDCClip(Obj As Object, Path As String, _
                              Optional ByVal PicType As Long = PIC_BMP, _
                              Optional ByVal ColorMode As Long = PIC_COLOR, _
                              Optional ByVal BitsPerPixel As Long = PIC_8BPP, _
                              Optional ByVal CMAPMode As Long = PIC_USE_MS_CMAP, _
                              Optional ByVal DitherMode As Long = PIC_DITHER_NONE, _
                              Optional ByVal ClipX0 As Long = -1, Optional ByVal ClipY0 As Long = -1, _
                              Optional ByVal ClipX1 As Long = -1, Optional ByVal ClipY1 As Long = -1) As Long
Attribute SaveObjDCClip.VB_Description = "The DO everything automativally version of saving a Device context (clip) to disk"
                              
  Dim rc As Long
  
  etElapsed = Timer
  SaveObjDCClip = 0 'assume the worst
  If Obj Is Nothing Or Len(Path) = 0 Then Exit Function
    
  'set the pipeline
  If SetPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode) = 0 Then Exit Function
  
  'get the basic image data
  If ImageFromObjDCCLip(Obj, ClipX0, ClipY0, ClipX1, ClipY1) = 0 Then Exit Function
  
  'we now have the whole thing in OUR Memory - so can manipulate it
'  Call FlipVert(zWidth, zHeight, zPixBits(), zReqBPP, zCMap(), zNMapColors)
  
  ' Post Processing
  If DoColorMapping() = 0 Then Exit Function
  
  ' File IO
  etSave = Timer
  Select Case zPicType
    Case PIC_BMP:       SaveObjDCClip = SaveAsBMP(Path)
    Case PIC_GIF:       SaveObjDCClip = SaveAsGIF(Path)
    Case PIC_GIF_LACED: SaveObjDCClip = SaveAsInterlacedGIF(Path)
  End Select
  etSave = Timer - etSave
  etElapsed = Timer - etElapsed
End Function

'==================================================================================================================
'HERE FOLLOW the individual steps of the Pipeline as applied by the SaveOBJDCCLip
'Some sanity checking is done at this level - the routines behind the scenes ASSUME THIS HAS BEEN DONE
'If you stuff up the parameters - the results are unpredictable as they say
'==================================================================================================================

Public Function SetPipeline(Optional ByVal PicType As Long = PIC_BMP, _
                            Optional ByVal ColorMode As Long = PIC_COLOR, _
                            Optional ByVal BitsPerPixel As Long = PIC_8BPP, _
                            Optional ByVal CMAPMode As Long = PIC_USE_MS_CMAP, _
                            Optional ByVal DitherMode As Long = PIC_DITHER_NONE) As Long
Attribute SetPipeline.VB_Description = "Set the control parameters for the processing steps to be done"

  SetPipeline = 0
  zPicState = zPicState And (Not IS_VALID_PIPELINE)
  
  'Validate the combinations from PICType down
  ' some (non comprehensive) sanity checking of PicType and BPP
  
  If PicType <> PIC_BMP And PicType <> PIC_GIF And PicType <> PIC_GIF_LACED Then PicType = PIC_BMP
  zPicType = PicType
  
   ' sanity checking of colormode
  If ColorMode <> PIC_COLOR And ColorMode <> PIC_BW And ColorMode <> PIC_GREY Then ColorMode = PIC_COLOR
  If ColorMode = PIC_BW Then BitsPerPixel = PIC_1BPP
  zColorMode = ColorMode
  
  zIntBPP = BitsPerPixel           'internal representation (assumed)
  zReqBPP = BitsPerPixel           'required representation
  
  If CMAPMode = PIC_FIXED_CMAP Then CMAPMode = PIC_FIXED_CMAP_C256
  'CMAP Mode and Dither Mode are not otherwise checked
  zCMAPMode = CMAPMode
  zDitherMode = DitherMode
  
  'now some cross validations
  Select Case zPicType
    Case PIC_BMP:              'BMPS can have 2,16,256,65536 and 16777216 colors
      If (zCMAPMode And PIC_USE_MS_CMAP) = PIC_USE_MS_CMAP Then 'we are limited to 1,4,8 planes of quantizing (BW/VGA/Internet)
        If zIntBPP <= PIC_1BPP Then
          zIntBPP = PIC_1BPP
        ElseIf zIntBPP <= PIC_4BPP Then
          zIntBPP = PIC_4BPP
        ElseIf zIntBPP <= PIC_8BPP Then
          zIntBPP = PIC_8BPP
        ElseIf zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP
        End If
      ElseIf (zCMAPMode And PIC_FIXED_CMAP) = PIC_FIXED_CMAP Then
        If zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP        ' and will need mapping and dithering of colours!!
        End If
      Else                   ' we will quantize from 16,24 or 32 bits down to the given sized palette
        If zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP        ' and will need quantizing,mapping and dithering of colours!!
        End If
      End If
      SetPipeline = 1
      
    Case PIC_GIF, PIC_GIF_LACED:            'GIFS can have 2,4,8,16,32,64,128,and 256 colors
      If (zCMAPMode And PIC_USE_MS_CMAP) = PIC_USE_MS_CMAP Then 'we are limited to 1,4,8 planes of quantizing (BW/VGA/Internet)
        If zIntBPP <= PIC_1BPP Then
          zIntBPP = PIC_1BPP
        ElseIf zIntBPP <= PIC_4BPP Then
          zIntBPP = PIC_4BPP
          zCMAPMode = zCMAPMode Or PIC_SMALLEST_CMAP
        ElseIf zIntBPP <= PIC_8BPP Then
          zIntBPP = PIC_8BPP
          zCMAPMode = zCMAPMode Or PIC_SMALLEST_CMAP
        ElseIf zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP        ' and will need quantizing,mapping and dithering of colours!!
          zCMAPMode = (zCMAPMode And (Not PIC_USE_MS_CMAP)) Or PIC_OPTIMAL_CMAP
        End If
      ElseIf (zCMAPMode And PIC_FIXED_CMAP) = PIC_FIXED_CMAP Then
        If zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP        ' and will need quantizing,mapping and dithering of colours!!
        End If
      Else                   ' we will quantize from 16,24 or 32 bits down to the given sized palette
        If zIntBPP <> PIC_16BPP And zIntBPP <> PIC_24BPP And zIntBPP <> PIC_32BPP Then
          zIntBPP = PIC_24BPP        ' and will need quantizing,mapping and dithering of colours!!
        End If
        zCMAPMode = zCMAPMode Or PIC_OPTIMAL_CMAP
      End If
      SetPipeline = 1
    
    Case Else:
  End Select      'PIC_TYPE

  
  'figure out the Nr of Colors
  If zIntBPP <= PIC_8BPP Then zNMapColors = 2 ^ zIntBPP Else zNMapColors = 0
  
  If SetPipeline = 1 Then zPicState = zPicState Or IS_VALID_PIPELINE
End Function

'==================================================================================================================
'public version of this function (uses the data provide by the object)
' Gets the Clip image into this objects memory (see BMPSave.bas)
'this can fail due to lack of memory  rc=0 indicates nomemory or no pipeline or worse
'==================================================================================================================

Public Function ImageFromObjDCCLip(ByVal Obj As Object, _
                                   Optional ByVal ClipX0 As Long = -1, _
                                   Optional ByVal ClipY0 As Long = -1, _
                                   Optional ByVal ClipX1 As Long = -1, _
                                   Optional ByVal ClipY1 As Long = -1) As Long
Attribute ImageFromObjDCCLip.VB_Description = "From the Device context attached to a GDI object (eg. Form, Picture Box) extract the image"
  
  Dim osm As Long   'saved scalemode
  
  If (zPicState And IS_VALID_PIPELINE) = 0 Then      'must have a valid pipeline to proceed
    ImageFromObjDCCLip = 0
  Else
    On Error GoTo InvalidhDC
    ' sanity checking of object - it must have a device context
    osm = Obj.hdc      'no device context will go to error found
    
    osm = Obj.ScaleMode: Obj.ScaleMode = vbPixels   'need to work in pixels, no scalemode is an error
  
    ' sanity checking of the clip data (overrides are taken for screwy data) Default is the whole DC Bitmap

    If ClipX0 < 0 Then ClipX0 = 0
    If ClipY0 < 0 Then ClipY0 = 0
    If ClipX1 < 0 Or ClipX1 >= Obj.ScaleWidth Then ClipX1 = Obj.ScaleWidth - 1
    If ClipY1 < 0 Or ClipY1 >= Obj.ScaleHeight Then ClipY1 = Obj.ScaleHeight - 1
    
    ImageFromObjDCCLip = ImageFromDCClip(Obj.hdc, zPicType, zPicState, _
                                         zWidth, zHeight, zPixBits(), zIntBPP, zCMap(), zNMapColors, _
                                         ClipX0, ClipY0, ClipX1, ClipY1)
    Obj.ScaleMode = osm
  End If
  On Error GoTo 0
  Exit Function

InvalidhDC:
  Obj.ScaleMode = osm
  ImageFromObjDCCLip = 0
  On Error GoTo 0
End Function

'========= ELEMENTERY IMAGE PROCESSING ===== New Options should be added Here =================================
'They should all follow the convention that rc=0 indicates failure, 1=success

Public Function FlipVert() As Long
Attribute FlipVert.VB_Description = "Flip an image from top to bottom"
  FlipVert = DoAPIOperations(PIC_FLIP_VERT)
End Function

Public Function FlipHorz() As Long
Attribute FlipHorz.VB_Description = "Flip an Image from Left to Right"
  'as I cant get this to work with StretchDIBits we will do it manually
  'FlipHorz = DoAPIOperations(PIC_FLIP_HORZ)
  FlipHorz = FlipImageHorz(zWidth, zHeight, zPixBits(), zIntBPP)  'deals with all MS formats
End Function

Public Function ResizeImage(ByVal NewWidth As Long, ByVal NewHeight As Long) As Long
Attribute ResizeImage.VB_Description = "Resize an image to new Width and Height"
  ResizeImage = DoAPIOperations(PIC_IMAGE_RESIZE, NewWidth, NewHeight)
End Function

'NOTE NOTE NOTE: Zoom parameters are scaled by 10000 ie. to zoom 3.6 times XZoom would be 36000
Public Function ZoomImage(ByVal XZoom As Long, ByVal YZoom As Long) As Long
Attribute ZoomImage.VB_Description = "Make the image Zoom times bigger (Zoom>0) or smaller (if Zoom <0). Either axis may be resized"
  ZoomImage = DoAPIOperations(PIC_IMAGE_ZOOM, XZoom, YZoom)
End Function

Public Function InvertColor() As Long
Attribute InvertColor.VB_Description = "Make all colors the complement of their present colors"
  InvertColor = DoAPIOperations(PIC_INVERT_COLOR)
End Function

Public Function UnMapColor() As Long
Attribute UnMapColor.VB_Description = "For a colormapped image remove colormapping and make a 24BPP image"
  UnMapColor = DoAPIOperations(PIC_UNMAP_COLOR)
End Function

'This one cannot be used in conjunction with Other Opcodes passing parameters or UnMapColor
Public Function MSMapColor(ByVal NewBPP As Long) As Long
Attribute MSMapColor.VB_Description = "Use MS methods to map an image into 1, 4, 8 16 or 24 BitsPerPixel (2,16,256,65536,16M colors resp.)"
  MSMapColor = DoAPIOperations(PIC_MSMAP_COLOR, NewBPP)
End Function

'============= THERE IS A SINGLE API Assisted Function Set - see VBIMGAPI for these options

Public Function DoAPIOperations(ByVal OpCodes As Long, _
                        Optional ByVal Parm1 As Long = -1, _
                        Optional ByVal Parm2 As Long = -1) As Long
Attribute DoAPIOperations.VB_Description = "A number of basic image processing options are carried out by this Method: FLIP Vertically and Horizontally, Resize, Zoom, Unmap and Colormapping by MS methods"
                                        
  'A temporary kludge because I cant get StretchDIBits to work for FLIP_HORZ
  If (OpCodes And PIC_FLIP_HORZ) = PIC_FLIP_HORZ Then Call FlipHorz
  
  DoAPIOperations = APIOperations(zWidth, zHeight, _
                                  zPixBits(), zIntBPP, zCMap(), zNMapColors, _
                                  OpCodes, Parm1, Parm2)
  If DoAPIOperations <> 0 Then
    If (OpCodes And PIC_FLIP_VERT) = PIC_FLIP_VERT Then
      zPicState = (zPicState And (Not IS_TOP_TO_BOTTOM)) Or IS_TOP_TO_BOTTOM  'toggle it
    End If
    zPicState = zPicState And (Not IS_CMAPPED)        'we need to remap
    If zNMapColors > 1 Then zPicState = zPicState Or IS_CMAPPED
  End If
End Function

'===== COLORMAPPING IN PREPARATION FOR IMAGE SAVING ==========================================================

Public Function DoColorMapping() As Long
Attribute DoColorMapping.VB_Description = "In turn, create a colormap,and  remap pixels by dithering in preparation for saving"
  If (zPicState And IS_VALID_PIPELINE) = 0 Then GoTo ErrorFound       'must have a validated pipeline
  
  '--------------------------------------------------------------------------------------------------------
  'A special case for GIFs is where we have Loaded a BMP from DISK and it is upside down.
  'In this case zPicState has IS_TOP_TO_BOTTOM unset so we need to flip the BMP Image before we carry on further
  
  If (zPicType = PIC_GIF) And (zPicState And IS_TOP_TO_BOTTOM) = 0 Then
    If FlipVert() = 0 Then GoTo ErrorFound
    zPicState = zPicState Or IS_TOP_TO_BOTTOM
  End If
      
  If (zCMAPMode And PIC_USE_MS_CMAP) = PIC_USE_MS_CMAP Then     'can be true for GIFs or BMPs
    'There are cases where Operations may have unmapped the Image and it is now 24BPP
    'In this case we use the API to do MS style of remapping  'this is in VBIMGGDI.bas
    etCMAP = Timer
    If zReqBPP <> zIntBPP Then
      If MSMapColor(zReqBPP) = 0 Then GoTo ErrorFound
    End If
    etCMAP = Timer - etCMAP
  Else                  '1,4,8,16,24,32 from a MS Bitmap
    etCMAP = Timer
      
    If zIntBPP >= PIC_16BPP Then      'we havent got a colormap so need to make one or use a fixed one
      '-------------------- COLORMAP DETERMINATION --------------------------------------------------
      'we are going to remap the BMP onto a best fit palette. We will use the 5 most significant bits
      'so we are effectively going to select from a gamut of 32x32x32 or 32768 shades
      'if youve got the memory you could make this 6 deep or 64^3=262144 shades, youll also have to up HistCmap
      'NOTE:::: MSBPP is the format given, BitsPerPixel is what we want to achieve

      If zCMAPMode = PIC_FIXED_CMAP_GREY Then
        Call GenGreyMap(zCMap(), zNMapColors, zReqBPP)
      ElseIf zCMAPMode = PIC_FIXED_CMAP_USER Then
        Call GenUserCMap(zCMap(), zNMapColors, zUserCMAP())
      ElseIf (zCMAPMode And PIC_FIXED_CMAP) = PIC_FIXED_CMAP Then
        Call GenFixedMap(zCMap(), zNMapColors, zReqBPP, zCMAPMode)
      Else   'we are going to make one ourselves (hooray)
        If zReqBPP >= PIC_16BPP Then zNMapColors = 256 Else zNMapColors = 2 ^ zReqBPP
        ReDim zCMap(0 To 3 * zNMapColors - 1) As Byte
      
        'Initialize Remapping variables  - Using OctTree Quantization    'PixelWidth=significant bits BPP
        'also the colormap will be resized to the minimum power of 2 greater or equal to the number of colors used
        zReqBPP = CreateOctTreeCMap(zWidth, zHeight, zPixBits(), zIntBPP, _
                                            zCMap(), zNMapColors, OctTreeDepthBits:=5&)
        'Cleanup after ourselves to minimise storage
        Call DeleteOctTreeCMap
      End If
      etCMAP = Timer - etCMAP
      
      '-------------------------- PIXEL REMAPPING  ---------------------------------------------------
      etRemap = Timer
      If zReqBPP = 0 Then    ' something screwed up in OctTreeCmap
        GoTo ErrorFound
      Else                        ' Remap the new Palette onto the Image by dithering - image is resized
        Select Case zDitherMode
          Case PIC_DITHER_FS1, PIC_DITHER_FS2, PIC_DITHER_FS3:
            Call FSDitherMapColors(zWidth, zHeight, _
                                   zPixBits(), zIntBPP, _
                                   zCMap(), zNMapColors, _
                                   zCMAPMode, zDitherMode)
          
          Case 0:
            Call SimpleMapColors(zWidth, zHeight, _
                                 zPixBits(), zIntBPP, _
                                 zCMap(), zNMapColors)
        
          Case Else:
            Call DitherMapColors(zWidth, zHeight, _
                                 zPixBits(), zIntBPP, _
                                 zCMap(), zNMapColors, _
                                 zCMAPMode, zDitherMode)
        End Select
      End If
      etRemap = Timer - etRemap
        
        'after colormapping has happened we are now a 8bit pixelmap resized properly
      zIntBPP = 8         ' we have made this into a newly colormapped 8BPP Bitmap
    End If
    'we now have a colormap in all cases, and the internal format is now PixelWidth 1,4 or 8 bits
    
    zPicState = zPicState Or IS_CMAPPED
  End If
  
  '----------------------------- GREY REMAPPING  -----------------------------------------------------------
  'we may have asked for a grey map - this is a cheap way to do it (but only if there is a colormap)
      
  If zColorMode = PIC_GREY Then
    If (zPicState And IS_CMAPPED) = IS_CMAPPED Then   'change the RGB palette to CIY greys
      Call CMaptoGrey(zCMap())
    Else                                                      'Grey requires individual pixels to be changed
      Call MakePixelsGrey(zWidth, zHeight, zPixBits(), zIntBPP)
    End If
  End If

  '-------------------------- PALETTE SHRINKING  ------------------------------------------------
  ' attempt to reduce the number of colors needed. This is done by counting their uses
  ' then shifting used ones to the front of the CMAP palette while at the same time fixing their
  ' references in the DIBits array
  
  If zPicType = PIC_GIF Then
    If (zPicState And IS_CMAPPED) = IS_CMAPPED Then
      If (zCMAPMode And PIC_SMALLEST_CMAP) = PIC_SMALLEST_CMAP Then
        If zIntBPP > 1 Then
          zReqBPP = ShrinkCMap(zPixBits(), zIntBPP, zCMap(), zReqBPP)  'no point if only 2 colours
        End If
      End If
    End If
  End If
  
  'For BMPs destined to be only 1BPP or 4BPP then we must pack it here
  If zPicType = PIC_BMP And zIntBPP = PIC_8BPP And (zReqBPP = PIC_1BPP Or zReqBPP = PIC_4BPP) Then
    Call PackBMPImage(zWidth, zHeight, zPixBits(), zIntBPP, zReqBPP)
  End If
  
  DoColorMapping = 1
  Exit Function
  
ErrorFound:
  DoColorMapping = 0
End Function

'================================================================================================================

Public Sub SetUserCMAP(UserCMap() As Byte)      'set aside a user prepared map - it should match NMapColors
Attribute SetUserCMAP.VB_Description = "Load a user supplied Colormap for remapping to. It must match the NCMapCOlors and be rgb tuples. Array size should be 3*NColors"
  Dim i As Long
  
  ReDim zUserCMAP(0 To UBound(UserCMap) - LBound(UserCMap)) As Byte
  For i = 0 To zUserCMAP
    zUserCMAP(i) = UserCMap(i + LBound(UserCMap))
  Next
End Sub

'=================================================================================================================

Public Function LoadAsBMP(Path As String) As Long
Attribute LoadAsBMP.VB_Description = "Load a BMP file into the image"
  zPicState = zPicState And IS_VALID_PIPELINE   'clear all other bits
  LoadAsBMP = LoadBMP(Path, zWidth, zHeight, zPixBits(), zReqBPP, zCMap(), zNMapColors)
  If LoadAsBMP = 1 Then
    If zNMapColors <> 0 Then zPicState = zPicState Or IS_CMAPPED
    If zHeight < 0 Then
      zHeight = -zHeight
      zPicState = zPicState Or IS_TOP_TO_BOTTOM
    End If
  End If
End Function

Public Function SaveAsBMP(Path As String) As Long
Attribute SaveAsBMP.VB_Description = "Save internal image as a BMP file"
  SaveAsBMP = SaveBMP(Path, zWidth, zHeight, zPixBits(), zReqBPP, zCMap(), zNMapColors)
End Function

Public Function SaveAsGIF(Path As String) As Long
Attribute SaveAsGIF.VB_Description = "Save internal image as a GIF file"
  If (zPicState And IS_CMAPPED) = 0 Then
    SaveAsGIF = 0
  Else
    SaveAsGIF = SaveGIF(Path, zWidth, zHeight, zReqBPP, zPixBits(), zIntBPP, zCMap(), 0)
  End If
End Function

Public Function SaveAsInterlacedGIF(Path As String) As Long
Attribute SaveAsInterlacedGIF.VB_Description = "Save internal image as an interlaced GIF  file"
  If (zPicState And IS_CMAPPED) = 0 Then
    SaveAsInterlacedGIF = 0
  Else
    SaveAsInterlacedGIF = SaveGIF(Path, zWidth, zHeight, zReqBPP, zPixBits(), zIntBPP, zCMap(), PIC_INTERLACED)
  End If
End Function


