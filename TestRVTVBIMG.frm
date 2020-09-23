VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "TestRVTVBIMG - Image Saving Pipeline Example - Use 640x480 images if possible"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   11520
   ScaleWidth      =   15360
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTestPat2 
      Caption         =   "TestPattern2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   55
      Top             =   9690
      Width           =   1305
   End
   Begin VB.CheckBox chkInvert 
      Caption         =   "Invert"
      Height          =   375
      Left            =   3870
      TabIndex        =   54
      Top             =   7350
      Width           =   885
   End
   Begin VB.CheckBox chkZoom 
      Caption         =   "2/3"
      Height          =   375
      Left            =   2910
      TabIndex        =   53
      Top             =   7350
      Width           =   825
   End
   Begin VB.CheckBox chkFlipH 
      Caption         =   "FlipH"
      Height          =   375
      Left            =   1950
      TabIndex        =   52
      Top             =   7350
      Width           =   825
   End
   Begin VB.CheckBox chkFlipV 
      Caption         =   "FlipV"
      Height          =   375
      Left            =   990
      TabIndex        =   51
      Top             =   7350
      Width           =   825
   End
   Begin VB.CheckBox chkClip 
      Caption         =   "Clip?"
      Height          =   375
      Left            =   4890
      TabIndex        =   47
      Top             =   7350
      Width           =   765
   End
   Begin MSComDlg.CommonDialog dlgFileLoad 
      Left            =   4020
      Top             =   10410
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load Picture"
      Filter          =   "*gif|*jpg|*jpeg|*.bmp"
      InitDir         =   "AppDir"
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   345
      Left            =   13980
      TabIndex        =   29
      Top             =   3900
      Width           =   1365
   End
   Begin VB.Frame frDither 
      Caption         =   "Dither"
      ForeColor       =   &H004A3010&
      Height          =   3825
      Left            =   13980
      TabIndex        =   28
      Top             =   60
      Width           =   1365
      Begin VB.OptionButton OptDither 
         Caption         =   "FS Equal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   60
         TabIndex        =   50
         Top             =   3480
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "FS Even"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   60
         TabIndex        =   39
         Top             =   3210
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "FS Odd"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   38
         Top             =   2940
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Vertical"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   60
         TabIndex        =   37
         Top             =   2550
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Horizontal"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   36
         Top             =   2250
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Bwd Diag"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   60
         TabIndex        =   35
         Top             =   1890
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Fwd Diag"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   60
         TabIndex        =   34
         Top             =   1620
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Halftone"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   60
         TabIndex        =   33
         Top             =   1230
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Ordered"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   60
         TabIndex        =   32
         Top             =   960
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "Binary"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   31
         Top             =   690
         Width           =   1245
      End
      Begin VB.OptionButton OptDither 
         Caption         =   "None"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   30
         Top             =   330
         Width           =   1245
      End
   End
   Begin VB.Frame frCMAP 
      Caption         =   "CMAP"
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   12900
      TabIndex        =   22
      Top             =   60
      Width           =   1065
      Begin VB.OptionButton OptCMAP 
         Caption         =   "Fixed Grey"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   90
         TabIndex        =   49
         Top             =   3630
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "MS256"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   90
         TabIndex        =   48
         Top             =   3120
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "iNet"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   2640
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "VGA"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   90
         TabIndex        =   26
         Top             =   2160
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "Fixed Color"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   90
         TabIndex        =   25
         Top             =   1620
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "New"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   1050
         Width           =   885
      End
      Begin VB.OptionButton OptCMAP 
         Caption         =   "MS Map"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame frNColors 
      Caption         =   "NColors"
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   11880
      TabIndex        =   11
      Top             =   60
      Width           =   1005
      Begin VB.OptionButton OptNColors 
         Caption         =   "16M"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   60
         TabIndex        =   21
         Top             =   3660
         Width           =   855
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "65536"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   60
         TabIndex        =   20
         Top             =   3270
         Width           =   855
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "256"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   60
         TabIndex        =   19
         Top             =   2850
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "128"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   60
         TabIndex        =   18
         Top             =   2490
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "64"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   60
         TabIndex        =   17
         Top             =   2130
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "32"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   16
         Top             =   1770
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "16"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   60
         TabIndex        =   15
         Top             =   1410
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   60
         TabIndex        =   14
         Top             =   1020
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   60
         TabIndex        =   13
         Top             =   660
         Width           =   735
      End
      Begin VB.OptionButton OptNColors 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Frame frColorMode 
      Caption         =   "ColorMode"
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   10650
      TabIndex        =   7
      Top             =   60
      Width           =   1215
      Begin VB.OptionButton optColorMode 
         Caption         =   "Grey"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   2700
         Width           =   945
      End
      Begin VB.OptionButton optColorMode 
         Caption         =   "B and W"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   1035
      End
      Begin VB.OptionButton optColorMode 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   750
         Width           =   885
      End
   End
   Begin VB.Frame frPicType 
      Caption         =   "PicType"
      ForeColor       =   &H004A3010&
      Height          =   4155
      Left            =   9690
      TabIndex        =   4
      Top             =   60
      Width           =   945
      Begin VB.OptionButton optPicType 
         Caption         =   "GIF"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   2040
         Width           =   645
      End
      Begin VB.OptionButton optPicType 
         Caption         =   "BMP"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   1260
         Width           =   675
      End
   End
   Begin VB.PictureBox picResult 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   5700
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   3
      Top             =   4260
      Width           =   9660
   End
   Begin VB.CommandButton cmdLoadPic 
      Caption         =   "LoadPic"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   2
      Top             =   8490
      Width           =   1305
   End
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   1
      Top             =   30
      Width           =   9660
   End
   Begin VB.CommandButton cmdTestPat 
      Caption         =   "TestPattern"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   390
      TabIndex        =   0
      Top             =   9090
      Width           =   1305
   End
   Begin VB.Label lblElapsed 
      Caption         =   "Elapsed"
      Height          =   285
      Left            =   2370
      TabIndex        =   46
      Top             =   9480
      Width           =   1725
   End
   Begin VB.Label lblSaveTime 
      Caption         =   "Save"
      Height          =   285
      Left            =   2370
      TabIndex        =   45
      Top             =   9150
      Width           =   1725
   End
   Begin VB.Label lblDitherTime 
      Caption         =   "Remap"
      Height          =   285
      Left            =   2370
      TabIndex        =   44
      Top             =   8820
      Width           =   1725
   End
   Begin VB.Label lblCMAPTime 
      Caption         =   "CMAP"
      Height          =   285
      Left            =   2370
      TabIndex        =   43
      Top             =   8490
      Width           =   1725
   End
   Begin VB.Label Label3 
      Caption         =   "Timings (sec)"
      Height          =   345
      Left            =   2580
      TabIndex        =   42
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblResult 
      Caption         =   "Result"
      Height          =   285
      Left            =   5010
      TabIndex        =   41
      Top             =   9390
      Width           =   645
   End
   Begin VB.Label lblOriginal 
      Caption         =   "Original"
      Height          =   285
      Left            =   60
      TabIndex        =   40
      Top             =   7350
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'- ©2001 Ron van Tilburg - All rights reserved  1.01.2001
'- Amateur reuse is permitted subject to Copyright notices being retained and Credits to author being quoted.
'- Commercial use not permitted - email author please

'TESTRVTVBIMG.frm
'This program is just a tester for the DLL which is where the real work gets done -
'however it may suffice as a starting point for a Image processing application - serious submissions welcome

' Require References to RVTVBGDI , and Components to COMMDLG32.OCX

Const VALID_BMP_COLOR_OPTIONS As Integer = &H7           '111
Const VALID_BMP_NCOLOR_OPTIONS As Integer = &H389        '11,1000,1001
Const VALID_BMP_CMAP_OPTIONS As Integer = &H1            '00,0001
Const VALID_BMP_DITHER_OPTIONS As Integer = &H1          '000,0000,0001

Const VALID_GIF_COLOR_OPTIONS As Integer = &H7           '111
Const VALID_GIF_NCOLOR_OPTIONS As Integer = &HFF         '00,1111,1111
Const VALID_GIF_CMAP_OPTIONS As Integer = &H3F           '11,1111
Const VALID_GIF_DITHER_OPTIONS As Integer = &H7FF        '111,1111,1111

Const VALID_BW_NCOLOR_OPTIONS As Integer = &H1           '00,0000,0001
Const VALID_GIF_BW_CMAP_OPTIONS As Integer = &H5         '000,0101
Const VALID_GIF_C4_CMAP_OPTIONS As Integer = &H4         '000,0100
Const VALID_GIF_C8_CMAP_OPTIONS As Integer = &H46        '100,0110
Const VALID_GIF_XX_CMAP_OPTIONS As Integer = &H4F        '100,1111   '16
Const VALID_GIF_YY_CMAP_OPTIONS As Integer = &H46        '100,0110   '32,64,128
Const VALID_GIF_ZZ_CMAP_OPTIONS As Integer = &H77        '111,0111   '256

Const NPICTYPE_OPTIONS As Integer = 2
Const NCOLOR_OPTIONS  As Integer = 3
Const NNCOLOR_OPTIONS As Integer = 10
Const NCMAP_OPTIONS As Integer = 7
Const NDITHER_OPTIONS As Integer = 11

Const TEMPFILE As String = "C:\TEMP\zzz.zzz"

Dim CurPicType As Integer
Dim CurColorMode As Integer
Dim CurNColors As Integer
Dim CurCMAPMode As Integer
Dim CurDitherMode As Integer

Dim FileDir As String
Dim FileName As String

Private Sub Form_Load()
  CurPicType = -1:  CurColorMode = -1:  CurNColors = -1:   CurCMAPMode = -1:  CurDitherMode = -1
  Call picOriginal.ZOrder(0)
  Call optPicType_Click(0)
End Sub

Private Sub cmdLoadPic_Click()
  Dim q() As String
  
  Call Form1.picOriginal.ZOrder(0)
  With dlgFileLoad
    .DialogTitle = "Select Picture File"
    .CancelError = False
    If FileDir = "" Then .InitDir = App.Path Else .InitDir = FileDir
    .FileName = FileName
    'ToDo: set the flags and attributes of the common dialog control
    .Filter = "GIF Files (*.gif)|*.gif|JPG Files (*.jp*)|*.jp*|BMP Files (*.bmp)|*.bmp"
    .FilterIndex = 2
    .ShowOpen
    If Len(.FileName) <> 0 Then
      q = Split(.FileName, "\")
      FileName = q(UBound(q))
      q(UBound(q)) = ""
      FileDir = Join$(q, "\")
      Form1.picOriginal.Picture = LoadPicture(.FileName)
    End If
  End With
End Sub

Private Sub cmdTestPat_Click()
  Call Form1.picOriginal.ZOrder(0)
  Form1.picOriginal.Picture = LoadPicture(App.Path & "\TestPattern.gif")
End Sub

Private Sub cmdTestPat2_Click()
  Call Form1.picOriginal.ZOrder(0)
  Form1.picOriginal.Picture = LoadPicture(App.Path & "\TestPattern2.bmp")
End Sub

'THis is the Guts Of it - we call the various Dll functions in turn
Private Sub cmdGo_Click()
  Dim PicType As Long
  Dim ColorMode As Long
  Dim BitsPerPixel As Long
  Dim CMAPMode As Long
  Dim DitherMode As Long
  Dim PICCB As cRVTVBIMG, z() As Variant
  
  Dim tget As Single, tcmap As Single, tremap As Single, telapsed As Single
  Dim rc As Long, OpCodes As Long, Parm1 As Long, Parm2 As Long
  
  Set PICCB = New cRVTVBIMG '("RVTVBIMG.cRVTVBIMG")
  
  'Lets Work out all of the parameters of the main (and only) call
  
  z = Array(PIC_BMP, PIC_GIF, PIC_GIF_LACED)
  PicType = z(CurPicType)
  
  z = Array(PIC_COLOR, PIC_BW, PIC_GREY)
  ColorMode = z(CurColorMode)
  
  z = Array(PIC_1BPP, PIC_2BPP, PIC_3BPP, PIC_4BPP, PIC_5BPP, PIC_6BPP, PIC_7BPP, PIC_8BPP, PIC_16BPP, PIC_24BPP)
  BitsPerPixel = z(CurNColors)
  
  z = Array(PIC_USE_MS_CMAP, PIC_OPTIMAL_CMAP, PIC_FIXED_CMAP, PIC_FIXED_CMAP_VGA, PIC_FIXED_CMAP_INET, PIC_FIXED_CMAP_MS256, PIC_FIXED_CMAP_GREY)
  CMAPMode = z(CurCMAPMode)
  If CMAPMode = PIC_FIXED_CMAP Then CMAPMode = CMAPMode Or (2& ^ BitsPerPixel)
  
  z = Array(PIC_DITHER_NONE, PIC_DITHER_BIN, PIC_DITHER_ORD, PIC_DITHER_HTC, _
            PIC_DITHER_FDIAG, PIC_DITHER_BDIAG, PIC_DITHER_HORZ, PIC_DITHER_VERT, _
            PIC_DITHER_FS1, PIC_DITHER_FS2, PIC_DITHER_FS3)
  DitherMode = z(CurDitherMode)
  
  Call picOriginal.ZOrder(0): DoEvents
  Me.MousePointer = vbHourglass
  OpCodes = 0
  If chkFlipV.Value = 1 Then OpCodes = OpCodes Or PIC_FLIP_VERT
  If chkFlipH.Value = 1 Then OpCodes = OpCodes Or PIC_FLIP_HORZ
  If chkZoom.Value = 1 Then OpCodes = OpCodes Or PIC_IMAGE_ZOOM: Parm1 = -12247: Parm2 = -12247 'sqrt(2/3)scaled*10000
  If chkInvert.Value = 1 Then OpCodes = OpCodes Or PIC_INVERT_COLOR
  
  If OpCodes = 0 Then 'we do it the easy way
    If chkClip.Value = False Then
      rc = PICCB.SaveObjDCClip(picOriginal, TEMPFILE, PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode)
    Else
      rc = PICCB.SaveObjDCClip(picOriginal, TEMPFILE, PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode, _
                               200, 200, 639, 479)
    End If
  Else        'we do it a step at a time
    rc = PICCB.SetPipeline(PicType, ColorMode, BitsPerPixel, CMAPMode, DitherMode)
    If rc <> 0 Then
      If chkClip.Value = False Then
        rc = PICCB.ImageFromObjDCCLip(picOriginal)
      Else
        rc = PICCB.ImageFromObjDCCLip(picOriginal, 200, 200, 639, 479)
      End If
    End If
    If rc <> 0 And OpCodes <> 0 Then rc = PICCB.DoAPIOperations(OpCodes, Parm1, Parm2)
    If rc <> 0 Then rc = PICCB.DoColorMapping()
    If rc <> 0 Then If PicType = PIC_BMP Then rc = PICCB.SaveAsBMP(TEMPFILE) Else rc = PICCB.SaveAsGIF(TEMPFILE)
  End If
  Me.MousePointer = vbDefault
  
  If rc <> 0 Then
    Call picResult.ZOrder(0)
    picResult.Picture = LoadPicture(TEMPFILE)
    lblCMAPTime.Caption = "CMAP        " & Format$(PICCB.etCMAP, "0.0s")
    lblDitherTime.Caption = "Remap       " & Format$(PICCB.etRemap, "0.0s")
    lblSaveTime.Caption = "Save         " & Format$(PICCB.etSave, "0.0s")
    lblElapsed.Caption = "Elapsed    " & Format$(PICCB.etElapsed, "0.0s")
    Kill TEMPFILE
  Else
    MsgBox "Something has gone wrong - but that may already be obvious - sorry"
  End If
  
  Set PICCB = Nothing
End Sub

Private Sub optPicType_Click(Index As Integer)
  If Index <> CurPicType Then
    If CurPicType <> -1 Then optPicType(CurPicType).Value = False
    CurPicType = Index
    optPicType(CurPicType).Value = True
    
    CurColorMode = -1:  CurNColors = -1:   CurCMAPMode = -1:  CurDitherMode = -1
    If CurPicType = 0 Then
      Call SetColorDefaults(VALID_BMP_COLOR_OPTIONS, 0)
    Else
      Call SetColorDefaults(VALID_GIF_COLOR_OPTIONS, 0)
    End If
  End If
End Sub

Private Sub optColorMode_Click(Index As Integer)
  Dim ValidNColors As Integer, NewNColors As Integer
  
  If Index <> CurColorMode Then
    If CurColorMode <> -1 Then optColorMode(CurColorMode).Value = False
    CurColorMode = Index
    optColorMode(CurColorMode).Value = True
  
    Select Case CurColorMode
      Case 1:
        ValidNColors = VALID_BW_NCOLOR_OPTIONS: NewNColors = 0
      Case Else:
        If CurPicType = 0 Then
          ValidNColors = VALID_BMP_NCOLOR_OPTIONS: NewNColors = 7
        Else
          ValidNColors = VALID_GIF_NCOLOR_OPTIONS: NewNColors = 7
        End If
    End Select
    Call SetNColorDefaults(ValidNColors, NewNColors)
  End If
End Sub

Private Sub optNColors_Click(Index As Integer)
  Dim ValidCMaps As Integer, NewCMap As Integer
  
  If Index <> CurNColors Then
    If CurNColors <> -1 Then OptNColors(CurNColors).Value = False
    CurNColors = Index
    OptNColors(CurNColors).Value = True
    
    Select Case CurNColors
      Case 0:       ValidCMaps = VALID_GIF_BW_CMAP_OPTIONS: NewCMap = 2  '2
      Case 1:       ValidCMaps = VALID_GIF_C4_CMAP_OPTIONS: NewCMap = 2  '4
      Case 2:       ValidCMaps = VALID_GIF_C8_CMAP_OPTIONS: NewCMap = 1  '8
      Case 3:       ValidCMaps = VALID_GIF_XX_CMAP_OPTIONS: NewCMap = 1  '16
      Case 4, 5, 6: ValidCMaps = VALID_GIF_YY_CMAP_OPTIONS: NewCMap = 1  '32,64,128
      Case 7:       ValidCMaps = VALID_GIF_ZZ_CMAP_OPTIONS: NewCMap = 1  '256
      Case 8, 9:    ValidCMaps = VALID_BMP_CMAP_OPTIONS: NewCMap = 0
    End Select
    If CurPicType = 0 Then NewCMap = 0
    Call SetCMAPDefaults(ValidCMaps, NewCMap)
  End If
End Sub

Private Sub OptCMAP_Click(Index As Integer)
  If Index <> CurCMAPMode Then
    If CurCMAPMode <> -1 Then OptCMAP(CurCMAPMode).Value = False
    CurCMAPMode = Index
    OptCMAP(CurCMAPMode).Value = True
    If CurCMAPMode = 0 Then 'If CurPicType = 0 Or CurCMAPMode = 0 Then
      Call SetDitherDefaults(VALID_BMP_DITHER_OPTIONS, 0)
    Else
      Call SetDitherDefaults(VALID_GIF_DITHER_OPTIONS, 0)
    End If
  End If
End Sub

Private Sub OptDither_Click(Index As Integer)
  If Index <> CurDitherMode Then
    If CurDitherMode <> -1 Then OptDither(CurDitherMode).Value = False
    CurDitherMode = Index
    OptDither(CurDitherMode).Value = True
  End If
End Sub

Private Sub SetColorDefaults(ColorOptions As Integer, NewColor As Integer)
                         
  Dim i As Integer, j As Integer
  
  If ColorOptions <> -1 Then
    j = 1
    For i = 0 To NCOLOR_OPTIONS - 1
      optColorMode(i).Value = False
      If (ColorOptions And j) <> 0 Then
        optColorMode(i).Enabled = True
      Else
        optColorMode(i).Enabled = False
      End If
      j = j + j
    Next
    If NewColor <> -1 Then CurColorMode = -1
  End If
  If NewColor <> -1 Then Call optColorMode_Click(NewColor)
End Sub

Private Sub SetNColorDefaults(NColorOptions As Integer, NewNColor As Integer)
                         
  Dim i As Integer, j As Integer
      
  If NColorOptions <> -1 Then
    j = 1
    For i = 0 To NNCOLOR_OPTIONS - 1
      OptNColors(i).Value = False
      If (NColorOptions And j) <> 0 Then
        OptNColors(i).Enabled = True
      Else
        OptNColors(i).Enabled = False
      End If
      j = j + j
    Next
    If NewNColor <> -1 Then CurNColors = -1
  End If
  If NewNColor <> -1 Then Call optNColors_Click(NewNColor)
End Sub

Private Sub SetCMAPDefaults(CMapOptions As Integer, NewCMap As Integer)
                         
  Dim i As Integer, j As Integer
      
  If CMapOptions <> -1 Then
    j = 1
    For i = 0 To NCMAP_OPTIONS - 1
      OptCMAP(i).Value = False
      If (CMapOptions And j) <> 0 Then
        OptCMAP(i).Enabled = True
      Else
        OptCMAP(i).Enabled = False
      End If
      j = j + j
    Next
    If NewCMap <> -1 Then CurCMAPMode = -1
  End If
  If NewCMap <> -1 Then Call OptCMAP_Click(NewCMap)
End Sub

Private Sub SetDitherDefaults(DitherOptions As Integer, NewDither As Integer)
                         
  Dim i As Integer, j As Integer
      
  If DitherOptions <> -1 Then
    j = 1
    For i = 0 To NDITHER_OPTIONS - 1
      OptDither(i).Value = False
      If (DitherOptions And j) <> 0 Then
        OptDither(i).Enabled = True
      Else
        OptDither(i).Enabled = False
      End If
      j = j + j
    Next
    If NewDither <> -1 Then CurDitherMode = -1
  End If
  If NewDither <> -1 Then Call OptDither_Click(NewDither)
End Sub

