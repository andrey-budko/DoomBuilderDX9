VERSION 5.00
Begin VB.Form frmPictureImport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picture Import"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7965
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmPictureImport.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   531
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picOriginal 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   1050
      Left            =   270
      ScaleHeight     =   70
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   108
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   210
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6360
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   4740
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5340
      Width           =   1485
   End
   Begin VB.Frame fraTrans 
      Caption         =   " Transparency "
      Height          =   1635
      Left            =   2280
      TabIndex        =   5
      Top             =   4050
      Width           =   2025
      Begin DoomBuilder.ctlValueBox txtTransRange 
         Height          =   375
         Left            =   840
         TabIndex        =   9
         Top             =   1080
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         Max             =   100
      End
      Begin VB.CommandButton cmdTransColor 
         BackColor       =   &H00FFFF00&
         Height          =   315
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "CLR_VERTEX"
         Top             =   570
         Width           =   1485
      End
      Begin VB.Label lblTransRange 
         AutoSize        =   -1  'True
         Caption         =   "Range:"
         Height          =   210
         Left            =   255
         TabIndex        =   8
         Top             =   1140
         Width           =   510
      End
      Begin VB.Label lblTrans 
         AutoSize        =   -1  'True
         Caption         =   "Transparency color:"
         Height          =   210
         Left            =   255
         TabIndex        =   6
         Top             =   330
         Width           =   1470
      End
   End
   Begin VB.Frame fraColors 
      Caption         =   " Color Conversion "
      Height          =   1635
      Left            =   105
      TabIndex        =   1
      Top             =   4050
      Width           =   2025
      Begin VB.OptionButton optLighter 
         Caption         =   "Lighter colors"
         Height          =   255
         Left            =   255
         TabIndex        =   4
         Top             =   1050
         Width           =   1485
      End
      Begin VB.OptionButton optDarker 
         Caption         =   "Darker colors"
         Height          =   255
         Left            =   255
         TabIndex        =   3
         Top             =   735
         Width           =   1485
      End
      Begin VB.OptionButton optNearest 
         Caption         =   "Nearest colors"
         Height          =   255
         Left            =   255
         TabIndex        =   2
         Top             =   420
         Value           =   -1  'True
         Width           =   1485
      End
   End
   Begin VB.PictureBox picPreview 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      Height          =   3900
      Left            =   105
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   75
      Width           =   7740
      Begin VB.Timer tmrCreatePreview 
         Interval        =   1
         Left            =   2370
         Top             =   240
      End
      Begin VB.Label lblBusy 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Analyzing source picture..."
         Height          =   210
         Left            =   3007
         TabIndex        =   13
         Top             =   1680
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1950
      End
   End
   Begin VB.TextBox txtLumpName 
      Height          =   315
      Left            =   5700
      MaxLength       =   8
      TabIndex        =   15
      Top             =   4155
      Width           =   1515
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "0 x 0 pixels"
      Height          =   210
      Left            =   5730
      TabIndex        =   17
      Top             =   4575
      Width           =   825
   End
   Begin VB.Label lblPictureSize 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Picture size:"
      Height          =   210
      Left            =   4710
      TabIndex        =   16
      Top             =   4575
      Width           =   885
   End
   Begin VB.Label lblLumpName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Lump name:"
      Height          =   210
      Left            =   4725
      TabIndex        =   14
      Top             =   4185
      Width           =   870
   End
End
Attribute VB_Name = "frmPictureImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    XODE Multimedia Doom Builder
'    by Pascal 'gherkin' vd Heiden
'    gherkin@xodemultimedia.com
'    www.xodemultimedia.com
'
'    Copyright (c) 2003 XODE Multimedia
'    This program is released under GNU General Public License
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'


'Do not allow any undeclared variables
Option Explicit

'Case sensitive comparisions
Option Compare Binary


Private Sub AnalyzeOriginalPicture()
     Dim x As Long, y As Long, c As Long
     
     'Show busyness
     lblBusy.Visible = True
     Refresh
     
     'Get picture size
     TexturePreviewWidth = picOriginal.ScaleWidth
     TexturePreviewHeight = picOriginal.ScaleHeight
     
     'Show sizes
     lblSize.Caption = TexturePreviewWidth & " x " & TexturePreviewHeight & " pixels"
     
     'Allocate memory for picture
     ReDim TexturePreviewColors(0 To (TexturePreviewWidth * TexturePreviewHeight - 1))
     
     'Go for all pixels in original picture
     For y = 0 To TexturePreviewHeight - 1
          For x = 0 To TexturePreviewWidth - 1
               
               'Get pixel color
               TexturePreviewColors(c) = picOriginal.POINT(x, y)
               
               'Next pixel
               c = c + 1
          Next x
     Next y
     
     'Remove busyness
     lblBusy.Visible = False
End Sub

Private Sub CreatePreviewPicture()
     Dim x As Long, y As Long, c As Long
     Dim BitmapHeader As BITMAPFILEHEADER
     Dim BitmapInfo As BITMAPINFOHEADER
     Dim RGBColors() As BITMAPRGB
     Dim BitmapColors() As Byte
     Dim RowsPadding As Long
     Dim Method As Long
     Dim TempFile As String
     Dim FileBuffer As Integer
     
     'Allocate memory for new picture
     ReDim RGBColors(0 To (TexturePreviewWidth * TexturePreviewHeight - 1))
     
     'Determine method
     If (optNearest.Value) Then
          Method = 0
     ElseIf (optDarker.Value) Then
          Method = 1
     ElseIf (optLighter.Value) Then
          Method = 2
     End If
     
     'Check if transparency is allowed
     If (txtTransRange.Enabled) Then
          
          'Create example picture with transparency
          DrawPalettedExample TexturePreviewColors(0), RGBColors(0), TexturePreviewWidth * TexturePreviewHeight, TexturePreviewWidth, playpal(0), 256, cmdTransColor.backcolor, (Val(txtTransRange.Text) + 1) * (2.56), GetSysColor(WCOLOR_APPWORKSPACE), Method
     Else
          
          'Create example picture
          DrawPalettedExample TexturePreviewColors(0), RGBColors(0), TexturePreviewWidth * TexturePreviewHeight, TexturePreviewWidth, playpal(0), 256, 0, 0, 0, Method
     End If
     
     'Calculate padding for rows
     'Rows must be 32bit aligned
     RowsPadding = 4 - ((TexturePreviewWidth * 3) Mod 4)
     If RowsPadding = 4 Then RowsPadding = 0
     
     'Allocate memory for bitmap picture
     ReDim BitmapColors(1 To (TexturePreviewWidth * 3 + RowsPadding), 1 To TexturePreviewHeight)
     
     'Create bitmap file header
     With BitmapHeader
          .bfType = "BM"
          .bfSize = Len(BitmapHeader) + Len(BitmapInfo) + (CLng(TexturePreviewWidth) * 3 + RowsPadding) * CLng(TexturePreviewHeight)
          .bfOffBits = Len(BitmapHeader) + Len(BitmapInfo)
     End With
     
     'Create bitmap info
     With BitmapInfo
          .biSize = Len(BitmapInfo)
          .biWidth = TexturePreviewWidth
          .biHeight = TexturePreviewHeight
          .biPlanes = 1
          .biBitCount = 24    '24 bits per pixel
          .biCompression = 0  'None
          .biSizeImage = (CLng(TexturePreviewWidth) * 3 + RowsPadding) * CLng(TexturePreviewHeight)
     End With
     
     'Copy pixel colors
     For y = TexturePreviewHeight To 1 Step -1
          For x = 0 To TexturePreviewWidth - 1
               
               With RGBColors(c)
                    BitmapColors(x * 3 + 1, y) = .rgbBlue
                    BitmapColors(x * 3 + 2, y) = .rgbGreen
                    BitmapColors(x * 3 + 3, y) = .rgbRed
               End With
               
               'Next pixel
               c = c + 1
          Next x
     Next y
     
     
     'Create temp file
     TempFile = MakeTempFile(False)
     
     'Open the bitmap file
     FileBuffer = FreeFile
     Open TempFile For Binary As #FileBuffer
     
     'Write the bitmap
     Put #FileBuffer, , BitmapHeader
     Put #FileBuffer, , BitmapInfo
     Put #FileBuffer, , BitmapColors
     
     'Close the bitmap file
     Close #FileBuffer
     
     'Show the picture
     Set picPreview.Picture = LoadPicture(TempFile)
     
     'Kill the temp file
     Kill TempFile
End Sub

Private Sub cmdCancel_Click()
     tag = 0
     Hide
End Sub

Private Sub cmdOK_Click()
     tag = 1
     Hide
End Sub


Private Sub cmdTransColor_Click()
     Dim NewColor As Long
     
     'Select new color
     NewColor = SelectColor(Me.hWnd, cmdTransColor.backcolor, cdlCCFullOpen Or cdlCCRGBInit, CustomColors())
     
     'Check if not cancelled
     If (NewColor <> -1) Then
          
          'Set the new color on the button
          cmdTransColor.backcolor = NewColor
          
          'Recreate picture
          tmrCreatePreview.Enabled = False
          tmrCreatePreview.Enabled = True
     End If
End Sub

Private Sub Form_Resize()
     
     'Analyze original
     AnalyzeOriginalPicture
     
     'Recreate picture
     CreatePreviewPicture
End Sub


Private Sub optDarker_Click()
     
     'Recreate picture
     tmrCreatePreview.Enabled = False
     tmrCreatePreview.Enabled = True
End Sub

Private Sub optLighter_Click()
     
     'Recreate picture
     tmrCreatePreview.Enabled = False
     tmrCreatePreview.Enabled = True
End Sub


Private Sub optNearest_Click()
     
     'Recreate picture
     tmrCreatePreview.Enabled = False
     tmrCreatePreview.Enabled = True
End Sub


Private Sub tmrCreatePreview_Timer()
     
     'Disable timer
     tmrCreatePreview.Enabled = False
     
     'Set the interval for next time
     tmrCreatePreview.Interval = 100
     
     'Create preview now
     CreatePreviewPicture
End Sub

Private Sub txtLumpName_Change()
     cmdOK.Enabled = (Trim$(txtLumpName) <> "")
End Sub

Private Sub txtLumpName_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtTransRange_Change()
     
     'Recreate picture
     tmrCreatePreview.Enabled = False
     tmrCreatePreview.Enabled = True
End Sub


