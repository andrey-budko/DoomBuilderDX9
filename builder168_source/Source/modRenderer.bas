Attribute VB_Name = "modRenderer"
'
'    Doom Builder
'    Copyright (c) 2003 Pascal vd Heiden, www.codeimp.com
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


'Export Picture border size
Private Const EXPORTPICTURE_BORDER As Long = 20

'Palette organization
Public Const PALETTE_16COLORS_OFFSET As Long = 32
Public Const PALETTE_16COLORSDIMMED_OFFSET As Long = 48
Private Const PALETTE_MIX_THINGSCOLOR As Single = 0.4
Private Const PALETTE_MIX_ORIGINALCOLOR As Single = 0.4
Private Const PALETTE_MIX_SELECTIONCOLOR As Single = 0.8
Public Enum ENUM_PALETTECOLORS
     CLR_BACKGROUND
     CLR_VERTEX
     CLR_VERTEXSELECTED
     CLR_VERTEXHIGHLIGHT
     CLR_LINE
     CLR_LINEDOUBLE
     CLR_LINESPECIAL
     CLR_LINESPECIALDOUBLE
     CLR_LINESELECTED
     CLR_LINEHIGHLIGHT
     CLR_LINEDRAG
     CLR_THINGTAG
     CLR_SECTORTAG
     CLR_THINGUNKNOWN
     CLR_THINGSELECTED
     CLR_THINGHIGHLIGHT
     CLR_MULTISELECT
     CLR_GRID
     CLR_GRID64
     CLR_LINEBLOCKSOUND
     CLR_MAPBOUNDARY
End Enum
Public Enum ENUM_PALETTES
     PAL_NORMAL = 0
     PAL_MULTISELECTION = 64
     PAL_THINGSELECTION = 128
     PAL_BACKGROUND = 192
End Enum

'Thing images
Public Enum ENUM_THINGIMAGES
     TI_ARROW0
     TI_ARROW45
     TI_ARROW90
     TI_ARROW135
     TI_ARROW180
     TI_ARROW225
     TI_ARROW270
     TI_ARROW315
     TI_DOT
     TI_UNKNOWN
End Enum

'Bitmap File Header
Public Type BITMAPFILEHEADER
     bfType As String * 2  'Integer
     bfSize As Long
     bfReserved1 As Integer
     bfReserved2 As Integer
     bfOffBits As Long
End Type

'Bitmap Info Header
Public Type BITMAPINFOHEADER
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

'Bitmap RGB Data
Public Type BITMAPRGB
     rgbBlue As Byte
     rgbGreen As Byte
     rgbRed As Byte
     rgbReserved As Byte
End Type

'Array bounds descriptor
Public Type SAFEARRAYBOUND
     cElements As Long
     lLbound As Long
End Type

'1 dimensions array descriptor
Public Type SAFEARRAY1D
     cDims As Integer
     fFeatures As Integer
     cbElements As Long
     cLocks As Long
     pvData As Long
     Bounds(0 To 0) As SAFEARRAYBOUND
End Type

'2 dimensions array descriptor
'Public Type SAFEARRAY2D
'     cDims As Integer
'     fFeatures As Integer
'     cbElements As Long
'     cLocks As Long
'     pvData As Long
'     Bounds(0 To 1) As SAFEARRAYBOUND
'End Type

'DC Bitmap header
Public Type DCBITMAPHEADER
     bmType As Long
     bmWidth As Long
     bmHeight As Long
     bmWidthBytes As Long
     bmPlanes As Integer
     bmBitsPixel As Integer
     bmBits As Long
End Type


'API Declarations
Public Declare Function GetObjectAPI Lib "gdi32.dll" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateBitmap Lib "gdi32.dll" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (Destination As Any, ByVal length As Long, ByVal Fill As Byte)
Public Declare Sub Render_Init Lib "builder.dll" (ByRef scdata As Byte, ByVal scwidth As Long, ByVal scheight As Long)
Public Declare Sub Render_Term Lib "builder.dll" ()
Public Declare Sub Render_Scale Lib "builder.dll" (ByVal left As Single, ByVal top As Single, ByVal Zoom As Single)
Public Declare Sub Render_Clear Lib "builder.dll" (ByVal c As Byte)
Public Declare Sub Render_Line Lib "builder.dll" (ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal c As Byte)
Public Declare Sub Render_LineSwitched Lib "builder.dll" (ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal high As Byte)
Public Declare Sub Render_LinedefLine Lib "builder.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal c As Byte, ByVal indicatorlength As Long)
Public Declare Sub Render_LinedefLineSwitched Lib "builder.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal high As Byte, ByVal indicatorlength As Long)
Public Declare Sub Render_DottedLine Lib "builder.dll" (ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal c As Byte)
Public Declare Sub Render_Box Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByVal diameter As Long, ByVal c As Byte, ByVal Fill As Long, ByVal fill_c As Byte)
Public Declare Sub Render_BoxSwitched Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByVal diameter As Long, ByVal high As Byte, ByVal Fill As Long, ByVal fill_high As Long)
Public Declare Sub Render_RectSwitched Lib "builder.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal high As Byte, ByVal BorderSize As Long)
Public Declare Sub Render_Bitmap Lib "builder.dll" (ByRef Bitmap As Byte, ByVal width As Long, ByVal height As Long, ByVal sourcex As Long, ByVal sourcey As Long, ByVal sourcewidth As Long, ByVal sourceheight As Long, ByVal targetx As Long, ByVal targety As Long, ByVal color1 As Byte, ByVal color2 As Byte)
Public Declare Sub Render_AllLinedefs Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal submode As Long, ByVal indicatorlength As Long)
Public Declare Sub Render_AllVertices Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal vertexsize As Long)
Public Declare Sub Render_AllThings Lib "builder.dll" (ByRef things As MAPTHING, ByVal StartIndex As Long, ByVal EndIndex As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal outlines As Long, ByVal outlinezoom As Single, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_AllThingsDarkened Lib "builder.dll" (ByRef things As MAPTHING, ByVal StartIndex As Long, ByVal EndIndex As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_TaggedLinedefs Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal argtag As Long, ByVal argmark As Long, ByVal c As Byte, ByVal indicatorlength As Long, ByVal rendervertices As Long, ByVal vertexsize As Long)
Public Declare Sub Render_TaggedSectors Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByVal ptr_sectors As Long, ByVal numsectors As Long, ByVal numlinedefs As Long, ByVal sectortag As Long, ByVal c As Byte, ByVal indicatorlength As Long, ByVal rendervertices As Long, ByVal vertexsize As Long)
Public Declare Sub Render_ChangingLengths Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByRef changelines As Long, ByVal numchangelines As Long, ByRef Bitmap As Byte, ByVal width As Long, ByVal height As Long, ByVal charwidth As Long, ByVal charheight As Long)
Public Declare Sub Render_NumberSwitched Lib "builder.dll" (ByVal number As Long, ByVal x As Long, ByVal y As Long, ByRef Bitmap As Byte, ByVal width As Long, ByVal height As Long, ByVal charwidth As Long, ByVal charheight As Long, ByVal palette1 As Byte, ByVal palette2 As Byte)
Public Declare Sub Render_TaggedThings Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByVal thingtag As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal outlines As Long, ByVal outlinezoom As Single, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_TaggedThingsNormal Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByVal thingtag As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal outlines As Long, ByVal outlinezoom As Single, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_TaggedArgThings Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByVal argtag As Long, ByVal argmark As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal outlines As Long, ByVal outlinezoom As Single, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_TaggedArgThingsNormal Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByVal argtag As Long, ByVal argmark As Long, ByRef thingbitmaps As Byte, ByVal bitmapswidth As Long, ByVal imagesize As Long, ByVal outlines As Long, ByVal outlinezoom As Single, ByVal filterthings As Long, ByRef Filter As THINGFILTERS)
Public Declare Sub Render_AllImpassableLinedefs Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal StartIndex As Long, ByVal EndIndex As Long, ByVal indicatorlength As Long)


'Current viewport
Public ViewLeft As Long
Public ViewTop As Long
Public ViewZoom As Single

'Map bitmap memory
Public ScreenPalette(0 To 255) As BITMAPRGB
Private ScreenDescriptor As SAFEARRAY1D
Private ScreenData() As Byte
Public ScreenWidth As Long
Public ScreenHeight As Long
Public ScreenTarget As PictureBox

'Bitmap pointers
Public ThingDescriptor As SAFEARRAY1D
Public ThingBitmapData() As Byte
Public NumbersDescriptor As SAFEARRAY1D
Public NumbersBitmapData() As Byte

Public Function GetImageFormat(ByRef pdata As String, ByVal FlatCandidate As Boolean, Optional ByRef width As Long, Optional ByRef height As Long) As Long
     Dim Columns As Long, Rows As Long
     Dim LastPtr As Long
     Dim Illegal As Boolean
     Dim BitmapHeader As BITMAPFILEHEADER
     Dim BitmapInfo As BITMAPINFOHEADER
     Dim i As Long
     
     'Unknown yet
     GetImageFormat = TF_UNKNOWN
     
     'IMG Image Specs
     '
     'Pos  Type     Description                   Recognize Value
     '1   <short>   Width
     '3   <short>   Height
     '5   <short>   X Offset
     '7   <short>   Y Offset
     '9   <long>    Start address of column 1     => (9+Width*4) and < Len(Data)
     '13  <long>    Start address of column 2     => (9+Width*4) and < Len(Data)
     '17  <long>    Start address of column 3     => (9+Width*4) and < Len(Data)
     '...
     '
     
     'PNG Specs
     '
     'Signature:    137, P, N, G, 13, 10, 26, 10
     'Header:       I, H, D, R, 4 bytes width, 4 bytes height, ...
     '
     
     'Check if long enough to check PNG
     If (Len(pdata) > 20) Then
          
          'Check for PNG Signature
          If (Asc(Mid$(pdata, 1)) = 137) And _
             (Mid$(pdata, 2, 1) = "P") And (Mid$(pdata, 3, 1) = "N") And (Mid$(pdata, 4, 1) = "G") And _
             (AscW(Mid$(pdata, 5)) = 13) And _
             (AscW(Mid$(pdata, 6)) = 10) And _
             (AscW(Mid$(pdata, 7)) = 26) And _
             (AscW(Mid$(pdata, 8)) = 10) Then
               
               'Format is PNG
               GetImageFormat = TF_PNG
               
               'Set width and height
               width = CVL(Mid$(pdata, 13, 4))
               height = CVL(Mid$(pdata, 17, 4))
          End If
     End If
     
     'Check if long enough to check BMP
     If (GetImageFormat = TF_UNKNOWN) And (Len(pdata) > Len(BitmapHeader) + Len(BitmapInfo)) Then
          
          'Fill structure with data
          CopyMemory BitmapHeader, ByVal pdata, Len(BitmapHeader)
          
          'Check type
          If (BitmapHeader.bfType = "BM") Then
               
               'Get the bitmap info from data
               CopyMemory BitmapInfo, ByVal Mid$(pdata, Len(BitmapHeader) + 1, Len(BitmapInfo)), Len(BitmapInfo)
               
               'Check the bitdepth
               If (BitmapInfo.biPlanes = 1) And (BitmapInfo.biBitCount = 8) Then
                    
                    '8 bit paletted
                    GetImageFormat = TF_BITMAP_P8
                    
                    'Set width and height
                    width = BitmapInfo.biWidth
                    height = BitmapInfo.biHeight
                    
               'Check the bitdepth
               ElseIf (BitmapInfo.biPlanes = 1) And (BitmapInfo.biBitCount = 16) Then
                    
                    '16 bit
                    GetImageFormat = TF_BITMAP_B5G6R5
                    
                    'Set width and height
                    width = BitmapInfo.biWidth
                    height = BitmapInfo.biHeight
                    
               'Check the bitdepth
               ElseIf (BitmapInfo.biPlanes = 1) And (BitmapInfo.biBitCount = 24) Then
                    
                    '24 bit
                    GetImageFormat = TF_BITMAP_B8G8R8
                    
                    'Set width and height
                    width = BitmapInfo.biWidth
                    height = BitmapInfo.biHeight
                    
               'Check the bitdepth
               ElseIf (BitmapInfo.biPlanes = 1) And (BitmapInfo.biBitCount = 32) Then
                    
                    '32 bit
                    GetImageFormat = TF_BITMAP_A8B8G8R8
                    
                    'Set width and height
                    width = BitmapInfo.biWidth
                    height = BitmapInfo.biHeight
               End If
          End If
     End If
     
     'If still unknown, check for Image format
     If (GetImageFormat = TF_UNKNOWN) Then
          
          'Get info
          Columns = CVI(Mid$(pdata, 1, 2))
          Rows = CVI(Mid$(pdata, 3, 2))
          LastPtr = (8 + Columns * 4)
          
          'Check if Width and Height as valid
          If (Columns > 0) And (Rows > 0) Then
               
               For i = 1 To Columns
                    
                    'Check if still within file
                    If ((9 + i * 4) < Len(pdata)) Then
                         
                         'Check if pointer is too low
                         If (CVL(Mid$(pdata, 5 + i * 4, 4)) < LastPtr) Then Illegal = True: Exit For
                         
                         'Check if pointer is too high
                         If (CVL(Mid$(pdata, 5 + i * 4, 4)) >= Len(pdata)) Then Illegal = True: Exit For
                         
                         'Last pointed here
                         'LastPtr = CVL(Mid$(pdata, 5 + i * 4, 4))
                         'CodeImp 6/19/2005: Doom specs dont say anything about the order of
                         'the columns so allow the columns to be in any order
                    Else
                         
                         'Illegal
                         Illegal = True: Exit For
                    End If
               Next i
               
               'Check if IMG is not considered illegal
               If (Illegal = False) Then
                    
                    'Legal image
                    GetImageFormat = TF_IMAGE
                    
                    'Set width and height
                    width = Columns
                    height = Rows
               End If
          End If
     End If
     
     'If still unknown, check for Flat format (only when Flat Candidate)
     If (GetImageFormat = TF_UNKNOWN) And (FlatCandidate = True) Then
          
          'Check if square
          If (Sqr(Len(pdata)) = Int(Sqr(Len(pdata)))) Then
               
               'Flat format
               GetImageFormat = TF_FLAT
               
               'Set width and height
               width = Sqr(Len(pdata))
               height = width
               
          'Check if this flat is larger than 4096
          ElseIf (Len(pdata) > 4096) Then
               
               'Flat format
               GetImageFormat = TF_FLAT
               
               'Set width and height
               width = 64
               height = 64
          End If
     End If
End Function




Public Function BITMAPRGBToLong(ByRef Color As BITMAPRGB) As Long
     
     'Make long color
     BITMAPRGBToLong = Color.rgbBlue Or (Color.rgbGreen * (2 ^ 8)) Or (Color.rgbRed * (2 ^ 16))
End Function

Public Function ColorValueToLong(ByRef Color As D3DCOLORVALUE) As Long
     
     'Make long color
     ColorValueToLong = (Color.b * 255) Or (Color.g * (2 ^ 8) * 255) Or (Color.r * (2 ^ 16) * 255)
End Function


Public Function BITMAPRGBToWinLong(ByRef Color As BITMAPRGB) As Long
     
     'Make long color
     BITMAPRGBToWinLong = Color.rgbRed Or (Color.rgbGreen * (2 ^ 8)) Or (Color.rgbBlue * (2 ^ 16))
End Function


Public Function ConvertPNGtoBitmap(ByRef Data As String) As String
     Dim TempFilebuffer As Integer
     Dim TempFilename As String
     Dim TempBitmapFilename As String
     Dim GDIBitmap As clsGDIBitmap
     
     'Create GDI Bitmap object
     Set GDIBitmap = New clsGDIBitmap
     
     'Make filenames
     TempBitmapFilename = App.Path & "\convert.bmp"
     TempFilename = App.Path & "\convert.png"
     
     'Make a temporary file
     TempFilebuffer = FreeFile
     Open TempFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     Put #TempFilebuffer, 1, Data
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Load the PNG picture from file
     GDIBitmap.LoadFromFile TempFilename
     
     'Get the data in Bitmap format
     GDIBitmap.SaveToFile TempBitmapFilename, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.bmp")), 0
     
     'Open new temporary file
     TempFilebuffer = FreeFile
     Open TempBitmapFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     ConvertPNGtoBitmap = Space$(LOF(TempFilebuffer))
     Get #TempFilebuffer, 1, ConvertPNGtoBitmap
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Clean up
     Set GDIBitmap = Nothing
     Kill TempFilename
     Kill TempBitmapFilename
End Function


Public Function ConvertJPGtoBitmap(ByRef Data As String) As String
     Dim TempFilebuffer As Integer
     Dim TempFilename As String
     Dim TempBitmapFilename As String
     Dim GDIBitmap As clsGDIBitmap
     
     'Create GDI Bitmap object
     Set GDIBitmap = New clsGDIBitmap
     
     'Make filenames
     TempBitmapFilename = App.Path & "\convert.bmp"
     TempFilename = App.Path & "\convert.jpg"
     
     'Make a temporary file
     TempFilebuffer = FreeFile
     Open TempFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     Put #TempFilebuffer, 1, Data
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Load the JPG picture from file
     GDIBitmap.LoadFromFile TempFilename
     
     'Get the data in Bitmap format
     GDIBitmap.SaveToFile TempBitmapFilename, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.bmp")), 0
     
     'Open new temporary file
     TempFilebuffer = FreeFile
     Open TempBitmapFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     ConvertJPGtoBitmap = Space$(LOF(TempFilebuffer))
     Get #TempFilebuffer, 1, ConvertJPGtoBitmap
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Clean up
     Set GDIBitmap = Nothing
     Kill TempFilename
     Kill TempBitmapFilename
End Function



Public Function ConvertGIFtoBitmap(ByRef Data As String) As String
     Dim TempFilebuffer As Integer
     Dim TempFilename As String
     Dim TempBitmapFilename As String
     Dim GDIBitmap As clsGDIBitmap
     
     'Create GDI Bitmap object
     Set GDIBitmap = New clsGDIBitmap
     
     'Make filenames
     TempBitmapFilename = App.Path & "\convert.bmp"
     TempFilename = App.Path & "\convert.gif"
     
     'Make a temporary file
     TempFilebuffer = FreeFile
     Open TempFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     Put #TempFilebuffer, 1, Data
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Load the GIF picture from file
     GDIBitmap.LoadFromFile TempFilename
     
     'Get the data in Bitmap format
     GDIBitmap.SaveToFile TempBitmapFilename, GDIBitmap.EncoderGuid(GDIBitmap.ExtensionExists("*.bmp")), 0
     
     'Open new temporary file
     TempFilebuffer = FreeFile
     Open TempBitmapFilename For Binary As #TempFilebuffer
     
     'Dump the data to a file
     ConvertGIFtoBitmap = Space$(LOF(TempFilebuffer))
     Get #TempFilebuffer, 1, ConvertGIFtoBitmap
     
     'Close the temporary file
     Close #TempFilebuffer
     
     'Clean up
     Set GDIBitmap = Nothing
     Kill TempFilename
     Kill TempBitmapFilename
End Function




Public Sub CreateBitmapPointer(ByRef source As PictureBox, ByRef Data() As Byte, ByRef Descriptor As SAFEARRAY1D)
     Dim BitmapDC As DCBITMAPHEADER
     
     'Get the DC bitmap header
     GetObjectAPI source.Picture, Len(BitmapDC), BitmapDC
     
     'Verify that the bitmap colordepth is 8 bits
     If (BitmapDC.bmPlanes <> 1) Or (BitmapDC.bmBitsPixel <> 8) Then
          
          'Show error
          MsgBox "Error in CreateBitmapPointer: " & "Cannot create data pointer, bitmap is not 256 color paletted!", vbCritical
     End If
     
     'Create the bitmap array info
     With Descriptor
         .cbElements = 1
         .cDims = 1
         .Bounds(0).lLbound = 0
         .Bounds(0).cElements = BitmapDC.bmHeight * BitmapDC.bmWidthBytes
         .pvData = BitmapDC.bmBits
     End With
     
     'Set the pointer for direct memory access
     CopyMemory ByVal VarPtrArray(Data), VarPtr(Descriptor), 4
End Sub

Private Function CreateExportPicture(ByRef MapZoom As Single, ByRef MapRect As RECT) As Boolean
     On Error GoTo errorhandler
     Dim ZoomWidth As Single, ZoomHeight As Single
     Dim PictureWidth As Long, PictureHeight As Long
     Dim PictureDesc As SAFEARRAY1D
     
     'Calculate map rect
     MapRect = CalculateMapRect
     
     'Check if making picture by size or mapscale
     If (frmExportPicture.optResolution.Value = True) Then
          
          'Check if the rect has a size > 0
          If (Abs(MapRect.right - MapRect.left) > 0) And (Abs(MapRect.bottom - MapRect.top) > 0) Then
               
               'Calculate both horizontal and vertical scale
               ZoomWidth = (frmExportPicture.txtWidth.Value - EXPORTPICTURE_BORDER * 2) / Abs(MapRect.right - MapRect.left)
               ZoomHeight = (frmExportPicture.txtHeight.Value - EXPORTPICTURE_BORDER * 2) / Abs(MapRect.bottom - MapRect.top)
               
               'Use the smallest
               If (ZoomWidth < ZoomHeight) Then MapZoom = ZoomWidth Else MapZoom = ZoomHeight
          Else
               
               'No size, scale 100%
               MapZoom = 1
          End If
          
          'Set the picture size
          PictureWidth = frmExportPicture.txtWidth.Value
          PictureHeight = frmExportPicture.txtHeight.Value
     Else
          
          'Resolution to given scale
          MapZoom = frmExportPicture.txtScale.Value / 100
          
          'Set the picture size
          PictureWidth = Abs(MapRect.right - MapRect.left) * MapZoom + EXPORTPICTURE_BORDER * 2
          PictureHeight = Abs(MapRect.bottom - MapRect.top) * MapZoom + EXPORTPICTURE_BORDER * 2
     End If
     
     'Erase anything on render target
     Set frmMain.picTexture.Picture = Nothing
     
     'Resize the render target
     frmMain.picTexture.width = PictureWidth
     frmMain.picTexture.height = PictureHeight
     
     'Create render target picture
     InitializeMapRenderer frmMain.picTexture
     
     'Leave
     CreateExportPicture = True
     Exit Function
     
     
errorhandler:
     
     'Erase anything on render target
     Set frmMain.picTexture.Picture = Nothing
     
     'No picture created
     CreateExportPicture = False
End Function

Public Function CreatePalettedBitmap(ByRef BitmapPalette() As BITMAPRGB, ByRef width As Long, ByRef height As Long, Optional defaultcolor As Byte) As StdPicture
     Dim TempFilebuffer As Integer
     Dim TempFilename As String
     Dim BitmapHeader As BITMAPFILEHEADER
     Dim BitmapInfo As BITMAPINFOHEADER
     Dim BitmapData() As Byte
     Dim RowsPadding As Long
     
     'Make the file header standards
     With BitmapHeader
          .bfType = "BM"
          .bfOffBits = 1078
     End With
     
     'Make the bitmap header standards
     With BitmapInfo
          .biBitCount = 8
          .biClrUsed = 256
          .biPlanes = 1
          .biSize = Len(BitmapInfo)
     End With
     
     'Calculate padding for data
     'We'll make the width a few pixels lagers to align with 32 bits and eliminate padding
     RowsPadding = 4 - (width Mod 4)
     If RowsPadding = 4 Then RowsPadding = 0
     
     'Update the width argument with the padding
     width = width + RowsPadding
     
     'Modify the width and height of bitmap to the screen
     BitmapInfo.biWidth = width
     BitmapInfo.biHeight = height
     BitmapInfo.biSizeImage = width * height
     
     'Allocate memory to write to file
     ReDim BitmapData(1 To BitmapInfo.biWidth, 1 To BitmapInfo.biHeight)
     
     'Set memory with default color
     FillMemory BitmapData(1, 1), BitmapInfo.biSizeImage, defaultcolor
     
     'Make a temporary file
     TempFilename = App.Path & "\renderer.tmp"
     TempFilebuffer = FreeFile
     Open TempFilename For Binary As #TempFilebuffer
     
     'Write the headers and bitmap data
     Put #TempFilebuffer, , BitmapHeader
     Put #TempFilebuffer, , BitmapInfo
     Put #TempFilebuffer, , BitmapPalette
     Put #TempFilebuffer, , BitmapData
     
     'Set length of the file in the header structure
     BitmapHeader.bfSize = Seek(TempFilebuffer) - 1
     
     'Overwrite the header with the new size
     Put #TempFilebuffer, 1, BitmapHeader
     
     'Close the file
     Close #TempFilebuffer
     
     'Load the file in the screen
     Set CreatePalettedBitmap = LoadPicture(TempFilename)
     
     'Remove the temporary file
     Kill TempFilename
End Function

Public Sub CreateRendererPalette()
     On Local Error Resume Next
     Dim i As Long
     
     'The 256 color palette is divided in 4 parts:
     '0 - 63    = Normal map colors & thing colors
     '64 - 127  = Transparent mutlti selection colors
     '128 - 191 = Transparent thing selection colors
     '192 - 255 = Background color
     
     'Make the palette from configuration
     If Not Config.Exists("palette") Then Config.Add "palette", New Dictionary
     ScreenPalette(CLR_BACKGROUND) = LongToBITMAPRGB(Config("palette")("CLR_BACKGROUND"))
     ScreenPalette(CLR_VERTEX) = LongToBITMAPRGB(Config("palette")("CLR_VERTEX"))
     ScreenPalette(CLR_VERTEXSELECTED) = LongToBITMAPRGB(Config("palette")("CLR_VERTEXSELECTED"))
     ScreenPalette(CLR_VERTEXHIGHLIGHT) = LongToBITMAPRGB(Config("palette")("CLR_VERTEXHIGHLIGHT"))
     ScreenPalette(CLR_LINE) = LongToBITMAPRGB(Config("palette")("CLR_LINE"))
     ScreenPalette(CLR_LINEDOUBLE) = LongToBITMAPRGB(Config("palette")("CLR_LINEDOUBLE"))
     ScreenPalette(CLR_LINESPECIAL) = LongToBITMAPRGB(Config("palette")("CLR_LINESPECIAL"))
     ScreenPalette(CLR_LINESPECIALDOUBLE) = LongToBITMAPRGB(Config("palette")("CLR_LINESPECIALDOUBLE"))
     ScreenPalette(CLR_LINESELECTED) = LongToBITMAPRGB(Config("palette")("CLR_LINESELECTED"))
     ScreenPalette(CLR_LINEHIGHLIGHT) = LongToBITMAPRGB(Config("palette")("CLR_LINEHIGHLIGHT"))
     ScreenPalette(CLR_LINEDRAG) = LongToBITMAPRGB(Config("palette")("CLR_LINEDRAG"))
     ScreenPalette(CLR_THINGTAG) = LongToBITMAPRGB(Config("palette")("CLR_THINGTAG"))
     ScreenPalette(CLR_SECTORTAG) = LongToBITMAPRGB(Config("palette")("CLR_SECTORTAG"))
     ScreenPalette(CLR_THINGUNKNOWN) = LongToBITMAPRGB(Config("palette")("CLR_THINGUNKNOWN"))
     ScreenPalette(CLR_THINGSELECTED) = LongToBITMAPRGB(Config("palette")("CLR_THINGSELECTED"))
     ScreenPalette(CLR_THINGHIGHLIGHT) = LongToBITMAPRGB(Config("palette")("CLR_THINGHIGHLIGHT"))
     ScreenPalette(CLR_MULTISELECT) = LongToBITMAPRGB(Config("palette")("CLR_MULTISELECT"))
     ScreenPalette(CLR_GRID) = LongToBITMAPRGB(Config("palette")("CLR_GRID"))
     ScreenPalette(CLR_GRID64) = LongToBITMAPRGB(Config("palette")("CLR_GRID64"))
     ScreenPalette(CLR_LINEBLOCKSOUND) = LongToBITMAPRGB(Config("palette")("CLR_LINEBLOCKSOUND"))
     ScreenPalette(CLR_MAPBOUNDARY) = LongToBITMAPRGB(Config("palette")("CLR_MAPBOUNDARY"))
     
     'Make the 16 color palette
     ScreenPalette(PALETTE_16COLORS_OFFSET + 0) = RGBToBITMAPRGB(72, 72, 72)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 1) = RGBToBITMAPRGB(0, 0, 144)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 2) = RGBToBITMAPRGB(0, 120, 0)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 3) = RGBToBITMAPRGB(0, 120, 120)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 4) = RGBToBITMAPRGB(120, 0, 0)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 5) = RGBToBITMAPRGB(120, 0, 120)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 6) = RGBToBITMAPRGB(120, 84, 0)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 7) = RGBToBITMAPRGB(144, 144, 144)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 8) = RGBToBITMAPRGB(96, 96, 96)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 9) = RGBToBITMAPRGB(72, 72, 216)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 10) = RGBToBITMAPRGB(72, 216, 72)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 11) = RGBToBITMAPRGB(72, 216, 216)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 12) = RGBToBITMAPRGB(216, 72, 72)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 13) = RGBToBITMAPRGB(216, 72, 216)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 14) = RGBToBITMAPRGB(216, 216, 72)
     ScreenPalette(PALETTE_16COLORS_OFFSET + 15) = RGBToBITMAPRGB(216, 216, 216)
     
     'Make the 16 dimmed colors
     For i = 0 To 15
          ScreenPalette(PALETTE_16COLORSDIMMED_OFFSET + i) = DimmedThingColor(ScreenPalette(PALETTE_16COLORS_OFFSET + i), ScreenPalette(CLR_BACKGROUND))
     Next i
     
     'Make the second part of the palette selection colors
     For i = PAL_MULTISELECTION To PAL_MULTISELECTION + 63
          ScreenPalette(i) = MixedSelectionColor(ScreenPalette(i - PAL_MULTISELECTION), ScreenPalette(CLR_MULTISELECT))
     Next i
     
     'Make the third part of the palette selection colors
     For i = PAL_THINGSELECTION To PAL_THINGSELECTION + 63
          ScreenPalette(i) = MixedSelectionColor(ScreenPalette(i - PAL_THINGSELECTION), ScreenPalette(CLR_THINGHIGHLIGHT))
     Next i
     
     'Make the fourth part of the palette selection colors
     For i = PAL_BACKGROUND To PAL_BACKGROUND + 63
          ScreenPalette(i) = ScreenPalette(CLR_BACKGROUND)
     Next i
End Sub

Public Sub CreateMaskPicture(ByRef SourcePicture As PictureBox, ByRef TargetMask As PictureBox)
     Dim btmp As Long
     Dim chdc As Long
     Dim x As Long
     
     'Make the target same in size as the source
     TargetMask.width = SourcePicture.width
     TargetMask.height = SourcePicture.height
     
     'Create the mask bitmap
     btmp = CreateBitmap(SourcePicture.ScaleWidth, SourcePicture.ScaleHeight, 1, 1, 0)
     chdc = CreateCompatibleDC(SourcePicture.hdc)
     x = SelectObject(chdc, btmp)
     BitBlt chdc, 0, 0, SourcePicture.ScaleWidth, SourcePicture.ScaleHeight, SourcePicture.hdc, 0, 0, vbSrcCopy
     BitBlt TargetMask.hdc, 0, 0, SourcePicture.ScaleWidth, SourcePicture.ScaleHeight, chdc, 0, 0, vbSrcCopy
     TargetMask.Refresh
     btmp = SelectObject(chdc, x)
     DeleteObject btmp
     DeleteDC chdc
End Sub


Public Sub DestroyBitmapPointer(ByRef Data() As Byte)
     
     'This will release the direct memory access pointer
     CopyMemory ByVal VarPtrArray(Data), 0&, 4
End Sub

Public Sub DetermineRenderScreenSize(ByRef Rendertarget As PictureBox)
     
     'Measure in pixels
     Rendertarget.ScaleMode = vbPixels
     
     'Keep the map screen width and height
     ScreenWidth = Rendertarget.ScaleWidth
     ScreenHeight = Rendertarget.ScaleHeight
End Sub

Private Function DimmedThingColor(ByRef Original As BITMAPRGB, ByRef Background As BITMAPRGB) As BITMAPRGB
     Dim nR As Long
     Dim nG As Long
     Dim nB As Long
     
     'Create mixed color
     nR = Original.rgbRed * PALETTE_MIX_THINGSCOLOR + Background.rgbRed * (1 - PALETTE_MIX_THINGSCOLOR)
     nG = Original.rgbGreen * PALETTE_MIX_THINGSCOLOR + Background.rgbGreen * (1 - PALETTE_MIX_THINGSCOLOR)
     nB = Original.rgbBlue * PALETTE_MIX_THINGSCOLOR + Background.rgbBlue * (1 - PALETTE_MIX_THINGSCOLOR)
     
     'Clip colors to byte range
     If (nR < 0) Then nR = 0
     If (nR > 255) Then nR = 255
     If (nG < 0) Then nG = 0
     If (nG > 255) Then nG = 255
     If (nB < 0) Then nB = 0
     If (nB > 255) Then nB = 255
     
     'Return colors
     With DimmedThingColor
          .rgbRed = nR
          .rgbGreen = nG
          .rgbBlue = nB
     End With
End Function

Public Sub InitializeMapRenderer(ByRef Rendertarget As PictureBox)
     Dim width As Long, height As Long
     
     'Make sure the renderer is terminated
     TerminateMapRenderer
     
     'Set background color
     If (mapfilename <> "") Then frmMain.picMap.BackColor = BITMAPRGBToWinLong(LongToBITMAPRGB(Val(Config("palette")("CLR_BACKGROUND"))))
     
     'Terminate last numbers pointer
     DestroyBitmapPointer NumbersBitmapData
     
     'Measure in pixels
     Rendertarget.ScaleMode = vbPixels
     
     'Check if we should set autoredraw (do we really need this?)
     frmMain.picMap.AutoRedraw = (Val(Config("autorerender")) <> 0)
     
     'Keep the map screen width and height
     ScreenWidth = Rendertarget.ScaleWidth
     ScreenHeight = Rendertarget.ScaleHeight
     
     'To use a palette and obtain direct memory access to the pixel data,
     'An image of the box's size must be loaded in the picturebox.
     width = Rendertarget.ScaleWidth
     height = Rendertarget.ScaleHeight
     Set Rendertarget.Picture = CreatePalettedBitmap(ScreenPalette, width, height, CLR_BACKGROUND)
     
     'Get the pointer for direct memory access
     CreateBitmapPointer Rendertarget, ScreenData, ScreenDescriptor
     
     'Get a pointer to the numbers bitmap
     CreateBitmapPointer frmMain.picNumbers, NumbersBitmapData, NumbersDescriptor
     
     'Pass the information on to the DLL
     Render_Init ScreenData(0), width, height
     
     'Keep the rendertarget for later use
     Set ScreenTarget = Rendertarget
End Sub

Public Function LongToBITMAPRGB(ByVal Color As Long) As BITMAPRGB
     With LongToBITMAPRGB
          .rgbRed = (Color And &HFF0000) / (2 ^ 16)
          .rgbGreen = (Color And &HFF00&) / (2 ^ 8)
          .rgbBlue = (Color And &HFF&)
     End With
End Function

Public Function LongToColorValue(ByVal Color As Long) As D3DCOLORVALUE
     With LongToColorValue
          .r = CSng((Color And &HFF0000) / (2 ^ 16)) / 255
          .g = CSng((Color And &HFF00&) / (2 ^ 8)) / 255
          .b = CSng(Color And &HFF&) / 255
     End With
End Function


Public Function LongToBGRLong(ByVal Color As Long) As Long
     LongToBGRLong = ((Color And &HFF0000) / (2 ^ 16)) Or (Color And &HFF00&) Or ((Color And &HFF&) * (2 ^ 16))
End Function


Private Function MixedSelectionColor(ByRef Original As BITMAPRGB, ByRef Selection As BITMAPRGB) As BITMAPRGB
     Dim nR As Long
     Dim nG As Long
     Dim nB As Long
     
     'Create mixed color
     nR = Original.rgbRed * PALETTE_MIX_ORIGINALCOLOR + Selection.rgbRed * PALETTE_MIX_SELECTIONCOLOR
     nG = Original.rgbGreen * PALETTE_MIX_ORIGINALCOLOR + Selection.rgbGreen * PALETTE_MIX_SELECTIONCOLOR
     nB = Original.rgbBlue * PALETTE_MIX_ORIGINALCOLOR + Selection.rgbBlue * PALETTE_MIX_SELECTIONCOLOR
     
     'Clip colors to byte range
     If (nR < 0) Then nR = 0
     If (nR > 255) Then nR = 255
     If (nG < 0) Then nG = 0
     If (nG > 255) Then nG = 255
     If (nB < 0) Then nB = 0
     If (nB > 255) Then nB = 255
     
     'Return colors
     With MixedSelectionColor
          .rgbRed = nR
          .rgbGreen = nG
          .rgbBlue = nB
     End With
End Function

Public Sub RedrawMap(Optional ByVal KeepCurrentHighlight As Boolean)
     Dim Indices As Variant
     Dim i As Long
     
     'Only render map when loaded
     If (mapfile <> "") And Not (ScreenTarget Is Nothing) And (mode <> EM_3D) Then
          
          'Clear map with background color
          Render_Clear CLR_BACKGROUND
          
          'Draw grid
          Render_Grid gridsizex, gridsizey, CLR_GRID
          If (gridsizex <= 64) Or (gridsizey <= 64) Then Render_Grid 64, 64, CLR_GRID64
          
          'Check the current mode
          Select Case mode
               
               Case EM_VERTICES
                    
                    If (Val(Config("modethings"))) Then Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
                    Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                    Render_AllVertices vertexes(0), 0, numvertexes - 1, vertexsize
                    If (submode = ESM_DRAGGING) Then
                         
                         'Render changing linedef lengths
                         Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                    End If
                    
               Case EM_LINES
                    
                    If (Val(Config("modethings"))) Then Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
                    Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                    If (Config("mode1vertices")) Then Render_AllVertices vertexes(0), 0, numvertexes - 1, vertexsize
                    If (submode = ESM_DRAWING) Then
                         
                         'Go for all selected vertices
                         Indices = selected.Items
                         For i = LBound(Indices) To UBound(Indices)
                              
                              'Redraw this vertex
                              Render_AllVertices vertexes(0), Indices(i), Indices(i), vertexsize
                         Next i
                         
                         'Render changing linedef lengths
                         Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                         
                    ElseIf (submode = ESM_DRAGGING) Then
                         
                         'Render changing linedef lengths
                         Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                    End If
                    
               Case EM_SECTORS, EM_MOVE
                    
                    If (Val(Config("modethings"))) Then Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
                    Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                    If (Config("mode2vertices")) Then Render_AllVertices vertexes(0), 0, numvertexes - 1, vertexsize
                    If (submode = ESM_DRAWING) Then
                         
                         'Go for all selected vertices
                         Indices = selected.Items
                         For i = LBound(Indices) To UBound(Indices)
                              
                              'Redraw this vertex
                              Render_AllVertices vertexes(0), Indices(i), Indices(i), vertexsize
                         Next i
                         
                         'Render changing linedef lengths
                         Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                         
                    ElseIf (submode = ESM_DRAGGING) Then
                         
                         'Render changing linedef lengths
                         Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numchangedlines, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                    End If
                    
               Case EM_THINGS
                    
                    'Render
                    Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
                    Render_AllThings things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), ViewZoom, filterthings, filtersettings
                    
          End Select
          
          'Highlight was reset (undrawn), reset its varaible
          If Not KeepCurrentHighlight Then currentselected = -1
          
          'Show new image
          frmMain.picMap.Refresh
     End If
End Sub

Public Sub Render_Grid(ByVal griddistancex As Long, ByVal griddistancey As Long, ByVal Color As ENUM_PALETTECOLORS)
     Dim i As Long
     Dim s As Long, e As Long
     Dim offset As Long
     
     'Check if grid should be rendered
     If (Config("gridshow")) Then
          
          'Only render grid if not screen-filling
          If (griddistancex * ViewZoom > 4) Then
               
               'Calculate offset in X
               offset = (gridx Mod griddistancex)
               
               'Determine horizontal start and end
               s = ((ScreenTarget.ScaleLeft - griddistancex - offset) \ griddistancex) * griddistancex + offset
               e = ((ScreenTarget.ScaleLeft + ScreenTarget.ScaleWidth + griddistancex - offset) \ griddistancex) * griddistancex + offset
               
               'Vertical Lines
               For i = s To e Step griddistancex
                    
                    'Draw line
                    Render_DottedLine i, -ScreenTarget.ScaleTop, i, -(ScreenTarget.ScaleTop + ScreenTarget.ScaleHeight), Color
               Next i
          End If
          
          'Only render grid if not screen-filling
          If (griddistancey * ViewZoom > 4) Then
               
               'Calculate offset in Y
               offset = (gridy Mod griddistancey)
               
               'Determine vertical start and end
               s = ((ScreenTarget.ScaleTop - griddistancey - offset) \ griddistancey) * griddistancey + offset
               e = ((ScreenTarget.ScaleTop + ScreenTarget.ScaleHeight + griddistancey - offset) \ griddistancey) * griddistancey + offset
               
               'Horizontal Lines
               For i = s To e Step griddistancey
                    
                    'Draw line
                    Render_DottedLine ScreenTarget.ScaleLeft, -i, ScreenTarget.ScaleLeft + ScreenTarget.ScaleWidth, -i, Color
               Next i
          End If
     End If
     
     'Render boundary lines
     Render_Line -32767, -32767, 32766, -32767, CLR_MAPBOUNDARY
     Render_Line 32766, -32767, 32766, 32766, CLR_MAPBOUNDARY
     Render_Line 32766, 32766, -32767, 32766, CLR_MAPBOUNDARY
     Render_Line -32767, 32766, -32767, -32767, CLR_MAPBOUNDARY
End Sub

Public Function RenderExportPicture() As Boolean
     Dim MapZoom As Single
     Dim MapRect As RECT
     Dim OrigViewZoom As Single
     Dim OrigViewLeft As Long
     Dim OrigViewTop As Long
     
     'Terminate map renderer
     TerminateMapRenderer
     
     'Change render target to picTexture with given dimensions
     If CreateExportPicture(MapZoom, MapRect) Then
          
          'Keep original viewport
          OrigViewZoom = ViewZoom
          OrigViewLeft = ViewLeft
          OrigViewTop = ViewTop
          
          'Set the viewport
          ChangeView MapRect.left - EXPORTPICTURE_BORDER / MapZoom, MapRect.top - EXPORTPICTURE_BORDER / MapZoom, MapZoom
          
          
          'Render the grid
          If (frmExportPicture.chkShowGrid.Value = vbChecked) Then
               
               'Render normal grid
               Render_Grid gridsizex, gridsizey, CLR_GRID
               
               'Render 64 mappixels grid
               If (frmExportPicture.chkGrid64.Value = vbChecked) Then Render_Grid 64, 64, CLR_GRID64
          End If
          
          
          'Render things dimmed?
          If (frmExportPicture.chkShowThings.Value = vbChecked) And _
             (frmExportPicture.chkThingDimmed.Value = vbChecked) Then
               
               'Render all things dimmed
               Render_AllThingsDarkened things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, filterthings, filtersettings
               
               'TODO: Coordinates?
               
          End If
          
          
          'Render lines
          If (frmExportPicture.chkShowLines.Value = vbChecked) Then
               
               'Turn off the line normals?
               If (frmExportPicture.chkShowLineNormals.Value = vbUnchecked) Then indicatorsize = 0
               
               'Render all linedefs
               Render_AllLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, submode, indicatorsize
               
               'Render all impassible linedefs
               Render_AllImpassableLinedefs vertexes(0), linedefs(0), 0, numlinedefs - 1, indicatorsize
               
               'Check if rendering lengths
               If (frmExportPicture.chkShowLengths.Value = vbChecked) Then
                    
                    'Make all lines "changing" so the lengths will be drawn
                    AllLinesChanging
                    
                    'Render linedef lengths
                    Render_ChangingLengths vertexes(0), linedefs(0), changedlines(0), numlinedefs, NumbersBitmapData(0), frmMain.picNumbers.width, frmMain.picNumbers.height, frmMain.picNumbers.width / 10, frmMain.picNumbers.height
                    
                    'Reset changing lines
                    ReDim changedlines(0)
                    numchangedlines = 0
               End If
          End If
          
          
          'Render vertices
          If (frmExportPicture.chkShowVertices.Value = vbChecked) Then
               
               'Render all vertices
               Render_AllVertices vertexes(0), 0, numvertexes - 1, (Val(frmExportPicture.txtVertexSize.Value) - 1) \ 2
               
               'TODO: Coordinates?
               
          End If
          
          
          'Render things bright?
          If (frmExportPicture.chkShowThings.Value = vbChecked) And _
             (frmExportPicture.chkThingDimmed.Value = vbUnchecked) Then
               
               'Render all things normal
               Render_AllThings things(0), 0, numthings - 1, ThingBitmapData(0), frmMain.picThings(thingsize).width, frmMain.picThings(thingsize).height, Val(Config("allthingsrects")), MapZoom, filterthings, filtersettings
               
               'TODO: Coordinates?
               
          End If
          
          
          'Terminate renderer
          TerminateMapRenderer
          
          'Restore renderer to map screen
          InitializeMapRenderer frmMain.picMap
          
          'Restore the viewport
          ChangeView OrigViewLeft, OrigViewTop, OrigViewZoom
          
          'Redraw entire map
          RedrawMap
          
          'Success
          RenderExportPicture = True
     Else
          
          'Show error
          MsgBox "An error occurred while creating the picture. Try making a smaller picture.", vbCritical
          
          'Failure
          RenderExportPicture = False
     End If
End Function

Public Function RGBToBITMAPRGB(ByVal r As Long, ByVal g As Long, ByVal b As Long) As BITMAPRGB
     With RGBToBITMAPRGB
          .rgbRed = r
          .rgbGreen = g
          .rgbBlue = b
     End With
End Function

Public Sub TerminateMapRenderer()
     
     'This will release the direct memory access pointer
     'CopyMemory ByVal VarPtrArray(ScreenData), 0&, 4
     DestroyBitmapPointer ScreenData
     
     'Do the same in the DLL
     Render_Term
     
     'Unreference rendertarget
     Set ScreenTarget = Nothing
End Sub

Public Function WinLongToBITMAPRGB(ByVal Color As Long) As BITMAPRGB
     With WinLongToBITMAPRGB
          .rgbBlue = (Color And &HFF0000) / (2 ^ 16)
          .rgbGreen = (Color And &HFF00&) / (2 ^ 8)
          .rgbRed = (Color And &HFF&)
     End With
End Function
