VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



'Patch structure
Private Type PATCHDESC
     x As Long                     'Textures only
     y As Long                     'Textures only
     width As Long                 'Flats only
     height As Long                'Flats only
     Index As Long
     File As ENUM_IMAGESOURCE
     Format As ENUM_IMAGEFORMAT
End Type


'Patches
Private Patches() As PATCHDESC
Private NumPatches As Long

'General properties
Public Name As String
Public height As Long
Public width As Long
Public IsLoaded As Boolean
Public FlatCandidate As Boolean
Public ScaleX As Single
Public ScaleY As Single
Public Flags As Long
Public d3dwidth As Long
Public d3dheight As Long
Public d3dscalewidth As Single          'd3dscalewidth x d3dwidth = width
Public d3dscaleheight As Single         'd3dscaleheight x d3dheight = height

'Direct3D cached texture (all Windows versions)
Public D3DTexture As Direct3DTexture8

'Caching
Private Bitmap As StdPicture  '<-- Windows 2000 and newer caches here (faster, but costs GDI)
Private BitmapData() As Byte  '<-- Windows NT, 98 and ME caches here (slower, but GDI is limited so that's not an option)
Private BitmapWidth As Long
Private BitmapHeight As Long


Public Function Picture(Optional ByVal NoCaching As Boolean) As StdPicture
     Dim UnloadCache As Boolean
     Dim TempData() As Byte
     Dim TempDesc As SAFEARRAY1D
     
     'Check if the image must be loaded
     If (IsLoaded = False) Then
          
          'Load the image now
          LoadImage
          
          'Remove it from cache afterwards if not chaching
          If (NoCaching) Then UnloadCache = True
     End If
     
     'Check if now loaded
     If IsLoaded Then
          
          'Check if we can use GDI objects (faster)
          If RunningWindows2000 Then
               
               'Return the bitmap
               Set Picture = Bitmap
          Else
               
               'Resize the texture picturebox
               frmMain.picTexture.width = width
               frmMain.picTexture.height = height
               
               'Create a paletted bitmap to render image on
               Set frmMain.picTexture.Picture = CreatePalettedBitmap(playpal, BitmapWidth, BitmapHeight, TRANSPARENCY_INDEX)
               
               'Get a pointer to the bitmap memory
               CreateBitmapPointer frmMain.picTexture, TempData(), TempDesc
               
               'Copy the image
               CopyMemory TempData(0), BitmapData(0), BitmapWidth * BitmapHeight
               
               'Destroy pointer
               DestroyBitmapPointer TempData()
               
               'Copy the image in its original size (without the padding)
               Set Picture = frmMain.picTexture.image
          End If
     End If
     
     'Remove the bitmap if not caching
     If ((Config("texturecaching") = 0) Or (UnloadCache)) Then
          
          'Clean up cache
          Set Bitmap = Nothing
          Erase BitmapData()
          IsLoaded = False
     End If
End Function


Public Function LoadImage() As Boolean
     On Local Error GoTo BadImage
     Dim BitmapDesc As SAFEARRAY1D
     Dim BitmapPointer() As Byte
     Dim lumpdata As String
     Dim lumpindex As Long
     Dim DataFormat As ENUM_IMAGEFORMAT
     Dim OldMousePointer As Long
     Dim PatchesDrawn As Long
     Dim ImageCreated As Boolean
     Dim p As Long
     Dim s As Single
     
     'Change mousepointer
     OldMousePointer = Screen.MousePointer
     If (Screen.MousePointer = vbDefault) Then Screen.MousePointer = vbArrowHourglass
     
     'Go for all patches in this image
     For p = 0 To (NumPatches - 1)
          
          'No data yet
          lumpdata = ""
          
          'Check if the patch is not marked as invalid
          If (Patches(p).Format <> TF_INVALID) Then
               
               'Check where this patch resides
               Select Case Patches(p).File
                    
                    'Load the lump from MapWAD file
                    Case TS_MAPWAD: lumpdata = MapWAD.GetLump(Patches(p).Index)
                    
                    'Load the lump from AddWAD file
                    Case TS_ADDWAD: lumpdata = AddWAD.GetLump(Patches(p).Index)
                    
                    'Load the lump from IWAD file
                    Case TS_IWAD: lumpdata = IWAD.GetLump(Patches(p).Index)
                    
                    'Load the data from textures directory
                    Case TS_TEXDIR: lumpdata = GetTextureFileData(Name)
                    
                    'Load the data from flats directory
                    Case TS_FLATDIR: lumpdata = GetFlatFileData(Name)
                    
               End Select
               
               'Check if patch is not zero length
               If (Len(lumpdata) > 0) Then
                    
                    'Check if the patch format must be determined
                    If (Patches(p).Format = TF_UNKNOWN) Then Patches(p).Format = GetImageFormat(lumpdata, FlatCandidate)
                    
                    'Check if the format is known or could be detected
                    If (Patches(p).Format <> TF_UNKNOWN) Then
                         
                         'Copy the data format (data is a copy as well)
                         DataFormat = Patches(p).Format
                         
                         'PNG Format must be converted and rechecked
                         If (DataFormat = TF_PNG) Then
                              
                              'Convert PNG to Bitmap
                              lumpdata = ConvertPNGtoBitmap(lumpdata)
                              
                              'Recheck format to get the type of Bitmap
                              DataFormat = GetImageFormat(lumpdata, FlatCandidate)
                         End If
                         
                         'When patch size is unknown, take it from data
                         If (Patches(p).width = 0) And (Patches(p).height = 0) Then
                              
                              'Get Width and Height from image data if possible
                              GetImageFormat lumpdata, FlatCandidate, Patches(p).width, Patches(p).height
                              
                              'When the image sizes are unknown, take them from patch
                              If (width = 0) And (height = 0) Then
                                   width = Patches(p).width
                                   height = Patches(p).height
                              End If
                         End If
                         
                         'Only continue when image size is known
                         If (width > 0) And (height > 0) Then
                              
                              'Create image if not yet created
                              If (ImageCreated = False) Then
                                   
                                   'Copy texture width and height
                                   BitmapWidth = width
                                   BitmapHeight = height
                                   
                                   'Resize the texture picturebox
                                   frmMain.picTexture.width = BitmapWidth
                                   frmMain.picTexture.height = BitmapHeight
                                   
                                   'Create a paletted bitmap to render texture on
                                   Set frmMain.picTexture.Picture = CreatePalettedBitmap(playpal, BitmapWidth, BitmapHeight, TRANSPARENCY_INDEX)
                                   
                                   'Get a pointer to the bitmap memory
                                   CreateBitmapPointer frmMain.picTexture, BitmapPointer, BitmapDesc
                                   
                                   'Image now created
                                   ImageCreated = True
                              End If
                              
                              'Check how to draw
                              Select Case DataFormat
                                   
                                   'Draw patch as texture image
                                   Case TF_IMAGE
                                        Draw_Image BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                                   'Draw patch as flat
                                   Case TF_FLAT
                                        Draw_Flat BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, Patches(p).width, Patches(p).height, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                                   'Draw patch as 8 bit paletted bitmap
                                   Case TF_BITMAP_P8
                                        Draw_BitmapP8 BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, Patches(p).width, Patches(p).height, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                                   'Draw patch as 16 bit bitmap
                                   Case TF_BITMAP_B5G6R5
                                        Draw_BitmapB5G6R5 BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, Patches(p).width, Patches(p).height, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                                   'Draw patch as 24 bit bitmap
                                   Case TF_BITMAP_B8G8R8
                                        Draw_BitmapB8G8R8 BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, Patches(p).width, Patches(p).height, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                                   'Draw patch as 32 bit bitmap
                                   Case TF_BITMAP_A8B8G8R8
                                        Draw_BitmapA8B8G8R8 BitmapPointer(0), BitmapWidth, BitmapHeight, lumpdata, Len(lumpdata), Patches(p).x, Patches(p).y, Patches(p).width, Patches(p).height, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                                        PatchesDrawn = PatchesDrawn + 1
                                        
                              End Select
                         End If
                    End If
               End If
          End If
     Next p
     
     'Was the image created?
     If (ImageCreated = True) Then
          
          'Copy the data
          BitmapData() = BitmapPointer()
          
          'Destroy bitmap pointer
          DestroyBitmapPointer BitmapPointer()
          
          'Copy the picture in its original size (without the padding)
          If RunningWindows2000 Then Set Bitmap = frmMain.picTexture.image
          
          'No problems
          LoadImage = True
          IsLoaded = True
     Else
          
          'Ai, problem!
          LoadImage = False
          IsLoaded = False
     End If
     
     'Change mousepointer back
     Screen.MousePointer = OldMousePointer
     Exit Function
     
     
BadImage:
     
     'Destroy bitmap pointer
     DestroyBitmapPointer BitmapPointer
     
     'Remove image
     Set frmMain.picTexture.Picture = Nothing
     
     'Change mousepointer back
     Screen.MousePointer = OldMousePointer
     
     'Not loaded
     LoadImage = False
     IsLoaded = False
End Function


Public Sub LoadD3DTexture(Optional ByVal Padding As Boolean = False)
     Dim texinfo As D3DSURFACE_DESC
     
     'Ensure image is loaded
     If (IsLoaded = False) Then LoadImage
     
     'Check if loaded
     If (IsLoaded = True) Then
          
          'Convert to D3D Texture
          Set D3DTexture = BitmapData_D3DTexture(BitmapData(), BitmapWidth, height, playpal(), Padding)
          
          'Get the D3D texture size
          D3DTexture.GetLevelDesc 0, texinfo
          d3dwidth = texinfo.width
          d3dheight = texinfo.height
          d3dscalewidth = CSng(width) / CSng(d3dwidth)
          d3dscaleheight = CSng(height) / CSng(d3dheight)
     End If
End Sub


Public Sub GetScale(ByVal maxwidth As Long, ByVal maxheight As Long, ByRef newwidth As Long, ByRef newheight As Long, ByVal NoCaching As Boolean)
     Dim UnloadCache As Boolean
     Dim s As Single
     
     'Check if the image must be loaded
     If (width = 0) And (height = 0) Then
          
          'Load the image now
          LoadImage
          
          'Remove it from cache afterwards if not chaching
          If (NoCaching) Then UnloadCache = True
     End If
     
     'Determine aspect
     If ((width <= maxwidth) And (height <= maxheight)) Then
          
          'Image fits in area, no scaling
          s = 1
          
     ElseIf ((width - maxwidth) > (height - maxheight)) Then
          
          'Image is wider than tall, scale by width
          s = maxwidth / width
          
     Else
          
          'Image is taller than wide, scale by height
          s = maxheight / height
     End If
     
     'Return scaled size
     newwidth = width * s
     newheight = height * s
     
     'Remove the bitmap if not caching
     If ((Config("texturecaching") = 0) Or (UnloadCache)) Then
          
          'Clean up cache
          Set Bitmap = Nothing
          Erase BitmapData()
          IsLoaded = False
     End If
End Sub


Public Sub AddPatch(ByVal x As Long, ByVal y As Long, ByVal pWidth As Long, ByVal pHeight As Long, ByVal PatchIndex As Long, ByVal File As ENUM_IMAGESOURCE, ByVal Format As ENUM_IMAGEFORMAT)
     
     'Reallocate memory, pereserving existing patches
     ReDim Preserve Patches(0 To NumPatches)
     
     'Set the patch data
     With Patches(NumPatches)
          .x = x
          .y = y
          .width = pWidth
          .height = pHeight
          .Index = PatchIndex
          .File = File
          .Format = Format
     End With
     
     'Increase number of patches
     NumPatches = NumPatches + 1
End Sub


