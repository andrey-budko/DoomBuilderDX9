VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlat"
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


'General properties
Public name As String
Public width As Long
Public height As Long
Public valid As Long

'Cache
Private bitmap As StdPicture
Private Data() As Byte
Private DataWidth As Long
Private DataHeight As Long
Public dxtexture As Direct3DTexture8
Public IsLoaded As Boolean

Public Function GetPicture(Optional ByVal NoCaching As Boolean) As StdPicture
     Dim UnloadCache As Boolean
     Dim BitmapWidth As Long, BitmapHeight As Long
     Dim BitmapData() As Byte
     Dim BitmapDesc As SAFEARRAY1D
     
     'Check if the picture must be loaded
     If (IsLoaded = False) Then
          
          'Load the flat now
          Load
          
          'Remove it from cache afterwards
          If (NoCaching) Then UnloadCache = True
     End If
     
     'Check if we can use GDI objects (faster)
     If Not RunningWindows2000 Then
          
          'Resize the texture picturebox
          BitmapWidth = width
          BitmapHeight = height
          frmMain.picTexture.width = BitmapWidth
          frmMain.picTexture.height = BitmapHeight
          
          'Create a paletted bitmap to render texture on
          Set frmMain.picTexture.Picture = CreatePalettedBitmap(playpal, BitmapWidth, BitmapHeight, TRANSPARENCY_INDEX)
          
          'Get a pointer to the bitmap memory
          CreateBitmapPointer frmMain.picTexture, BitmapData(), BitmapDesc
          
          'Copy the picture
          CopyMemory BitmapData(0), Data(0), BitmapWidth * BitmapHeight
          
          'Destroy pointer
          DestroyBitmapPointer BitmapData()
          
          'Copy the picture in its original size (without the padding)
          Set GetPicture = frmMain.picTexture.image
     Else
          
          'Return the bitmap
          Set GetPicture = bitmap
     End If
     
     'Remove the bitmap if not caching
     If ((Config("texturecaching") = 0) Or (UnloadCache)) Then
          
          Set bitmap = Nothing
          Erase Data()
          IsLoaded = False
     End If
End Function

Public Sub GetScale(ByVal maxwidth As Long, ByVal maxheight As Long, ByRef newwidth As Long, ByRef newheight As Long)
     Dim s As Single
     
     'Determine scale
     If ((width <= maxwidth) And (height <= maxheight)) Then
          s = 1
     ElseIf ((width - maxwidth) > (height - maxheight)) Then
     'ElseIf (width > height) Then
          s = maxwidth / width
     Else
          s = maxheight / height
     End If
     
     'Return scaled size
     newwidth = width * s
     newheight = height * s
End Sub

Public Function IsDXTextureLoaded() As Boolean
     
     'Return loaded status
     IsDXTextureLoaded = Not (dxtexture Is Nothing)
End Function

Public Function IsTexture() As Boolean
     IsTexture = False
End Function

Public Function IsFlat() As Boolean
     IsFlat = True
End Function


Public Function Load() As Boolean
     Dim FileBuffer As Integer
     Dim BitmapWidth As Long, BitmapHeight As Long
     Dim BitmapData() As Byte
     Dim BitmapDesc As SAFEARRAY1D
     Dim LumpIndex As Long
     Dim i As Long
     
     'Change mousepointer
     If (Screen.MousePointer = vbDefault) Then Screen.MousePointer = vbArrowHourglass
     
     'Make the flat width and height
     BitmapWidth = width
     BitmapHeight = height
     
     'Resize the texture picturebox
     frmMain.picTexture.width = BitmapWidth
     frmMain.picTexture.height = BitmapHeight
     
     'Create a paletted bitmap to render flat on
     Set frmMain.picTexture.Picture = CreatePalettedBitmap(playpal, BitmapWidth, BitmapHeight)
     
     'Get a pointer to the bitmap memory
     CreateBitmapPointer frmMain.picTexture, BitmapData, BitmapDesc
     
     'Check if a valid flat
     If valid Then
          
          'Try getting the lump index from wad file
          LumpIndex = FindLumpIndex(MapWAD, WADFILEfstart, name, WADFILEfend - WADFILEfstart)
          If (LumpIndex > 0) Then
               
               'Load the bitmap data from wad file
               FileBuffer = MapWAD.FileBuffer
               Get #FileBuffer, MapWAD.LumpAddress(LumpIndex) + 1, BitmapData
          Else
               
               'Try getting lump index from addwad file
               LumpIndex = FindLumpIndex(ADDWAD, ADDWADfstart, name, ADDWADfend - ADDWADfstart)
               If (LumpIndex > 0) Then
                    
                    'Load the bitmap data from addwad file
                    FileBuffer = ADDWAD.FileBuffer
                    Get #FileBuffer, ADDWAD.LumpAddress(LumpIndex) + 1, BitmapData
               Else
                    
                    'Try getting the lump index from IWAD
                    LumpIndex = FindLumpIndex(IWAD, IWADfstart, name, IWADfend - IWADfstart)
                    If (LumpIndex > 0) Then
                         
                         'Load the bitmap data from IWAD
                         FileBuffer = IWAD.FileBuffer
                         Get #FileBuffer, IWAD.LumpAddress(LumpIndex) + 1, BitmapData
                    End If
               End If
          End If
          
          'Go for all pixels to remap reserved colors
          For i = LBound(BitmapData) To UBound(BitmapData)
               
               'Remap if color is reserved for transparent
               If (BitmapData(i) = TRANSPARENCY_INDEX) Then BitmapData(i) = ALTERNATE_INDEX
          Next i
          
          'Flip up-side down because it was loaded up-side down
          Flip_Flat BitmapData(0), BitmapWidth, BitmapHeight
     End If
     
     'Copy the data
     Data() = BitmapData()
     DataWidth = BitmapWidth
     DataHeight = BitmapHeight
     
     'Destroy bitmap pointer
     DestroyBitmapPointer BitmapData
     
     'Copy the picture in its original size (without the padding)
     If RunningWindows2000 Then Set bitmap = frmMain.picTexture.image
     
     'Change mousepointer
     If (Screen.MousePointer = vbArrowHourglass) Then Screen.MousePointer = vbDefault
     
     'No problems
     Load = True
     IsLoaded = True
End Function

Public Sub MakeDirect3DTexture()
     
     'Ensure image is loaded
     If (IsLoaded = False) Then Load
     
     'Convert to D3D Texture
     'Set dxtexture = StdPicture_D3DTexture(bitmap)
     Set dxtexture = BitmapData_D3DTexture(Data(), DataWidth, width, height, playpal())
End Sub

Public Sub UnloadDirect3DTexture()
     
     'Destroy texture
     Set dxtexture = Nothing
End Sub
