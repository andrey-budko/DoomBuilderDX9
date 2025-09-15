VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexture"
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


'PATCH DESCRIPTOR
Private Type PATCHDESC
     x As Long
     y As Long
     Index As Long
     CustomFile As Long
     StepDir As Long
     Colormap As Long
End Type


'General properties
Public Name As String
Public Height As Long
Public Width As Long

'Patches
Private Patches() As PATCHDESC
Private NumPatches As Long

'Cache
Private Bitmap As StdPicture
Private Data() As Byte
Private DataWidth As Long
Private DataHeight As Long
Public IsLoaded As Boolean
Public dxtexture As Direct3DTexture8

Public Sub AddPatch(ByVal x As Long, ByVal y As Long, ByVal PatchIndex As Long, ByVal StepDir As Long, ByVal Colormap As Long, ByVal CustomFile As Long)
     
     'Allocate memory
     ReDim Preserve Patches(0 To NumPatches)
     
     'Set the patch data
     With Patches(NumPatches)
          .x = x
          .y = y
          .Index = PatchIndex
          .CustomFile = CustomFile
          .StepDir = StepDir
          .Colormap = Colormap
     End With
     
     'Increase number of patches
     NumPatches = NumPatches + 1
End Sub

Public Function GetPicture(Optional ByVal NoCaching As Boolean) As StdPicture
     Dim UnloadCache As Boolean
     Dim BitmapWidth As Long, BitmapHeight As Long
     Dim BitmapData() As Byte
     Dim BitmapDesc As SAFEARRAY1D
     
     'Check if the picture must be loaded
     If (IsLoaded = False) Then
          
          'Load the texture now
          Load
          
          'Remove it from cache afterwards
          If (NoCaching) Then UnloadCache = True
     End If
     
     'Check if we can use GDI objects (faster)
     If Not RunningWindows2000 Then
          
          'Resize the texture picturebox
          BitmapWidth = Width
          BitmapHeight = Height
          frmMain.picTexture.Width = BitmapWidth
          frmMain.picTexture.Height = BitmapHeight
          
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
          Set GetPicture = Bitmap
     End If
     
     'Remove the bitmap if not caching
     If ((Config("texturecaching") = 0) Or (UnloadCache)) Then
          
          Set Bitmap = Nothing
          Erase Data()
          IsLoaded = False
     End If
End Function

Public Sub GetScale(ByVal maxwidth As Long, ByVal maxheight As Long, ByRef newwidth As Long, ByRef newheight As Long)
     Dim s As Single
     
     'Determine scale
     If ((Width <= maxwidth) And (Height <= maxheight)) Then
          s = 1
     ElseIf ((Width - maxwidth) > (Height - maxheight)) Then
     'ElseIf (width > height) Then
          s = maxwidth / Width
     Else
          s = maxheight / Height
     End If
     
     'Return scaled size
     newwidth = Width * s
     newheight = Height * s
End Sub

Public Function IsDXTextureLoaded() As Boolean
     
     'Return loaded status
     IsDXTextureLoaded = Not (dxtexture Is Nothing)
End Function

Public Function Load() As Boolean
     On Local Error GoTo BadPicture
     Dim BitmapWidth As Long, BitmapHeight As Long
     Dim BitmapData() As Byte
     Dim BitmapDesc As SAFEARRAY1D
     Dim pdata As String
     Dim LumpIndex As Long
     Dim p As Long
     Dim s As Single
     
     'Change mousepointer
     If (Screen.MousePointer = vbDefault) Then Screen.MousePointer = vbArrowHourglass
     
     'Copy texture width and height
     BitmapWidth = Width
     BitmapHeight = Height
     
     'Resize the texture picturebox
     frmMain.picTexture.Width = BitmapWidth
     frmMain.picTexture.Height = BitmapHeight
     
     'Create a paletted bitmap to render texture on
     Set frmMain.picTexture.Picture = CreatePalettedBitmap(playpal, BitmapWidth, BitmapHeight, TRANSPARENCY_INDEX)
     
     'Get a pointer to the bitmap memory
     CreateBitmapPointer frmMain.picTexture, BitmapData, BitmapDesc
     
     'Go for all patches
     For p = 0 To (NumPatches - 1)
          
          'Check if the patch resides in PWAD or IWAD
          pdata = ""
          Select Case Patches(p).CustomFile
               
               'Load the lump from wad file
               Case 1: pdata = MapWAD.GetLump(Patches(p).Index)
               
               'Load the lump from AddWad file
               Case 2: pdata = AddWAD.GetLump(Patches(p).Index)
               
               'Load the lump from IWAD
               Case 3: pdata = IWAD.GetLump(Patches(p).Index)
               
          End Select
          
          'Draw the patch
          If (pdata <> "") Then
               
               'Check how to draw
               Select Case GetTextureFormat(pdata)
                    
                    'Draw patch as texture image
                    Case TF_IMAGE: Draw_Patch BitmapData(0), BitmapWidth, BitmapHeight, pdata, Len(pdata), Patches(p).x, Patches(p).y, TRANSPARENCY_INDEX, ALTERNATE_INDEX
                    
                    'Convert to bitmap and draw
                    Case TF_PNG
                    
                    'Otherwise, clear the texture white
                    Case Else: FillMemory BitmapData(0), UBound(BitmapData) - LBound(BitmapData) + 1, DEFAULT_INDEX
                    
               End Select
          End If
     Next p
     
     'Copy the data
     Data() = BitmapData()
     DataWidth = BitmapWidth
     DataHeight = BitmapHeight
     
     'Destroy bitmap pointer
     DestroyBitmapPointer BitmapData
     
     'Copy the picture in its original size (without the padding)
     If RunningWindows2000 Then Set Bitmap = frmMain.picTexture.image
     
     'Change mousepointer
     If (Screen.MousePointer = vbArrowHourglass) Then Screen.MousePointer = vbDefault
     
     'No problems
     Load = True
     IsLoaded = True
     Exit Function
     
     
BadPicture:
     
     'Destroy bitmap pointer
     DestroyBitmapPointer BitmapData
     
     'Remove image
     Set frmMain.picTexture.Picture = Nothing
     frmMain.picTexture.AutoRedraw = True
     
     'Change mousepointer
     If (Screen.MousePointer = vbArrowHourglass) Then Screen.MousePointer = vbDefault
     
     'Leave now
     Exit Function
End Function

Public Function IsTexture() As Boolean
     IsTexture = True
End Function


Public Function IsFlat() As Boolean
     IsFlat = False
End Function


Public Sub MakeDirect3DTexture()
     
     'Ensure image is loaded
     If (IsLoaded = False) Then Load
     
     'Convert to D3D Texture
     'Set dxtexture = StdPicture_D3DTexture(bitmap)
     Set dxtexture = BitmapData_D3DTexture(Data(), DataWidth, Width, Height, playpal())
End Sub

Public Sub UnloadDirect3DTexture()
     
     'Destroy texture
     Set dxtexture = Nothing
End Sub

Private Sub Class_Initialize()

End Sub


