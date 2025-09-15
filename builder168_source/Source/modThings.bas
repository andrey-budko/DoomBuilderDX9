Attribute VB_Name = "modThings"
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



'Sprites
Public sprites As Dictionary                'clsImage objects with sprite name as key


Public Sub UnloadDirect3DSprites()
     Dim i As Long
     Dim SpriteKeys As Variant
     Dim Sprite As clsImage
     
     'Go for all flats
     SpriteKeys = sprites.Keys
     For i = LBound(SpriteKeys) To UBound(SpriteKeys)
          
          'Get the sprite object
          Set Sprite = sprites(SpriteKeys(i))
          
          'Unload the Direct3D texture
          Set Sprite.D3DTexture = Nothing
          
          'Clean up
          Set Sprite = Nothing
     Next i
End Sub

Public Sub CleanUpSpriteImages()
     
     'Unload all images
     Set sprites = New Dictionary
End Sub

Public Function GetSpriteForThingType(ByVal thingtype As Long, Optional ByVal LoadNow As Boolean = True) As clsImage
     Dim ThingDef As Dictionary
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Get the thing definition
          Set ThingDef = mapconfig("__things")(CStr(thingtype))
          
          'Has a sprite been set?
          If (ThingDef.Exists("sprite") = True) Then
               
               'Return the image for this sprite name
               Set GetSpriteForThingType = LoadSpriteImage(ThingDef("sprite"), LoadNow)
          End If
     End If
End Function

Public Function TestSpriteForThingType(ByVal thingtype As Long) As Boolean
     Dim ThingDef As Dictionary
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Get the thing definition
          Set ThingDef = mapconfig("__things")(CStr(thingtype))
          
          'Has a sprite been set?
          If (ThingDef.Exists("sprite") = True) Then
               
               'Return the result of image lookup
               TestSpriteForThingType = TestSpriteImage(ThingDef("sprite"))
          Else
               
               'No sprite set
               TestSpriteForThingType = False
          End If
     Else
          
          'No sprite either
          TestSpriteForThingType = False
     End If
End Function


Public Sub GetScaledSpritePicture(ByVal thingtype As Long, ByRef target As image, Optional ByVal width As Long, Optional ByVal height As Long, Optional ByVal NoCaching As Boolean)
     Dim Sprite As clsImage
     Dim pic As StdPicture
     Dim sw As Long, sh As Long
     
     'Get the sprite for this type
     Set Sprite = GetSpriteForThingType(thingtype)
     
     'Check if we have a sprite
     If Not (Sprite Is Nothing) Then
          
          'Get the sprite picture
          Set pic = Sprite.Picture(NoCaching)
          
          'Check if we have a picture
          If Not (pic Is Nothing) Then
               
               'Determine scale
               Sprite.GetScale width - 4, height - 4, sw, sh, NoCaching
               
               'Resize picture
               frmMain.picTexture.BackColor = RGB(0, 0, 0)
               frmMain.picTexture.width = sw
               frmMain.picTexture.height = sh
               frmMain.picTexture.PaintPicture pic, 0, 0, sw, sh, , , , , vbSrcCopy
               Set pic = frmMain.picTexture.image
               
               'Create mask picture
               CreateMaskPicture frmMain.picTexture, frmMain.picMask
               
               'Erase the temp picturebox
               Set frmMain.picTexture.Picture = Nothing
               frmMain.picTexture.BackColor = vbApplicationWorkspace
               
               'Draw the mask over the target
               frmMain.picTexture.PaintPicture frmMain.picMask.image, 0, 0, , , , , , , vbSrcAnd
               
               'Invert the mask picture
               frmMain.picMask.DrawMode = vbInvert
               frmMain.picMask.Line (0, 0)-(Sprite.width, Sprite.height), 0, BF
               frmMain.picMask.DrawMode = vbCopyPen
               
               'Draw the pic over the mask
               frmMain.picMask.PaintPicture pic, 0, 0, , , , , , , vbSrcAnd
               
               'Draw the pic (in mask) over target
               frmMain.picTexture.PaintPicture frmMain.picMask.image, 0, 0, , , , , , , vbSrcPaint
               
               'Set the picture on the imagebox
               Set target.Picture = frmMain.picTexture.image
               
               'Clean up
               Set frmMain.picTexture.Picture = Nothing
               Set frmMain.picMask.Picture = Nothing
               
               'Move the image box depending on scale
               target.Move (width - sw) \ 2, (height - sh) \ 2, sw, sh
          Else
               
               'Set nothing
               Set target.Picture = Nothing
               
               'Move the box
               target.Move 0, 0, 64, 64
          End If
     Else
          
          'Set nothing
          Set target.Picture = Nothing
          
          'Move the box
          target.Move 0, 0, 64, 64
     End If
End Sub


Public Sub InitializeSprites()
     
     'Create collections
     Set sprites = New Dictionary
End Sub

Public Function LoadSpriteImage(ByRef spritename As String, Optional ByVal LoadNow As Boolean = True) As clsImage
     Dim lumpindex As Long
     Dim Sprite As clsImage
     Dim source As ENUM_IMAGESOURCE
     Dim lumpdata As String
     Dim offsetx As Long
     Dim offsety As Long
     Dim width As Long
     Dim height As Long
     Dim imgformat As Long
     
     'No empty spritename allowed
     If (LenB(Trim$(spritename)) = 0) Then Exit Function
     
     'Check if available from cache
     If (sprites.Exists(spritename) = True) Then
          
          'Use image from cache
          Set LoadSpriteImage = sprites(spritename)
     Else
          
          'First find the sprite lump name in the open WAD file
          If Not (MapWAD Is Nothing) Then lumpindex = FindLumpIndex(MapWAD, 1, spritename)
          If (lumpindex > 0) Then
               
               'Load from MapWAD
               source = TS_MAPWAD
               lumpdata = MapWAD.GetLump(lumpindex)
          Else
               
               'Second find the sprite lump in the additional WAD file
               If Not (AddWAD Is Nothing) Then lumpindex = FindLumpIndex(AddWAD, 1, spritename)
               If (lumpindex > 0) Then
                    
                    'Load from AddWAD
                    source = TS_ADDWAD
                    lumpdata = AddWAD.GetLump(lumpindex)
               Else
                    
                    'Last find the sprite lump in the IWAD file
                    If Not (IWAD Is Nothing) Then lumpindex = FindLumpIndex(IWAD, 1, spritename)
                    If (lumpindex > 0) Then
                         
                         'Load from IWAD
                         source = TS_IWAD
                         lumpdata = IWAD.GetLump(lumpindex)
                    End If
               End If
          End If
          
          'Check if the sprite is found
          If (lumpindex > 0) Then
               
               'Determine format of image data
               imgformat = GetImageFormat(lumpdata, False, width, height)
               
               'Check if offsets can be read from lump data
               If (imgformat = TF_IMAGE) Then
                    
                    'Read offsets
                    offsetx = CVI(Mid$(lumpdata, 5, 2))
                    offsety = CVI(Mid$(lumpdata, 7, 2))
               End If
               
               'Sprite is in the MapWAD
               Set Sprite = New clsImage
               With Sprite
                    .Name = spritename
                    .width = width
                    .height = height
                    .ScaleX = offsetx
                    .ScaleY = offsety
                    .FlatCandidate = False
                    .AddPatch 0, 0, 0, 0, lumpindex, source, TF_UNKNOWN
               End With
               
               'Load the sprite now
               If (LoadNow) Then Sprite.LoadImage
               
               'Add to the list of known sprites
               sprites.Add spritename, Sprite
               
               'Use this sprite
               Set LoadSpriteImage = Sprite
          End If
     End If
End Function


Public Function TestSpriteImage(ByRef spritename As String) As Boolean
     Dim lumpindex As Long
     
     'No empty spritename allowed
     If (LenB(Trim$(spritename)) = 0) Then
          TestSpriteImage = False
          Exit Function
     End If
     
     'Check if available from cache
     If (sprites.Exists(spritename) = True) Then
          
          'Yes, image has a sprite
          TestSpriteImage = True
     Else
          
          'Not found yet
          TestSpriteImage = False
          
          'First find the sprite lump name in the open WAD file
          If Not (MapWAD Is Nothing) Then lumpindex = FindLumpIndex(MapWAD, 1, spritename)
          If (lumpindex > 0) Then
               
               'Load from MapWAD
               TestSpriteImage = True
          Else
               
               'Second find the sprite lump in the additional WAD file
               If Not (AddWAD Is Nothing) Then lumpindex = FindLumpIndex(AddWAD, 1, spritename)
               If (lumpindex > 0) Then
                    
                    'Load from AddWAD
                    TestSpriteImage = True
               Else
                    
                    'Last find the sprite lump in the IWAD file
                    If Not (IWAD Is Nothing) Then lumpindex = FindLumpIndex(IWAD, 1, spritename)
                    If (lumpindex > 0) Then
                         
                         'Load from IWAD
                         TestSpriteImage = True
                    End If
               End If
          End If
     End If
End Function



