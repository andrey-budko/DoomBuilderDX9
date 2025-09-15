Attribute VB_Name = "modEditing"
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


'Unlimited distance for nearest linedef checking
Public Const ENDLESS_DISTANCE As Long = 2147483640

'Editing modes
Public Enum ENUM_EDITMODE
     EM_MOVE
     EM_VERTICES
     EM_LINES
     EM_SECTORS
     EM_THINGS
     EM_3D
End Enum

'Editing sub modes
Public Enum ENUM_EDITSUBMODE
     ESM_NONE
     ESM_DRAGGING
     ESM_DRAWING
     ESM_SELECTING
     ESM_PASTING
     ESM_MOVING
End Enum

'Find and Replace searches
Public Enum ENUM_FINDREPLACE
     FR_VERTEXNUMBER
     FR_LINEDEFNUMBER
     FR_LINEDEFACTION
     FR_LINEDEFSECTORTAG
     FR_LINEDEFTHINGTAG
     FR_LINEDEFTEXTURE
     FR_SECTORNUMBER
     FR_SECTOREFFECT
     FR_SECTORTAG
     FR_SECTORFLAT
     FR_THINGNUMBER
     FR_THINGACTION
     FR_THINGTAG
     FR_THINGSECTORTAG
     FR_THINGTHINGTAG
     FR_THINGTYPE
End Enum

'Thing filter settings
Public Type THINGFILTERS
     filtermode As Long       '0 = any of settings, 1 = all of settings
     category As Long
     Flags As Long
End Type


'API Declarations
Public Declare Function NearestVertex Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByRef distance As Long) As Long
Public Declare Function NearestSelectedVertex Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByRef distance As Long) As Long
Public Declare Function NearestUnselectedVertex Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByRef distance As Long) As Long
Public Declare Function NearestOtherVertex Lib "builder.dll" (ByVal v As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByRef distance As Long) As Long
Public Declare Function NearestThing Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef things As MAPTHING, ByVal numthings As Long, ByRef distance As Long, ByVal filterthings As Long, ByRef Filter As THINGFILTERS) As Long
Public Declare Function NearestSelectedThing Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef things As MAPTHING, ByVal numthings As Long, ByRef distance As Long) As Long
Public Declare Function NearestUnselectedThing Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef things As MAPTHING, ByVal numthings As Long, ByRef distance As Long, ByVal filterthings As Long, ByRef Filter As THINGFILTERS) As Long
Public Declare Function NearestLinedef Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByRef distance As Long) As Long
Public Declare Function NearestSelectedLinedef Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByRef SelectedLines As Long, ByVal numselectedlinedefs As Long, ByRef distance As Long, ByVal maxdistance As Long) As Long
Public Declare Function NearestUnselectedLinedef Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByRef distance As Long, ByVal maxdistance As Long) As Long
Public Declare Function NearestUnselectedUnreferencedLinedef Lib "builder.dll" (ByVal v As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByRef distance As Long, ByVal maxdistance As Long) As Long
Public Declare Function IntersectSector Lib "builder.dll" (ByVal x As Long, ByVal y As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByVal numlinedefs As Long, ByVal unselectedonly As Long) As Long
Public Declare Function LinedefBetweenVertices Lib "builder.dll" (ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal startlinedef As Long, ByVal vertex1 As Long, ByVal vertex2 As Long, ByVal excludeline As Long) As Long
Public Declare Sub RoundVertices Lib "builder.dll" (ByRef vertexes As MAPVERTEX, ByVal numvertexes As Long)
Public Declare Sub ResetSelections Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByVal ptr_sectors As Long, ByVal numsectors As Long)
Public Declare Function IntersectLineA Lib "builder.dll" (ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal x3 As Single, ByVal y3 As Single, ByVal x4 As Single, ByVal y4 As Single) As Single

'Interface
Public LastX As Single
Public LastY As Single
Public MouseInside As Boolean

'Current editing mode
Public mode As ENUM_EDITMODE
Public submode As ENUM_EDITSUBMODE
Public PreviousMode As ENUM_EDITMODE

'Grid
Public gridsizex As Long, gridsizey As Long
Public gridx As Long, gridy As Long

'Options
Public vertexsize As Long
Public indicatorsize As Long
Public thingsize As Long
Public snapmode As Boolean
Public stitchmode As Boolean

'Things filtering
Public filterthings As Boolean
Public filtersettings As THINGFILTERS

'Current mousepointer selection (what actually is selected depends on the mode)
Public currentselected As Long

'Selected objects (what actually is selected depends on the mode)
Public selected As Dictionary       'Index by CStr(Index) as key
Public numselected As Long
Public selectedtype As ENUM_EDITMODE

'Dragging objects (what is temporarely selection during drag mode. only vertices or things)
Public dragselected As Dictionary
Public dragnumselected As Long

'Grabbed object (object that will be synched with the mouse in drag mode. only a vertex or thing)
Public grabobject As Long

'Linedefs beign modified in a drag operation
Public changedlines() As Long
Public numchangedlines As Long

'Defaults
Public LastThing As MAPTHING

'Properties Copying
Public CopiedLinedef As MAPLINEDEF
Public CopiedSidedef1 As MAPSIDEDEF
Public CopiedSidedef2 As MAPSIDEDEF
Public CopiedSidedef As MAPSIDEDEF
Public CopiedSector As MAPSECTOR
Public CopiedThing As MAPTHING

'Recursive trace
Public TerminateRecursion As Long
Public SectorSplitLinesList() As Long
Public SectorSplitNumLines As Long

Public Sub AddSidedefTextures(ByVal SourceSidedef As Long, ByVal TargetSidedef As Long)
     Dim CopyOffsets As Long
     
     'Cant copy when either the source or target is nothing
     If ((SourceSidedef > -1) And (TargetSidedef > -1)) Then
          
          'Copy upper texture if any set
          If ((StrComp(sidedefs(SourceSidedef).upper, "-", vbBinaryCompare) <> 0) And _
              (LenB(sidedefs(SourceSidedef).upper) <> 0)) Then
               
               'Copy upper texture
               sidedefs(TargetSidedef).upper = sidedefs(SourceSidedef).upper
               
               'Count as half the choice for copying offsets
               CopyOffsets = CopyOffsets + 1
          End If
          
          'Copy middle texture if any set
          If ((StrComp(sidedefs(SourceSidedef).middle, "-", vbBinaryCompare) <> 0) And _
              (LenB(sidedefs(SourceSidedef).middle) <> 0)) Then
               
               'Copy middle texture
               sidedefs(TargetSidedef).middle = sidedefs(SourceSidedef).middle
               
               'Counts for copying offsets
               CopyOffsets = CopyOffsets + 2
          End If
          
          'Copy lower texture if any set
          If ((StrComp(sidedefs(SourceSidedef).lower, "-", vbBinaryCompare) <> 0) And _
              (LenB(sidedefs(SourceSidedef).lower) <> 0)) Then
               
               'Copy lower texture
               sidedefs(TargetSidedef).lower = sidedefs(SourceSidedef).lower
               
               'Count as half the choice for copying offsets
               CopyOffsets = CopyOffsets + 1
          End If
          
          'Should offsets be copied as well?
          If (CopyOffsets >= 2) Then
               
               'Copy offsets
               sidedefs(TargetSidedef).tx = sidedefs(SourceSidedef).tx
               sidedefs(TargetSidedef).ty = sidedefs(SourceSidedef).ty
          End If
     End If
End Sub

Public Sub AlignTexturesX(ByVal v As Long, ByVal v_offset As Long, ByVal texturename As String, ByVal firstfront As Boolean, ByVal firstlinedef As Long)
     Dim ld As Long
     Dim sd As Long
     Dim startoffset As Long
     Dim endoffset As Long
     Dim length As Long
     Dim DX As Single, dy As Single
     Dim texturesize As Long
     Dim upper As String, lower As String, middle As String
     Dim nextfirstfront As Boolean
     Dim texturescale As Single
     
     'v = vertex to align linedefs from
     'v_offset = offset at the vertex that the textures should be aligned to
     'texturename = only align textures with this name
     'firstfront = if v is linedef's first vertex, then align its front sidedef
     'firstlinedef = linedef where the firstfront parameter will switch
     
     'Leave if the texture does not exist
     If (textures.Exists(texturename) = False) Then Exit Sub
     
     'Get the texture width
     texturesize = textures(texturename).width
     texturescale = textures(texturename).ScaleX
     
     'Go for all lines
     For ld = 0 To (numlinedefs - 1)
          
          'Start without textures
          upper = ""
          middle = ""
          lower = ""
          
          'Check if linedef is unselected
          If (linedefs(ld).selected = 0) Then
               
               'Check if this is vertex 1
               If (linedefs(ld).v1 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Calculate the length of the line
                    'This is multiplied by the texture scale for correct alignment
                    DX = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
                    dy = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
                    length = Sqr(DX * DX + dy * dy) * texturescale
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the front sidedef
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Start offset is same as offset given for this vertex
                                   startoffset = v_offset
                                   
                                   'Calculate the end offset
                                   endoffset = v_offset + length
                                   
                                   'Wrap the offset to the texture length
                                   endoffset = (endoffset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).tx = startoffset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesX linedefs(ld).v2, endoffset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    Else
                         
                         'Do the back sidedef with reverse alignment
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'End offset is same as offset given for this vertex
                                   endoffset = v_offset
                                   
                                   'Calculate the start offset
                                   startoffset = v_offset - length
                                   
                                   'Wrap the offset to the texture length
                                   startoffset = (startoffset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).tx = startoffset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesX linedefs(ld).v2, startoffset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    End If
                    
               'Check if this is vertex 2
               ElseIf (linedefs(ld).v2 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Calculate the length of the line
                    DX = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
                    dy = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
                    length = Sqr(DX * DX + dy * dy)
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the back sidedef
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Start offset is same as offset given for this vertex
                                   startoffset = v_offset
                                   
                                   'Calculate the end offset
                                   endoffset = v_offset + length
                                   
                                   'Wrap the offset to the texture length
                                   endoffset = (endoffset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).tx = startoffset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesX linedefs(ld).v1, endoffset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    Else
                         
                         'Do the front sidedef with reverse alignment
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'End offset is same as offset given for this vertex
                                   endoffset = v_offset
                                   
                                   'Calculate the start offset
                                   startoffset = v_offset - length
                                   
                                   'Wrap the offset to the texture length
                                   startoffset = (startoffset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).tx = startoffset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesX linedefs(ld).v1, startoffset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    End If
               End If
          End If
     Next ld
End Sub

Public Sub FloodFillTexture(ByVal v As Long, ByVal texturename As String, ByVal firstfront As Boolean, ByVal firstlinedef As Long, ByVal painttexture As String)
     Dim ld As Long
     Dim sd As Long
     Dim upper As String, lower As String, middle As String
     Dim nextfirstfront As Boolean
     
     'v = vertex to paint texture from
     'texturename = only paint over textures with this name
     'painttexture = paint with this texture
     'firstfront = if v is linedef's first vertex, then paint its front sidedef
     'firstlinedef = linedef where the firstfront parameter will switch
     
     'Go for all lines
     For ld = 0 To (numlinedefs - 1)
          
          'Start without textures
          upper = ""
          middle = ""
          lower = ""
          
          'Check if linedef is unselected
          If (linedefs(ld).selected = 0) Then
               
               'Check if this is vertex 1
               If (linedefs(ld).v1 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the front sidedef
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint textures
                                   If (StrComp(upper, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).upper = painttexture
                                   If (StrComp(middle, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).middle = painttexture
                                   If (StrComp(lower, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).lower = painttexture
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillTexture linedefs(ld).v2, texturename, nextfirstfront, firstlinedef, painttexture
                              End If
                         End If
                    Else
                         
                         'Do the back sidedef with reverse alignment
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint textures
                                   If (StrComp(upper, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).upper = painttexture
                                   If (StrComp(middle, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).middle = painttexture
                                   If (StrComp(lower, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).lower = painttexture
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillTexture linedefs(ld).v2, texturename, nextfirstfront, firstlinedef, painttexture
                              End If
                         End If
                    End If
                    
               'Check if this is vertex 2
               ElseIf (linedefs(ld).v2 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the back sidedef
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint textures
                                   If (StrComp(upper, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).upper = painttexture
                                   If (StrComp(middle, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).middle = painttexture
                                   If (StrComp(lower, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).lower = painttexture
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillTexture linedefs(ld).v1, texturename, nextfirstfront, firstlinedef, painttexture
                              End If
                         End If
                    Else
                         
                         'Do the front sidedef with reverse alignment
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint textures
                                   If (StrComp(upper, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).upper = painttexture
                                   If (StrComp(middle, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).middle = painttexture
                                   If (StrComp(lower, texturename, vbBinaryCompare) = 0) Then sidedefs(sd).lower = painttexture
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillTexture linedefs(ld).v1, texturename, nextfirstfront, firstlinedef, painttexture
                              End If
                         End If
                    End If
               End If
          End If
     Next ld
End Sub


Public Sub FloodFillFlats(ByVal s As Long, ByVal texturename As String, ByVal painttexture As String, ByVal floors As Boolean)
     Dim ld As Long
     Dim sd As Long
     Dim osd As Long
     Dim os As Long
     Dim upper As String, lower As String, middle As String
     
     's = sector to start from
     'texturename = only paint over textures with this name
     'painttexture = paint with this texture
     'floors = true when doing floors
     
     'Go for all sides
     For sd = 0 To (numsidedefs - 1)
          
          'Check if the sidedef is adjacent to this sector
          If (sidedefs(sd).sector = s) Then
               
               'Get sidedef on the other side
               ld = sidedefs(sd).linedef
               If (linedefs(ld).s1 = sd) Then osd = linedefs(ld).s2
               If (linedefs(ld).s2 = sd) Then osd = linedefs(ld).s1
               
               'Check if there is another sidedef
               If (osd > -1) Then
                    
                    'Get the other sector
                    os = sidedefs(osd).sector
                    
                    'Check if sector is unselected
                    If (sectors(os).selected = 0) Then
                              
                         'Check if doing floor
                         If floors Then
                              
                              'Floor has the same texture?
                              If (StrComp(sectors(os).tfloor, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint texture
                                   sectors(os).tfloor = painttexture
                                   
                                   'Select the sector to indicate its done
                                   sectors(os).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillFlats os, texturename, painttexture, floors
                              End If
                         Else
                              
                              'Ceiling has the same texture?
                              If (StrComp(sectors(os).tceiling, texturename, vbBinaryCompare) = 0) Then
                                   
                                   'Paint texture
                                   sectors(os).tceiling = painttexture
                                   
                                   'Select the sector to indicate its done
                                   sectors(os).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   FloodFillFlats os, texturename, painttexture, floors
                              End If
                         End If
                    End If
               End If
          End If
     Next sd
End Sub



Public Sub AlignTexturesY(ByVal v As Long, ByVal base_offset As Long, ByVal texturename As String, ByVal firstfront As Boolean, ByVal firstlinedef As Long)
     Dim ld As Long
     Dim sd As Long
     Dim offset As Long
     Dim texturesize As Long
     Dim upper As String, lower As String, middle As String
     Dim nextfirstfront As Boolean
     Dim texturescale As Single
     
     'v = vertex to align linedefs from
     'v_offset = offset at the vertex that the textures should be aligned to
     'texturename = only align textures with this name
     'firstfront = if v is linedef's first vertex, then align its front sidedef
     'firstlinedef = linedef where the firstfront parameter will switch
     
     'Leave if the texture does not exist
     If (textures.Exists(texturename) = False) Then Exit Sub
     
     'Get the texture height
     texturesize = textures(texturename).height
     texturescale = textures(texturename).ScaleY
     
     'Go for all lines
     For ld = 0 To (numlinedefs - 1)
          
          'Start without textures
          upper = ""
          middle = ""
          lower = ""
          
          'Check if linedef is unselected
          If (linedefs(ld).selected = 0) Then
               
               'Check if this is vertex 1
               If (linedefs(ld).v1 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the front sidedef
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   offset = base_offset - sectors(sidedefs(sd).sector).hceiling
                                   
                                   'Wrap the offset to the texture length
                                   offset = (offset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).ty = offset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesY linedefs(ld).v2, base_offset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    Else
                         
                         'Do the back sidedef with reverse alignment
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   offset = base_offset - sectors(sidedefs(sd).sector).hceiling
                                   
                                   'Wrap the offset to the texture length
                                   offset = (offset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).ty = offset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesY linedefs(ld).v2, base_offset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    End If
                    
               'Check if this is vertex 2
               ElseIf (linedefs(ld).v2 = v) Then
                    
                    'Determine next firstfront
                    nextfirstfront = firstfront Xor (ld = firstlinedef)
                    
                    'Check if doing front or back
                    If nextfirstfront Then
                         
                         'Do the back sidedef
                         sd = linedefs(ld).s2
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS2Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS2Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   offset = base_offset - sectors(sidedefs(sd).sector).hceiling
                                   
                                   'Wrap the offset to the texture length
                                   offset = (offset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).ty = offset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesY linedefs(ld).v1, base_offset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    Else
                         
                         'Do the front sidedef with reverse alignment
                         sd = linedefs(ld).s1
                         
                         'Check if a sidedef exists
                         If (sd > -1) Then
                              
                              'Check if the sidedef has any texture that matches
                              If RequiresS1Upper(ld) Then upper = left$(sidedefs(sd).upper, Len(texturename))
                              middle = left$(sidedefs(sd).middle, Len(texturename))
                              If RequiresS1Lower(ld) Then lower = left$(sidedefs(sd).lower, Len(texturename))
                              If (StrComp(upper, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(middle, texturename, vbBinaryCompare) = 0) Or _
                                 (StrComp(lower, texturename, vbBinaryCompare) = 0) Then
                                   
                                   offset = base_offset - sectors(sidedefs(sd).sector).hceiling
                                   
                                   'Wrap the offset to the texture length
                                   offset = (offset Mod texturesize)
                                   
                                   'Apply offset to sidedef
                                   sidedefs(sd).ty = offset
                                   
                                   'Select the lindef to indicate its done
                                   linedefs(ld).selected = 1
                                   
                                   'Pass routine on to the next vertex
                                   AlignTexturesY linedefs(ld).v1, base_offset, texturename, nextfirstfront, firstlinedef
                              End If
                         End If
                    End If
               End If
          End If
     Next ld
End Sub



Public Sub AllLinesChanging()
     Dim s As Long
     
     'Allocate memory for array
     ReDim changedlines(0 To (numlinedefs - 1))
     
     'Go for all linedefs
     For s = 0 To (numlinedefs - 1)
          
          'Set in array
          changedlines(s) = s
     Next s
     
     'Count
     numchangedlines = numlinedefs
End Sub

Public Sub ApplyParentSectors()
     Dim sd As Long, s As Long
     Dim sx As Single, sy As Single
     Dim Merged As Long
     Dim floorheight As Long
     Dim ceilheight As Long
     Dim SectorsMerged As Boolean
     
     'Start with very high/low floor/ceiling
     floorheight = -2147483640
     ceilheight = 2147483640
     
     'Go for all sidedefs
     sd = (numsidedefs - 1)
     Do While (sd >= 0)
          
          'Check if sidedef sector is supposed to be the parent sector
          If (sidedefs(sd).sector = -1) Then
               
               'Get sector check spot
               GetLineSideSpot sidedefs(sd).linedef, 1, (linedefs(sidedefs(sd).linedef).s1 <> sd), sx, sy
               
               'Get the sector where this sidedef will be at
               s = IntersectSector(sx, sy, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 1)
               
               'Yes, merged
               Merged = True
          Else
               
               'Own sector provided
               s = sidedefs(sd).sector
               
               'Normal
               Merged = False
          End If
          
          'Check if a sector for this sidedef is found
          If (s > -1) Then
               
               'Set sidedef sector
               sidedefs(sd).sector = s
               
               'Was it merged?
               If (Merged) Then
                    
                    'At least one merged
                    SectorsMerged = True
                    
                    'Measure the floor/ceiling heights
                    If (sectors(s).hfloor > floorheight) Then floorheight = sectors(s).hfloor
                    If (sectors(s).hceiling < ceilheight) Then ceilheight = sectors(s).hceiling
                    
                    'Remove middle texture when doublesided
                    If (linedefs(sidedefs(sd).linedef).s1 > -1) And (linedefs(sidedefs(sd).linedef).s2 > -1) Then
                         sidedefs(linedefs(sidedefs(sd).linedef).s1).middle = "-"
                         sidedefs(linedefs(sidedefs(sd).linedef).s2).middle = "-"
                    End If
               End If
          Else
               
               'Remove the sidedef
               RemoveSidedef sd, True, False, True
          End If
          
          'Next sidedef
          sd = sd - 1
     Loop
     
     'Do we use sector height adjustment?
     If (PrefabAdjustHeights = True) Then
          
          'If sectors were merged
          If (SectorsMerged = True) Then
               
               'Apply heights to start heights
               'so that when these are applied to sectors,
               'they will be adjusted to match their target sector heights
               PrefabFloorHeight = PrefabFloorHeight - floorheight
               PrefabCeilHeight = PrefabCeilHeight - ceilheight
          Else
               
               'Nothing merges, must have been pasted into the void
               PrefabFloorHeight = 0
               PrefabCeilHeight = 0
          End If
     End If
End Sub

Public Sub ApplySectorHeightAdjustments()
     Dim SectorKeys As Variant
     Dim s As Long
     
     'Get list of sectors to adjust
     SectorKeys = selected.Items
     
     'Go for all selected sectors
     For s = LBound(SectorKeys) To UBound(SectorKeys)
          
          'Adjust sector ceiling and floor
          With sectors(SectorKeys(s))
               .hfloor = .hfloor - PrefabFloorHeight
               .hceiling = .hceiling - PrefabCeilHeight
          End With
     Next s
End Sub


Public Function AutoStitchDraggedSelection() As Boolean
     Dim s As Long
     Dim sv As Long
     Dim sel As Variant
     Dim nv As Long
     Dim distance As Long
     Dim sc As Long
     Dim stdnc As Long
     
     'Auto Stitch Vertices:
     'Checks selected vertices versus non-selected vertices for
     'ones that are close together (autostitchdistance) and stitch
     'selected vertex to the non-selected vertex
     'Returns True when changes are made
     
     'Get selection
     sel = dragselected.Items
     
     'Get stitch distance
     stdnc = Config("autostitchdistance")
     
     'Go for all selected vertices
     For s = LBound(sel) To UBound(sel)
          
          'Get vertex index
          sv = CLng(sel(s))
          
          'Get the nearest, non-selected vertex
          'nv = NearestUnselectedVertex(vertexes(sv).x, -vertexes(sv).y, vertexes(0), numvertexes, distance)
          nv = NearestOtherVertex(sv, vertexes(0), numvertexes, distance)
          
          'Check if close enough for stitching
          If (distance <= stdnc) Then
               
               'Stitch the bitch
               StitchVertices nv, sv
               
               'Update vertices in selection array
               For sc = LBound(sel) To UBound(sel)
                    
                    'Last vertex moved to stitched vertex
                    'so apply this to selection too
                    If (CLng(sel(sc)) = numvertexes) Then sel(sc) = sv
               Next sc
               
               'Changes were made
               AutoStitchDraggedSelection = True
          End If
     Next s
End Function

Public Sub CurveLines(ByVal Verts As Long, ByVal distance As Single, ByVal Theta As Single, ByVal circlesegment As Boolean)
     Dim lines As Variant
     Dim l1x As Single, l1y As Single, l2x As Single, l2y As Single
     Dim lx As Single, ly As Single
     Dim la As Single
     Dim ld As Long
     
     Dim a As Single
     Dim c As Single
     Dim h As Single
     Dim r As Single
     Dim d As Single
               
     Dim i As Long, v As Long
     Dim nv As Long
     Dim nl As Long
     Dim fx As Single, fy As Single
     Dim fa As Single
     Dim fd As Single
     
     'Get the selected linedefs
     lines = selected.Items
     
     'Go for all selected linedefs
     For i = LBound(lines) To UBound(lines)
          
          'Get the linedef
          ld = CLng(lines(i))
          
          'Get line coordinates
          l1x = vertexes(linedefs(ld).v1).x
          l1y = vertexes(linedefs(ld).v1).y
          l2x = vertexes(linedefs(ld).v2).x
          l2y = vertexes(linedefs(ld).v2).y
          
          'Get line difference
          lx = l2x - l1x
          ly = l2y - l1y
          
          'Get the line angle
          la = ATan2(lx, ly)
          
          'Calc stuff
          'ADDED BY ANDERS ASTRAND 12/27/2003
          c = Sqr(lx * lx + ly * ly)
          d = (c / Tan(Theta / 2)) / 2
          r = d / Cos(Theta / 2)
          h = r - d
          If circlesegment Then distance = h * Sgn(distance)
          
          'Start splitting with this line
          nl = ld
          
          'Go for all split vertices
          For v = 1 To Verts
               
               'Create a vertex
               nv = CreateVertex
               
               'Calculate angle
               'MODIFIED BY ANDERS ASTRAND 12/27/2003
               a = (v * Theta / (Verts + 1)) + (pi - Theta) / 2
               
               'Create x and y on a horizontal ellipse
               'MODIFIED BY ANDERS ASTRAND 12/27/2003
               fx = Cos(a) * r
               fy = (sIn(a) * r - d) * distance / h
               
               'Get angle and distance
               fa = ATan2(fx, fy)
               fd = Sqr(fx * fx + fy * fy)
               
               'Rotate the angle for this vertex
               fa = fa + la
               
               'Coordinate the new vertex
               vertexes(nv).x = l1x + lx * 0.5 + Cos(fa) * fd
               vertexes(nv).y = l1y + ly * 0.5 + sIn(fa) * fd
               
               'Split the linedef with this vertex
               nl = SplitLinedef(ld, nv)
               
               'Select the new linedef too
               linedefs(nl).selected = 1
          Next v
     Next i
End Sub


Public Sub DEBUG_FindUnusedSectors()
     Dim i As Long
     Dim AnyFound As Long
     Dim SectorUsed() As Long
     
     ReDim SectorUsed(0 To (numsectors - 1))
     
     'Go for all linedefs
     For i = 0 To (numlinedefs - 1)
          
          'Mark sector on sidedef 1
          If (linedefs(i).s1 > -1) Then If (sidedefs(linedefs(i).s1).sector > -1) Then SectorUsed(sidedefs(linedefs(i).s1).sector) = 1
          
          'Mark sector on sidedef 2
          If (linedefs(i).s2 > -1) Then If (sidedefs(linedefs(i).s2).sector > -1) Then SectorUsed(sidedefs(linedefs(i).s2).sector) = 1
     Next i
     
     'Check for unused sectors
     For i = 0 To (numsectors - 1)
          
          'Check if sector is unused
          If (SectorUsed(i) = 0) Then
               
               'Ouput debug
               Debug.Print "Sector " & i & "   Linedefs: 0   Sidedefs: " & CountSectorSidedefs(VarPtr(sidedefs(0)), numsidedefs, i)
               AnyFound = True
          End If
     Next i
     
     'Halt when any found
     If AnyFound Then Stop
End Sub

Public Sub DeleteSelectedLinedefs()
     Dim ld As Long
     
     'Go for all selected linedefs
     Do While (ld < numlinedefs)
          
          'Check if the linedef is selected
          If (linedefs(ld).selected) Then
               
               'Delete the linedef (keep same index, will now point to moved linedef)
               RemoveLinedef ld, , , True
               
               'Map changed
               mapchanged = True
               mapnodeschanged = True
          Else
               
               'Go to next linedef index
               ld = ld + 1
          End If
     Loop
End Sub

Public Sub DeleteSelectedSectors()
     Dim s As Long
     
     'Go for all selected sectors
     Do While (s < numsectors)
          
          'Check if the sector is selected
          If (sectors(s).selected) Then
               
               'Delete the sector (keep same index, will now point to moved sector)
               RemoveSector s, True
               
               'Map changed
               mapchanged = True
               mapnodeschanged = True
          Else
               
               'Go to next sector index
               s = s + 1
          End If
     Loop
     
     'Remove looped linedefs
     RemoveLoopedLinedefs
End Sub

Public Sub DeleteSelectedThings()
     Dim t As Long
     
     'Go for all selected things
     Do While (t < numthings)
          
          'Check if the thing is selected
          If (things(t).selected) Then
               
               'Delete the thing (keep same index, will now point to moved thing)
               RemoveThing t
               
               'Map changed
               mapchanged = True
          Else
               
               'Go to next thing index
               t = t + 1
          End If
     Loop
End Sub

Public Sub DeleteSelectedVertices(ByVal MergeLinedefs As Boolean)
     Dim v As Long, nld As Long, ld As Long
     Dim foundld1 As Long
     Dim foundld2 As Long
     Dim changed As New collection
     
     'Go for all selected vertices
     Do While (v < numvertexes)
          
          'Check if the vertex is selected
          If (vertexes(v).selected) Then
               
               'Check if we should merge linedefs
               If MergeLinedefs Then
                    
                    'Get number of attached linedefs
                    nld = CountVertexLinedefs(linedefs(0), numlinedefs, v)
                    
                    'Merge when only 2 linedefs
                    If (nld = 2) Then
                         
                         'Find the first linedef
                         For ld = 0 To (numlinedefs - 1)
                              
                              'Check if this linedef is attached to this vertex
                              If (linedefs(ld).v1 = v) Then foundld1 = ld: Exit For
                              If (linedefs(ld).v2 = v) Then foundld1 = ld: Exit For
                         Next ld
                         
                         'Find the second linedef
                         For ld = (foundld1 + 1) To (numlinedefs - 1)
                              
                              'Check if this linedef is attached to this vertex
                              If (linedefs(ld).v1 = v) Then foundld2 = ld: Exit For
                              If (linedefs(ld).v2 = v) Then foundld2 = ld: Exit For
                         Next ld
                         
                         'Check if Linedef 1 points away from this vertex
                         If (linedefs(foundld1).v1 = v) Then
                              
                              'Linedef 1 points away,
                              'use its To vertex to move Linedef 2
                              If (linedefs(foundld2).v1 = v) Then
                                   linedefs(foundld2).v1 = linedefs(foundld1).v2
                              Else
                                   linedefs(foundld2).v2 = linedefs(foundld1).v2
                              End If
                              
                              'Linedef 2 changed
                              changed.Add foundld2
                              
                         'Check if Linedef 2 points away from this vertex
                         ElseIf (linedefs(foundld2).v1 = v) Then
                              
                              'Linedef 2 points away,
                              'use its To vertex to move Linedef 1
                              If (linedefs(foundld1).v1 = v) Then
                                   linedefs(foundld1).v1 = linedefs(foundld2).v2
                              Else
                                   linedefs(foundld1).v2 = linedefs(foundld2).v2
                              End If
                              
                              'Linedef 1 changed
                              changed.Add foundld1
                         Else
                              
                              'Both point towards this vertex,
                              'use From vertex of Linedef 2 to move Linedef 1
                              If (linedefs(foundld1).v1 = v) Then
                                   linedefs(foundld1).v1 = linedefs(foundld2).v1
                              Else
                                   linedefs(foundld1).v2 = linedefs(foundld2).v1
                              End If
                              
                              'Linedef 1 changed
                              changed.Add foundld1
                         End If
                    End If
               End If
               
               'Go for all linedefs
               ld = numlinedefs - 1
               Do While ld >= 0
                    
                    'Check if this line is in some way attached to this vertex
                    If ((linedefs(ld).v1 = v) Or (linedefs(ld).v2 = v)) Then
                         
                         'Remove this linedef
                         RemoveLinedef ld, True, False, True
                    End If
                    
                    'Go to next linedef
                    ld = ld - 1
               Loop
               
               'Delete the vertex (keep same index, will now point to moved vertex)
               RemoveVertex v
               
               'Map changed
               mapchanged = True
               mapnodeschanged = True
          Else
               
               'Go to next vertex index
               v = v + 1
          End If
     Loop
     
     'Remove looped linedefs
     RemoveLoopedLinedefs
     
'     'Get number of changing lines
'     numchangedlines = changed.Count
'     If (numchangedlines > 0) Then
'
'          'Allocate memory for array
'          ReDim changedlines(0 To (numchangedlines - 1))
'
'          'Go for all changed linedefs
'          For s = 1 To changed.Count
'
'               'Set in array
'               changedlines(s - 1) = changed(s)
'          Next s
'
'          'Due to linedef merging, linedefs could be overlapping
'          'Combine these into one now
'          MergeDoubleLinedefs
'     End If
End Sub

Public Sub DeleteSelection(ByVal DeleteDescription As String)
     
     'Check for selection of highlight
     If (currentselected > -1) Or (numselected > 0) Then
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'Check what to delete
          Select Case mode
               
               Case EM_VERTICES
                    
                    'Make undo backup
                    CreateUndo "vertex " & DeleteDescription
                    
                    'Turn highlight into selection when needed
                    If (numselected = 0) Then SelectCurrentVertex
                    
                    'Delete selected objects
                    DeleteSelectedVertices True '(numselected = 1)
                    
               Case EM_LINES
                    
                    'Make undo backup
                    CreateUndo "linedef " & DeleteDescription
                    
                    'Turn highlight into selection when needed
                    If (numselected = 0) Then SelectCurrentLine
                    
                    'Delete selected objects
                    DeleteSelectedLinedefs
                    
               Case EM_SECTORS
                    
                    'Make undo backup
                    CreateUndo "sector " & DeleteDescription
                    
                    'Turn highlight into selection when needed
                    If (numselected = 0) Then SelectCurrentSector
                    
                    'Delete selected objects
                    DeleteSelectedSectors
                    
               Case EM_THINGS
                    
                    'Make undo backup
                    CreateUndo "thing " & DeleteDescription
                    
                    'Turn highlight into selection when needed
                    If (numselected = 0) Then SelectCurrentThing
                    
                    'Delete selected objects
                    DeleteSelectedThings
                    
          End Select
          
          'No selected objects!
          ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
          selected.RemoveAll
          numselected = 0
          
          'Reset mousepointer
          Screen.MousePointer = vbNormal
     End If
End Sub

Public Sub FindChangingLines(ByVal IncludeStableLines As Boolean, ByVal IncludeUnstableLines As Boolean)
     Dim ld As Long, s As Long
     Dim changed As New collection
     
     'This will find all changing linedefs and add them to an array
     
     'Check if only adding stable and unstable lines
     If (IncludeStableLines And IncludeUnstableLines) Then
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if one or more of the vertices are selected
               If ((vertexes(linedefs(ld).v1).selected <> 0) Or _
                   (vertexes(linedefs(ld).v2).selected <> 0)) Then
                    
                    'Add to list
                    changed.Add ld
               End If
          Next ld
          
     'Check if only adding unstable lines
     ElseIf IncludeUnstableLines Then
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if one of the vertices are selected
               If ((vertexes(linedefs(ld).v1).selected <> 0) Xor _
                   (vertexes(linedefs(ld).v2).selected <> 0)) Then
                    
                    'Add to list
                    changed.Add ld
               End If
          Next ld
          
     'Check if only adding stable lines
     ElseIf IncludeStableLines Then
          
          'Go for all linedefs
          For ld = 0 To (numlinedefs - 1)
               
               'Check if one of the vertices are selected
               If ((vertexes(linedefs(ld).v1).selected <> 0) And _
                   (vertexes(linedefs(ld).v2).selected <> 0)) Then
                    
                    'Add to list
                    changed.Add ld
               End If
          Next ld
     End If
     
     'Get number of changing lines
     numchangedlines = changed.Count
     If (numchangedlines > 0) Then
          
          'Allocate memory for array
          ReDim changedlines(0 To (numchangedlines - 1))
          
          'Go for all changed linedefs
          For s = 1 To changed.Count
               
               'Set in array
               changedlines(s - 1) = changed(s)
          Next s
     End If
End Sub

Public Function FindFirstSequenceLine(ByRef FirstVertex As Boolean) As Long
     Dim ldIndices As Variant           'Selection indices
     Dim sldi As Long, sldk As Long     'Selected Linedef Index
     Dim sld As Long                    'Selected Linedef
     Dim v1i As Long
     Dim v2i As Long
     Dim v1 As Long
     Dim v2 As Long
     
     'This will search for the first selected linedef,
     'which has only 1 vertex to which adjacent linedefs are selected.
     
     
     'Go for all selected linedefs
     ldIndices = selected.Items
     For sldi = LBound(ldIndices) To UBound(ldIndices)
          
          'Get selected linedef
          sld = ldIndices(sldi)
          
          'Get its vertices
          v1i = linedefs(sld).v1
          v2i = linedefs(sld).v2
          
          'These are still unreferred
          v1 = 0
          v2 = 0
          
          'This linedef has 2 vertices where other selected lines can connect to.
          'Check all selected lines to see if one vertex stays unreferred by those.
          
          'Go for all select linedefs again
          For sldk = LBound(ldIndices) To UBound(ldIndices)
               
               'Check if this is not the same linedef
               If (sldk <> sldi) Then
                    
                    'Check if referring to the same vertex 1
                    If (linedefs(sldk).v1 = v1i) Then v1 = 1
                    If (linedefs(sldk).v2 = v1i) Then v1 = 1
                    
                    'Check if referring to the same vertex 2
                    If (linedefs(sldk).v1 = v2i) Then v2 = 1
                    If (linedefs(sldk).v2 = v2i) Then v2 = 1
               End If
          Next sldk
          
          'Check if only exactly one vertex is referred to
          If (v1 Xor v2) Then
               
               'This is the line we're looking for,
               'return the selection index
               FindFirstSequenceLine = sldi
               
               'Set if the first vertex is unreferenced
               FirstVertex = (v1 = 0)
               
               'And leave here now
               Exit For
          End If
     Next sldi
End Function

Public Function FindSelectAndReplace(ByVal SearchType As ENUM_FINDREPLACE, ByVal Find As String, ByVal WithinSelection As Long, ByVal Replace As String, ByVal ReplaceOnly As Long) As Long
     Dim i As Long, s As Long, e As Long
     Dim SelectionItems As Variant
     Dim ii As Long
     Dim Qualifies As Long
     Dim LongFind As Long
     Dim CurFind As String
     Dim LongReplace As Long
     Dim CurReplace As String
     Dim DoReplace As Long
     
     'Make find and replace in other datatypes for faster comparision
     LongFind = Val(Find)
     CurFind = UCase$(Trim$(Find))
     LongReplace = Val(Replace)
     CurReplace = UCase$(Trim$(Replace))
     DoReplace = (Replace <> "")
     
     'Check if searching in selection
     If (WithinSelection) Then
          
          'Go through selection
          s = 0
          e = selected.Count - 1
          
          'Get items
          SelectionItems = selected.Items
     Else
          
          'Check what array to go through
          Select Case SearchType
               
               'Lines
               Case FR_LINEDEFACTION, FR_LINEDEFNUMBER, _
                    FR_LINEDEFSECTORTAG, FR_LINEDEFTHINGTAG, _
                    FR_LINEDEFTEXTURE
                    s = 0
                    e = numlinedefs - 1
               
               'Sectors
               Case FR_SECTOREFFECT, FR_SECTORFLAT, _
                    FR_SECTORNUMBER, FR_SECTORTAG
                    s = 0
                    e = numsectors - 1
               
               'Things
               Case FR_THINGACTION, FR_THINGNUMBER, _
                    FR_THINGTAG, FR_THINGSECTORTAG, _
                    FR_THINGTHINGTAG, FR_THINGTYPE
                    s = 0
                    e = numthings - 1
               
               'Vertices
               Case FR_VERTEXNUMBER
                    s = 0
                    e = numvertexes - 1
                    
          End Select
     End If
     
     'Check if current selection must be cleared
     If (ReplaceOnly = 0) Then
          
          'Reset selection list
          Set selected = New Dictionary
          
          'Remove selection from arrays
          ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     End If
     
     'Go throught the array
     For i = s To e
          
          'Get the real item index
          If (WithinSelection) Then ii = SelectionItems(i) Else ii = i
          
          'Qualify this item
          Select Case SearchType
               
               Case FR_SECTORFLAT
                    Qualifies = (sectors(ii).tceiling Like Find) Or _
                                (sectors(ii).tfloor Like Find)
                              
               Case FR_SECTOREFFECT: Qualifies = (sectors(ii).special = LongFind)
               Case FR_SECTORNUMBER: Qualifies = (ii = LongFind)
               Case FR_SECTORTAG: Qualifies = (sectors(ii).tag = LongFind)
               Case FR_THINGACTION: Qualifies = (things(ii).effect = LongFind)
               Case FR_THINGNUMBER: Qualifies = (ii = LongFind)
               Case FR_THINGTAG: Qualifies = (things(ii).tag = LongFind)
               Case FR_THINGTYPE: Qualifies = (things(ii).thing = LongFind)
               Case FR_VERTEXNUMBER: Qualifies = (ii = LongFind)
               Case FR_LINEDEFACTION: Qualifies = (linedefs(ii).effect = LongFind)
               Case FR_LINEDEFNUMBER: Qualifies = (ii = LongFind)
               
               Case FR_THINGSECTORTAG
                    Qualifies = ((things(ii).arg0 = LongFind) And (things(ii).argref0 = 1)) Or _
                                ((things(ii).arg1 = LongFind) And (things(ii).argref1 = 1)) Or _
                                ((things(ii).arg2 = LongFind) And (things(ii).argref2 = 1)) Or _
                                ((things(ii).arg3 = LongFind) And (things(ii).argref3 = 1)) Or _
                                ((things(ii).arg4 = LongFind) And (things(ii).argref4 = 1))
                              
               Case FR_THINGTHINGTAG
                    Qualifies = ((things(ii).arg0 = LongFind) And (things(ii).argref0 = 2)) Or _
                                ((things(ii).arg1 = LongFind) And (things(ii).argref1 = 2)) Or _
                                ((things(ii).arg2 = LongFind) And (things(ii).argref2 = 2)) Or _
                                ((things(ii).arg3 = LongFind) And (things(ii).argref3 = 2)) Or _
                                ((things(ii).arg4 = LongFind) And (things(ii).argref4 = 2))
                              
               Case FR_LINEDEFSECTORTAG
                    Qualifies = (linedefs(ii).tag = LongFind) Or _
                                ((linedefs(ii).arg0 = LongFind) And (linedefs(ii).argref0 = 1)) Or _
                                ((linedefs(ii).arg1 = LongFind) And (linedefs(ii).argref1 = 1)) Or _
                                ((linedefs(ii).arg2 = LongFind) And (linedefs(ii).argref2 = 1)) Or _
                                ((linedefs(ii).arg3 = LongFind) And (linedefs(ii).argref3 = 1)) Or _
                                ((linedefs(ii).arg4 = LongFind) And (linedefs(ii).argref4 = 1))
               
               Case FR_LINEDEFTHINGTAG
                    Qualifies = ((linedefs(ii).arg0 = LongFind) And (linedefs(ii).argref0 = 2)) Or _
                                ((linedefs(ii).arg1 = LongFind) And (linedefs(ii).argref1 = 2)) Or _
                                ((linedefs(ii).arg2 = LongFind) And (linedefs(ii).argref2 = 2)) Or _
                                ((linedefs(ii).arg3 = LongFind) And (linedefs(ii).argref3 = 2)) Or _
                                ((linedefs(ii).arg4 = LongFind) And (linedefs(ii).argref4 = 2))
                              
               Case FR_LINEDEFTEXTURE
                    If (linedefs(ii).s1 > -1) Then
                         Qualifies = (sidedefs(linedefs(ii).s1).lower Like Find) Or _
                                     (sidedefs(linedefs(ii).s1).middle Like Find) Or _
                                     (sidedefs(linedefs(ii).s1).upper Like Find)
                    End If
                    If (linedefs(ii).s2 > -1) Then
                         Qualifies = Qualifies Or _
                                     (sidedefs(linedefs(ii).s2).lower Like Find) Or _
                                     (sidedefs(linedefs(ii).s2).middle Like Find) Or _
                                     (sidedefs(linedefs(ii).s2).upper Like Find)
                    End If
               
          End Select
          
          'Check if qualifies
          If (Qualifies) Then
               
               'Check it must be selected
               If (ReplaceOnly = 0) Then
                    
                    'Add to list of selected items
                    selected.Add CStr(ii), ii
                    
                    'Check in what array to select
                    Select Case selectedtype
                         Case EM_VERTICES: vertexes(ii).selected = 1
                         Case EM_LINES: linedefs(ii).selected = 1
                         Case EM_SECTORS: SelectSectorRef ii: sectors(ii).selected = 1
                         Case EM_THINGS: things(ii).selected = 1
                    End Select
               End If
               
               'Check it must be replaced
               If (DoReplace) Then
                    
                    'Replace the item
                    Select Case SearchType
                         
                         Case FR_SECTORFLAT
                              If (sectors(ii).tceiling = CurFind) Then sectors(ii).tceiling = CurReplace
                              If (sectors(ii).tfloor = CurFind) Then sectors(ii).tfloor = CurReplace
                              
                         Case FR_SECTOREFFECT: sectors(ii).special = LongReplace
                         Case FR_SECTORTAG: sectors(ii).tag = LongReplace
                         Case FR_THINGACTION: things(ii).effect = LongReplace
                         Case FR_THINGTAG: things(ii).tag = LongReplace
                         
                         Case FR_THINGTYPE
                              things(ii).thing = LongReplace
                              
                              'Update thing image, color and size
                              UpdateThingImageColor ii
                              UpdateThingSize ii
                              UpdateThingCategory ii
                              
                         Case FR_LINEDEFACTION: linedefs(ii).effect = LongReplace
                         
                         Case FR_LINEDEFSECTORTAG
                              If (linedefs(ii).tag = LongFind) Then linedefs(ii).tag = LongReplace
                              If (linedefs(ii).arg0 = LongFind) And (linedefs(ii).argref0 = 1) Then linedefs(ii).arg0 = LongReplace
                              If (linedefs(ii).arg1 = LongFind) And (linedefs(ii).argref1 = 1) Then linedefs(ii).arg1 = LongReplace
                              If (linedefs(ii).arg2 = LongFind) And (linedefs(ii).argref2 = 1) Then linedefs(ii).arg2 = LongReplace
                              If (linedefs(ii).arg3 = LongFind) And (linedefs(ii).argref3 = 1) Then linedefs(ii).arg3 = LongReplace
                              If (linedefs(ii).arg4 = LongFind) And (linedefs(ii).argref4 = 1) Then linedefs(ii).arg4 = LongReplace
                              
                         Case FR_LINEDEFTHINGTAG
                              If (linedefs(ii).arg0 = LongFind) And (linedefs(ii).argref0 = 2) Then linedefs(ii).arg0 = LongReplace
                              If (linedefs(ii).arg1 = LongFind) And (linedefs(ii).argref1 = 2) Then linedefs(ii).arg1 = LongReplace
                              If (linedefs(ii).arg2 = LongFind) And (linedefs(ii).argref2 = 2) Then linedefs(ii).arg2 = LongReplace
                              If (linedefs(ii).arg3 = LongFind) And (linedefs(ii).argref3 = 2) Then linedefs(ii).arg3 = LongReplace
                              If (linedefs(ii).arg4 = LongFind) And (linedefs(ii).argref4 = 2) Then linedefs(ii).arg4 = LongReplace
                              
                         Case FR_LINEDEFTEXTURE
                              If (linedefs(ii).s1 > -1) Then
                                   If (sidedefs(linedefs(ii).s1).lower = CurFind) Then sidedefs(linedefs(ii).s1).lower = CurReplace
                                   If (sidedefs(linedefs(ii).s1).middle = CurFind) Then sidedefs(linedefs(ii).s1).middle = CurReplace
                                   If (sidedefs(linedefs(ii).s1).upper = CurFind) Then sidedefs(linedefs(ii).s1).upper = CurReplace
                              End If
                              If (linedefs(ii).s2 > -1) Then
                                   If (sidedefs(linedefs(ii).s2).lower = CurFind) Then sidedefs(linedefs(ii).s2).lower = CurReplace
                                   If (sidedefs(linedefs(ii).s2).middle = CurFind) Then sidedefs(linedefs(ii).s2).middle = CurReplace
                                   If (sidedefs(linedefs(ii).s2).upper = CurFind) Then sidedefs(linedefs(ii).s2).upper = CurReplace
                              End If
                         
                         Case FR_THINGSECTORTAG
                              If (things(ii).arg0 = LongFind) And (things(ii).argref0 = 1) Then things(ii).arg0 = LongReplace
                              If (things(ii).arg1 = LongFind) And (things(ii).argref1 = 1) Then things(ii).arg1 = LongReplace
                              If (things(ii).arg2 = LongFind) And (things(ii).argref2 = 1) Then things(ii).arg2 = LongReplace
                              If (things(ii).arg3 = LongFind) And (things(ii).argref3 = 1) Then things(ii).arg3 = LongReplace
                              If (things(ii).arg4 = LongFind) And (things(ii).argref4 = 1) Then things(ii).arg4 = LongReplace
                              
                         Case FR_THINGTHINGTAG
                              If (things(ii).arg0 = LongFind) And (things(ii).argref0 = 2) Then things(ii).arg0 = LongReplace
                              If (things(ii).arg1 = LongFind) And (things(ii).argref1 = 2) Then things(ii).arg1 = LongReplace
                              If (things(ii).arg2 = LongFind) And (things(ii).argref2 = 2) Then things(ii).arg2 = LongReplace
                              If (things(ii).arg3 = LongFind) And (things(ii).argref3 = 2) Then things(ii).arg3 = LongReplace
                              If (things(ii).arg4 = LongFind) And (things(ii).argref4 = 2) Then things(ii).arg4 = LongReplace
                              
                    End Select
               End If
               
               'Count the item
               FindSelectAndReplace = FindSelectAndReplace + 1
          End If
     Next i
     
     'Check if map changed
     If (DoReplace <> 0) And (FindSelectAndReplace > 0) Then
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
     End If
     
     'Currect the new selection
     numselected = selected.Count
End Function

Public Sub FixMissingSidedefTextures(ByVal SidedefIndex As Long)
     Dim ld As Long
     
     'Fix any incorrect user input
     If (Trim$(Config("defaulttexture")("upper")) = "") Then Config("defaulttexture")("upper") = "-"
     If (Trim$(Config("defaulttexture")("middle")) = "") Then Config("defaulttexture")("middle") = "-"
     If (Trim$(Config("defaulttexture")("lower")) = "") Then Config("defaulttexture")("lower") = "-"
     
     'Get the linedef
     ld = sidedefs(SidedefIndex).linedef
     
     'Check which sidedef it is
     If (linedefs(ld).s1 = SidedefIndex) Then
          
          'Sidedef 1
          
          'Set upper texture if required
          If (RequiresS1Upper(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).upper, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).upper) = 0)) Then sidedefs(SidedefIndex).upper = UCase$(Config("defaulttexture")("upper"))
          
          'Set middle texture if required
          If (RequiresS1Middle(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).middle, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).middle) = 0)) Then sidedefs(SidedefIndex).middle = UCase$(Config("defaulttexture")("middle"))
          
          'Set lower texture if required
          If (RequiresS1Lower(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).lower, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).lower) = 0)) Then sidedefs(SidedefIndex).lower = UCase$(Config("defaulttexture")("lower"))
     Else
          
          'Sidedef 2
          
          'Set upper texture if required
          If (RequiresS2Upper(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).upper, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).upper) = 0)) Then sidedefs(SidedefIndex).upper = UCase$(Config("defaulttexture")("upper"))
          
          'Set middle texture if required
          If (RequiresS2Middle(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).middle, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).middle) = 0)) Then sidedefs(SidedefIndex).middle = UCase$(Config("defaulttexture")("middle"))
          
          'Set lower texture if required
          If (RequiresS2Lower(ld) = True) And _
             ((StrComp(sidedefs(SidedefIndex).lower, "-", vbBinaryCompare) = 0) Or _
              (LenB(sidedefs(SidedefIndex).lower) = 0)) Then sidedefs(SidedefIndex).lower = UCase$(Config("defaulttexture")("lower"))
     End If
End Sub

Public Sub FlipLinedefSidedefs(ByVal LinedefIndex As Long)
     Dim s As Long
     
     'Flip linedef sidedefs
     s = linedefs(LinedefIndex).s1
     linedefs(LinedefIndex).s1 = linedefs(LinedefIndex).s2
     linedefs(LinedefIndex).s2 = s
End Sub

Public Sub FlipLinedefVertices(ByVal LinedefIndex As Long)
     Dim v As Long
     
     'Flip linedef vertices
     v = linedefs(LinedefIndex).v1
     linedefs(LinedefIndex).v1 = linedefs(LinedefIndex).v2
     linedefs(LinedefIndex).v2 = v
End Sub

Public Sub FlipThingsHorizontal()
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim FlipLine As Long
     Dim i As Long
     Dim na As Long
     
     'Go for all items
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    If (things(i).x < selrect.left) Then selrect.left = things(i).x
                    If (things(i).x > selrect.right) Then selrect.right = things(i).x
                    'If (things(i).Y < selrect.top) Then selrect.top = things(i).Y
                    'If (things(i).Y > selrect.bottom) Then selrect.bottom = things(i).Y
               Else
                    selrect.left = things(i).x
                    selrect.right = things(i).x
                    'selrect.top = things(i).Y
                    'selrect.bottom = things(i).Y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine flip line
     FlipLine = selrect.left + (selrect.right - selrect.left) / 2
     
     'Start flipping them
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Flip over flip line
               things(i).x = FlipLine + (FlipLine - things(i).x)
               
               'Crop angle
               While (things(i).angle >= 360): things(i).angle = things(i).angle - 360: Wend
               While (things(i).angle < 0): things(i).angle = things(i).angle + 360: Wend
               
               'Check quadrant
               na = things(i).angle
               If (na >= 0) And (na < 90) Then
                    
                    'NE
                    things(i).angle = na + (180 - na * 2)
                    
               ElseIf (na >= 90) And (na <= 180) Then
                    
                    'NW
                    things(i).angle = na - (na - 90) * 2
                    
               ElseIf (na >= 180) And (na < 270) Then
                    
                    'SW
                    things(i).angle = na + (180 - (na - 180) * 2)
                    
               Else
                    
                    'SE
                    things(i).angle = na - (na - 270) * 2
               End If
               
               'Crop angle
               While (things(i).angle >= 360): things(i).angle = things(i).angle - 360: Wend
               While (things(i).angle < 0): things(i).angle = things(i).angle + 360: Wend
               
               'Set new image for thing
               UpdateThingImageColor i
          End If
     Next i
End Sub

Public Sub FlipThingsVertical()
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim FlipLine As Long
     Dim i As Long
     Dim na As Long
     
     'Go for all items
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    'If (things(i).X < selrect.left) Then selrect.left = things(i).X
                    'If (things(i).X > selrect.right) Then selrect.right = things(i).X
                    If (things(i).y < selrect.top) Then selrect.top = things(i).y
                    If (things(i).y > selrect.bottom) Then selrect.bottom = things(i).y
               Else
                    'selrect.left = things(i).X
                    'selrect.right = things(i).X
                    selrect.top = things(i).y
                    selrect.bottom = things(i).y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine flip line
     FlipLine = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start flipping them
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Flip over flip line
               things(i).y = FlipLine + (FlipLine - things(i).y)
               
               'Crop angle
               While (things(i).angle >= 360): things(i).angle = things(i).angle - 360: Wend
               While (things(i).angle < 0): things(i).angle = things(i).angle + 360: Wend
               
               'Check quadrant
               na = things(i).angle
               If (na >= 0) And (na < 90) Then
                    
                    'NE
                    things(i).angle = na - (na * 2)
                    
               ElseIf (na >= 90) And (na <= 180) Then
                    
                    'NW
                    things(i).angle = na + (180 - na) * 2
                    
               ElseIf (na >= 180) And (na < 270) Then
                    
                    'SW
                    things(i).angle = na - (na - 180) * 2
                    
               Else
                    
                    'SE
                    things(i).angle = na + (360 - na) * 2
               End If
               
               'Crop angle
               While (things(i).angle >= 360): things(i).angle = things(i).angle - 360: Wend
               While (things(i).angle < 0): things(i).angle = things(i).angle + 360: Wend
               
               'Set new image for thing
               UpdateThingImageColor i
          End If
     Next i
End Sub

Public Sub FlipVerticesHorizontal()
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim FlipLine As Long
     Dim i As Long
     Dim v As Long
     
     'Find all lines that will flip (stable)
     FindChangingLines True, False
     
     'Go for all items
     For i = 0 To (numvertexes - 1)
          
          'Check if selected
          If (vertexes(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    If (vertexes(i).x < selrect.left) Then selrect.left = vertexes(i).x
                    If (vertexes(i).x > selrect.right) Then selrect.right = vertexes(i).x
                    'If (vertexes(i).Y < selrect.top) Then selrect.top = vertexes(i).Y
                    'If (vertexes(i).Y > selrect.bottom) Then selrect.bottom = vertexes(i).Y
               Else
                    selrect.left = vertexes(i).x
                    selrect.right = vertexes(i).x
                    'selrect.top = vertexes(i).Y
                    'selrect.bottom = vertexes(i).Y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine flip line
     FlipLine = selrect.left + (selrect.right - selrect.left) / 2
     
     'Start flipping them
     For i = 0 To (numvertexes - 1)
          
          'Check if selected
          If (vertexes(i).selected > 0) Then
               
               'Flip over flip line
               vertexes(i).x = FlipLine + (FlipLine - vertexes(i).x)
          End If
     Next i
     
     'Go for all lines to flip vertices
     'because lines were flipped over
     For i = 0 To (numchangedlines - 1)
          
          'Flip this lines vertices
          v = linedefs(changedlines(i)).v1
          linedefs(changedlines(i)).v1 = linedefs(changedlines(i)).v2
          linedefs(changedlines(i)).v2 = v
     Next i
End Sub

Public Sub FlipVerticesVertical()
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim FlipLine As Long
     Dim i As Long
     Dim v As Long
     
     'Find all lines that will flip (stable)
     FindChangingLines True, False
     
     'Go for all items
     For i = 0 To (numvertexes - 1)
          
          'Check if selected
          If (vertexes(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    'If (vertexes(i).X < selrect.left) Then selrect.left = vertexes(i).X
                    'If (vertexes(i).X > selrect.right) Then selrect.right = vertexes(i).X
                    If (vertexes(i).y < selrect.top) Then selrect.top = vertexes(i).y
                    If (vertexes(i).y > selrect.bottom) Then selrect.bottom = vertexes(i).y
               Else
                    'selrect.left = vertexes(i).X
                    'selrect.right = vertexes(i).X
                    selrect.top = vertexes(i).y
                    selrect.bottom = vertexes(i).y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine flip line
     FlipLine = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start flipping them
     For i = 0 To (numvertexes - 1)
          
          'Check if selected
          If (vertexes(i).selected > 0) Then
               
               'Flip over flip line
               vertexes(i).y = FlipLine + (FlipLine - vertexes(i).y)
          End If
     Next i
     
     'Go for all lines to flip vertices
     'because lines were flipped over
     For i = 0 To (numchangedlines - 1)
          
          'Flip this lines vertices
          v = linedefs(changedlines(i)).v1
          linedefs(changedlines(i)).v1 = linedefs(changedlines(i)).v2
          linedefs(changedlines(i)).v2 = v
     Next i
End Sub

Public Function InsertVertex(ByVal x As Long, ByVal y As Long) As Long
     Dim v As Long
     
     'InsertVertex:
     'Create new vertex and place it at position x, y
     'Returns the index of the new vertex
     
     'Create vertex
     v = CreateVertex
     
     'Set position
     vertexes(v).x = x
     vertexes(v).y = y
     
     'Not selected
     vertexes(v).selected = 0
     
     'Return index
     InsertVertex = v
End Function

Public Sub JoinSelectedSectors()
     Dim TargetSector As Long
     Dim s As Long
     Dim SelectedItems As Variant
     
     'The target sector will be the first selected
     TargetSector = selected.Items(0)
     
     'Go for all sidedefs
     For s = 0 To (numsidedefs - 1)
          
          'Check if the sector, to which this sidedef belongs, is selected
          If selected.Exists(CStr(sidedefs(s).sector)) Then
               
               'Change the sidedef sector
               sidedefs(s).sector = TargetSector
          End If
     Next s
     
     'Get indices
     SelectedItems = selected.Items
     
     'All selected sectors beside the first one are now unused
     'Go for all other selected sectors
     For s = (selected.Count - 1) To 1 Step -1
          
          'Remove this sector
          RemoveSector SelectedItems(s), True
     Next s
     
     'Selection is now different
     Set selected = New Dictionary
     selected.Add CStr(TargetSector), TargetSector
     numselected = 1
End Sub

Public Sub MergeDoubleLinedefs()
     Dim i As Long
     Dim cl As Long
     Dim ol As Long
     
     'MergeDoubleLinedefs:
     '> Finds linedefs in the changed lines array which go from/to
     '  same vertices and merges these together.
     
     'Go for all changed lines
     i = (numchangedlines - 1)
     Do While (i >= 0)
          
          'Get the line index
          cl = changedlines(i)
          
          'Get the overlapping line
          ol = LinedefBetweenVertices(linedefs(0), numlinedefs, 0, linedefs(cl).v1, linedefs(cl).v2, cl)
          
          'Check if overlapping
          Do While (ol > -1)
               
               'Now that both overlapping linedefs are found,
               'try to merge the two linedefs
               MergeLinedefs cl, ol
               
               'Find next overlapping
               cl = ol
               ol = LinedefBetweenVertices(linedefs(0), numlinedefs, 0, linedefs(cl).v1, linedefs(cl).v2, cl)
          Loop
          
          'Next changed line
          i = i - 1
     Loop
End Sub

Public Function FindAdjoiningSidedef(ByVal PreferredX As Long, ByVal PreferredY As Long) As Long
     Dim ldIndices As Variant
     Dim i As Long
     Dim cl As Long
     Dim ol As Long
     Dim v1 As MAPVERTEX
     Dim v2 As MAPVERTEX
     
     'FindAdjoiningSidedef:
     '> Finds a sidedef on the first line that overlaps one of the
     '> selected linedefs witht the preferred side given
     
     'Return -1 when none can be found
     FindAdjoiningSidedef = -1
     
     'Go for all selected linedefs
     ldIndices = selected.Items
     For i = LBound(ldIndices) To UBound(ldIndices)
          
          'Get the line index
          cl = ldIndices(i)
          
          'Get the overlapping line
          ol = LinedefBetweenVertices(linedefs(0), numlinedefs, 0, linedefs(cl).v1, linedefs(cl).v2, cl)
          
          'Check if there is an overlaping line here
          If (ol > -1) Then
               
               'Get line vertices
               v1 = vertexes(linedefs(ol).v1)
               v2 = vertexes(linedefs(ol).v2)
               
               'Get the side on which the preferred point is
               If (side_of_line(v1.x, v1.y, v2.x, v2.y, PreferredX, PreferredY) < 0) Then
                    
                    'Front Side
                    'Does this linedef have a front sidedef?
                    If (linedefs(ol).s1 > -1) Then
                         
                         'Return this sidedef index
                         FindAdjoiningSidedef = linedefs(ol).s1
                    Else
                         
                         'Return the other side index
                         FindAdjoiningSidedef = linedefs(ol).s2
                    End If
               Else
                    
                    'Back Side
                    'Does this linedef have a back sidedef?
                    If (linedefs(ol).s2 > -1) Then
                         
                         'Return this sidedef index
                         FindAdjoiningSidedef = linedefs(ol).s2
                    Else
                         
                         'Return the other side index
                         FindAdjoiningSidedef = linedefs(ol).s1
                    End If
               End If
               
               'And leave the search
               Exit For
          End If
     Next i
End Function


Public Sub MergeLinedefs(ByVal cl As Long, ByVal ol As Long)
     Dim bothsinglesided As Long
     Dim l1s1s As Long
     Dim l1s2s As Long
     Dim l2s1s As Long
     Dim l2s2s As Long
     
     'check if they both some way refer to the same sector
     '(either on sidedef 1 or sidedef 2)
     
     'First get the sector references (where -1 is no sector)
     If (linedefs(ol).s1 > -1) Then l1s1s = sidedefs(linedefs(ol).s1).sector Else l1s1s = -1
     If (linedefs(ol).s2 > -1) Then l1s2s = sidedefs(linedefs(ol).s2).sector Else l1s2s = -1
     If (linedefs(cl).s1 > -1) Then l2s1s = sidedefs(linedefs(cl).s1).sector Else l2s1s = -1
     If (linedefs(cl).s2 > -1) Then l2s2s = sidedefs(linedefs(cl).s2).sector Else l2s2s = -1
     
     'Compare L1S1 and L2S1
     If (l1s1s = l2s1s) Then
          
          'Copy texture from removing sidedefs
          AddSidedefTextures linedefs(ol).s1, linedefs(cl).s2
          AddSidedefTextures linedefs(cl).s1, linedefs(ol).s2
          
          'Remove L1S1
          If (linedefs(ol).s1 > -1) Then RemoveSidedef linedefs(ol).s1, False, True, False
          
          'Move L2S2 to L1S1
          linedefs(ol).s1 = linedefs(cl).s2
          If (linedefs(ol).s1 > -1) Then sidedefs(linedefs(ol).s1).linedef = ol
          linedefs(cl).s2 = -1
          
     'Compare L1S2 and L2S2
     ElseIf (l1s2s = l2s2s) Then
          
          'Copy texture from removing sidedefs
          AddSidedefTextures linedefs(ol).s2, linedefs(cl).s1
          AddSidedefTextures linedefs(cl).s2, linedefs(ol).s1
          
          'Remove L1S2
          If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
          
          'Move L2S1 to L1S2
          linedefs(ol).s2 = linedefs(cl).s1
          If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
          linedefs(cl).s1 = -1
          
     'Compare L1S1 and L2S2
     ElseIf (l1s1s = l2s2s) Then
          
          'Copy texture from removing sidedefs
          AddSidedefTextures linedefs(ol).s1, linedefs(cl).s1
          AddSidedefTextures linedefs(cl).s2, linedefs(ol).s2
          
          'Remove L1S1
          If (linedefs(ol).s1 > -1) Then RemoveSidedef linedefs(ol).s1, False, True, False
          
          'Move L2S1 to L1S1
          linedefs(ol).s1 = linedefs(cl).s1
          If (linedefs(ol).s1 > -1) Then sidedefs(linedefs(ol).s1).linedef = ol
          linedefs(cl).s1 = -1
          
     'Compare L1S2 and L2S1
     ElseIf (l1s2s = l2s1s) Then
          
          'Copy texture from removing sidedefs
          AddSidedefTextures linedefs(ol).s2, linedefs(cl).s2
          AddSidedefTextures linedefs(cl).s1, linedefs(ol).s1
          
          'Remove L1S2
          If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
          
          'Move L2S2 to L1S2
          linedefs(ol).s2 = linedefs(cl).s2
          If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
          linedefs(cl).s2 = -1
     
     'When no valid pair could be found
     Else
          
          'Check if L1 has no S2
          If (l1s2s = 1) Then
               
               'Check if L2 is with his back to this line
               If (linedefs(cl).v1 = linedefs(ol).v2) Then
                    
                    'Use the S1 of L2 to make S2 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s2, linedefs(cl).s1
                    AddSidedefTextures linedefs(cl).s2, linedefs(ol).s1
                    
                    'Remove L1S2
                    If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
                    
                    'Move L2S1 to L1S2
                    linedefs(ol).s2 = linedefs(cl).s1
                    If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
                    linedefs(cl).s1 = -1
               Else
                    
                    'Use the S2 of L2 to make S2 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s2, linedefs(cl).s2
                    AddSidedefTextures linedefs(cl).s1, linedefs(ol).s1
                    
                    'Remove L1S2
                    If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
                    
                    'Move L2S2 to L1S2
                    linedefs(ol).s2 = linedefs(cl).s2
                    If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
                    linedefs(cl).s2 = -1
               End If
               
          'Check if L2 has no S2
          ElseIf (l2s2s = -1) Then
               
               'Check if L1 is with his back to this line
               If (linedefs(ol).v1 = linedefs(cl).v2) Then
                    
                    'Use S1 of L2 to make S2 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s2, linedefs(cl).s1
                    AddSidedefTextures linedefs(cl).s2, linedefs(ol).s1
                    
                    'Remove L1S2
                    If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
                    
                    'Move L2S1 to L1S2
                    linedefs(ol).s2 = linedefs(cl).s1
                    If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
                    linedefs(cl).s1 = -1
               Else
                    
                    'Use S1 of L2 to make S1 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s1, linedefs(cl).s1
                    AddSidedefTextures linedefs(cl).s2, linedefs(ol).s2
                    
                    'Remove L1S1
                    If (linedefs(ol).s1 > -1) Then RemoveSidedef linedefs(ol).s1, False, True, False
                    
                    'Move L2S1 to L1S1
                    linedefs(ol).s1 = linedefs(cl).s1
                    If (linedefs(ol).s1 > -1) Then sidedefs(linedefs(ol).s1).linedef = ol
                    linedefs(cl).s1 = -1
               End If
               
          'When no line without second sidedef could be found
          Else
               
               'Check if L2 is with his back to this line
               If (linedefs(cl).v1 = linedefs(ol).v2) Then
                    
                    'Use the S1 of L2 to make S2 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s2, linedefs(cl).s1
                    AddSidedefTextures linedefs(cl).s2, linedefs(ol).s1
                    
                    'Remove L1S2
                    If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
                    
                    'Move L2S1 to L1S2
                    linedefs(ol).s2 = linedefs(cl).s1
                    If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
                    linedefs(cl).s1 = -1
               Else
                    
                    'Use the S2 of L2 to make S2 of L1
                    
                    'Copy texture from removing sidedefs
                    AddSidedefTextures linedefs(ol).s2, linedefs(cl).s2
                    AddSidedefTextures linedefs(cl).s1, linedefs(ol).s1
                    
                    'Remove L1S2
                    If (linedefs(ol).s2 > -1) Then RemoveSidedef linedefs(ol).s2, False, True, False
                    
                    'Move L2S2 to L1S2
                    linedefs(ol).s2 = linedefs(cl).s2
                    If (linedefs(ol).s2 > -1) Then sidedefs(linedefs(ol).s2).linedef = ol
                    linedefs(cl).s2 = -1
               End If
          End If
     End If
     
     'If either of two lines were selected, keep this selected
     linedefs(ol).selected = (linedefs(ol).selected Or linedefs(cl).selected)
     If (selected.Exists(CStr(ol)) = False) And (linedefs(ol).selected <> 0) Then
          selected.Add CStr(ol), ol
          numselected = selected.Count
     End If
     
     'Check if both line are singlesided
     'because in that case we are always allowed to clear middle texture later on
     bothsinglesided = ((linedefs(ol).Flags And LDF_TWOSIDED) = 0) And ((linedefs(cl).Flags And LDF_TWOSIDED) = 0)
     
     'Set doublesided flag
     If ((linedefs(ol).s1 > -1) And (linedefs(ol).s2 > -1)) Then
          linedefs(ol).Flags = linedefs(ol).Flags Or LDF_TWOSIDED
          linedefs(ol).Flags = linedefs(ol).Flags And Not LDF_IMPASSIBLE
     Else
          linedefs(ol).Flags = linedefs(ol).Flags And Not LDF_TWOSIDED
          linedefs(ol).Flags = linedefs(ol).Flags Or LDF_IMPASSIBLE
     End If
     
     'Copy linedef action and tags
     If (linedefs(ol).effect = 0) Then
          linedefs(ol).effect = linedefs(cl).effect
          linedefs(ol).arg0 = linedefs(cl).arg0
          linedefs(ol).arg1 = linedefs(cl).arg1
          linedefs(ol).arg2 = linedefs(cl).arg2
          linedefs(ol).arg3 = linedefs(cl).arg3
          linedefs(ol).arg4 = linedefs(cl).arg4
          linedefs(ol).tag = linedefs(cl).tag
     End If
     
     'Remove unused textures
     If (linedefs(ol).s1 > -1) Then RemoveUnusedSidedefTextures linedefs(ol).s1, (((linedefs(ol).Flags And LDF_TWOSIDED) = LDF_TWOSIDED) And (submode = ESM_PASTING)) Or bothsinglesided
     If (linedefs(ol).s2 > -1) Then RemoveUnusedSidedefTextures linedefs(ol).s2, (((linedefs(ol).Flags And LDF_TWOSIDED) = LDF_TWOSIDED) And (submode = ESM_PASTING)) Or bothsinglesided
     
     'If ol was the last line, it now took place of cl
     If (ol = numlinedefs - 1) Then ol = cl
     
     'Remove L2
     RemoveLinedef cl, True, False, True
End Sub

Public Sub ReapplyVerticesSelection()
     Dim i As Long
     Dim Indices As Variant
     
     ' Go for all items in selection
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Select vertex
          vertexes(Indices(i)).selected = 1
     Next i
End Sub

Public Sub RemoveLinesSelection()
     Dim i As Long
     Dim Indices As Variant
     
     'Go for all selected objects
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Remove selected flag from object
          linedefs(Indices(i)).selected = 0
     Next i
     
     'No more selected objects
     selected.RemoveAll
     numselected = 0
End Sub

Public Sub RemoveLoopedLinedefs()
     Dim ld As Long
     
     'Start with last linedef
     ld = (numlinedefs - 1)
     
     'Continue until all done
     Do Until (ld < 0)
          
          'Remove the linedef if this linedef loops back
          If (linedefs(ld).v1 = linedefs(ld).v2) Then RemoveLinedef ld, True, True, True
          
          'Next linedef
          ld = ld - 1
     Loop
End Sub

Public Sub RemoveSectorsSelection()
     Dim i As Long
     Dim Indices As Variant
     
     'Go for all selected objects
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Remove selected flag from object
          sectors(Indices(i)).selected = 0
     Next i
     
     'Go for all sidedefs
     For i = 0 To (numsidedefs - 1)
          
          'Mark the linedef as normal
          linedefs(sidedefs(i).linedef).selected = 0
     Next i
     
     'No more selected objects
     selected.RemoveAll
     numselected = 0
End Sub

Public Sub RemoveSelectedSharedLinedefs()
     Dim ld As Long
     
     'For each linedef, both sectors of the sidedefs must be selected AND must be different
     
     'Go for all linedefs
     ld = numlinedefs - 1
     Do While ld >= 0
          
          'Check for a right sidedef
          If (linedefs(ld).s1 >= 0) Then
               
               'Check if the right sidedef's sector is selected
               If (selected.Exists(CStr(sidedefs(linedefs(ld).s1).sector))) Then
                    
                    'Check for a left sidedef
                    If (linedefs(ld).s2 >= 0) Then
                         
                         'Check if the left sidedef's sector is selected
                         If (selected.Exists(CStr(sidedefs(linedefs(ld).s2).sector))) Then
                              
                              'Check if the sectors are different
                              If (sidedefs(linedefs(ld).s1).sector <> sidedefs(linedefs(ld).s2).sector) Then
                                   
                                   'Remove this linedef now
                                   RemoveLinedef ld, True, True, False
                              End If
                         End If
                    End If
               End If
          End If
          
          'Next (previous) linedef
          ld = ld - 1
     Loop
End Sub

Public Sub RemoveSelection(ByVal Redraw As Boolean)
     
     'Deselect everything
     ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     
     'No more selected objects
     'selected.RemoveAll
     Set selected = New Dictionary
     numselected = 0
     
     'Redraw entire map
     If Redraw Then RedrawMap True
End Sub

Public Sub RemoveThingsSelection()
     Dim i As Long
     Dim Indices As Variant
     
     'Go for all selected objects
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Remove selected flag from object
          things(Indices(i)).selected = 0
     Next i
     
     'No more selected objects
     selected.RemoveAll
     numselected = 0
End Sub

Public Sub RemoveUnusedSidedefTextures(ByVal SidedefIndex As Long, Optional ByVal CleanMiddleToo As Boolean = False)
     Dim ld As Long
     
     'Get the linedef
     ld = sidedefs(SidedefIndex).linedef
     
     'Check which sidedef it is
     If (linedefs(ld).s1 = SidedefIndex) Then
          
          'Sidedef 1
          
          'Remove upper texture if not needed
          If (RequiresS1Upper(ld) = False) Then sidedefs(SidedefIndex).upper = "-"
          
          'Remove middle texture if not needed
          If (RequiresS1Middle(ld) = False) And (CleanMiddleToo = True) Then sidedefs(SidedefIndex).middle = "-"
          
          'Remove lower texture if not needed
          If (RequiresS1Lower(ld) = False) Then sidedefs(SidedefIndex).lower = "-"
     Else
          
          'Sidedef 2
          
          'Remove upper texture if not needed
          If (RequiresS2Upper(ld) = False) Then sidedefs(SidedefIndex).upper = "-"
          
          'Remove middle texture if not needed
          If (RequiresS2Middle(ld) = False) And (CleanMiddleToo = True) Then sidedefs(SidedefIndex).middle = "-"
          
          'Remove lower texture if not needed
          If (RequiresS2Lower(ld) = False) Then sidedefs(SidedefIndex).lower = "-"
     End If
End Sub

Public Sub RemoveUnusedVertices()
     Dim i As Long
     Dim UsedVertices As New Dictionary
     
     'Go for all linedefs
     For i = 0 To (numlinedefs - 1)
          
          'Add the used vertices if not already added
          If (UsedVertices.Exists(linedefs(i).v1) = False) Then UsedVertices.Add linedefs(i).v1, 0
          If (UsedVertices.Exists(linedefs(i).v2) = False) Then UsedVertices.Add linedefs(i).v2, 0
     Next i
     
     'Go for all vertices
     i = numvertexes - 1
     Do While i >= 0
          
          'Check if this vertex is not used
          If (UsedVertices.Exists(i) = False) Then
               
               'Remove vertex
               RemoveVertex i
          End If
          
          'Go check the previous vertex
          i = i - 1
     Loop
End Sub

Public Sub RemoveVertexSelection()
     Dim i As Long
     Dim Indices As Variant
     
     'Go for all selected objects
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Remove selected flag from object
          vertexes(Indices(i)).selected = 0
     Next i
     
     'No more selected objects
     selected.RemoveAll
     numselected = 0
End Sub

Public Sub ReselectLinedefs(Optional ByVal FirstSelected As Long = -1)
     Dim i As Long
     
     'Make new selection dictionary
     'based on selected flags
     
     'Clear list
     Set selected = New Dictionary
     
     'Check if we should add a first selected linedef
     If (FirstSelected > -1) Then
          
          'Add this linedef first, if selected
          If (linedefs(FirstSelected).selected <> 0) Then selected.Add CStr(FirstSelected), FirstSelected
     End If
     
     'Go for all
     For i = 0 To (numlinedefs - 1)
          
          'Add to list if selected
          If (linedefs(i).selected <> 0) And (selected.Exists(CStr(i)) = False) Then selected.Add CStr(i), i
     Next i
     
     'Set number of selected items
     numselected = selected.Count
     selectedtype = EM_LINES
End Sub

Public Sub ReselectSectors()
     Dim i As Long
     
     'Make new selection dictionary
     'based on selected flags
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all
     For i = 0 To (numsectors - 1)
          
          'Add to list if selected
          If (sectors(i).selected <> 0) And (selected.Exists(CStr(i)) = False) Then selected.Add CStr(i), i
     Next i
     
     'Set number of selected items
     numselected = selected.Count
     selectedtype = EM_SECTORS
End Sub

Public Sub ReselectThings()
     Dim i As Long
     
     'Make new selection dictionary
     'based on selected flags
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all
     For i = 0 To (numthings - 1)
          
          'Add to list if selected
          If (things(i).selected <> 0) And (selected.Exists(CStr(i)) = False) Then selected.Add CStr(i), i
     Next i
     
     'Set number of selected items
     numselected = selected.Count
     selectedtype = EM_THINGS
End Sub

Public Sub ReselectVertices()
     Dim i As Long
     
     'Make new selection dictionary
     'based on selected flags
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all
     For i = 0 To (numvertexes - 1)
          
          'Add to list if selected
          If (vertexes(i).selected <> 0) And (selected.Exists(CStr(i)) = False) Then selected.Add CStr(i), i
     Next i
     
     'Set number of selected items
     numselected = selected.Count
     selectedtype = EM_VERTICES
End Sub

Public Sub RotateThings(ByVal Amount As Double)
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim i As Long
     Dim tx As Single, ty As Single
     Dim DX As Single, dy As Single
     Dim a As Single, d As Single
     
     'Go for all items
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    If (things(i).x < selrect.left) Then selrect.left = things(i).x
                    If (things(i).x > selrect.right) Then selrect.right = things(i).x
                    If (things(i).y < selrect.top) Then selrect.top = things(i).y
                    If (things(i).y > selrect.bottom) Then selrect.bottom = things(i).y
               Else
                    selrect.left = things(i).x
                    selrect.right = things(i).x
                    selrect.top = things(i).y
                    selrect.bottom = things(i).y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine rotation x and y
     tx = selrect.left + (selrect.right - selrect.left) / 2
     ty = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start rotating them
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Get the difference
               DX = tx - things(i).x
               dy = ty - things(i).y
               
               'Get angle from rotation point to thing
               a = ATan2(DX, dy) + pi
               
               'Get the distance
               d = Sqr(DX * DX + dy * dy)
               
               'Add rotation
               a = a + Amount
               
               'Move the thing
               things(i).x = tx + Cos(a) * d + 0.4
               things(i).y = ty + sIn(a) * d + 0.4
               
               'Change the thing angle
               things(i).angle = things(i).angle + Amount * PiDiv
               
               'Crop angle
               While (things(i).angle >= 360): things(i).angle = things(i).angle - 360: Wend
               While (things(i).angle < 0): things(i).angle = things(i).angle + 360: Wend
               
               'Update thing image
               UpdateThingImageColor i
          End If
     Next i
End Sub

Public Sub ScaleThings(ByVal Percent As Single)
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim i As Long
     Dim tx As Single, ty As Single
     Dim DX As Single, dy As Single
     
     'Go for all items
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Check if first item
               If FirstDone Then
                    If (things(i).x < selrect.left) Then selrect.left = things(i).x
                    If (things(i).x > selrect.right) Then selrect.right = things(i).x
                    If (things(i).y < selrect.top) Then selrect.top = things(i).y
                    If (things(i).y > selrect.bottom) Then selrect.bottom = things(i).y
               Else
                    selrect.left = things(i).x
                    selrect.right = things(i).x
                    selrect.top = things(i).y
                    selrect.bottom = things(i).y
                    FirstDone = True
               End If
          End If
     Next i
     
     'Determine rotation x and y
     tx = selrect.left + (selrect.right - selrect.left) / 2
     ty = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start rotating them
     For i = 0 To (numthings - 1)
          
          'Check if selected
          If (things(i).selected > 0) Then
               
               'Get the difference
               DX = tx - things(i).x
               dy = ty - things(i).y
               
               'Scale difference
               DX = DX * Percent * 0.01
               dy = dy * Percent * 0.01
               
               'Move the thing
               things(i).x = tx - DX
               things(i).y = ty - dy
          End If
     Next i
End Sub


Public Sub RotateVertices(ByVal Amount As Double)
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim i As Long, v As Long
     Dim tx As Single, ty As Single
     Dim DX As Single, dy As Single
     Dim a As Single, d As Single
     Dim Verts As Dictionary
     Dim VertsKeys As Variant
     
     'Find rotating vertices
     If (mode = EM_VERTICES) Then
          Set Verts = selected
     Else
          ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
          Set Verts = SelectVerticesFromSelection
     End If
     
     'Get indices
     VertsKeys = Verts.Keys
     
     'Go for all items
     For i = 0 To (Verts.Count - 1)
          
          'Get vertex
          v = CLng(VertsKeys(i))
          
          'Check if first item
          If FirstDone Then
               If (vertexes(v).x < selrect.left) Then selrect.left = vertexes(v).x
               If (vertexes(v).x > selrect.right) Then selrect.right = vertexes(v).x
               If (vertexes(v).y < selrect.top) Then selrect.top = vertexes(v).y
               If (vertexes(v).y > selrect.bottom) Then selrect.bottom = vertexes(v).y
          Else
               selrect.left = vertexes(v).x
               selrect.right = vertexes(v).x
               selrect.top = vertexes(v).y
               selrect.bottom = vertexes(v).y
               FirstDone = True
          End If
     Next i
     
     'Determine rotation x and y
     tx = selrect.left + (selrect.right - selrect.left) / 2
     ty = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start rotating them
     For i = 0 To (Verts.Count - 1)
          
          'Get vertex
          v = CLng(VertsKeys(i))
          
          'Get the difference
          DX = tx - vertexes(v).x
          dy = ty - vertexes(v).y
          
          'Get angle from rotation point to vertex
          a = ATan2(DX, dy) + pi
          
          'Get the distance
          d = Sqr(DX * DX + dy * dy)
          
          'Add rotation
          a = a + Amount
          
          'Move the vertex
          vertexes(v).x = tx + Cos(a) * d + 0.4
          vertexes(v).y = ty + sIn(a) * d + 0.4
     Next i
     
     'Remove vertex selection if not in vertices mode
     If (mode <> EM_VERTICES) Then ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
End Sub

Public Sub ScaleVertices(ByVal Percent As Single)
     Dim selrect As SRECT
     Dim FirstDone As Long
     Dim i As Long, v As Long
     Dim tx As Single, ty As Single
     Dim DX As Single, dy As Single
     Dim Verts As Dictionary
     Dim VertsKeys As Variant
     
     'Find rotating vertices
     If (mode = EM_VERTICES) Then
          Set Verts = selected
     Else
          ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
          Set Verts = SelectVerticesFromSelection
     End If
     
     'Get indices
     VertsKeys = Verts.Keys
     
     'Go for all items
     For i = 0 To (Verts.Count - 1)
          
          'Get vertex
          v = CLng(VertsKeys(i))
          
          'Check if first item
          If FirstDone Then
               If (vertexes(v).x < selrect.left) Then selrect.left = vertexes(v).x
               If (vertexes(v).x > selrect.right) Then selrect.right = vertexes(v).x
               If (vertexes(v).y < selrect.top) Then selrect.top = vertexes(v).y
               If (vertexes(v).y > selrect.bottom) Then selrect.bottom = vertexes(v).y
          Else
               selrect.left = vertexes(v).x
               selrect.right = vertexes(v).x
               selrect.top = vertexes(v).y
               selrect.bottom = vertexes(v).y
               FirstDone = True
          End If
     Next i
     
     'Determine center x and y
     tx = selrect.left + (selrect.right - selrect.left) / 2
     ty = selrect.top + (selrect.bottom - selrect.top) / 2
     
     'Start scaling them
     For i = 0 To (Verts.Count - 1)
          
          'Get vertex
          v = CLng(VertsKeys(i))
          
          'Get the difference
          DX = tx - vertexes(v).x
          dy = ty - vertexes(v).y
          
          'Scale difference
          DX = DX * Percent * 0.01
          dy = dy * Percent * 0.01
          
          'Move the vertex
          vertexes(v).x = tx - DX
          vertexes(v).y = ty - dy
     Next i
     
     'Remove vertex selection if not in vertices mode
     If (mode <> EM_VERTICES) Then ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
End Sub


Public Sub SelectCurrentLine()
     
     'Check if a line is highlighted
     If (currentselected > -1) Then
          
          'Check if the line is not yet selected
          If (linedefs(currentselected).selected = 0) Then
               
               'Make selection
               linedefs(currentselected).selected = 1
               
               'Add index to selected objects
               selected.Add CStr(currentselected), currentselected
               
               'Increase counter
               numselected = numselected + 1
          Else
               
               'Remove selection
               linedefs(currentselected).selected = 0
               
               'Remove index from selected objects
               selected.Remove CStr(currentselected)
               
               'Decrease counter
               numselected = numselected - 1
          End If
     End If
     
     'Redraw the line so it gets his
     'own color until mouse is released
     frmMain.RemoveHighlight False
     selectedtype = EM_LINES
End Sub

Public Sub SelectCurrentSector()
     Dim sd As Long, ld As Long
     
     'Check if a sector is highlighted
     If (currentselected > -1) Then
          
          'Check if the sector is not yet selected
          If (sectors(currentselected).selected = 0) Then
               
               'Make selection
               sectors(currentselected).selected = 1
               
               'Go for all sidedefs
               For sd = 0 To (numsidedefs - 1)
                    
                    'Check if this sidedef belongs to this sector
                    If (sidedefs(sd).sector = currentselected) Then
                         
                         'Get the linedef
                         ld = sidedefs(sd).linedef
                         
                         'Increase selection reference count
                         linedefs(ld).selected = linedefs(ld).selected + 1
                         
                         'Render the linedef to selected (also vertices, those have been overdrawn)
                         Render_AllLinedefs vertexes(0), linedefs(0), ld, ld, submode, indicatorsize
                         'Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
                         'Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
                    End If
               Next sd
               
               'Add index to selected objects
               selected.Add CStr(currentselected), currentselected
               
               'Increase counter
               numselected = numselected + 1
          Else
               
               'Remove selection
               sectors(currentselected).selected = 0
               
               'Go for all sidedefs
               For sd = 0 To (numsidedefs - 1)
                    
                    'Check if this sidedef belongs to this sector
                    If (sidedefs(sd).sector = currentselected) Then
                         
                         'Get the linedef
                         ld = sidedefs(sd).linedef
                         
                         'Decrease selection reference count
                         linedefs(ld).selected = linedefs(ld).selected - 1
                         
                         'Check if we can reset this line's color
                         If (linedefs(ld).selected = 0) Then
                              
                              'Render the linedef back to normal (also vertices, those have been overdrawn)
                              Render_AllLinedefs vertexes(0), linedefs(0), ld, ld, submode, indicatorsize
                              'Render_AllVertices vertexes(0), linedefs(ld).v1, linedefs(ld).v1, vertexsize
                              'Render_AllVertices vertexes(0), linedefs(ld).v2, linedefs(ld).v2, vertexsize
                         End If
                    End If
               Next sd
               
               'Remove index from selected objects
               selected.Remove CStr(currentselected)
               
               'Decrease counter
               numselected = numselected - 1
          End If
     End If
     
     'Redraw the sector so it gets his
     'own color until mouse is released
     frmMain.RemoveHighlight False
     selectedtype = EM_SECTORS
End Sub

Public Sub SelectCurrentThing()
     
     'Check if a thing is highlighted
     If (currentselected > -1) Then
          
          'Check if the thing is not yet selected
          If (things(currentselected).selected = 0) Then
               
               'Make selection
               things(currentselected).selected = 1
               
               'Add index to selected objects
               selected.Add CStr(currentselected), currentselected
               
               'Increase counter
               numselected = numselected + 1
          Else
               
               'Remove selection
               things(currentselected).selected = 0
               
               'Remove index from selected objects
               selected.Remove CStr(currentselected)
               
               'Decrease counter
               numselected = numselected - 1
          End If
     End If
     
     'Redraw the vertex so it gets his
     'own color until mouse is released
     frmMain.RemoveHighlight False
     selectedtype = EM_THINGS
End Sub

Public Sub SelectCurrentVertex()
     
     'Check if a vertex is highlighted
     If (currentselected > -1) Then
          
          'Check if the vertex is not yet selected
          If (vertexes(currentselected).selected = 0) Then
               
               'Make selection
               vertexes(currentselected).selected = 1
               
               'Add index to selected objects
               selected.Add CStr(currentselected), currentselected
               
               'Increase counter
               numselected = numselected + 1
          Else
               
               'Remove selection
               vertexes(currentselected).selected = 0
               
               'Remove index from selected objects
               selected.Remove CStr(currentselected)
               
               'Decrease counter
               numselected = numselected - 1
          End If
     End If
     
     'Redraw the vertex so it gets his
     'own color until mouse is released
     frmMain.RemoveHighlight False
     selectedtype = EM_VERTICES
End Sub

Public Function SelectedVerticesSplitLinedefs() As Boolean
     Dim v As Long
     Dim ld As Long, nld As Long
     Dim distance As Long
     Dim SplitDistance As Long
     
     'Get split distance
     SplitDistance = Config("linesplitdistance")
     
     'Go for all vertices
     For v = 0 To (numvertexes - 1)
          
          'Check if selected
          If (vertexes(v).selected <> 0) Then
               
               'Get the nearest linedef
               ld = NearestUnselectedUnreferencedLinedef(v, vertexes(0), linedefs(0), numlinedefs, distance, SplitDistance)
               'ld = NearestUnselectedLinedef(vertexes(v).X, -vertexes(v).Y, vertexes(0), linedefs(0), numlinedefs, distance)
               
               'Linedef found?
               If (ld > -1) Then
                    
                    'Check if distance is close enough for linedef split
                    If (distance <= SplitDistance) Then
                         
                         'Split the linedef
                         nld = SplitLinedef(ld, v)
                         
                         'Indicate we made changes
                         SelectedVerticesSplitLinedefs = True
                    End If
               End If
               
          Else
               
               'Get the nearest selected linedef
               ld = NearestSelectedLinedef(vertexes(v).x, -vertexes(v).y, vertexes(0), linedefs(0), changedlines(0), numchangedlines, distance, SplitDistance)
               
               'Linedef found?
               If (ld > -1) Then
                    
                    'Check if distance is close enough for linedef split
                    If (distance <= SplitDistance) Then
                         
                         'Split the linedef
                         nld = SplitLinedef(ld, v)
                         
                         'Find the linedef in the changing linedefs
                         If (FindValueInArray(changedlines(), ld) > -1) Then
                              
                              'Add to changed linedefs
                              ReDim Preserve changedlines(0 To numchangedlines)
                              changedlines(numchangedlines) = nld
                              numchangedlines = numchangedlines + 1
                         End If
                         
                         'Select the vertex now, because its part of changing selection now
                         vertexes(v).selected = 1
                         
                         'Indicate we made changes
                         SelectedVerticesSplitLinedefs = True
                    End If
               End If
          End If
     Next v
     
     'Deselect all linedefs
'     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, sectors(0), 0
End Function

Public Sub SelectLinedefsFromSectors()
     Dim ld As Long
     
     'Linedefs are already set to selected, but with a reference number
     'go for all linedefs to set the selected to 1 and add it to the list
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if selected
          If (linedefs(ld).selected > 0) Then
               
               'Normal selection, not a reference count
               linedefs(ld).selected = 1
               
               'Add to list
               selected.Add CStr(ld), ld
          End If
     Next ld
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_LINES
End Sub

Public Sub SelectLinedefsFromVertices()
     Dim ld As Long
     
     'Select all linedefs that have both vertices selected
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if both vertices selected
          If ((vertexes(linedefs(ld).v1).selected > 0) And _
              (vertexes(linedefs(ld).v2).selected > 0)) Then
               
               'Select the line
               linedefs(ld).selected = 1
               
               'Add to list
               selected.Add CStr(ld), ld
          End If
     Next ld
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_LINES
     
     'Deselect all vertices
     ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
End Sub

Public Sub SelectSector(ByVal s As Long)
     Dim sd As Long
     
     'Go for all sidedefs
     For sd = 0 To (numsidedefs - 1)
          
          'Check if on the given sector
          If (sidedefs(sd).sector = s) Then
               
               'Normal selection, not a reference count
               linedefs(sidedefs(sd).linedef).selected = 1
          End If
     Next sd
End Sub

Public Sub SelectSectorRef(ByVal s As Long)
     Dim sd As Long
     
     'Go for all sidedefs
     For sd = 0 To (numsidedefs - 1)
          
          'Check if on the given sector
          If (sidedefs(sd).sector = s) Then
               
               'Normal selection, not a reference count
               linedefs(sidedefs(sd).linedef).selected = linedefs(sidedefs(sd).linedef).selected + 1
          End If
     Next sd
End Sub


Public Sub SelectSectorsFromLinedefs(Optional ListOnly As Boolean = False)
     Dim sd As Long, ld As Long, s As Long
     Dim sectorlist As New Dictionary        '1 = do select, 2 = dont select
     
     'Reselect Linedefs by the list to give them a reference count selection
     'and set all sector selected variables
     
     'Go for all sidedefs
     For sd = 0 To (numsidedefs - 1)
          
          'Get the linedef
          ld = sidedefs(sd).linedef
          
          'Get the sector
          s = sidedefs(sd).sector
          
          'Check if the linedef is selected
          If (linedefs(ld).selected) Then
               
               'Add sector to list if not already added
               If (sectorlist.Exists(CStr(s)) = False) Then sectorlist.Add CStr(s), 1
          Else
               
               'Check if in the list already
               If (sectorlist.Exists(CStr(s)) = True) Then
                    
                    'Set to dont select
                    sectorlist(CStr(s)) = 2
               Else
                    
                    'Add as dont select
                    sectorlist.Add CStr(s), 2
               End If
          End If
     Next sd
     
     'Deselect all linedefs
     If (ListOnly = False) Then ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Make new selection list
     Set selected = New Dictionary
     
     'Go for all sectors
     For s = 0 To (numsectors - 1)
          
          'Check if sector is listed
          If (sectorlist.Exists(CStr(s)) = True) Then
               
               'Check if not set to dont select
               If (sectorlist(CStr(s)) <> 2) Then
                    
                    'Select the sector
                    If (ListOnly = False) Then sectors(s).selected = 1
                    
                    'Add to list
                    selected.Add CStr(s), s
               Else
                    
                    'Deselect the sector
                    If (ListOnly = False) Then sectors(s).selected = 0
               End If
          Else
               
               'Deselect the sector
               If (ListOnly = False) Then sectors(s).selected = 0
          End If
     Next s
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_SECTORS
     
     'Check if allowed to select indices
     If (ListOnly = False) Then
          
          'Go for all sidedefs
          For sd = 0 To (numsidedefs - 1)
               
               'Get the linedef
               ld = sidedefs(sd).linedef
               
               'Get the sector
               s = sidedefs(sd).sector
               
               'Check if sector is assigned
               If (s > -1) Then
                    
                    'Check if the sector is selected
                    If (sectors(s).selected) Then
                         
                         'Increase linedef selection reference count
                         linedefs(ld).selected = linedefs(ld).selected + 1
                    End If
               End If
          Next sd
     End If
End Sub

Public Function SelectThingsFromRect(ByRef r As RECT, ByVal Shift As Long) As Long
     Dim t As Long
     Dim c As Long
     
     'Check if we should clear selection first
     If (Config("additiveselect") = vbUnchecked) And (Shift = False) Then
          
          'Deselect all things
          ResetSelections things(0), numthings, linedefs(0), 0, vertexes(0), 0, VarPtr(sectors(0)), 0
          
          'Clear list
          Set selected = New Dictionary
     End If
     
     'Go for all things
     For t = 0 To (numthings - 1)
          With things(t)
               
               'Check if the thing is within rect
               If ((.x >= r.left) And (.x <= r.right) And (.y >= r.top) And (.y <= r.bottom)) Then
                    
                    'Check if this thing is shown through the filter
                    If (ThingFiltered(t)) Then
                         
                         'Count thing
                         c = c + 1
                         
                         'Check if vertex is selected already
                         If (.selected <> 0) Then
                              
                              'Check if we should deselect
                              If Shift Then
                                   
                                   'Deselect thing
                                   .selected = 0
                                   
                                   'Remove from list if exists
                                   If selected.Exists(CStr(t)) Then selected.Remove CStr(t)
                              End If
                         Else
                              
                              'Select thing
                              .selected = 1
                              
                              'Add to list if not already added
                              If (selected.Exists(CStr(t)) = False) Then selected.Add CStr(t), t
                         End If
                    End If
               End If
          End With
     Next t
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_THINGS
     
     'Return number of things found in rect
     SelectThingsFromRect = c
End Function

Public Sub SelectThingsFromSectors()
     Dim th As Long
     Dim s As Long
     Dim newlist As New Dictionary
     
     'Things will be selected when inside a selected sector
     
     'Go for all things
     For th = 0 To (numthings - 1)
          
          'Get sector in which the thing is
          s = IntersectSector(things(th).x, -things(th).y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
          
          'Check if not outside
          If (s > -1) Then
               
               'Check if sector is selected
               If (sectors(s).selected) Then
                    
                    'Select the thing
                    things(th).selected = 1
                    
                    'Add to list
                    newlist.Add CStr(th), th
               End If
          End If
     Next th
     
     'Deselect all linedefs and sectors
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), numsectors
     
     'Set new list
     Set selected = newlist
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_THINGS
End Sub

Public Sub SelectVerticesFromLinedefs()
     Dim ld As Long
     
     'Select all vertices that are referred to by selected linedefs
     'This also work for sector to vertices selection
     
     'Clear list
     Set selected = New Dictionary
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if selected
          If (linedefs(ld).selected > 0) Then
               
               'Check if first vertex is not already selected
               If (vertexes(linedefs(ld).v1).selected = 0) Then
                    
                    'Select vertex
                    vertexes(linedefs(ld).v1).selected = 1
                    
                    'Add to list
                    selected.Add CStr(linedefs(ld).v1), linedefs(ld).v1
               End If
               
               'Check if first vertex is not already selected
               If (vertexes(linedefs(ld).v2).selected = 0) Then
                    
                    'Select vertex
                    vertexes(linedefs(ld).v2).selected = 1
                    
                    'Add to list
                    selected.Add CStr(linedefs(ld).v2), linedefs(ld).v2
               End If
          End If
     Next ld
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_VERTICES
     
     'Deselect all linedefs and sectors
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), numsectors
End Sub

Public Function SelectVerticesFromRect(ByRef r As RECT, ByVal Shift As Long) As Long
     Dim v As Long
     Dim c As Long
     
     'Check if we should clear selection first
     If (Config("additiveselect") = vbUnchecked) And (Shift = False) Then
          
          'Deselect all vertices
          ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
          
          'Clear list
          Set selected = New Dictionary
     End If
     
     'Go for all vertices
     For v = 0 To (numvertexes - 1)
          With vertexes(v)
               
               'Check if the vertex is within rect
               If ((.x >= r.left) And (.x <= r.right) And (.y >= r.top) And (.y <= r.bottom)) Then
                    
                    'Count vertex
                    c = c + 1
                    
                    'Check if vertex is selected already
                    If (.selected <> 0) Then
                         
                         'Check if we should deselect
                         If Shift Then
                              
                              'Deselect vertex
                              .selected = 0
                              
                              'Remove from list if exists
                              If selected.Exists(CStr(v)) Then selected.Remove CStr(v)
                         End If
                    Else
                         
                         'Select vertex
                         .selected = 1
                         
                         'Add to list if not already added
                         If (selected.Exists(CStr(v)) = False) Then selected.Add CStr(v), v
                    End If
               End If
          End With
     Next v
     
     'Count selected items
     numselected = selected.Count
     selectedtype = EM_VERTICES
     
     'Return number of vertices found within rec
     SelectVerticesFromRect = c
End Function

Public Function SelectVerticesFromSelection() As Dictionary
     Dim NewSelection As Dictionary
     Dim ld As Long
     
     'Create new dictionary
     Set NewSelection = New Dictionary
     
     'Vertices selection can only be made from
     'lines, which are selected in both Lines and Sectors mode
     If ((mode = EM_LINES) Or (mode = EM_SECTORS)) Then
          
          'Got for all lines
          For ld = 0 To (numlinedefs - 1)
               
               'Check if this line is selected
               If (linedefs(ld).selected) Then
                    
                    'First vertex, check if already selected
                    If (vertexes(linedefs(ld).v1).selected = 0) Then
                         
                         'Add to new selection
                         NewSelection.Add CStr(linedefs(ld).v1), linedefs(ld).v1
                         
                         'Select vertex
                         vertexes(linedefs(ld).v1).selected = 1
                    End If
                    
                    'Second vertex, check if already selected
                    If (vertexes(linedefs(ld).v2).selected = 0) Then
                         
                         'Add to new selection
                         NewSelection.Add CStr(linedefs(ld).v2), linedefs(ld).v2
                         
                         'Select vertex
                         vertexes(linedefs(ld).v2).selected = 1
                    End If
               End If
          Next ld
          
          'Return the selected vertices
          Set SelectVerticesFromSelection = NewSelection
     End If
End Function

Public Sub SolveGlitches()
     Dim ld As Long
     Dim s1 As Long, s2 As Long
     
     'Some routines may leave glitches, like linedefs with
     'only a second sidedef or sidedefs with NULL textures
     'This routines fixes them.
     
     'Go for all linedefs to check its sidedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Get the sidedefs
          s1 = linedefs(ld).s1
          s2 = linedefs(ld).s2
          
          'Check if only a second sidedef
          If (s1 = -1) And (s2 > -1) Then
               
               'Flip this linedef
               FlipLinedefVertices ld
               FlipLinedefSidedefs ld
          End If
          
          'Check if s1 is valid
          If (s1 > -1) Then
               
               'Check for zero textures and fix them
               If (LenB(sidedefs(s1).upper) = 0) Then sidedefs(s1).upper = "-"
               If (LenB(sidedefs(s1).middle) = 0) Then sidedefs(s1).middle = "-"
               If (LenB(sidedefs(s1).lower) = 0) Then sidedefs(s1).lower = "-"
          End If
          
          'Check if s2 is valid
          If (s2 > -1) Then
               
               'Check for zero textures and fix them
               If (LenB(sidedefs(s2).upper) = 0) Then sidedefs(s2).upper = "-"
               If (LenB(sidedefs(s2).middle) = 0) Then sidedefs(s2).middle = "-"
               If (LenB(sidedefs(s2).lower) = 0) Then sidedefs(s2).lower = "-"
          End If
     Next ld
End Sub

Public Function SplitLinedef(ByVal SourceLinedef As Long, ByVal TargetVertex As Long) As Long
     Dim nl As Long
     Dim ns1 As Long
     Dim ns2 As Long
     
     'SplitLinedef:
     '> Copy SourceLinedef to a new linedef,
     '  Vertex1 set to TargetVertex
     '> Copy the Sidedefs of SourceLinedef to new sidedefs and
     '  let the new linedef refer to them.
     '> Modify SourceLinedef, set Vertex2 to TargetVertex
     '
     'Returns the index of the new linedef
     
     'Create new objects
     nl = CreateLinedef
     If (linedefs(SourceLinedef).s1 > -1) Then ns1 = CreateSidedef Else ns1 = -1
     If (linedefs(SourceLinedef).s2 > -1) Then ns2 = CreateSidedef Else ns2 = -1
     
     'Copy the properties
     linedefs(nl) = linedefs(SourceLinedef)
     If (ns1 > -1) Then sidedefs(ns1) = sidedefs(linedefs(SourceLinedef).s1)
     If (ns2 > -1) Then sidedefs(ns2) = sidedefs(linedefs(SourceLinedef).s2)
     
     'Set sidedef references
     linedefs(nl).s1 = ns1
     linedefs(nl).s2 = ns2
     If (ns1 > -1) Then sidedefs(ns1).linedef = nl
     If (ns2 > -1) Then sidedefs(ns2).linedef = nl
     
     'Set vertex references
     linedefs(nl).v1 = TargetVertex
     linedefs(SourceLinedef).v2 = TargetVertex
     
     'Return new linedef
     SplitLinedef = nl
End Function

Public Sub StitchVertices(ByVal TargetVertex As Long, ByVal StitchVertex As Long)
     Dim ld As Long
     Dim v1ref As Long, v2ref As Long
     
     'Stitch Vertices:
     '> Re-refer all linedef that refer to StitchVertex to TargetVertex.
     '> In case one of the linedefs refers to both TargetVertex and StitchVertex,
     '  remove the linedef and its sidedefs, clean up unused sectors.
     '> Remove StitchVertex.
     
     'Go for all linedefs (using Do Loop because linedefs may change)
     Do While (ld < numlinedefs)
          
          'Get the vertex references
          v1ref = linedefs(ld).v1
          v2ref = linedefs(ld).v2
          
          'Check if both referring to both of the given vertices
          If (((v1ref = TargetVertex) And (v2ref = StitchVertex)) Or _
             ((v2ref = TargetVertex) And (v1ref = StitchVertex))) Then
               
               'Remove this linedef now
               RemoveLinedef ld, True, False, True
               
          'Check if v1 refers to stitching vertex
          ElseIf (v1ref = StitchVertex) Then
               
               'Change to TargetVertex
               linedefs(ld).v1 = TargetVertex
               
               'Next linedef
               ld = ld + 1
               
          'Check if v2 refers to stitching vertex
          ElseIf (v2ref = StitchVertex) Then
               
               'Change to TargetVertex
               linedefs(ld).v2 = TargetVertex
               
               'Next linedef
               ld = ld + 1
          Else
               
               'Next linedef
               ld = ld + 1
          End If
     Loop
     
     'Modify selection of target vertex
     vertexes(TargetVertex).selected = vertexes(TargetVertex).selected Or vertexes(StitchVertex).selected
     
     'Remove the StitchVertex now
     RemoveVertex StitchVertex
End Sub

Public Sub StitchSelectedVertices()
     Dim ld As Long, v As Long
     Dim v1ref As Long, v2ref As Long
     Dim v1ex As Long, v2ex As Long
     Dim Indices As Variant
     Dim TargetVertex As Long
     
     'Stitch Vertices:
     '> Re-refer all linedef that refer to StitchVertex to TargetVertex.
     '> In case one of the linedefs refers to both TargetVertex and StitchVertex,
     '  remove the linedef and its sidedefs, clean up unused sectors.
     '> Remove StitchVertex.
     
     'First selected will be the target
     Indices = selected.Items
     TargetVertex = Indices(LBound(Indices))
     
     'Go for all linedefs (using Do Loop because linedefs may change)
     Do While (ld < numlinedefs)
          
          'Get the vertex references
          v1ref = linedefs(ld).v1
          v2ref = linedefs(ld).v2
          v1ex = selected.Exists(CStr(v1ref))
          v2ex = selected.Exists(CStr(v2ref))
          
          'Check if both referring to both of the given vertices
          If v1ex And v2ex Then
               
               'Remove this linedef now
               RemoveLinedef ld, True, False, True
               
          'Check if v1 refers to stitching vertex
          ElseIf v1ex Then
               
               'Change to TargetVertex
               linedefs(ld).v1 = TargetVertex
               
               'Next linedef
               ld = ld + 1
               
          'Check if v2 refers to stitching vertex
          ElseIf v2ex Then
               
               'Change to TargetVertex
               linedefs(ld).v2 = TargetVertex
               
               'Next linedef
               ld = ld + 1
          Else
               
               'Next linedef
               ld = ld + 1
          End If
     Loop
     
     'Deselect the target vertex
     vertexes(TargetVertex).selected = 0
     
     'Go for all vertices (using Do Loop because vertices may change)
     Do While (v < numvertexes)
          
          'Is this vertex selected?
          If (vertexes(v).selected) Then
               
               'Remove it
               RemoveVertex v
          Else
               
               'Next
               v = v + 1
          End If
     Loop
     
     'Reselect the target vertex
     vertexes(TargetVertex).selected = 1
     Set selected = New Dictionary
     selected.Add CStr(TargetVertex), TargetVertex
     numselected = 1
End Sub


Public Sub TraceSectorSplitVertex(ByVal start As Long, ByVal target As Long, ByVal sector As Long, ByRef lines() As Long, ByVal numlines As Long)
     Dim ld As Long
     Dim nextlines() As Long
     Dim onsector As Boolean
     
     'Allocate memory for new lines list
     ReDim nextlines(0 To numlines)
     
     'Copy previous lines
     If (numlines > 0) Then CopyMemory nextlines(0), lines(0), numlines * 4
     
     'One line will be added before next call
     numlines = numlines + 1
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Leave when terminated
          If TerminateRecursion Then Exit Sub
          
          'Check if this linedef is not selected
          If (linedefs(ld).selected = 0) Then
               
               'Check if connected to this vertex
               If (linedefs(ld).v1 = start) Or (linedefs(ld).v2 = start) Then
                    
                    'Check if either front or back side are on this sector
                    onsector = False
                    If (linedefs(ld).s1 > -1) Then onsector = (sidedefs(linedefs(ld).s1).sector = sector)
                    If (linedefs(ld).s2 > -1) Then onsector = onsector Or (sidedefs(linedefs(ld).s2).sector = sector)
                    
                    'Check if on this sector
                    If onsector Then
                         
                         'Select the line so it wont be used twice
                         linedefs(ld).selected = 1
                         
                         'Add line on nextlines
                         nextlines(numlines - 1) = ld
                         
                         'Check what vertex is the next vertex in trace
                         If (linedefs(ld).v1 = start) Then
                              
                              'Check if v2 is the target
                              If (linedefs(ld).v2 = target) Then
                                   
                                   'Save the lines list
                                   SectorSplitLinesList() = nextlines()
                                   SectorSplitNumLines = numlines
                                   
                                   'Terminate recursion
                                   TerminateRecursion = True
                              Else
                                   
                                   'Trace from v2
                                   TraceSectorSplitVertex linedefs(ld).v2, target, sector, nextlines(), numlines
                              End If
                         Else
                              
                              'Check if v1 is the target
                              If (linedefs(ld).v1 = target) Then
                                   
                                   'Save the lines list
                                   SectorSplitLinesList() = nextlines()
                                   SectorSplitNumLines = numlines
                                   
                                   'Terminate recursion
                                   TerminateRecursion = True
                              Else
                                        
                                   'Trace from v1
                                   TraceSectorSplitVertex linedefs(ld).v1, target, sector, nextlines(), numlines
                              End If
                         End If
                    End If
               End If
          End If
     Next ld
End Sub
