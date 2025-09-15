Attribute VB_Name = "mod3DMode"
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


'=============== BSP TREE WALKING AND RENDERING =====================
' This code can be found in the C++ source files
' bsp.cpp, clip.cpp and referenced files (see Builder VS.NET workspace)
'
' ProcessBSP:
'
' - Run through the tree (front to back)
'    - When arriving at a ssector:
'         - Calculate the range on the clipbuffer
'         - Check if any of it is visible
'         - Store ssector it the renderarray
'         - Store the ssector splits if not already stored
'         - Make the polygon if not done yet
'         - Apply clipping to the clipbuffer
'
'    - When clipbuffer is enitely set:
'         - Leave the BSP walk
'
' PickObject returns the aimed target object using
' ray intersection testing. This code can be found in the
' C++ source file pick.cpp (see Builder VS.NET workspace)
'====================================================================


'Aspect for 320x200 fixed
Public Const VIDEO_FIXED_ASPECT As Single = 0.625

'Camera and Physics
Private Const EYESHEIGHT As Single = 40                'Eyes height distance from floor
Private Const HEADHEIGHT As Single = 10                'Eyes height distance from ceiling
Private Const GRAVITYWEIGHT As Single = 2              'Gravity multiplier

'Map rendering scale
Public Const MAP_RENDER_SCALE As Single = 0.01
Public Const INV_MAP_RENDER_SCALE As Single = 100      '1 / MAP_RENDER_SCALE

'Max to render at once
Public Const MAX_SSECTORS_VERTICES As Long = 200
Public Const MAX_VISIBLE_SSECTORS As Long = 2000
Public Const MAX_VISIBLE_SIDEDEFS As Long = 4000
Public Const MAX_VISIBLE_THINGS As Long = 500

'Mouse button keys
Public Const MOUSE_BUTTON_0 As Long = 4000
Public Const MOUSE_BUTTON_1 As Long = 4001
Public Const MOUSE_BUTTON_2 As Long = 4002
Public Const MOUSE_BUTTON_3 As Long = 4003
Public Const MOUSE_BUTTON_4 As Long = 4004
Public Const MOUSE_BUTTON_5 As Long = 4005
Public Const MOUSE_BUTTON_6 As Long = 4006
Public Const MOUSE_BUTTON_7 As Long = 4007
Public Const MOUSE_SCROLL_UP As Long = 4008
Public Const MOUSE_SCROLL_DOWN As Long = 4009

'Textures/Things selector
Private Const TEXTURE_COLS As Long = 5
Private Const TEXTURE_ROWS As Long = 5
Private Const TEXTURE_SPACING As Single = 0.1
Private Const TEXTURE_TEXTHEIGHT As Single = 0.1
Private Const TEXTURE_DESC As String = "Select or enter a texture:"
Private Const THING_DESC As String = "Select a thing:"
Private Const TEXTURE_CHARS As String = "abcdefghijklmnopqrstuvwxyz_-0123456789=+;:,.<>[]{}!@#$%^&*()'""/\?|"
Private Const CURSOR_FLASH_INTERVAL As Long = 150

'Text (for colors use D3DColorMake())
Private Const TEXT_MAXCHARS As Long = 128         'Maximum characters in 1 text buffer
Private Const TEXT_SIZE As Single = 2
Private Const TEXT_C1 As Long = -205
Private Const TEXT_C2 As Long = -19968
Private Const TEXT_C3 As Long = -46080
Private Const TEXT_C4 As Long = -65536
Private Const TEXT_SHOWTIME As Long = 4000

'Info panel
Private Const INFO_C1 As Long = -1
Private Const INFO_C2 As Long = -10027162
Private Const INFO_C3 As Long = -10027162
Private Const INFO_C4 As Long = -11776948
Private Const INFO_UPDATEDELAY As Long = 100
Private Const INFO_COORDS_UPDATEDELAY As Long = 100

'Misc
Private Const STATUP_TITLE As String = "Doom Builder  3D Editing Mode"
Private Const STATUP_SUBTITLE As String = "Have a peachy day and be well!"

'Texture filter modes
Public Enum ENUM_TEXTUREFILTERS
     TF_NONE
     TF_LINEAR_MIPMAP_NEAREST
     TF_LINEAR_MIPMAP_LINEAR
End Enum

'Enum for horizontal alignment
Public Enum ENUM_HALIGN
     ALIGN_LEFT
     ALIGN_RIGHT
     ALIGN_CENTER
End Enum

'Enum for vertical alignment
Public Enum ENUM_VALIGN
     ALIGN_TOP
     ALIGN_BOTTOM
     ALIGN_MIDDLE
End Enum

'Aimed object types
Private Enum ENUM_OBJECTTYPES
     OBJ_NOTHING
     OBJ_SECTORFLOOR
     OBJ_SECTORCEILING
     OBJ_SIDEDEFUPPER
     OBJ_SIDEDEFLOWER
     OBJ_SIDEDEFMIDDLE
     OBJ_THING
End Enum

'Map SEGS type
Private Type MAPSEG
     v1 As Long
     v2 As Long
     angle As Long
     linedef As Long
     side As Long
     offset As Long
End Type

'Float Vertices
Private Type FPOINT
     x As Single
     y As Single
End Type

'Map SSECTOR type
Private Type MAPSSECTOR
     startseg As Long
     numsegs As Long
     
     'Optimization variables
     sector As Long
     numvertices As Long
     vertices(0 To (MAX_SSECTORS_VERTICES - 1)) As FPOINT
End Type

'Map NODES type
Private Type MAPNODE
     x As Long
     y As Long
     DX As Long
     dy As Long
     
     rtop As Long
     rbottom As Long
     rleft As Long
     rright As Long
     
     ltop As Long
     lbottom As Long
     lleft As Long
     lright As Long
     
     right As Long
     left As Long
End Type

'Structure for character information
Public Type CHARRECTYPE
     char As Byte
     u1 As Single
     v1 As Single
     u2 As Single
     v2 As Single
     width As Byte
     height As Byte
End Type

'Structure for Rectangle
Public Type SRECT
     left As Single
     right As Single
     top As Single
     bottom As Single
End Type



'Declarations
Private Declare Sub SetStructurePointers Lib "builder.dll" (ByRef vertices As POINT, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByRef segs As MAPSEG, ByVal ptr_sectors As Long, ByRef ssectors As MAPSSECTOR, ByRef things As MAPTHING, ByRef nodes As MAPNODE, ByVal numnodes As Long, ByVal numsectors As Long, ByVal numssectors As Long, ByVal numthings As Long)
Private Declare Sub DestroyStructurePointers Lib "builder.dll" ()
Private Declare Sub CreateSSectorReferences Lib "builder.dll" ()
Private Declare Sub GetMissingEntries Lib "builder.dll" (ByRef array1 As Long, ByVal count1 As Long, ByRef array2 As Long, ByVal count2 As Long, ByRef resultarray As Long, ByRef resultcount As Long)
Private Declare Sub PrepareAllSSectors Lib "builder.dll" ()
Private Declare Sub ProcessBSP Lib "builder.dll" (ByRef renderarray As Long, ByVal maxssectors As Long, ByRef sidedefsarray As Long, ByVal maxsidedefs As Long, ByRef r_numssectors As Long, ByRef r_numsidedefs As Long, ByVal x As Long, ByVal y As Long, ByVal Z As Long, ByVal angle As Single, ByVal FOV As Long, ByVal renderdistance As Long, ByRef thingsarray As Long, ByRef r_numthings As Long, ByVal maxthings As Long)
Private Declare Function PickObject Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByVal ptr_sectors As Long, ByRef ssectors As MAPSSECTOR, ByRef things As MAPTHING, ByRef r_sidedefs As Long, ByVal r_numsidedefs As Long, ByVal numlinedefs As Long, ByRef r_subsectors As Long, ByVal r_numsubsectors As Long, ByRef r_things As Long, ByVal r_numthings As Long, ByRef r1 As D3DVECTOR, ByRef r2 As D3DVECTOR, ByRef hit_point As D3DVECTOR, ByRef hit_index As Long) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Sub SetFontChar Lib "builder.dll" (ByVal char As String, ByVal width As Single, ByVal height As Single, ByVal u1 As Single, ByVal u2 As Single, ByVal v1 As Single, ByVal v2 As Single)
Private Declare Sub CreateText Lib "builder.dll" (ByVal Text As String, ByRef Position As SRECT, ByVal hAlign As Long, ByVal vAlign As Long, ByVal c_lt As Long, ByVal c_rt As Long, ByVal c_lb As Long, ByVal c_rb As Long, ByVal CharScale As Single, ByRef TextVertex As TLVERTEX, ByVal ScreenWidth As Long, ByVal ScreenHeight As Long)
Private Declare Sub SetAllThingSectors Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal ptr_sidedefs As Long)


'Map structure
Private m_vertices() As POINT
Private m_segs() As MAPSEG
Private m_subsectors() As MAPSSECTOR
Private m_nodes() As MAPNODE
Private numsegs As Long
Private numsubsectors As Long
Private numnodes As Long

'Vertex Buffers
Private SubSectorFloors() As Direct3DVertexBuffer9
Private SubSectorCeilings() As Direct3DVertexBuffer9
Private SidedefUpper() As Direct3DVertexBuffer9
Private SidedefMiddle() As Direct3DVertexBuffer9
Private SidedefLower() As Direct3DVertexBuffer9
Private d_SubSectorFloors() As Long
Private d_SubSectorCeilings() As Long
Private d_SidedefUpper() As Long
Private d_SidedefMiddle() As Long
Private d_SidedefLower() As Long
Private i_SectorFloors() As Direct3DTexture9
Private i_SectorCeilings() As Direct3DTexture9
Private i_SidedefUpper() As Direct3DTexture9
Private i_SidedefMiddle() As Direct3DTexture9
Private i_SidedefLower() As Direct3DTexture9

'Rendering
Private r_subsectors(MAX_VISIBLE_SSECTORS) As Long
Private r_sidedefs(MAX_VISIBLE_SIDEDEFS) As Long
Private r_things(MAX_VISIBLE_THINGS) As Long
Private r_thingwindow(MAX_VISIBLE_THINGS) As Long
Private r_prevsubsectors(MAX_VISIBLE_SSECTORS) As Long
Private r_prevsidedefs(MAX_VISIBLE_SIDEDEFS) As Long
Private r_numsubsectors As Long
Private r_numsidedefs As Long
Private r_numthings As Long
Private r_numprevsubsectors As Long
Private r_numprevsidedefs As Long
Private r_discards(MAX_VISIBLE_SIDEDEFS) As Long
Private r_numdiscards As Long
Private r_maintext As Direct3DVertexBuffer9
Private r_nummaintextfaces As Long
Private r_subtext As Direct3DVertexBuffer9
Private r_numsubtextfaces As Long
Private r_crosshair As Direct3DVertexBuffer9
Private r_texpoly(0 To (TEXTURE_COLS * TEXTURE_ROWS - 1)) As Direct3DVertexBuffer9
Private r_texclass(0 To (TEXTURE_COLS * TEXTURE_ROWS - 1)) As clsImage
Private r_texdesc As Direct3DVertexBuffer9
Private r_numtexdescfaces As Long
Private r_texname As Direct3DVertexBuffer9
Private r_numtexnamefaces As Long
Private r_infopanel As Direct3DVertexBuffer9
Private r_infotexts(0 To 9) As Direct3DVertexBuffer9
Private r_numinfotextfaces(0 To 9) As Long
Private r_thingboxvb As Direct3DVertexBuffer9
Private r_thingboxlines As Direct3DVertexBuffer9
Private r_thingarrow As Direct3DVertexBuffer9
Private r_thingsprite As Direct3DVertexBuffer9

'Configuration settings
Private c_videowidth As Long
Private c_videoheight As Long
Private c_movespeed As Single
Private c_mixresource As Long
Private c_belowceiling As Long
Private c_videoviewdistance As Long
Private c_videofov As Long
Private c_invertmousey As Long
Private c_mousespeed As Single

'Lighting Tables
Private t_brightness(0 To 255) As Long
Private t_fogness(0 To 255) As Long

'Textures
Private tex_font As Direct3DTexture9
Private tex_crosshair As Direct3DTexture9
Private tex_unknown As Direct3DTexture9
Private tex_missing As Direct3DTexture9
Private tex_thingbox As Direct3DTexture9
Private tex_thingarrow As Direct3DTexture9

'Matrices
Private matrixView As D3DMATRIX
Private matrixProject As D3DMATRIX
Private matrixWorld As D3DMATRIX

'Modes and info
Private CrosshairInfo As D3DXIMAGE_INFO
Public TextureSelecting As Long
Public ThingSelecting As Long
Public TextureRowOffset As Long
Public TextureSelectedIndex As Long
Public TextureSelectCancelled As Long
Public TextureUseFlats As Long
Public TextureEraseOnType As Boolean
Public TextureCount As Long
Public SelectedName As String
Public IgnoreInput As Long              'Set to true to ignore input for 1 frame
Public CopiedTexture As String
Public CopiedFlat As String
Public ApplyGravity As Boolean
Public FullBrightness As Boolean
Public ShowTextCursor As Boolean
Public ShowInfo As Boolean
Public CopiedX As Long
Public CopiedY As Long
Public LastInfoObject As Long
Public LastInfoObjectType As Long
Public DelayVideoFrames As Long
Public ShowThings As Long
Public ShowAllTextures As Boolean
Public HasProcessed As Boolean

'Keys
Public Key3DForward As Boolean
Public Key3DBackward As Boolean
Public Key3DStrafeLeft As Boolean
Public Key3DStrafeRight As Boolean
Public Key3DStrafeUp As Boolean
Public Key3DStrafeDown As Boolean

'Timing
Private CurrentTime As Long
Private FrameTime As Long
Private LastTime As Long
Private TextRemoveTime As Long
Private InfoUpdateTime As Long
Private InfoCoordsUpdateTime As Long

'Physics
Public HAngle As Single
Public VAngle As Single
Public Position As D3DVECTOR
Public Gravity As Single
Public TLastX As Long, TLastY As Long

'Texture browsing
Private itemnames() As String
Private useditemnames() As String
Private numitems As Long
Private numuseditems As Long
Private curitemnames() As String
Private curnumitems As Long
Private collection As Dictionary
Public Sub CaptureMouse()
     Dim client As RECT
     Dim upperleft As POINT
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo"))) Then
          
          'Make window rectangle
          frmMain.picMap.ScaleMode = vbPixels
          client.left = 0
          client.top = 0
          client.right = client.left + ScreenWidth - 1
          client.bottom = client.top + ScreenHeight - 1
          upperleft.x = client.left
          upperleft.y = client.top
          
          'Convert window coordinates to screen coordinates
          ClientToScreen frmMain.picMap.hWnd, upperleft
          
          'Move rectangle
          OffsetRect client, upperleft.x, upperleft.y
          
          'Limit the cursor movement
          ClipCursor client
          
          'And stay using pixels
          frmMain.picMap.ScaleMode = vbPixels
     End If
     
     'Remove hourglass and hide the cursor
     Screen.MousePointer = vbNormal
     While ShowCursor(False) >= 0: Wend
     
     'Start mouse events polling
     InitMouse
End Sub


Private Sub ApplyPhysics()
     Dim Velocity As D3DVECTOR
     Dim MoveSpeedFactor As Single
     Dim s As Long
     Dim hceiling As Single
     Dim hfloor As Single
     
     'Calculate movement speed factor
     MoveSpeedFactor = c_movespeed * MAP_RENDER_SCALE
     
     'Get the current sector
     s = IntersectSector(Position.x * INV_MAP_RENDER_SCALE, Position.y * INV_MAP_RENDER_SCALE, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
     
     'Set time for updating info panel when moving
     'If Key3DForward Or Key3DBackward Or Key3DStrafeLeft Or Key3DStrafeRight Then InfoUpdateTime = CurrentTime + INFO_UPDATEDELAY
     
     'Check if we should use gravity
     If ApplyGravity Then
          
          'Modify velocity with pressed keys
          If Key3DForward Then
               Velocity.x = Velocity.x + sIn(HAngle)
               Velocity.y = Velocity.y + Cos(HAngle)
               Velocity.Z = Velocity.Z + sIn(VAngle)
          End If
          If Key3DBackward Then
               Velocity.x = Velocity.x - sIn(HAngle)
               Velocity.y = Velocity.y - Cos(HAngle)
               Velocity.Z = Velocity.Z - sIn(VAngle)
          End If
          If Key3DStrafeLeft Then
               Velocity.x = Velocity.x + sIn(HAngle + pi * 0.5)
               Velocity.y = Velocity.y + Cos(HAngle + pi * 0.5)
               Velocity.Z = Velocity.Z + sIn(VAngle)
          End If
          If Key3DStrafeRight Then
               Velocity.x = Velocity.x + sIn(HAngle - pi * 0.5)
               Velocity.y = Velocity.y + Cos(HAngle - pi * 0.5)
               Velocity.Z = Velocity.Z + sIn(VAngle)
          End If
          
          'Apply velocity over time
          Position.x = Position.x + Velocity.x * 0.001 * FrameTime * MoveSpeedFactor
          Position.y = Position.y + Velocity.y * 0.001 * FrameTime * MoveSpeedFactor
          
          'Check if we can check for height
          If (s > -1) Then
               
               
               'Check if above ceiling and if we should stay below
               If (Position.Z >= (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE) And (c_belowceiling = vbChecked) Then
                    
                    'Begin at ceiling top if above real ceiling
                    If (Position.Z > sectors(s).hceiling * MAP_RENDER_SCALE) Then Position.Z = sectors(s).hceiling * MAP_RENDER_SCALE
                    
                    'Become heavier
                    Gravity = Gravity + GRAVITYWEIGHT
                    
                    'Go down
                    Position.Z = Position.Z - (0.01 * FrameTime * (Position.Z - (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE))
                         
                    'Check if above the floor
                    If (Position.Z > (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then
                         
                         'Become heavier
                         Gravity = Gravity + GRAVITYWEIGHT
                         
                         'Go down
                         Position.Z = Position.Z - Gravity * 0.0001 * FrameTime
                         
                         'Recheck and adjust
                         If (Position.Z - 0.01 < (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then Position.Z = (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE
                    Else
                         
                         'Recheck and adjust
                         If (Position.Z - 0.01 < (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE) Then Position.Z = (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE
                    End If
                    
               'Check if below the floor
               ElseIf (Position.Z < (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then
                    
                    'No more gravity to apply down
                    Gravity = 0
                    
                    'Begin at floor bottom if below real floor
                    If (Position.Z < sectors(s).hfloor * MAP_RENDER_SCALE) Then Position.Z = sectors(s).hfloor * MAP_RENDER_SCALE
                    
                    'Go up
                    Position.Z = Position.Z + 0.01 * FrameTime * ((sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE - Position.Z)
                    
                    'Recheck and adjust if below ceiling
                    If (Position.Z > (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE) And (c_belowceiling = vbChecked) Then Position.Z = (sectors(s).hceiling - HEADHEIGHT) * MAP_RENDER_SCALE
                    
                    'Recheck and adjust
                    If (Position.Z + 0.005 > (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then Position.Z = (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE
                    
               'Check if above the floor
               ElseIf (Position.Z > (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then
                    
                    'Become heavier
                    Gravity = Gravity + GRAVITYWEIGHT
                    
                    'Go down
                    Position.Z = Position.Z - Gravity * 0.0001 * FrameTime
                    
                    'Recheck and adjust
                    If (Position.Z - 0.01 < (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE) Then
                         
                         'Adjust position to stop moving
                         Position.Z = (sectors(s).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE
                    End If
               Else
                    
                    'No more gravity to apply down
                    Gravity = 0
               End If
          Else
               
               'Cant determine floor height,
               'do not apply gravity
               Gravity = 0
          End If
          
     Else
          
          'No more gravity to apply down
          Gravity = 0
          
          'Modify velocity with pressed keys
          If Key3DForward Then
               Velocity.x = Velocity.x + sIn(HAngle) * Cos(VAngle)
               Velocity.y = Velocity.y + Cos(HAngle) * Cos(VAngle)
               Velocity.Z = Velocity.Z + sIn(VAngle)
          End If
          If Key3DBackward Then
               Velocity.x = Velocity.x - sIn(HAngle) * Cos(VAngle)
               Velocity.y = Velocity.y - Cos(HAngle) * Cos(VAngle)
               Velocity.Z = Velocity.Z - sIn(VAngle)
          End If
          If Key3DStrafeLeft Then
               Velocity.x = Velocity.x + sIn(HAngle + pi * 0.5)
               Velocity.y = Velocity.y + Cos(HAngle + pi * 0.5)
               Velocity.Z = Velocity.Z
          End If
          If Key3DStrafeRight Then
               Velocity.x = Velocity.x + sIn(HAngle - pi * 0.5)
               Velocity.y = Velocity.y + Cos(HAngle - pi * 0.5)
               Velocity.Z = Velocity.Z
          End If
          If Key3DStrafeUp Then
               Velocity.Z = Velocity.Z + 1
          End If
          If Key3DStrafeDown Then
               Velocity.Z = Velocity.Z - 1
          End If
          
          'Apply velocity over time
          Position.x = Position.x + Velocity.x * 0.001 * FrameTime * MoveSpeedFactor
          Position.y = Position.y + Velocity.y * 0.001 * FrameTime * MoveSpeedFactor
          Position.Z = Position.Z + Velocity.Z * 0.001 * FrameTime * MoveSpeedFactor
     End If
End Sub

Public Sub ApplyPositionFromThing(ByVal t As Long)
     Dim ts As Long
     
     'Set position
     Position.x = things(t).x * MAP_RENDER_SCALE
     Position.y = -things(t).y * MAP_RENDER_SCALE
     
     'Get sector in which this thing is
     ts = IntersectSector(things(t).x, -things(t).y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
     
     'Set height only when sector could be found
     If (ts > -1) Then
          Position.Z = (sectors(ts).hfloor + EYESHEIGHT) * MAP_RENDER_SCALE
     'Else
     '     Position.Z = 64 * MAP_RENDER_SCALE
     End If
     
     'Set angle
     HAngle = (things(t).angle + 90) * PiDivMul
     VAngle = 0
     
     'Keep thing index
     PositionThing = t
End Sub

Private Sub AutoAlignLowerTextures(ByVal sd As Long, ByVal yoffsets As Boolean)
     Dim texturename As String
     Dim backside As Boolean
     
     'Show we're doing alignments
     ShowMainText "Performing alignments, please wait..."
     RunSingleFrame
     
     'Make undo
     CreateUndo "autoalign textures", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Get the texture name
     texturename = sidedefs(sd).lower
     
     'Determine back side
     backside = (linedefs(sidedefs(sd).linedef).s2 = sd)
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Perform alignment
     If (yoffsets) Then
          AlignTexturesY linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).ty, texturename, backside, sidedefs(sd).linedef
     Else
          AlignTexturesX linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).tx, texturename, backside, sidedefs(sd).linedef
     End If
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Remove all vertexbuffers
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Done
     ShowMainText "Texture alignment done"
End Sub

Private Sub AutoAlignMiddleTextures(ByVal sd As Long, ByVal yoffsets As Boolean)
     Dim texturename As String
     Dim backside As Boolean
     
     'Show we're doing alignments
     ShowMainText "Performing alignments, please wait..."
     RunSingleFrame
     
     'Make undo
     CreateUndo "autoalign textures", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Get the texture name
     texturename = sidedefs(sd).middle
     
     'Determine back side
     backside = (linedefs(sidedefs(sd).linedef).s2 = sd)
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Perform alignment
     If (yoffsets) Then
          AlignTexturesY linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).ty, texturename, backside, sidedefs(sd).linedef
     Else
          AlignTexturesX linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).tx, texturename, backside, sidedefs(sd).linedef
     End If
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Remove all vertexbuffers
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Done
     ShowMainText "Texture alignment done"
End Sub

Private Sub AutoAlignUpperTextures(ByVal sd As Long, ByVal yoffsets As Boolean)
     Dim texturename As String
     Dim backside As Boolean
     
     'Show we're doing alignments
     ShowMainText "Performing alignments, please wait..."
     RunSingleFrame
     
     'Make undo
     CreateUndo "autoalign textures", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Get the texture name
     texturename = sidedefs(sd).upper
     
     'Determine back side
     backside = (linedefs(sidedefs(sd).linedef).s2 = sd)
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Perform alignment
     If (yoffsets) Then
          AlignTexturesY linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).ty, texturename, backside, sidedefs(sd).linedef
     Else
          AlignTexturesX linedefs(sidedefs(sd).linedef).v1, sidedefs(sd).tx, texturename, backside, sidedefs(sd).linedef
     End If
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Remove all vertexbuffers
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Done
     ShowMainText "Texture alignment done"
End Sub

Private Sub DoFloodfillTextures(ByVal sd As Long, ByVal texturename As String)
     Dim backside As Boolean
     
     'Show we're doing floodfill
     ShowMainText "Performing floodfill, please wait..."
     RunSingleFrame
     
     'Make undo
     CreateUndo "floodfill texture", UGRP_NONE, 0, True
     
     'Determine back side
     backside = (linedefs(sidedefs(sd).linedef).s2 = sd)
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Perform alignment
     FloodFillTexture linedefs(sidedefs(sd).linedef).v1, texturename, backside, sidedefs(sd).linedef, CopiedTexture
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
     
     'Remove all vertexbuffers
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     
     'Remove all texture references
     ReDim i_SidedefUpper(-1 To numsidedefs - 1)
     ReDim i_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim i_SidedefLower(-1 To numsidedefs - 1)
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Done
     ShowMainText "Texture floodfill done"
End Sub



Private Sub DoFloodfillFlats(ByVal s As Long, ByVal floors As Boolean)
     
     'Show we're doing floodfill
     ShowMainText "Performing floodfill, please wait..."
     RunSingleFrame
     
     'Make undo
     CreateUndo "floodfill sectors", UGRP_NONE, 0, True
     
     'Remove sector selections
     ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), 0, VarPtr(sectors(0)), numsectors
     
     'Floors or ceilings?
     If floors Then
          FloodFillFlats s, sectors(s).tfloor, CopiedTexture, True
     Else
          FloodFillFlats s, sectors(s).tceiling, CopiedTexture, False
     End If
     
     'Remove linedef selections
     ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), 0, VarPtr(sectors(0)), numsectors
     
     'Remove all vertexbuffers
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     
     'Remove all texture references
     If floors Then
          ReDim i_SectorFloors(0 To numsectors - 1)
     Else
          ReDim i_SectorCeilings(0 To numsectors - 1)
     End If
     
     'Map has changed
     mapchanged = True
     mapnodeschanged = True
     
     'Done
     ShowMainText "Sectors floodfill done"
End Sub




Private Sub ChangeBrightness(ByVal sector As Long, ByVal Amount As Long)
     
     'Leave if exceeding the limits
     If (sectors(sector).Brightness = 0) And (Amount < 0) Then Exit Sub
     If (sectors(sector).Brightness = 255) And (Amount > 0) Then Exit Sub
     
     'Make undo
     CreateUndo "change brightness", UGRP_BRIGHNESSCHANGE, sector, True
     
     'This will fix the 16 changes offset
     If (sectors(sector).Brightness = 255) Then Amount = Amount + 1
     
     'Change the sector brightness
     sectors(sector).Brightness = sectors(sector).Brightness + Amount
     If (sectors(sector).Brightness < 0) Then sectors(sector).Brightness = 0
     If (sectors(sector).Brightness > 255) Then sectors(sector).Brightness = 255
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Show message
     ShowMainText "Brightness:  " & sectors(sector).Brightness
End Sub

Private Sub ChangeTextureOffset(ByVal sd As Long, ByVal x As Long, ByVal y As Long)
     
     'Make undo
     CreateUndo "texture offsets change", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Change texture offset
     sidedefs(sd).tx = sidedefs(sd).tx + x
     sidedefs(sd).ty = sidedefs(sd).ty + y
     
     'Show message
     ShowMainText "Texture offset:  " & sidedefs(sd).tx & ", " & sidedefs(sd).ty
     
     'Remove vertexbuffers so the sides will be recreated
     d_SidedefLower(sd) = False
     d_SidedefMiddle(sd) = False
     d_SidedefUpper(sd) = False
End Sub

Private Sub CleanUp3DMode()
     On Error Resume Next
     
     '3D Mode should stop now!
     Running3D = False
     TextureSelecting = False
     
     'Free the mouse
     FreeMouse
     
     'Terminate DirectX
     TerminateDirectX
     
     'Release pointers
     DestroyStructurePointers
     
     'Vertex buffers
     Set r_crosshair = Nothing
     Set r_maintext = Nothing
     Set r_texdesc = Nothing
     Set r_texname = Nothing
     Set r_infopanel = Nothing
     Set r_infotexts(0) = Nothing
     Set r_infotexts(1) = Nothing
     Set r_infotexts(2) = Nothing
     Set r_infotexts(3) = Nothing
     Set r_infotexts(4) = Nothing
     Set r_infotexts(5) = Nothing
     Set r_infotexts(6) = Nothing
     Set r_infotexts(7) = Nothing
     Set r_infotexts(8) = Nothing
     Set r_infotexts(9) = Nothing
     Set r_thingboxlines = Nothing
     Set r_thingboxvb = Nothing
     Set r_thingsprite = Nothing
     Set r_thingarrow = Nothing
     Set collection = Nothing
     
     'Texture previews
     Erase r_texclass()
     Erase r_texpoly()
     
     'Erase arrays
     Erase m_vertices()
     Erase m_segs()
     Erase m_subsectors()
     Erase m_nodes()
     Erase r_sidedefs()
     Erase r_discards()
     Erase r_subsectors()
     Erase r_things()
     Erase itemnames()
     Erase useditemnames()
     Erase curitemnames()
     
     'Destroy databases
     Erase SubSectorCeilings()
     Erase SubSectorFloors()
     Erase SidedefUpper()
     Erase SidedefMiddle()
     Erase SidedefLower()
     Erase d_SubSectorCeilings()
     Erase d_SubSectorFloors()
     Erase d_SidedefUpper()
     Erase d_SidedefMiddle()
     Erase d_SidedefLower()
     Erase i_SectorCeilings()
     Erase i_SectorFloors()
     Erase i_SidedefUpper()
     Erase i_SidedefMiddle()
     Erase i_SidedefLower()
     
     'Textures
     Set tex_crosshair = Nothing
     Set tex_font = Nothing
     Set tex_unknown = Nothing
     Set tex_missing = Nothing
     Set tex_thingbox = Nothing
     Set tex_thingarrow = Nothing
     
     'Unload textures/flats
     UnloadDirect3DFlats
     UnloadDirect3DTextures
     UnloadDirect3DSprites
     
     'Check if main form is loaded
     If (IsLoaded(frmMain)) Then
          
          'Stop redraw timer
          frmMain.tmr3DRedraw.Enabled = False
          
          'Enable menus
          frmMain.mnuEdit.Enabled = True
          frmMain.mnuFile.Enabled = True
          frmMain.mnuHelp.Enabled = True
          frmMain.mnuLines.Enabled = True
          frmMain.mnuPrefabs.Enabled = True
          frmMain.mnuScripts.Enabled = True
          frmMain.mnuSectors.Enabled = True
          frmMain.mnuThings.Enabled = True
          frmMain.mnuTools.Enabled = True
          frmMain.mnuVertices.Enabled = True
     End If
     
     'Check if fullscreen
     If (Val(Config("windowedvideo")) = 0) Then
          
          'Unload 3D rendering form
          Unload frm3D
          Set frm3D = Nothing
          
          'Focus there
          If (IsLoaded(frmMain)) Then frmMain.Show
     End If
End Sub

Private Sub CleanUpDiscards()
     Dim i As Long
     Dim d As Long
     
     'Get subsectors to discard
     GetMissingEntries r_prevsubsectors(0), r_numprevsubsectors, r_subsectors(0), r_numsubsectors, r_discards(0), r_numdiscards
     
     'Discard vertex buffers
     For i = 0 To (r_numdiscards - 1)
          
          'Get the index
          d = r_discards(i)
          
          'Discard
          d_SubSectorCeilings(d) = False
          d_SubSectorFloors(d) = False
          Set SubSectorCeilings(d) = Nothing
          Set SubSectorFloors(d) = Nothing
     Next i
     
     'Get sidedefs to discard
     GetMissingEntries r_prevsidedefs(0), r_numprevsidedefs, r_sidedefs(0), r_numsidedefs, r_discards(0), r_numdiscards
     
     'Discard vertex buffers
     For i = 0 To (r_numdiscards - 1)
          
          'Get the index
          d = r_discards(i)
          
          'Discard
          d_SidedefLower(d) = False
          d_SidedefMiddle(d) = False
          d_SidedefUpper(d) = False
          Set SidedefLower(d) = Nothing
          Set SidedefMiddle(d) = Nothing
          Set SidedefUpper(d) = Nothing
          'Set i_SidedefLower(d) = Nothing
          'Set i_SidedefMiddle(d) = Nothing
          'Set i_SidedefUpper(d) = Nothing
     Next i
     
     'Copy the currents to the previous
     CopyMemory r_prevsubsectors(0), r_subsectors(0), r_numsubsectors * 4
     CopyMemory r_prevsidedefs(0), r_sidedefs(0), r_numsidedefs * 4
     r_numprevsubsectors = r_numsubsectors
     r_numprevsidedefs = r_numsidedefs
End Sub

Private Sub CopySectorProperties(ByVal sector As Long)
     
     'Copy properties
     CopiedSector = sectors(sector)
     
     'Show message
     ShowMainText "Copied sector properties"
End Sub

Private Sub CopySidedefProperties(ByVal sd As Long)
     
     'Copy properties
     CopiedSidedef = sidedefs(sd)
     
     'Show message
     ShowMainText "Copied sidedef properties"
End Sub

Private Sub CopySidedefTexture(ByRef texturename As String)
     
     'Copy texture name
     CopiedTexture = texturename
     If (c_mixresource = vbChecked) Then CopiedFlat = texturename
     
     'Show info
     ShowMainText "Copied texture:  " & texturename
End Sub


Private Sub CopySidedefOffsets(ByVal sd As Long)
     
     'Copy offsets
     CopiedX = sidedefs(sd).tx
     CopiedY = sidedefs(sd).ty
     
     'Show info
     ShowMainText "Copied offsets:  " & CopiedX & ", " & CopiedY
End Sub



Private Sub CopySectorFlat(ByRef FlatName As String)
     
     'Copy texture name
     CopiedFlat = FlatName
     If (c_mixresource = vbChecked) Then CopiedTexture = FlatName
     
     'Show info
     ShowMainText "Copied flat:  " & FlatName
End Sub


Private Sub CopyThing(ByVal th As Long)
     
     'Copy thing properties
     CopiedThing = things(th)
     
     'Show info
     ShowMainText "Copied thing:  " & GetThingTypeDesc(things(th).thing) & " (" & things(th).thing & ")"
End Sub



Public Sub CreateSelectedTextureText()
     Dim TextRect As SRECT
     Dim ShownName As String
     
     'Check if we should show cursor
     If ShowTextCursor Then
          
          'Make the name to show
          ShownName = SelectedName & "_"
     Else
          
          'Make the name to show
          ShownName = SelectedName
     End If
     
     'Check if there is text to make
     If (LenB(ShownName) > 0) Then
          
          'Determine area
          With TextRect
               .left = 0.65
               .top = 0.9
               .right = 1
               .bottom = 1
          End With
          
          'Set the text
          Set r_texname = VertexBufferFromText(ShownName, TextRect, ALIGN_LEFT, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, TEXT_SIZE)
          r_numtexnamefaces = Len(ShownName) * 4 - 2
     Else
          
          'Erase
          Set r_texname = Nothing
          r_numtexnamefaces = 0
     End If
End Sub

Public Sub CreateSelectedThingText()
     Dim TextRect As SRECT
     Dim ShownName As String
     
     'Make the name to show
     ShownName = GetThingTypeDesc(Val(SelectedName)) & " (" & SelectedName & ")"
     
     'Determine area
     With TextRect
          .left = 0.4
          .top = 0.9
          .right = 1
          .bottom = 1
     End With
     
     'Set the text
     Set r_texname = VertexBufferFromText(ShownName, TextRect, ALIGN_LEFT, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, TEXT_SIZE)
     r_numtexnamefaces = Len(ShownName) * 4 - 2
End Sub


Private Function CreateSidedefLower(ByVal sd As Long, ByRef SidedefPoly() As VERTEX) As Long
     Dim ld As Long
     Dim side As Long
     Dim thissector As Long
     Dim othersector As Long
     Dim xl As Long, yl As Long
     Dim tx As Single, ty As Single
     Dim sx As Single, sy As Single
     Dim length As Long
     Dim floordifference As Long
     Dim texturename As String
     Dim TextureWidth As Long
     Dim TextureHeight As Long
     Dim Texture As clsImage
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Get this sector
     thissector = sidedefs(sd).sector
     
     'Get side and other sector
     If (linedefs(ld).s1 = sd) Then
          side = 1
          If (linedefs(ld).s2 > -1) Then othersector = sidedefs(linedefs(ld).s2).sector Else othersector = -1
     Else
          side = 2
          If (linedefs(ld).s1 > -1) Then othersector = sidedefs(linedefs(ld).s1).sector Else othersector = -1
     End If
     
     'Only continue if there is another sector
     If (othersector > -1) Then
          
          
          
          'Only continue if the other floor is higher
          If (sectors(othersector).hfloor > sectors(thissector).hfloor) Then
               
               'If (ld = 30) And (sd = linedefs(ld).s1) Then Stop
                    
                         
                         
                         
                    
                    
                         
                         
                         
                    
                    
               
               'Calculate linedef length
               xl = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
               yl = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
               length = CLng(Sqr(xl * xl + yl * yl))
               
               'Floor difference
               floordifference = (sectors(othersector).hfloor - sectors(thissector).hfloor)
               
               'Check if texture exists
               texturename = sidedefs(sd).lower
               If alltextures.Exists(texturename) Then
                    
                    'Convert needed texture if not done so yet
                    Set Texture = alltextures(texturename)
                    If (Texture.D3DTexture Is Nothing) Then Texture.LoadD3DTexture
                    
                    'Get the texture sizes
                    TextureWidth = Texture.width
                    TextureHeight = Texture.height
                    If (TextureWidth = 0) Then TextureWidth = 64
                    If (TextureHeight = 0) Then TextureHeight = 64
                    sx = Texture.ScaleX
                    sy = Texture.ScaleY
                    
                    'Check if unpegged
                    If (linedefs(ld).Flags And LDF_LOWERUNPEGGED) = LDF_LOWERUNPEGGED Then
                         
                         'Align texture to the facing ceiling
                         ty = (sectors(thissector).hceiling - sectors(othersector).hfloor) / TextureHeight
                    Else
                         
                         'Align texture to the top
                         ty = 0
                    End If
                    
                    'Apply texture coordinates
                    tx = tx + sidedefs(sd).tx / TextureWidth
                    ty = ty + sidedefs(sd).ty / TextureHeight
                    
                    'Check if coordinates must be adjusted
                    'to scale for world coordinates
                    If (Texture.Flags And IF_WORLDCOORDS) Then
                         tx = tx * sx
                         ty = ty * sy
                    End If
                    
                    
                    'Clean up
                    Set Texture = Nothing
               Else
                    
                    'No texture
                    TextureWidth = 64
                    TextureHeight = 64
                    sx = 1
                    sy = 1
               End If
               
               
               'Create first vertex
               With SidedefPoly(0)
                    .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
                    .Z = sectors(othersector).hfloor * MAP_RENDER_SCALE
                    .tu = tx
                    .tv = ty
               End With
               
               'Create second vertex
               With SidedefPoly(1)
                    .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
                    .Z = sectors(thissector).hfloor * MAP_RENDER_SCALE
                    .tu = tx
                    .tv = ty + (floordifference / TextureHeight) * sy
               End With
               
               'Create third vertex
               With SidedefPoly(2)
                    .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
                    .Z = sectors(othersector).hfloor * MAP_RENDER_SCALE
                    .tu = tx + (length / TextureWidth) * sx
                    .tv = ty
               End With
               
               'Create fourth vertex
               With SidedefPoly(3)
                    .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
                    .Z = sectors(thissector).hfloor * MAP_RENDER_SCALE
                    .tu = tx + (length / TextureWidth) * sx
                    .tv = ty + (floordifference / TextureHeight) * sy
               End With
               
               'Check if on the back side
               If (linedefs(ld).s2 = sd) Then SwitchSidedefPolygon SidedefPoly()
               
               'Polygon created
               CreateSidedefLower = True
          End If
     End If
End Function


Private Function CreateSidedefMiddle(ByVal sd As Long, ByRef SidedefPoly() As VERTEX) As Long
     Dim ld As Long
     Dim side As Long
     Dim thissector As Long
     Dim othersector As Long
     Dim zc As Long, zf As Long
     Dim xl As Long, yl As Long
     Dim tx As Single, ty As Single
     Dim sx As Single, sy As Single
     Dim length As Long
     Dim floordifference As Long
     Dim texturename As String
     Dim TextureWidth As Long
     Dim TextureHeight As Long
     Dim Texture As clsImage
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Get this sector
     thissector = sidedefs(sd).sector
     
     'Get side and other sector
     If (linedefs(ld).s1 = sd) Then
          side = 1
          If (linedefs(ld).s2 > -1) Then othersector = sidedefs(linedefs(ld).s2).sector Else othersector = thissector
     Else
          side = 2
          If (linedefs(ld).s1 > -1) Then othersector = sidedefs(linedefs(ld).s1).sector Else othersector = thissector
     End If
     
     'Get texture name
     texturename = sidedefs(sd).middle
     
     'Check if a middle texture is set or is singlesided
     If textures.Exists(texturename) Or (linedefs(ld).s2 = -1) Then
          
          'Determine top and bottom
          zc = sectors(thissector).hceiling
          zf = sectors(thissector).hfloor
          If (sectors(othersector).hceiling < zc) Then zc = sectors(othersector).hceiling
          If (sectors(othersector).hfloor > zf) Then zf = sectors(othersector).hfloor
          
          'Calculate linedef length
          xl = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
          yl = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
          length = CLng(Sqr(xl * xl + yl * yl))
          
          'Floor difference
          floordifference = (zc - zf)
          
          'Check if a middle texture is set
          If alltextures.Exists(texturename) Then
               
               'Convert needed texture if not done so yet
               Set Texture = alltextures(texturename)
               If (Texture.D3DTexture Is Nothing) Then Texture.LoadD3DTexture
               
               'Get the texture sizes
               TextureWidth = Texture.width
               TextureHeight = Texture.height
               If (TextureWidth = 0) Then TextureWidth = 64
               If (TextureHeight = 0) Then TextureHeight = 64
               sx = Texture.ScaleX
               sy = Texture.ScaleY
               
               'Check if unpegged
               If (linedefs(ld).Flags And LDF_LOWERUNPEGGED) = LDF_LOWERUNPEGGED Then
                    
                    'Align texture to floor
                    ty = (TextureHeight - floordifference) / TextureHeight
               Else
                    
                    'Align texture to the top
                    ty = 0
               End If
               
               'Apply texture coordinates
               tx = tx + sidedefs(sd).tx / TextureWidth
               ty = ty + sidedefs(sd).ty / TextureHeight
               
               'Check if coordinates must be adjusted
               'to scale for world coordinates
               If (Texture.Flags And IF_WORLDCOORDS) Then
                    tx = tx * sx
                    ty = ty * sy
               End If
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'No texture
               TextureWidth = 64
               TextureHeight = 64
               sx = 1
               sy = 1
          End If
          
          'Crop the heights to texture?
          If (linedefs(ld).s2 > -1) Then
               
               'Adjust the heights to texture
               zc = zc + ty * TextureHeight: ty = 0
               zf = zc - TextureHeight
               floordifference = (zc - zf)
               
               'Cut the heights with this sectors ceiling
               If (zc > sectors(thissector).hceiling) And (sectors(thissector).hceiling < sectors(othersector).hceiling) Then
                    
                    'Cut ceiling to this ceiling
                    ty = (zc - sectors(thissector).hceiling) / floordifference
                    zc = sectors(thissector).hceiling
                    floordifference = (zc - zf)
               
               'Cut the heights with other sectors ceiling
               ElseIf (zc > sectors(othersector).hceiling) Then
                    
                    'Cut ceiling to other ceiling
                    ty = (zc - sectors(othersector).hceiling) / floordifference
                    zc = sectors(othersector).hceiling
                    floordifference = (zc - zf)
               End If
               
               'Cut the heights with this sectors floor
               If (zf < sectors(thissector).hfloor) And (sectors(thissector).hfloor > sectors(othersector).hfloor) Then
                    
                    'Cut floor to this floor
                    zf = sectors(thissector).hfloor
                    floordifference = (zc - zf)
               
               'Cut the heights with other sectors floor
               ElseIf (zf < sectors(othersector).hfloor) Then
                    
                    'Cut floor to other floor
                    zf = sectors(othersector).hfloor
                    floordifference = (zc - zf)
               End If
               
               'Store heights in the sidedef data
               'this is used for better object picking
               sidedefs(sd).MiddleTop = zc
               sidedefs(sd).MiddleBottom = zf
          End If
          
          
          'Create first vertex
          With SidedefPoly(0)
               .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
               .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
               .Z = zc * MAP_RENDER_SCALE
               .tu = tx
               .tv = ty
          End With
          
          'Create second vertex
          With SidedefPoly(1)
               .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
               .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
               .Z = zf * MAP_RENDER_SCALE
               .tu = tx
               .tv = ty + (floordifference / TextureHeight) * sy
          End With
          
          'Create third vertex
          With SidedefPoly(2)
               .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
               .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
               .Z = zc * MAP_RENDER_SCALE
               .tu = tx + (length / TextureWidth) * sx
               .tv = ty
          End With
          
          'Create fourth vertex
          With SidedefPoly(3)
               .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
               .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
               .Z = zf * MAP_RENDER_SCALE
               .tu = tx + (length / TextureWidth) * sx
               .tv = ty + (floordifference / TextureHeight) * sy
          End With
          
          'Check if on the back side
          If (linedefs(ld).s2 = sd) Then SwitchSidedefPolygon SidedefPoly()
          
          'Polygon created
          CreateSidedefMiddle = True
     End If
End Function

Private Function CreateSidedefUpper(ByVal sd As Long, ByRef SidedefPoly() As VERTEX) As Long
     Dim ld As Long
     Dim side As Long
     Dim thissector As Long
     Dim othersector As Long
     Dim xl As Long, yl As Long
     Dim tx As Single, ty As Single
     Dim sx As Single, sy As Single
     Dim length As Long
     Dim floordifference As Long
     Dim texturename As String
     Dim TextureWidth As Long
     Dim TextureHeight As Long
     Dim Texture As clsImage
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Get this sector
     thissector = sidedefs(sd).sector
     
     'Get side and other sector
     If (linedefs(ld).s1 = sd) Then
          side = 1
          If (linedefs(ld).s2 > -1) Then othersector = sidedefs(linedefs(ld).s2).sector Else othersector = -1
     Else
          side = 2
          If (linedefs(ld).s1 > -1) Then othersector = sidedefs(linedefs(ld).s1).sector Else othersector = -1
     End If
     
     'Only continue if there is another sector
     If (othersector > -1) Then
          
          'Only continue if the other sector is lower
          If (sectors(othersector).hceiling < sectors(thissector).hceiling) Then
               
               'Calculate linedef length
               xl = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
               yl = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
               length = CLng(Sqr(xl * xl + yl * yl))
               
               'Floor difference
               floordifference = (sectors(thissector).hceiling - sectors(othersector).hceiling)
               
               'Check if texture exists
               texturename = sidedefs(sd).upper
               If alltextures.Exists(texturename) Then
                    
                    'Convert needed texture if not done so yet
                    Set Texture = alltextures(texturename)
                    If (Texture.D3DTexture Is Nothing) Then Texture.LoadD3DTexture
                    
                    'Get the texture sizes
                    TextureWidth = Texture.width
                    TextureHeight = Texture.height
                    If (TextureWidth = 0) Then TextureWidth = 64
                    If (TextureHeight = 0) Then TextureHeight = 64
                    sx = Texture.ScaleX
                    sy = Texture.ScaleY
                    
                    'Check if unpegged
                    If (linedefs(ld).Flags And LDF_UPPERUNPEGGED) = LDF_UPPERUNPEGGED Then
                         
                         'Align texture to the ceiling top
                         ty = 0
                    Else
                         
                         'Align texture to the other ceilings top
                         ty = TextureHeight - floordifference / TextureHeight
                    End If
                    
                    'Apply texture coordinates
                    tx = tx + sidedefs(sd).tx / TextureWidth
                    ty = ty + sidedefs(sd).ty / TextureHeight
                    
                    'Check if coordinates must be adjusted
                    'to scale for world coordinates
                    If (Texture.Flags And IF_WORLDCOORDS) Then
                         tx = tx * sx
                         ty = ty * sy
                    End If
                    
                    'Clean up
                    Set Texture = Nothing
               Else
                    
                    'No texture
                    TextureWidth = 64
                    TextureHeight = 64
                    sx = 1
                    sy = 1
               End If
               
               
               'Create first vertex
               With SidedefPoly(0)
                    .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
                    .Z = sectors(thissector).hceiling * MAP_RENDER_SCALE
                    .tu = tx
                    .tv = ty
               End With
               
               'Create second vertex
               With SidedefPoly(1)
                    .x = m_vertices(linedefs(ld).v1).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v1).y * MAP_RENDER_SCALE
                    .Z = sectors(othersector).hceiling * MAP_RENDER_SCALE
                    .tu = tx
                    .tv = ty + (floordifference / TextureHeight) * sy
               End With
               
               'Create third vertex
               With SidedefPoly(2)
                    .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
                    .Z = sectors(thissector).hceiling * MAP_RENDER_SCALE
                    .tu = tx + (length / TextureWidth) * sx
                    .tv = ty
               End With
               
               'Create fourth vertex
               With SidedefPoly(3)
                    .x = m_vertices(linedefs(ld).v2).x * MAP_RENDER_SCALE
                    .y = -m_vertices(linedefs(ld).v2).y * MAP_RENDER_SCALE
                    .Z = sectors(othersector).hceiling * MAP_RENDER_SCALE
                    .tu = tx + (length / TextureWidth) * sx
                    .tv = ty + (floordifference / TextureHeight) * sy
               End With
               
               'Check if on the back side
               If (linedefs(ld).s2 = sd) Then SwitchSidedefPolygon SidedefPoly()
               
               'Polygon created
               CreateSidedefUpper = True
          End If
     End If
End Function

Private Sub CreateSubSectorCeiling(ByVal ss As Long, ByRef SSectorPoly() As VERTEX)
     Dim v As Long, height As Long
     Dim s As Long
     Dim Flat As clsImage
     Dim FlatWidth As Long, FlatHeight As Long
     Dim sx As Single, sy As Single
     
     'Get the sector
     s = m_subsectors(ss).sector
     
     'Get the height
     height = sectors(s).hceiling
     
     'Get the flat sizes
     If allflats.Exists(sectors(s).tceiling) Then
          Set Flat = allflats(sectors(s).tceiling)
          If (Flat.D3DTexture Is Nothing) Then Flat.LoadD3DTexture
          FlatWidth = Flat.width
          FlatHeight = Flat.height
          sx = Flat.ScaleX
          sy = Flat.ScaleY
     Else
          sx = 1
          sy = 1
     End If
     If (FlatWidth = 0) Then FlatWidth = 64
     If (FlatHeight = 0) Then FlatHeight = 64
     
     'Reserve some memory
     ReDim SSectorPoly(0 To m_subsectors(ss).numvertices - 1)
     
     'Go for each vertex
     For v = 0 To (m_subsectors(ss).numvertices - 1)
          
          'Create the D3DVERTEX
          With SSectorPoly(v)
               .x = m_subsectors(ss).vertices(v).x * MAP_RENDER_SCALE
               .y = -m_subsectors(ss).vertices(v).y * MAP_RENDER_SCALE
               .Z = height * MAP_RENDER_SCALE
               .tu = (m_subsectors(ss).vertices(v).x) / FlatWidth * sx
               .tv = -((m_subsectors(ss).vertices(v).y) / FlatHeight) * sy
               
          End With
     Next v
End Sub

Private Sub CreateSubSectorFloor(ByVal ss As Long, ByRef SSectorPoly() As VERTEX)
     Dim v As Long, height As Long
     Dim s As Long
     Dim Flat As clsImage
     Dim FlatWidth As Long, FlatHeight As Long
     Dim sx As Single, sy As Single
     
     'Get the sector
     s = m_subsectors(ss).sector
     
     'Get the height
     height = sectors(s).hfloor
     
     'Get the flat sizes
     If allflats.Exists(sectors(s).tfloor) Then
          Set Flat = allflats(sectors(s).tfloor)
          If (Flat.D3DTexture Is Nothing) Then Flat.LoadD3DTexture
          FlatWidth = Flat.width
          FlatHeight = Flat.height
          sx = Flat.ScaleX
          sy = Flat.ScaleY
     Else
          sx = 1
          sy = 1
     End If
     If (FlatWidth = 0) Then FlatWidth = 64
     If (FlatHeight = 0) Then FlatHeight = 64
     
     'Reserve some memory
     ReDim SSectorPoly(0 To m_subsectors(ss).numvertices - 1)
     
     'Go for each vertex
     For v = 0 To (m_subsectors(ss).numvertices - 1)
          
          'Create the D3DVERTEX
          With SSectorPoly(v)
               .x = m_subsectors(ss).vertices(v).x * MAP_RENDER_SCALE
               .y = -m_subsectors(ss).vertices(v).y * MAP_RENDER_SCALE
               .Z = height * MAP_RENDER_SCALE
               .tu = (m_subsectors(ss).vertices(v).x) / FlatWidth * sx
               .tv = -((m_subsectors(ss).vertices(v).y) / FlatHeight) * sy
               
          End With
     Next v
End Sub

Private Sub CreateTexturePreviews()
     Dim Shown As Long
     Dim offset As Long
     Dim Texture As clsImage
     Dim TexturePoly(3) As TLVERTEX
     Dim w As Long, h As Long
     Dim cw As Single, ch As Single
     Dim sw As Single, sh As Single
     Dim tw As Single, th As Single
     Dim x As Long, y As Long
     Dim i As Long
     Dim ci As Long
     
     'Calculate number of textures we can show
     Shown = TEXTURE_COLS * TEXTURE_ROWS
     
     'Calculate index offset
     offset = TextureRowOffset * TEXTURE_COLS
     
     'Calculate cell width and height
     cw = (VideoParams.BackBufferWidth * (1 - TEXTURE_SPACING)) / TEXTURE_COLS
     ch = (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT) * (1 - TEXTURE_SPACING)) / TEXTURE_ROWS
     
     'Calculate cell spacing
     sw = (VideoParams.BackBufferWidth * TEXTURE_SPACING) / TEXTURE_COLS
     sh = (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT) * TEXTURE_SPACING) / TEXTURE_ROWS
     
     'Go for all textures to be shown
     For i = offset To (offset + Shown - 1)
          
          'Get visual index
          ci = i - offset
          
          'Determine x an y
          y = ci \ TEXTURE_COLS
          x = ci - y * TEXTURE_COLS
          
          'Check if within bounds
          If (i < curnumitems) Then
               
               'Get texture object
               Set Texture = collection(curitemnames(i))
               
               'Ensure the texture is loaded
               If (Texture.D3DTexture Is Nothing) Then Texture.LoadD3DTexture ThingSelecting
               
               'Check if anything
               If Not (Texture Is Nothing) Then
                    
                    'Set texture scale
                    Texture.GetScale cw, ch, w, h, False
                    
                    'Check if making previews for Things
                    If (ThingSelecting) Then
                         tw = Texture.d3dscalewidth
                         th = Texture.d3dscaleheight
                    Else
                         tw = 0.99
                         th = 0.99
                    End If
               Else
                    
                    'Standard scale
                    w = cw
                    h = ch
                    tw = 0.99
                    th = 0.99
               End If
               
               'Create Polgon
               With TexturePoly(0)
                    .Color = D3DColorMake(1, 1, 1, 1)
                    .rhw = 1
                    .sx = (cw + sw) * x + (cw - w + sw) * 0.5
                    .sy = (ch + sh) * y + (ch - h + sh) * 0.5
                    .tu = 0
                    .tv = 0
               End With
               
               With TexturePoly(1)
                    .Color = D3DColorMake(1, 1, 1, 1)
                    .rhw = 1
                    .sx = (cw + sw) * x + (cw - w + sw) * 0.5
                    .sy = (ch + sh) * y + (ch + h + sh) * 0.5
                    .tu = 0
                    .tv = th
               End With
               
               With TexturePoly(2)
                    .Color = D3DColorMake(1, 1, 1, 1)
                    .rhw = 1
                    .sx = (cw + sw) * x + (cw + w + sw) * 0.5
                    .sy = (ch + sh) * y + (ch - h + sh) * 0.5
                    .tu = tw
                    .tv = 0
               End With
               
               With TexturePoly(3)
                    .Color = D3DColorMake(1, 1, 1, 1)
                    .rhw = 1
                    .sx = (cw + sw) * x + (cw + w + sw) * 0.5
                    .sy = (ch + sh) * y + (ch + h + sh) * 0.5
                    .tu = tw
                    .tv = th
               End With
               
               'Make vertexbuffer for the preview item
               Set r_texpoly(ci) = CreateTLVertexBuffer(TexturePoly(), 4)
               Set r_texclass(ci) = Texture
          Else
               
               'Clear preview item
               Set r_texpoly(ci) = Nothing
               Set r_texclass(ci) = Nothing
          End If
     Next i
End Sub

Public Function CreateTLVertexBuffer(ByRef polygon() As TLVERTEX, ByVal VertCount As Long) As Direct3DVertexBuffer9
     Dim BUFFERSIZE As Long
     
     'Calculate buffer size in bytes
     BUFFERSIZE = TLVERTEXSTRIDE * VertCount
     
     'Create the vertex buffer
     Set CreateTLVertexBuffer = D3DD.CreateVertexBuffer(BUFFERSIZE, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, TLVERTEXFVF, D3DPOOL_DEFAULT)
     
     'Copy the vertices to the buffer
     CreateTLVertexBuffer.SetData 0, BUFFERSIZE, VarPtr(polygon(0)), 0
End Function

Public Function CreateVertexBuffer(ByRef polygon() As VERTEX, ByVal VertCount As Long) As Direct3DVertexBuffer9
     Dim BUFFERSIZE As Long
     
     'Calculate buffer size in bytes
     BUFFERSIZE = VERTEXSTRIDE * VertCount
     
     'Create the vertex buffer
     Set CreateVertexBuffer = D3DD.CreateVertexBuffer(BUFFERSIZE, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, VERTEXFVF, D3DPOOL_DEFAULT)
     
     'Copy the vertices to the buffer
     CreateVertexBuffer.SetData 0, BUFFERSIZE, VarPtr(polygon(0)), 0
End Function

     
     
     


Private Sub DeleteLowerTexture(ByVal sd As Long)
     
     'Make undo
     CreateUndo "remove lower texture", UGRP_LOWERTEXTUREDELETE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Delete it
     sidedefs(sd).lower = "-"
     
     'Show message
     ShowMainText "Removed lower texture"
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefLower(sd) = False
     Set SidedefLower(sd) = Nothing
     Set i_SidedefLower(sd) = Nothing
End Sub

Private Sub DeleteMiddleTexture(ByVal sd As Long)
     
     'Make undo
     CreateUndo "remove middle texture", UGRP_MIDDLETEXTUREDELETE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Delete it
     sidedefs(sd).middle = "-"
     
     'Show message
     ShowMainText "Removed middle texture"
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefMiddle(sd) = False
     Set SidedefMiddle(sd) = Nothing
     Set i_SidedefMiddle(sd) = Nothing
End Sub

Private Sub DeleteUpperTexture(ByVal sd As Long)
     
     'Make undo
     CreateUndo "remove upper texture", UGRP_UPPERTEXTUREDELETE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Delete it
     sidedefs(sd).upper = "-"
     
     'Show message
     ShowMainText "Removed upper texture"
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefUpper(sd) = False
     Set SidedefUpper(sd) = Nothing
     Set i_SidedefUpper(sd) = Nothing
End Sub

Private Sub DeleteThing(ByVal th As Long)
     Dim tt As Long
     
     'Make undo
     CreateUndo "remove thing", UGRP_NONE, 0, True
     
     'Map changed
     mapchanged = True
     
     'Destroy pointers
     DestroyStructurePointers
     
     'Delete it
     tt = things(th).thing
     RemoveThing th
     
     'Give pointers to the DLL
     SetStructurePointers m_vertices(0), linedefs(0), VarPtr(sidedefs(0)), m_segs(0), VarPtr(sectors(0)), m_subsectors(0), things(0), m_nodes(0), numnodes, numsectors, numsubsectors, numthings
     
     'Show info
     ShowMainText "Removed thing:  " & GetThingTypeDesc(tt) & " (" & tt & ")"
End Sub


Public Sub DirectXPrecache()
     Dim i As Long
     Dim Names As Variant
     Dim Texture As clsImage
     Dim Flat As clsImage
     
     'Show the precache process
     ShowMainText "Precaching texture resources..."
     RunSingleFrame False, False
     
     'Go for all textures
     Names = alltextures.Keys
     For i = 0 To (alltextures.Count - 1)
          
          'Get object
          Set Texture = alltextures(Names(i))
          
          'Load if not yet loaded
          If (Texture.D3DTexture Is Nothing) Then Texture.LoadD3DTexture
          
          'Clean up
          Set Texture = Nothing
     Next i
     
     'Show the precache process
     ShowMainText "Precaching flat resources..."
     RunSingleFrame False, False
     
     'Go for all flats
     Names = allflats.Keys
     For i = 0 To (allflats.Count - 1)
          
          'Get object
          Set Flat = allflats(Names(i))
          
          'Load if not yet loaded
          If (Flat.D3DTexture Is Nothing) Then Flat.LoadD3DTexture
          
          'Clean up
          Set Flat = Nothing
     Next i
End Sub

Public Sub Init3DModeDefaults()
     Dim t As Long, th As Long
     Dim t_found As Boolean
     
     'Default limits
     r_numdiscards = 0
     r_numprevsidedefs = 0
     r_numprevsubsectors = 0
     r_numsidedefs = 0
     r_numsubsectors = 0
     HasProcessed = False
     
     
     'Check if the position thing is within bounds
     If (PositionThing >= 0) And (PositionThing < numthings) Then
          
          'Check if the position thing is correct
          If (things(PositionThing).thing = mapconfig("start3dmode")) Then t_found = True
     End If
     
     'If no thing could be found, find a new one
     If (t_found = False) Then
          
          'Go for all things to find another positioning thing
          For t = 0 To (numthings - 1)
               
               'Check if this is a 3D start position
               If (things(t).thing = mapconfig("start3dmode")) Then
                    
                    'Use this
                    ApplyPositionFromThing t
                    
                    'Found one
                    t_found = True
                    Exit For
               End If
          Next t
     End If
     
     'If no thing could be found, find player 1 start
     If (t_found = False) Then
          
          'Go for all things to find another positioning thing
          For t = 0 To (numthings - 1)
               
               'Check if this is a player 1 start
               If (things(t).thing = 1) Then
                    
                    'Use this
                    ApplyPositionFromThing th
                    
                    'Found one
                    t_found = True
                    Exit For
               End If
          Next t
     End If
End Sub

Private Sub InitTextureSelect(ByVal CurrentTexture As String, ByVal UseFlats As Boolean)
     Dim useditems As Dictionary
     Dim Keys As Variant
     Dim i As Long
     
     'Defaults
     TextureSelectCancelled = False
     TextureSelectedIndex = -1
     TextureRowOffset = 0
     TextureUseFlats = UseFlats
     TextureEraseOnType = True
     IgnoreInput = True
     ShowAllTextures = False
     
     'Check if using flats
     If UseFlats Then
          
          'Set information for flats
          Set collection = flats
          numitems = collection.Count
     Else
          
          'Set information for textures
          Set collection = textures
          numitems = collection.Count
     End If
     
     'Get the key names
     Keys = collection.Keys
     
     'Allocate memory for string names
     ReDim itemnames(0 To numitems - 1)
     
     'Make string array from names
     For i = 0 To numitems - 1
          itemnames(i) = Keys(i)
     Next i
     
     'Clear collection
     Set useditems = New Dictionary
     
     'Check if we should select used names from sidedefs (textures)
     If (UseFlats = False) Or (Val(mapconfig("mixtexturesflats")) = vbChecked) Then
          
          'Go for all sidedefs
          For i = 0 To numsidedefs - 1
               If (useditems.Exists(sidedefs(i).upper) = False) Then If (collection.Exists(sidedefs(i).upper)) Then useditems.Add sidedefs(i).upper, 1
               If (useditems.Exists(sidedefs(i).middle) = False) Then If (collection.Exists(sidedefs(i).middle)) Then useditems.Add sidedefs(i).middle, 1
               If (useditems.Exists(sidedefs(i).lower) = False) Then If (collection.Exists(sidedefs(i).lower)) Then useditems.Add sidedefs(i).lower, 1
          Next i
     End If
     
     'Check if we should select used names from sectors (flats)
     If (UseFlats = True) Or (Val(mapconfig("mixtexturesflats")) = vbChecked) Then
          
          'Go for all sector
          For i = 0 To numsectors - 1
               If (useditems.Exists(sectors(i).tfloor) = False) Then If (collection.Exists(sectors(i).tfloor)) Then useditems.Add sectors(i).tfloor, 1
               If (useditems.Exists(sectors(i).tceiling) = False) Then If (collection.Exists(sectors(i).tceiling)) Then useditems.Add sectors(i).tceiling, 1
          Next i
     End If
     
     'Any items used?
     If (useditems.Count > 0) And (Val(Config("alwaysalltextures")) = vbUnchecked) Then
          
          'Sort used items
          Set useditems = SortDictionary(useditems)
          numuseditems = useditems.Count
          
          'Allocate memory for string names
          ReDim useditemnames(0 To numuseditems - 1)
          Keys = useditems.Keys
          
          'Make string array from texture names
          For i = 0 To numuseditems - 1
               useditemnames(i) = Keys(i)
          Next i
          
          'Set the current collection
          curitemnames() = useditemnames()
          curnumitems = numuseditems
     Else
          
          'Show all textures now
          ShowAllTextures = True
          curitemnames() = itemnames()
          curnumitems = numitems
     End If
     
     'Go for all items
     For i = 0 To (curnumitems - 1)
          
          'Check if this is the current texture
          If (StrComp(CurrentTexture, curitemnames(i), vbTextCompare) = 0) Then
               
               'Found, keep it and blow this joint
               TextureSelectedIndex = i
               TextureRowOffset = (TextureSelectedIndex \ TEXTURE_COLS) - 2
               Exit For
          End If
     Next i
     
     'Limit the scroll
     If (TextureRowOffset > (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)) Then TextureRowOffset = (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)
     If (TextureRowOffset < 0) Then TextureRowOffset = 0
End Sub

Private Sub InitThingSelect(ByVal CurrentThingType As Long)
     Dim Keys As Variant
     Dim useditems As Dictionary
     Dim ThingCollection As Dictionary
     Dim th As Dictionary
     Dim ThingKeys As Variant
     Dim CurKey As String
     Dim Sprite As clsImage
     Dim i As Long
     
     'Defaults
     TextureSelectCancelled = False
     TextureSelectedIndex = -1
     TextureRowOffset = 0
     TextureUseFlats = False
     TextureEraseOnType = False
     IgnoreInput = True
     'ShowAllTextures = False
     ShowAllTextures = True
     
     'Create collection
     Set ThingCollection = New Dictionary
     Set collection = New Dictionary
     
     'Get things
     ThingKeys = mapconfig("__things").Keys
     
     'Go for all things
     For i = LBound(ThingKeys) To UBound(ThingKeys)
          
          'Get the key
          CurKey = ThingKeys(i)
          
          'Check if not one of the category properties
          If IsNumeric(CurKey) Then
               
               'Get the thing
               Set th = mapconfig("__things")(CurKey)
               
               'Check if the thing will be visible in 3D Mode
               If (th("width") <> 0) And (th("height") <> 0) And (th("sprite") <> "") Then
                    
                    'Get image
                    Set Sprite = GetSpriteForThingType(Val(CurKey), False)
                    
                    'Check if image found
                    If Not (Sprite Is Nothing) Then
                         
                         'Add thing to list
                         ThingCollection.Add CurKey, th("title")
                         collection.Add CurKey, Sprite
                    End If
               End If
          End If
     Next i
     
     'Sort all items
     'Set ThingCollection = SortDictionaryByValue(ThingCollection)
     
     'Allocate memory for string names
     numitems = ThingCollection.Count
     ReDim itemnames(0 To numitems - 1)
     
     'Make string array from names
     ThingKeys = ThingCollection.Keys
     For i = 0 To numitems - 1
          itemnames(i) = ThingKeys(i)
     Next i
     
'     'Create collection
'     Set useditems = New Dictionary
'
'     'Go for all things
'     For i = 0 To numthings - 1
'          If (useditems.Exists(CStr(things(i).thing)) = False) Then If (ThingCollection.Exists(CStr(things(i).thing))) Then useditems.Add CStr(things(i).thing), GetThingTypeDesc(things(i).thing)
'     Next i
'
'     'Sort used items
'     Set useditems = SortDictionaryByValue(useditems)
'
'     'Allocate memory for string names
'     numuseditems = useditems.Count
'     ReDim useditemnames(0 To numuseditems - 1)
'     Keys = useditems.Keys
'
'     'Make string array from texture names
'     For i = 0 To numuseditems - 1
'          useditemnames(i) = Keys(i)
'     Next i
     
'     'Set the current collection
'     curitemnames() = useditemnames()
'     curnumitems = numuseditems
     curitemnames() = itemnames()
     curnumitems = numitems
     
     'Go for all items
     For i = 0 To (curnumitems - 1)
          
          'Check if this is the current texture
          If (CurrentThingType = CLng(curitemnames(i))) Then
               
               'Found, keep it and blow this joint
               TextureSelectedIndex = i
               TextureRowOffset = (TextureSelectedIndex \ TEXTURE_COLS) - 2
               Exit For
          End If
     Next i
     
     'Limit the scroll
     If (TextureRowOffset > (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)) Then TextureRowOffset = (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)
     If (TextureRowOffset < 0) Then TextureRowOffset = 0
End Sub


Public Sub Keydown3D(ByVal ShortcutCode As Long)
     Dim Obj As Long
     Dim ObjType As ENUM_OBJECTTYPES
     Dim ObjSpot As D3DVECTOR
     
     'Check if Mode switch keys must be used
     If (Val(Config("modekeys3d"))) Then
          
          'Check if one of the mode keys is used
          Select Case ShortcutCode
               Case Config("shortcuts")("editvertices"): frmMain.mnuEdit_Click: If frmMain.itmEditMode(1).Enabled Then frmMain.itmEditMode_Click 1
               Case Config("shortcuts")("editlines"): frmMain.mnuEdit_Click: If frmMain.itmEditMode(2).Enabled Then frmMain.itmEditMode_Click 2
               Case Config("shortcuts")("editsectors"): frmMain.mnuEdit_Click: If frmMain.itmEditMode(3).Enabled Then frmMain.itmEditMode_Click 3
               Case Config("shortcuts")("editthings"): frmMain.mnuEdit_Click: If frmMain.itmEditMode(4).Enabled Then frmMain.itmEditMode_Click 4
          End Select
     End If
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo"))) Then
          
          'Check if one of the 2D keys is used
          Select Case ShortcutCode
               
               Case Config("shortcuts")("togglebar"): frmMain.InfoBarToggle
               
          End Select
     End If
     
     'Check what key
     Select Case ShortcutCode
          
          'These are copied from general/menu keys
          Case Config("shortcuts")("filetest")
               
               'Check file menu
               frmMain.mnuFile_Click
               If frmMain.itmFile(12).Enabled Then
                    
                    'Return to previous mode
                    frmMain.itmEditMode_Click CInt(PreviousMode)
                    
                    'Test the map
                    frmMain.itmFileTest_Click False
               End If
               
          Case Config("shortcuts")("filetest2")
               
               'Check file menu
               frmMain.mnuFile_Click
               If frmMain.itmFile(12).Enabled Then
                    
                    'Return to previous mode
                    frmMain.itmEditMode_Click CInt(PreviousMode)
                    
                    'Test the map
                    frmMain.itmFileTest_Click True
               End If
               
          Case Config("shortcuts")("fileconfig")
               
               'Check tools menu
               frmMain.mnuTools_Click
               If frmMain.itmToolsConfiguration.Enabled Then
                    
                    'Return to previous mode
                    frmMain.itmEditMode_Click CInt(PreviousMode)
                    
                    'Show configuration
                    frmMain.itmToolsConfiguration_Click
               End If
               
          'These are for navigation
          Case Config("shortcuts")("mode3dforward"): Key3DForward = True
          Case Config("shortcuts")("mode3dbackward"): Key3DBackward = True
          Case Config("shortcuts")("mode3dstrafeleft"): Key3DStrafeLeft = True
          Case Config("shortcuts")("mode3dstraferight"): Key3DStrafeRight = True
          Case Config("shortcuts")("mode3dstrafeup"): Key3DStrafeUp = True
          Case Config("shortcuts")("mode3dstrafedown"): Key3DStrafeDown = True
          
          'This leaves 3D mode
          Case Config("shortcuts")("mode3dexit")
               
               'Show unloading message
               ShowMainText "Switching to previous mode..."
               RunSingleFrame False, False
               
               'Switch to previous mode
               frmMain.itmEditMode_Click CInt(PreviousMode)
          
          'Other
          Case Config("shortcuts")("mode3dgravity"): ApplyGravity = Not ApplyGravity: ShowMainText "Gravity:  " & OnOff(ApplyGravity)
          Case Config("shortcuts")("mode3dfullbright"): FullBrightness = Not FullBrightness: ShowMainText "Lighting:  " & OnOff(Not FullBrightness)
               
          Case Config("shortcuts")("editundo")
               
               'Check if we can undo
               If frmMain.itmEditUndo.Enabled Then
                    
                    'Check if can be done during 3D mode
                    If AllowThis3DUndo Then
                         
                         'Show loading message
                         ShowMainText "Performing undo, please wait..."
                         RunSingleFrame
                         
                         'Destroy pointers
                         DestroyStructurePointers
                         
                         'Remove selection and highlight
                         'frmMain.RemoveHighlight
                         RemoveSelection False
                         
                         'Do the undo
                         PerformUndo
                         
                         'Remove all vertexbuffers
                         ReDim SubSectorFloors(0 To numsubsectors - 1)
                         ReDim SubSectorCeilings(0 To numsubsectors - 1)
                         ReDim SidedefUpper(-1 To numsidedefs - 1)
                         ReDim SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim SidedefLower(-1 To numsidedefs - 1)
                         ReDim d_SubSectorFloors(0 To numsubsectors - 1)
                         ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
                         ReDim d_SidedefUpper(-1 To numsidedefs - 1)
                         ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim d_SidedefLower(-1 To numsidedefs - 1)
                         ReDim i_SectorFloors(0 To numsectors - 1)
                         ReDim i_SectorCeilings(0 To numsectors - 1)
                         ReDim i_SidedefUpper(-1 To numsidedefs - 1)
                         ReDim i_SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim i_SidedefLower(-1 To numsidedefs - 1)
                         
                         'Give pointers to the DLL
                         SetStructurePointers m_vertices(0), linedefs(0), VarPtr(sidedefs(0)), m_segs(0), VarPtr(sectors(0)), m_subsectors(0), things(0), m_nodes(0), numnodes, numsectors, numsubsectors, numthings
                         
                         'Show undo message
                         ShowMainText RedoDescription & " undone"
                    Else
                         
                         'Show its not possible now
                         ShowMainText "Previous change cannot be undone in 3D mode"
                    End If
               End If
               
          Case Config("shortcuts")("editredo")
               
               'Check if we can redo
               If frmMain.itmEditRedo.Enabled Then
                    
                    'Check if can be done during 3D mode
                    If AllowThis3DRedo Then
                         
                         'Show loading message
                         ShowMainText "Performing redo, please wait..."
                         RunSingleFrame
                         
                         'Destroy pointers
                         DestroyStructurePointers
                         
                         'Remove selection and highlight
                         'frmMain.RemoveHighlight
                         RemoveSelection False
                         
                         'Do the redo
                         PerformRedo
                         
                         'Remove all vertexbuffers
                         ReDim SubSectorFloors(0 To numsubsectors - 1)
                         ReDim SubSectorCeilings(0 To numsubsectors - 1)
                         ReDim SidedefUpper(-1 To numsidedefs - 1)
                         ReDim SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim SidedefLower(-1 To numsidedefs - 1)
                         ReDim d_SubSectorFloors(0 To numsubsectors - 1)
                         ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
                         ReDim d_SidedefUpper(-1 To numsidedefs - 1)
                         ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim d_SidedefLower(-1 To numsidedefs - 1)
                         ReDim i_SectorFloors(0 To numsectors - 1)
                         ReDim i_SectorCeilings(0 To numsectors - 1)
                         ReDim i_SidedefUpper(-1 To numsidedefs - 1)
                         ReDim i_SidedefMiddle(-1 To numsidedefs - 1)
                         ReDim i_SidedefLower(-1 To numsidedefs - 1)
                         
                         'Give pointers to the DLL
                         SetStructurePointers m_vertices(0), linedefs(0), VarPtr(sidedefs(0)), m_segs(0), VarPtr(sectors(0)), m_subsectors(0), things(0), m_nodes(0), numnodes, numsectors, numsubsectors, numthings
                         
                         'Show undo message
                         ShowMainText UndoDescription & " redone"
                    Else
                         
                         'Show its not possible now
                         ShowMainText "Next change cannot be redone in 3D mode"
                    End If
               End If
               
          Case Config("shortcuts")("mode3draise")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING: LiftCeiling Obj, 8
                    Case OBJ_SECTORFLOOR: LiftFloor Obj, 8
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: If (Config("raiselowerceiling") = vbChecked) Then LiftCeiling sidedefs(Obj).sector, 8
                    Case OBJ_THING: LiftThing Obj, 8
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dlower")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING: LiftCeiling Obj, -8
                    Case OBJ_SECTORFLOOR: LiftFloor Obj, -8
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: If (Config("raiselowerceiling") = vbChecked) Then LiftCeiling sidedefs(Obj).sector, -8
                    Case OBJ_THING: LiftThing Obj, -8
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3draisefine")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING: LiftCeiling Obj, 1
                    Case OBJ_SECTORFLOOR: LiftFloor Obj, 1
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: If (Config("raiselowerceiling") = vbChecked) Then LiftCeiling sidedefs(Obj).sector, 1
                    Case OBJ_THING: LiftThing Obj, 1
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dlowerfine")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING: LiftCeiling Obj, -1
                    Case OBJ_SECTORFLOOR: LiftFloor Obj, -1
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: If (Config("raiselowerceiling") = vbChecked) Then LiftCeiling sidedefs(Obj).sector, -1
                    Case OBJ_THING: LiftThing Obj, -1
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dincbright")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING, OBJ_SECTORFLOOR: ChangeBrightness Obj, 16
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeBrightness sidedefs(Obj).sector, 16
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3ddecbright")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORCEILING, OBJ_SECTORFLOOR: ChangeBrightness Obj, -16
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeBrightness sidedefs(Obj).sector, -16
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dmiddle")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ToggleMiddleTexture Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3duunpeg")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ToggleUpperUnpegged Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dlunpeg")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ToggleLowerUnpegged Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexselect")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: SelectUpperTexture Obj
                    Case OBJ_SIDEDEFMIDDLE: SelectMiddleTexture Obj
                    Case OBJ_SIDEDEFLOWER: SelectLowerTexture Obj
                    Case OBJ_SECTORCEILING: SelectCeilingTexture Obj
                    Case OBJ_SECTORFLOOR: SelectFloorTexture Obj
                    Case OBJ_THING: SelectNewThing Obj
               End Select
               
          Case Config("shortcuts")("mode3dtexcopy")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: CopySidedefTexture sidedefs(Obj).upper
                    Case OBJ_SIDEDEFMIDDLE: CopySidedefTexture sidedefs(Obj).middle
                    Case OBJ_SIDEDEFLOWER: CopySidedefTexture sidedefs(Obj).lower
                    Case OBJ_SECTORCEILING: CopySectorFlat sectors(Obj).tceiling
                    Case OBJ_SECTORFLOOR: CopySectorFlat sectors(Obj).tfloor
                    Case OBJ_THING: CopyThing Obj
               End Select
               
          Case Config("shortcuts")("mode3dtexpaste")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: PasteUpperTexture Obj
                    Case OBJ_SIDEDEFMIDDLE: PasteMiddleTexture Obj
                    Case OBJ_SIDEDEFLOWER: PasteLowerTexture Obj
                    Case OBJ_SECTORCEILING: PasteCeilingTexture Obj
                    Case OBJ_SECTORFLOOR: PasteFloorTexture Obj
                    Case OBJ_THING: PasteThing Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignleft")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, 1, 0
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignright")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, -1, 0
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignup")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, 0, 1
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexaligndown")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, 0, -1
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignreset")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, -sidedefs(Obj).tx, -sidedefs(Obj).ty
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignresetx")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, -sidedefs(Obj).tx, 0
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexalignresety")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: ChangeTextureOffset Obj, 0, -sidedefs(Obj).ty
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dtexrem")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: DeleteUpperTexture Obj
                    Case OBJ_SIDEDEFMIDDLE: DeleteMiddleTexture Obj
                    Case OBJ_SIDEDEFLOWER: DeleteLowerTexture Obj
                    Case OBJ_THING: DeleteThing Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dpainttexture")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: DoFloodfillTextures Obj, sidedefs(Obj).upper
                    Case OBJ_SIDEDEFMIDDLE: DoFloodfillTextures Obj, sidedefs(Obj).middle
                    Case OBJ_SIDEDEFLOWER: DoFloodfillTextures Obj, sidedefs(Obj).lower
                    Case OBJ_SECTORCEILING: DoFloodfillFlats Obj, False
                    Case OBJ_SECTORFLOOR: DoFloodfillFlats Obj, True
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dautoalign")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: AutoAlignUpperTextures Obj, False
                    Case OBJ_SIDEDEFMIDDLE: AutoAlignMiddleTextures Obj, False
                    Case OBJ_SIDEDEFLOWER: AutoAlignLowerTextures Obj, False
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dautoaligny")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: AutoAlignUpperTextures Obj, True
                    Case OBJ_SIDEDEFMIDDLE: AutoAlignMiddleTextures Obj, True
                    Case OBJ_SIDEDEFLOWER: AutoAlignLowerTextures Obj, True
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("copyprops")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORFLOOR, OBJ_SECTORCEILING: CopySectorProperties Obj
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: CopySidedefProperties Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("pasteprops")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SECTORFLOOR, OBJ_SECTORCEILING: PasteSectorProperties Obj
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER: PasteSidedefProperties Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dinfopanel")
               
               'Check if already displayed
               If (ShowInfo) Then
                    
                    'Close it
                    ShowInfo = False
               Else
                    
                    'Update directly
                    UpdateInfoPanel
                    
                    'Show it
                    ShowInfo = True
               End If
               
          Case Config("shortcuts")("mode3dcopyoffsets")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: CopySidedefOffsets Obj
                    Case OBJ_SIDEDEFMIDDLE: CopySidedefOffsets Obj
                    Case OBJ_SIDEDEFLOWER: CopySidedefOffsets Obj
               End Select
               
          Case Config("shortcuts")("mode3dpasteoffsets")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Call function depending on object type
               Select Case ObjType
                    Case OBJ_SIDEDEFUPPER: PasteSidedefOffsets Obj
                    Case OBJ_SIDEDEFMIDDLE: PasteSidedefOffsets Obj
                    Case OBJ_SIDEDEFLOWER: PasteSidedefOffsets Obj
               End Select
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dthingstoggle")
               
               'Toggle things on and off
               If (ShowThings) Then
                    
                    'Hide things
                    ShowThings = 0
                    ShowMainText "Things:  Off"
               Else
                    
                    'Show things
                    ShowThings = 1
                    ShowMainText "Things:  On"
               End If
               
          Case Config("shortcuts")("mode3dthingheightreset")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Works only on things
               If (ObjType = OBJ_THING) Then
                    
                    'Change depending on hanging
                    If (things(Obj).hangs) Then LiftThing Obj, things(Obj).Z Else LiftThing Obj, -things(Obj).Z
               End If
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dthingrotatecw")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Works only on things
               If (ObjType = OBJ_THING) Then RotateThing Obj, -45
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dthingrotateccw")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Works only on things
               If (ObjType = OBJ_THING) Then RotateThing Obj, 45
               
               'Show changes immediately
               RunSingleFrame
               
          Case Config("shortcuts")("mode3dinsert")
               
               'Get the targeted object
               ObjType = PickAimedObject(Obj, ObjSpot)
               
               'Works on sectors only
               If (ObjType = OBJ_SECTORCEILING) Or (ObjType = OBJ_SECTORFLOOR) Then InsertThing ObjSpot
               
               'Show changes immediately
               RunSingleFrame
               
               
     End Select
     
     'Make sure info refreshes
     LastInfoObject = -2
End Sub

Public Sub KeydownTextureSelect(ByVal ShortcutCode As Long)
     Dim ci As Long
     Dim c As Long, r As Long
     Dim MousePoint As POINT
     Dim NewIndex As Long
     
     'Check if we should ignore input
     If IgnoreInput Then Exit Sub
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo"))) Then
          
          'Check if one of the 2D keys is used
          Select Case ShortcutCode
               Case Config("shortcuts")("togglebar"): frmMain.InfoBarToggle
          End Select
     End If
     
     'Check what key
     Select Case ShortcutCode
          
          Case vbKeyTab
               
               'Switch to all textures
               If (ShowAllTextures = False) Then
                    
                    'Now show all
                    ShowAllTextures = True
                    
                    'Change collections
                    curitemnames() = itemnames()
                    curnumitems = numitems
                    CreateTexturePreviews
                    
                    'Reselect texture
                    TextureSelectedIndex = -1
                    SetTextureSelection
                    NewIndex = TextureSelectedIndex
               Else
                    
                    'Leave now
                    Exit Sub
               End If
               
          Case vbKeyUp
               
               'Move selection up
               TextureSelectedIndex = TextureSelectedIndex - TEXTURE_COLS
               If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
               If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
               NewIndex = TextureSelectedIndex
               SelectedName = curitemnames(TextureSelectedIndex)
               CreateSelectedTextureText
          
          Case vbKeyDown
               
               'Move selection down
               TextureSelectedIndex = TextureSelectedIndex + TEXTURE_COLS
               If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
               If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
               NewIndex = TextureSelectedIndex
               SelectedName = curitemnames(TextureSelectedIndex)
               CreateSelectedTextureText
          
          Case vbKeyRight
               
               'Move selection right
               TextureSelectedIndex = TextureSelectedIndex + 1
               If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
               If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
               NewIndex = TextureSelectedIndex
               SelectedName = curitemnames(TextureSelectedIndex)
               CreateSelectedTextureText
          
          Case vbKeyLeft
               
               'Move selection left
               TextureSelectedIndex = TextureSelectedIndex - 1
               If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
               If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
               NewIndex = TextureSelectedIndex
               SelectedName = curitemnames(TextureSelectedIndex)
               CreateSelectedTextureText
               
          Case MOUSE_BUTTON_0
               
               'Check if in windowed mode
               If (Val(Config("windowedvideo"))) Then
                    
                    'Get mouse coords from form
                    MousePoint.x = frmMain.LastMouseX
                    MousePoint.y = frmMain.LastMouseY
               Else
                    
                    'Get mouse coords
                    GetCursorPos MousePoint
               End If
               
               'Calculate the row and col
               c = Int((MousePoint.x / VideoParams.BackBufferWidth) * TEXTURE_COLS)
               r = Int((MousePoint.y / (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT))) * TEXTURE_ROWS)
               
               'Change selection index
               TextureSelectedIndex = (TextureRowOffset + r) * TEXTURE_COLS + c
               If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
               If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
               
               'And apply immediately
               TextureSelectCancelled = False
               TextureSelecting = False
               
          Case MOUSE_SCROLL_UP, vbKeyPageUp
               
               'Move selection up by 4
               NewIndex = TextureRowOffset * TEXTURE_COLS - TEXTURE_COLS * 4
               If (NewIndex > (curnumitems - 1)) Then NewIndex = curnumitems - 1
               If (NewIndex < 0) Then NewIndex = 0
               
          Case MOUSE_SCROLL_DOWN, vbKeyPageDown
               
               'Move selection down by 4
               NewIndex = (TextureRowOffset + TEXTURE_ROWS - 1) * TEXTURE_COLS + TEXTURE_COLS * 4
               If (NewIndex > (curnumitems - 1)) Then NewIndex = curnumitems - 1
               If (NewIndex < 0) Then NewIndex = 0
               
          Case vbKeyReturn, vbKeySpace
               
               'Apply
               TextureSelectCancelled = False
               TextureSelecting = False
               
          Case vbKeyEscape, MOUSE_BUTTON_1
               
               'Cancel
               TextureSelectCancelled = True
               TextureSelecting = False
               
          Case Else
               
               'Leave now
               Exit Sub
               
     End Select
     
     'Check if a valid selection is made
     If (NewIndex >= 0) Then
          
          'Check if the selection is above view
          ci = NewIndex - TextureRowOffset * TEXTURE_COLS
          If (ci < 0) Then
               
               'Scroll to selection
               TextureRowOffset = NewIndex \ TEXTURE_COLS
               If (TextureRowOffset > (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)) Then TextureRowOffset = (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)
               If (TextureRowOffset < 0) Then TextureRowOffset = 0
               
               'Recreate previews
               CreateTexturePreviews
               
               'Ignore any more input for this frame
               'so the new textures will be rendered for sure
               IgnoreInput = True
               
          'Check if the selection is below view
          ElseIf (ci > TEXTURE_COLS * (TEXTURE_ROWS - 1)) Then
               
               'Scroll to selection
               TextureRowOffset = NewIndex \ TEXTURE_COLS - (TEXTURE_ROWS - 1)
               If (TextureRowOffset > (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)) Then TextureRowOffset = (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)
               If (TextureRowOffset < 0) Then TextureRowOffset = 0
               
               'Recreate previews
               CreateTexturePreviews
               
               'Ignore any more input for this frame
               'so the new textures will be rendered for sure
               IgnoreInput = True
          End If
     End If
End Sub

Public Sub KeypressTextureSelect(ByVal KeyAscii As Long)
     
     'Check if we should ignore input
     If IgnoreInput Then Exit Sub
     
     'Check if key is valid
     If (InStr(1, TEXTURE_CHARS, Chr$(KeyAscii), vbTextCompare) > 0) Or (KeyAscii = 8) Then
          
          'Show the cursor
          ShowTextCursor = True
          
          'Check if the texture name should be erased
          If (TextureEraseOnType) Then
               
               'Erase the name
               SelectedName = ""
               TextureEraseOnType = False
          End If
          
          'Check if a char should be removed
          If (KeyAscii = 8) Then
               
               'Remove last character if possible
               If (Len(SelectedName) > 0) Then SelectedName = left$(SelectedName, Len(SelectedName) - 1)
          Else
               
               'Add a character if possible
               If (Len(SelectedName) < 8) Then SelectedName = SelectedName & UCase$(Chr$(KeyAscii))
          End If
          
          'Remake the text
          CreateSelectedTextureText
          
          'Reflect selection
          SetTextureSelection
     End If
End Sub

Public Sub Keyrelease3D(ByVal ShortcutCode As Long)
     
     'Check what key
     Select Case ShortcutCode
          
          Case Config("shortcuts")("mode3dforward"): Key3DForward = False
          Case Config("shortcuts")("mode3dbackward"): Key3DBackward = False
          Case Config("shortcuts")("mode3dstrafeleft"): Key3DStrafeLeft = False
          Case Config("shortcuts")("mode3dstraferight"): Key3DStrafeRight = False
          Case Config("shortcuts")("mode3dstrafeup"): Key3DStrafeUp = False
          Case Config("shortcuts")("mode3dstrafedown"): Key3DStrafeDown = False
          
     End Select
End Sub

Private Sub LiftCeiling(ByVal sector As Long, ByVal Amount As Long)
     Dim ld As Long
     Dim ss As Long
     Dim RemoveThis As Long
     
     'Make undo
     CreateUndo "ceiling height change", UGRP_CEILINGHEIGHTCHANGE, sector, True
     
     'Move the sector ceiling
     sectors(sector).hceiling = sectors(sector).hceiling + Amount
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Show message
     ShowMainText "Ceiling height:  " & sectors(sector).hceiling
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Dont assume remove
          RemoveThis = False
          
          'Check side 1
          If (linedefs(ld).s1 > -1) Then
               
               'Check if sidedef refers to this sector
               If (sidedefs(linedefs(ld).s1).sector = sector) Then RemoveThis = True
          End If
          
          'Check side 2
          If (linedefs(ld).s2 > -1) Then
               
               'Check if sidedef refers to this sector
               If (sidedefs(linedefs(ld).s2).sector = sector) Then RemoveThis = True
          End If
          
          'Check if should be removed
          If RemoveThis Then
               
               'Remove vertexbuffers so they will be recreated
               d_SidedefLower(linedefs(ld).s1) = False
               d_SidedefMiddle(linedefs(ld).s1) = False
               d_SidedefUpper(linedefs(ld).s1) = False
               d_SidedefLower(linedefs(ld).s2) = False
               d_SidedefMiddle(linedefs(ld).s2) = False
               d_SidedefUpper(linedefs(ld).s2) = False
               Set SidedefLower(linedefs(ld).s1) = Nothing
               Set SidedefMiddle(linedefs(ld).s1) = Nothing
               Set SidedefUpper(linedefs(ld).s1) = Nothing
               Set SidedefLower(linedefs(ld).s2) = Nothing
               Set SidedefMiddle(linedefs(ld).s2) = Nothing
               Set SidedefUpper(linedefs(ld).s2) = Nothing
          End If
     Next ld
     
     'Go for all subsectors
     For ss = 0 To (numsubsectors - 1)
          
          'Check if subsector is part of this sector
          If (m_subsectors(ss).sector = sector) Then
               
               'Remove vertexbuffer so it will be recreated
               d_SubSectorCeilings(ss) = False
               Set SubSectorCeilings(ss) = Nothing
          End If
     Next ss
End Sub

Private Sub LiftFloor(ByVal sector As Long, ByVal Amount As Long)
     Dim sd As Long
     Dim ss As Long
     Dim RemoveThis As Long
     Dim ld As MAPLINEDEF
     
     'Make undo
     CreateUndo "floor height change", UGRP_FLOORHEIGHTCHANGE, sector, True
     
     'Move the sector ceiling
     sectors(sector).hfloor = sectors(sector).hfloor + Amount
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Show message
     ShowMainText "Floor height:  " & sectors(sector).hfloor
     
     'Go for all sidedefs
     For sd = 0 To (numsidedefs - 1)
          
          'Check if should be removed
          If (sidedefs(sd).sector = sector) Then
               
               'Remove vertexbuffers so they will be recreated
               ld = linedefs(sidedefs(sd).linedef)
               d_SidedefLower(ld.s1) = False
               d_SidedefMiddle(ld.s1) = False
               d_SidedefUpper(ld.s1) = False
               d_SidedefLower(ld.s2) = False
               d_SidedefMiddle(ld.s2) = False
               d_SidedefUpper(ld.s2) = False
               Set SidedefLower(ld.s1) = Nothing
               Set SidedefMiddle(ld.s1) = Nothing
               Set SidedefUpper(ld.s1) = Nothing
               Set SidedefLower(ld.s2) = Nothing
               Set SidedefMiddle(ld.s2) = Nothing
               Set SidedefUpper(ld.s2) = Nothing
          End If
     Next sd
     
     'Go for all subsectors
     For ss = 0 To (numsubsectors - 1)
          
          'Check if subsector is part of this sector
          If (m_subsectors(ss).sector = sector) Then
               
               'Remove vertexbuffer so it will be recreated
               d_SubSectorFloors(ss) = False
               Set SubSectorFloors(ss) = Nothing
          End If
     Next ss
End Sub

Private Sub LiftThing(ByVal thing As Long, ByVal Amount As Long)
     
     'Check if in hexen format
     If (mapconfig("mapformat") = 2) Then
          
          'Make undo
          CreateUndo "thing height change", UGRP_THINGHEIGHTCHANGE, thing, True
          
          'Check if hanging from ceiling
          If (things(thing).hangs) Then
               
               'Move the thing backwards
               things(thing).Z = things(thing).Z - Amount
          Else
               
               'Move the thing
               things(thing).Z = things(thing).Z + Amount
          End If
          
          'Show message
          ShowMainText "Thing height:  " & things(thing).Z
          
          'Map changed
          mapchanged = True
     End If
End Sub

Private Sub RotateThing(ByVal thing As Long, ByVal Amount As Long)
     
     'Make undo
     CreateUndo "rotate thing ", UGRP_THINGANGLECHANGE, thing, True
     
     'Rotate the thing angle
     things(thing).angle = things(thing).angle + Amount
     If (things(thing).angle >= 360) Then things(thing).angle = things(thing).angle - 360
     If (things(thing).angle < 0) Then things(thing).angle = things(thing).angle + 360
     
     'Update thing
     UpdateThingImageColor thing
     
     'Show message
     ShowMainText "Thing angle:  " & things(thing).angle & " (" & GetThingAngleDesc(things(thing).angle) & ")"
     
     'Map changed
     mapchanged = True
End Sub



Private Function LinedefLength(ByVal ld As Long) As Long
     Dim xl As Long, yl As Long
     
     'Calculate linedef length
     xl = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
     yl = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
     LinedefLength = CLng(Sqr(xl * xl + yl * yl))
End Function

Private Function LinedefFrontHeight(ByVal ld As Long) As String
     
     If (linedefs(ld).s1 > -1) Then
          LinedefFrontHeight = CStr(sectors(sidedefs(linedefs(ld).s1).sector).hceiling - sectors(sidedefs(linedefs(ld).s1).sector).hfloor)
     Else
          LinedefFrontHeight = "-"
     End If
End Function

Private Function LinedefBackHeight(ByVal ld As Long) As String
     
     If (linedefs(ld).s2 > -1) Then
          LinedefBackHeight = CStr(sectors(sidedefs(linedefs(ld).s2).sector).hceiling - sectors(sidedefs(linedefs(ld).s2).sector).hfloor)
     Else
          LinedefBackHeight = "-"
     End If
End Function


Private Sub MakeThingResources()
     Dim tv(0 To 35) As VERTEX
     Dim TextureFile As String
     Dim x0 As Single, x1 As Single
     Dim y0 As Single, y1 As Single
     Dim z0 As Single, z1 As Single
     Dim u0 As Single, u1 As Single
     Dim v0 As Single, v1 As Single
     
     'Create coordinates
     x0 = -0.5 * MAP_RENDER_SCALE
     x1 = 0.5 * MAP_RENDER_SCALE
     y0 = -0.5 * MAP_RENDER_SCALE
     y1 = 0.5 * MAP_RENDER_SCALE
     z0 = 0 * MAP_RENDER_SCALE
     z1 = 1 * MAP_RENDER_SCALE
     u0 = 0
     u1 = 1 - 1 / 64
     v0 = 0
     v1 = 1 - 1 / 64
     
     'Create vertices for box
     
     'Front
     With tv(0): .x = x0: .y = y0: .Z = z0: .tu = u0: .tv = v0: End With
     With tv(1): .x = x0: .y = y0: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(2): .x = x1: .y = y0: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(3): .x = x1: .y = y0: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(4): .x = x0: .y = y0: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(5): .x = x1: .y = y0: .Z = z1: .tu = u1: .tv = v1: End With
     
     'Right
     With tv(6): .x = x1: .y = y0: .Z = z0: .tu = u0: .tv = v0: End With
     With tv(7): .x = x1: .y = y0: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(8): .x = x1: .y = y1: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(9): .x = x1: .y = y1: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(10): .x = x1: .y = y0: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(11): .x = x1: .y = y1: .Z = z1: .tu = u1: .tv = v1: End With
     
     'Back
     With tv(12): .x = x1: .y = y1: .Z = z0: .tu = u0: .tv = v0: End With
     With tv(13): .x = x1: .y = y1: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(14): .x = x0: .y = y1: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(15): .x = x0: .y = y1: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(16): .x = x1: .y = y1: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(17): .x = x0: .y = y1: .Z = z1: .tu = u1: .tv = v1: End With
     
     'Left
     With tv(18): .x = x0: .y = y1: .Z = z0: .tu = u0: .tv = v0: End With
     With tv(19): .x = x0: .y = y1: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(20): .x = x0: .y = y0: .Z = z1: .tu = u1: .tv = v0: End With
     With tv(21): .x = x0: .y = y1: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(22): .x = x0: .y = y0: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(23): .x = x0: .y = y0: .Z = z0: .tu = u1: .tv = v1: End With
     
     'Top
     With tv(24): .x = x0: .y = y0: .Z = z1: .tu = u0: .tv = v0: End With
     With tv(25): .x = x0: .y = y1: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(26): .x = x1: .y = y0: .Z = z1: .tu = u1: .tv = v0: End With
     With tv(27): .x = x1: .y = y0: .Z = z1: .tu = u1: .tv = v0: End With
     With tv(28): .x = x0: .y = y1: .Z = z1: .tu = u0: .tv = v1: End With
     With tv(29): .x = x1: .y = y1: .Z = z1: .tu = u1: .tv = v1: End With
     
     'Bottom
     With tv(30): .x = x1: .y = y0: .Z = z0: .tu = u0: .tv = v0: End With
     With tv(31): .x = x0: .y = y1: .Z = z0: .tu = u0: .tv = v1: End With
     With tv(32): .x = x0: .y = y0: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(33): .x = x1: .y = y0: .Z = z0: .tu = u1: .tv = v0: End With
     With tv(34): .x = x1: .y = y1: .Z = z0: .tu = u0: .tv = v1: End With
     With tv(35): .x = x0: .y = y1: .Z = z0: .tu = u1: .tv = v1: End With
     
     'Create vertex buffer
     Set r_thingboxvb = CreateVertexBuffer(tv(), 36)
     
     'Create vertices for box lines
     
     'Top
     With tv(0): .x = x0: .y = y0: .Z = z1: End With
     With tv(1): .x = x1: .y = y0: .Z = z1: End With
     With tv(2): .x = x1: .y = y0: .Z = z1: End With
     With tv(3): .x = x1: .y = y1: .Z = z1: End With
     With tv(4): .x = x1: .y = y1: .Z = z1: End With
     With tv(5): .x = x0: .y = y1: .Z = z1: End With
     With tv(6): .x = x0: .y = y1: .Z = z1: End With
     With tv(7): .x = x0: .y = y0: .Z = z1: End With
     
     'Bottom
     With tv(8): .x = x0: .y = y0: .Z = z0: End With
     With tv(9): .x = x1: .y = y0: .Z = z0: End With
     With tv(10): .x = x1: .y = y0: .Z = z0: End With
     With tv(11): .x = x1: .y = y1: .Z = z0: End With
     With tv(12): .x = x1: .y = y1: .Z = z0: End With
     With tv(13): .x = x0: .y = y1: .Z = z0: End With
     With tv(14): .x = x0: .y = y1: .Z = z0: End With
     With tv(15): .x = x0: .y = y0: .Z = z0: End With
     
     'Spokes
     With tv(16): .x = x0: .y = y0: .Z = z1: End With
     With tv(17): .x = x0: .y = y0: .Z = z0: End With
     With tv(18): .x = x1: .y = y0: .Z = z1: End With
     With tv(19): .x = x1: .y = y0: .Z = z0: End With
     With tv(20): .x = x1: .y = y1: .Z = z1: End With
     With tv(21): .x = x1: .y = y1: .Z = z0: End With
     With tv(22): .x = x0: .y = y1: .Z = z1: End With
     With tv(23): .x = x0: .y = y1: .Z = z0: End With
     
     'Create vertex buffer
     Set r_thingboxlines = CreateVertexBuffer(tv(), 24)
     
     'Create vertices for sprite
     tv(0).x = -0.5 * MAP_RENDER_SCALE:     tv(0).y = 0:    tv(0).Z = 1 * MAP_RENDER_SCALE: tv(0).tu = 0: tv(0).tv = 0
     tv(1).x = -0.5 * MAP_RENDER_SCALE:     tv(1).y = 0:    tv(1).Z = 0 * MAP_RENDER_SCALE: tv(1).tu = 0: tv(1).tv = 1
     tv(2).x = 0.5 * MAP_RENDER_SCALE:      tv(2).y = 0:    tv(2).Z = 1 * MAP_RENDER_SCALE: tv(2).tu = 1: tv(2).tv = 0
     tv(3).x = 0.5 * MAP_RENDER_SCALE:      tv(3).y = 0:    tv(3).Z = 0 * MAP_RENDER_SCALE: tv(3).tu = 1: tv(3).tv = 1
     
     'Create vertex buffer
     Set r_thingsprite = CreateVertexBuffer(tv(), 4)
     
     'Create vertices for arrow tile
     tv(0).x = -0.5 * MAP_RENDER_SCALE:     tv(0).Z = 0:    tv(0).y = -0.5 * MAP_RENDER_SCALE: tv(0).tu = u0: tv(0).tv = v1
     tv(1).x = -0.5 * MAP_RENDER_SCALE:     tv(1).Z = 0:    tv(1).y = 0.5 * MAP_RENDER_SCALE: tv(1).tu = u0: tv(1).tv = v0
     tv(2).x = 0.5 * MAP_RENDER_SCALE:      tv(2).Z = 0:    tv(2).y = -0.5 * MAP_RENDER_SCALE: tv(2).tu = u1: tv(2).tv = v1
     tv(3).x = 0.5 * MAP_RENDER_SCALE:      tv(3).Z = 0:    tv(3).y = 0.5 * MAP_RENDER_SCALE: tv(3).tu = u1: tv(3).tv = v0
     
     'Create vertex buffer
     Set r_thingarrow = CreateVertexBuffer(tv(), 4)
     
     'Load thingbox texture
     TextureFile = App.Path & "\Thingbox.tga"
     Set tex_thingbox = CreateTextureFromFileEx(D3DD, TextureFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  0, ByVal 0, ByVal 0)
     
     'Load thingarrow texture
     TextureFile = App.Path & "\Thingarrow.tga"
     Set tex_thingarrow = CreateTextureFromFileEx(D3DD, TextureFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  0, ByVal 0, ByVal 0)
End Sub

Public Sub RunSingleLoop()
     
     'Calculate time
     CurrentTime = timeExactTime
     FrameTime = CurrentTime - LastTime
     LastTime = CurrentTime
     
     'Check if we should remove the main text
     If (TextRemoveTime < GetTickCount) Then
          
          'Clear main and sub text
          Set r_maintext = Nothing
          Set r_subtext = Nothing
     End If
     
     'Poll the mouse
     PollMouse
     
     'Mouse events can do anything, also terminating 3d mode
     If Not Running3D Then Exit Sub
     
     'Apply Physics
     ApplyPhysics
     
     'Check if we should update the info panel
     If (InfoUpdateTime < CurrentTime) Then
          
          'Update info
          If (ShowInfo = True) Then UpdateInfoPanel
          
          'Update line info
          UpdateLinesSectorsInfo
          
          'Change update time
          InfoUpdateTime = CurrentTime + INFO_UPDATEDELAY
     End If
     
     
     'Run a single frame
     RunSingleFrame
     
     
     'Delay frames?
     If (DelayVideoFrames) Then Sleep 50
     
End Sub

Private Sub SetTextureFilters(ByVal ForceBilinear As Boolean)
     
     'Check if bilinear is forced
     If (ForceBilinear = True) Then
          
          'Set up bilinear texture filtering
          D3DD.SetSamplerState 0, D3DSAMP_MIPFILTER, D3DTEXF_NONE
          D3DD.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
          D3DD.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
     Else
          
          'Check if trilinear is configured
          If (Val(Config("texturefilter")) = TF_LINEAR_MIPMAP_LINEAR) Then
               
               'Set up bilinear texture filtering
               D3DD.SetSamplerState 0, D3DSAMP_MIPFILTER, D3DTEXF_LINEAR
               D3DD.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
               D3DD.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
               
          'Check if bilinear is configured
          ElseIf (Val(Config("texturefilter")) = TF_LINEAR_MIPMAP_NEAREST) Then
               
               'Set up bilinear texture filtering
               D3DD.SetSamplerState 0, D3DSAMP_MIPFILTER, D3DTEXF_NONE
               D3DD.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_LINEAR
               D3DD.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_LINEAR
               
          'Dont use texture filtering
          Else
               
               'Disable texture filtering
               D3DD.SetSamplerState 0, D3DSAMP_MIPFILTER, D3DTEXF_NONE
               D3DD.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_POINT
               D3DD.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_POINT
               D3DD.SetSamplerState 0, D3DSAMP_MINFILTER, D3DTEXF_NONE
               D3DD.SetSamplerState 0, D3DSAMP_MAGFILTER, D3DTEXF_NONE
          End If
     End If
End Sub

Private Sub SetTextureSelection()
     Dim ci As Long
     Dim CurrentName As String
     Dim OldRowOffset As Long
          
     'Get the current selected texture name
     If (TextureSelectedIndex >= 0) Then CurrentName = curitemnames(TextureSelectedIndex)
     
     'Check if anything typed at all
     If Len(SelectedName) Then
          
          'Check if this name no longer matches the selection
          If (StrComp(SelectedName, left$(CurrentName, Len(SelectedName)), vbTextCompare) <> 0) Or (Val(Config("autocompletetex")) = 0) Then
               
               'Keep previous row offset
               OldRowOffset = TextureRowOffset
               
               'When nothing will be found, default to nothing
               TextureSelectedIndex = -1
               
               'Check if all of the typed name must match
               If (Val(Config("autocompletetex")) = 0) Then
                    
                    'Find the first that exactly matches
                    For ci = 0 To (curnumitems - 1)
                         
                         'Check if it matches
                         If (StrComp(SelectedName, curitemnames(ci), vbTextCompare) = 0) Then
                              
                              'Go here
                              TextureSelectedIndex = ci
                              Exit For
                         End If
                    Next ci
               Else
                    
                    'Find the first that partly matches
                    For ci = 0 To (curnumitems - 1)
                         
                         'Check if it matches
                         If (StrComp(SelectedName, left$(curitemnames(ci), Len(SelectedName)), vbTextCompare) = 0) Then
                              
                              'Go here
                              TextureSelectedIndex = ci
                              Exit For
                         End If
                    Next ci
               End If
               
               'Check if anything found
               If (TextureSelectedIndex >= 0) Then
                    
                    'Scroll to the selection
                    TextureRowOffset = (TextureSelectedIndex \ TEXTURE_COLS) - 2
                    
                    'Limit the scroll
                    If (TextureRowOffset > (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)) Then TextureRowOffset = (curnumitems \ TEXTURE_COLS + 1 - TEXTURE_ROWS)
                    If (TextureRowOffset < 0) Then TextureRowOffset = 0
                    
                    'Check if scrolled
                    If (TextureRowOffset <> OldRowOffset) Then
                         
                         'Recreate previews
                         CreateTexturePreviews
                    End If
               End If
          End If
     Else
          
          'Select nothing
          TextureSelectedIndex = -1
     End If
End Sub

Private Function SidedefUpperHeight(ByVal sd As Long) As String
     Dim sc As Long      'Sector
     Dim osd As Long     'Sidedef on other side of line
     Dim osc As Long     'Sector on other side of line
     
     'Get references
     sc = sidedefs(sd).sector
     osd = linedefs(sidedefs(sd).linedef).s2
     If (osd > -1) Then osc = sidedefs(osd).sector
     
     'There must be another side
     If (osd > -1) Then
          
          'Other sector must have lower ceiling
          If (sectors(osc).hceiling < sectors(sc).hceiling) Then
               SidedefUpperHeight = CStr(sectors(sc).hceiling - sectors(osc).hceiling)
          Else
               SidedefUpperHeight = "-"
          End If
     Else
          SidedefUpperHeight = "-"
     End If
End Function

Private Function SidedefLowerHeight(ByVal sd As Long) As String
     Dim sc As Long      'Sector
     Dim osd As Long     'Sidedef on other side of line
     Dim osc As Long     'Sector on other side of line
     
     'Get references
     sc = sidedefs(sd).sector
     osd = linedefs(sidedefs(sd).linedef).s2
     If (osd > -1) Then osc = sidedefs(osd).sector
     
     'There must be another side
     If (osd > -1) Then
          
          'Other sector must have higher floor
          If (sectors(osc).hfloor > sectors(sc).hfloor) Then
               SidedefLowerHeight = CStr(sectors(osc).hfloor - sectors(sc).hfloor)
          Else
               SidedefLowerHeight = "-"
          End If
     Else
          SidedefLowerHeight = "-"
     End If
End Function

Private Function SidedefMiddleHeight(ByVal sd As Long) As String
     Dim sc As Long      'Sector
     Dim osd As Long     'Sidedef on other side of line
     Dim osc As Long     'Sector on other side of line
     Dim lc As Long
     Dim hf As Long
     
     'Get references
     sc = sidedefs(sd).sector
     osd = linedefs(sidedefs(sd).linedef).s2
     If (osd > -1) Then osc = sidedefs(osd).sector
     
     'Check for another side
     If (osd > -1) Then
          
          'Check if ceiling or floor crosses
          If (sectors(osc).hfloor > sectors(sc).hceiling) Or (sectors(sc).hfloor > sectors(osc).hceiling) Then
               SidedefMiddleHeight = "-"
          Else
               'Get lowest ceiling and highest floor
               If (sectors(osc).hceiling < sectors(sc).hceiling) Then lc = sectors(osc).hceiling Else lc = sectors(sc).hceiling
               If (sectors(osc).hfloor > sectors(sc).hfloor) Then hf = sectors(osc).hfloor Else hf = sectors(sc).hfloor
               SidedefMiddleHeight = CStr(lc - hf)
          End If
     Else
          'Total height
          SidedefMiddleHeight = CStr(sectors(sc).hceiling - sectors(sc).hfloor)
     End If
End Function

Public Sub LoadFontFile(ByRef Filename As String)
     Dim FontFilebuffer As Integer
     Dim Chars As Long
     Dim FontChar As CHARRECTYPE
     Dim i As Long
     
     'Open the Font file
     FontFilebuffer = FreeFile
     Open Filename For Binary As #FontFilebuffer Len = Len(FontChar)
     
     'Read the number of characters
     Get #FontFilebuffer, 1, Chars
     
     'Go for all characters
     For i = 1 To Chars
          
          'Read character
          Get #FontFilebuffer, Len(FontChar) * i + 1, FontChar
          
          'Add the character to database
          SetFontChar Chr$(FontChar.char), CSng(FontChar.width) / 1600, CSng(FontChar.height) / 1200, _
                      FontChar.u1, FontChar.u2, FontChar.v1, FontChar.v2
     Next i
     
     'Close the Font file
     Close #FontFilebuffer
End Sub

Private Sub LoadNodes(ByVal FileBuffer As Integer, ByVal LumpAddress As Long, ByVal Count As Long, ByRef m_nodes() As MAPNODE)
     Dim ShortValue As Integer
     Dim i As Long
     
     'Allocate memory for ssectors
     ReDim m_nodes(0 To Count - 1)
     
     'Go for all ssectors to load
     Seek #FileBuffer, LumpAddress + 1
     For i = 0 To (Count - 1)
          Get #FileBuffer, , ShortValue: m_nodes(i).x = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).y = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).DX = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).dy = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).rtop = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).rbottom = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).rleft = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).rright = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).ltop = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).lbottom = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).lleft = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).lright = ShortValue
          Get #FileBuffer, , ShortValue: m_nodes(i).right = ItoL(ShortValue)
          Get #FileBuffer, , ShortValue: m_nodes(i).left = ItoL(ShortValue)
     Next i
End Sub

Private Sub LoadSegs(ByVal FileBuffer As Integer, ByVal LumpAddress As Long, ByVal Count As Long)
     Dim ShortValue As Integer
     Dim i As Long
     
     'Allocate memory for segs
     ReDim m_segs(0 To Count - 1)
     
     'Go for all segs to load
     Seek #FileBuffer, LumpAddress + 1
     For i = 0 To (Count - 1)
          Get #FileBuffer, , ShortValue: m_segs(i).v1 = ItoL(ShortValue)
          Get #FileBuffer, , ShortValue: m_segs(i).v2 = ItoL(ShortValue)
          Get #FileBuffer, , ShortValue: m_segs(i).angle = ShortValue
          Get #FileBuffer, , ShortValue: m_segs(i).linedef = ItoL(ShortValue)
          Get #FileBuffer, , ShortValue: m_segs(i).side = ShortValue
          Get #FileBuffer, , ShortValue: m_segs(i).offset = ShortValue
     Next i
End Sub

Private Sub LoadSSectors(ByVal FileBuffer As Integer, ByVal LumpAddress As Long, ByVal Count As Long)
     Dim ShortValue As Integer
     Dim i As Long
     
     'Allocate memory for ssectors
     ReDim m_subsectors(0 To Count - 1)
     
     'Go for all ssectors to load
     Seek #FileBuffer, LumpAddress + 1
     For i = 0 To (Count - 1)
          Get #FileBuffer, , ShortValue: m_subsectors(i).numsegs = ItoL(ShortValue)
          Get #FileBuffer, , ShortValue: m_subsectors(i).startseg = ItoL(ShortValue)
     Next i
End Sub

Private Sub LoadVertices(ByVal FileBuffer As Integer, ByVal LumpAddress As Long, ByVal Count As Long)
     Dim ShortValue As Integer
     Dim i As Long
     
     'Allocate memory for vertexes
     ReDim m_vertices(0 To Count - 1)
     
     'Go for all vertexes to load
     Seek #FileBuffer, LumpAddress + 1
     For i = 0 To (Count - 1)
          Get #FileBuffer, , ShortValue: m_vertices(i).x = ShortValue
          Get #FileBuffer, , ShortValue: m_vertices(i).y = ShortValue
     Next i
End Sub

Private Sub MakeCrosshair()
     Dim CrosshairPoly(0 To 3) As TLVERTEX
     Dim CrosshairFile As String
     Dim CrosshairWidth As Long
     Dim CrosshairHeight As Long
     Dim CrosshairFileInfo As D3DXIMAGE_INFO
     Dim BUFFERSIZE As Long
     
     'Make filename
     CrosshairFile = App.Path & "\Crosshair.bmp"
     
     'Create Direct3D Texture from file
     Set tex_crosshair = CreateTextureFromFileEx(D3DD, CrosshairFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  &HFF000000, VarPtr(CrosshairFileInfo), ByVal 0)
     
     'Determine crosshair size
     CrosshairInfo = CrosshairFileInfo
     CrosshairWidth = (VideoParams.BackBufferWidth / 25) * (CSng(CrosshairInfo.width) / 32)
     CrosshairHeight = (VideoParams.BackBufferWidth / 25) * (CSng(CrosshairInfo.height) / 32)
     
     'Create Polgon
     With CrosshairPoly(0)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = CLng((VideoParams.BackBufferWidth - CrosshairWidth) / 2)
          .sy = CLng((VideoParams.BackBufferHeight - CrosshairHeight) / 2)
          .tu = 0
          .tv = 0
     End With
     
     With CrosshairPoly(1)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = CLng((VideoParams.BackBufferWidth - CrosshairWidth) / 2)
          .sy = CLng((VideoParams.BackBufferHeight - CrosshairHeight) / 2) + CrosshairHeight
          .tu = 0
          .tv = 1
     End With
     
     With CrosshairPoly(2)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = CLng((VideoParams.BackBufferWidth - CrosshairWidth) / 2) + CrosshairWidth
          .sy = CLng((VideoParams.BackBufferHeight - CrosshairHeight) / 2)
          .tu = 1
          .tv = 0
     End With
     
     With CrosshairPoly(3)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = CLng((VideoParams.BackBufferWidth - CrosshairWidth) / 2) + CrosshairWidth
          .sy = CLng((VideoParams.BackBufferHeight - CrosshairHeight) / 2) + CrosshairHeight
          .tu = 1
          .tv = 1
     End With
     
     'Calculate buffer size in bytes
     BUFFERSIZE = TLVERTEXSTRIDE * 4
     
     'Create the vertex buffer
     Set r_crosshair = D3DD.CreateVertexBuffer(BUFFERSIZE, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, TLVERTEXFVF, D3DPOOL_DEFAULT)
     
     'Copy the vertices to the buffer
     r_crosshair.SetData 0, BUFFERSIZE, VarPtr(CrosshairPoly(0)), 0
End Sub

Private Sub MakeExtraTextures()
     Dim TextureFile As String
     
     'Make filename
     TextureFile = App.Path & "\Unknown.bmp"
     
     'Create Direct3D Texture from file
     Set tex_unknown = CreateTextureFromFileEx(D3DD, TextureFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  &HFF000000, ByVal 0, ByVal 0)
     
     'Make filename
     TextureFile = App.Path & "\Missing.bmp"
     
     'Create Direct3D Texture from file
     Set tex_missing = CreateTextureFromFileEx(D3DD, TextureFile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  &HFF000000, ByVal 0, ByVal 0)
End Sub

Private Sub MakeInfoPanel()
     Const PanelLeft As Single = 0.02
     Const PanelTop As Single = 0.6
     Const PanelRight As Single = 0.98
     Const PanelBottom As Single = 0.98
     Dim PanelPoly(0 To 3) As TLVERTEX
     Dim BUFFERSIZE As Long
     
     'Create Polgon
     With PanelPoly(0)
          .Color = D3DColorMake(0, 0, 0, 0.6)
          .rhw = 1
          .sx = VideoParams.BackBufferWidth * PanelLeft
          .sy = VideoParams.BackBufferHeight * PanelTop
          .tu = 0
          .tv = 0
     End With
     
     With PanelPoly(1)
          .Color = D3DColorMake(0, 0, 0, 0.6)
          .rhw = 1
          .sx = VideoParams.BackBufferWidth * PanelLeft
          .sy = VideoParams.BackBufferHeight * PanelBottom
          .tu = 0
          .tv = 0
     End With
     
     With PanelPoly(2)
          .Color = D3DColorMake(0, 0, 0, 0.6)
          .rhw = 1
          .sx = VideoParams.BackBufferWidth * PanelRight
          .sy = VideoParams.BackBufferHeight * PanelTop
          .tu = 0
          .tv = 0
     End With
     
     With PanelPoly(3)
          .Color = D3DColorMake(0, 0, 0, 0.6)
          .rhw = 1
          .sx = VideoParams.BackBufferWidth * PanelRight
          .sy = VideoParams.BackBufferHeight * PanelBottom
          .tu = 0
          .tv = 0
     End With
     
     'Calculate buffer size in bytes
     BUFFERSIZE = TLVERTEXSTRIDE * 4
     
     'Create the vertex buffer
     Set r_infopanel = D3DD.CreateVertexBuffer(BUFFERSIZE, D3DUSAGE_DYNAMIC Or D3DUSAGE_WRITEONLY, TLVERTEXFVF, D3DPOOL_DEFAULT)
     
     'Copy the vertices to the buffer
     r_infopanel.SetData 0, BUFFERSIZE, VarPtr(PanelPoly(0)), 0
End Sub

Private Sub MakeLightingTables()
     Dim i As Long
     Dim b As Single
     Dim f As Single
     
     'Go for all light levels
     For i = 0 To 255
          
          'Adjust the light so that it represents the doom light better
          b = i * i * 0.4 * 0.011 + 20
          
          'Convert 0-255 scale to 0-1 scale
          b = b * 0.00392
          
          'Limit the light
          If (b > 1) Then b = 1
          If (b < 0) Then b = 0
          
          'Calculate the fog from brightness
          f = (1 - b * b * 2) * 0.2
          
          'Limit the fog
          If (f < 0) Then f = 0
          
          'Set the table entries
          t_brightness(i) = D3DColorMake(b, b, b, 1)
          t_fogness(i) = CVL(MKS(f))
     Next i
End Sub

Private Sub MakeTextFont()
     Dim FontTexture As String
     Dim FontData As String
     
     'Make filenames
     FontTexture = App.Path & "\Font.tga"
     FontData = App.Path & "\Font.fnt"
     
     'Create Direct3D Texture from file
     Set tex_font = CreateTextureFromFileEx(D3DD, FontTexture, D3DX_DEFAULT, D3DX_DEFAULT, _
                                                  D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                  D3DPOOL_MANAGED, D3DX_DEFAULT, _
                                                  D3DX_FILTER_LINEAR Or D3DX_FILTER_DITHER, _
                                                  0, ByVal 0, ByVal 0)
     'Load the font data
     LoadFontFile FontData
End Sub

Private Sub MakeVertexBuffers()
     Dim ssi As Long
     Dim ss As Long
     Dim sd As Long
     Dim polygon() As VERTEX
     Dim sdpolygon(0 To 3) As VERTEX
     
     'Go for all visible subsectors
     For ssi = 0 To (r_numsubsectors - 1)
          
          'Get subsector index
          ss = r_subsectors(ssi)
          
          'Check if no Floor already created
          If (d_SubSectorFloors(ss) = False) Then
               
               'Check if we can create it
               If (m_subsectors(ss).numvertices > 2) And _
                  (m_subsectors(ss).sector > -1) Then
                    
                    'Create polygon vertices
                    CreateSubSectorFloor ss, polygon()
                    
                    'Create vertexbuffer
                    Set SubSectorFloors(ss) = CreateVertexBuffer(polygon(), m_subsectors(ss).numvertices)
                    
                    'Create texture reference
                    If (i_SectorFloors(m_subsectors(ss).sector) Is Nothing) Then
                         If allflats.Exists(sectors(m_subsectors(ss).sector).tfloor) Then
                              Set i_SectorFloors(m_subsectors(ss).sector) = allflats(sectors(m_subsectors(ss).sector).tfloor).D3DTexture
                         Else
                              Set i_SectorFloors(m_subsectors(ss).sector) = tex_unknown
                         End If
                    End If
               End If
               
               'Created
               d_SubSectorFloors(ss) = True
          End If
          
          'Check if no Ceiling already created
          If (d_SubSectorCeilings(ss) = False) Then
               
               'Check if we can create it
               If (m_subsectors(ss).numvertices > 2) And _
                  (m_subsectors(ss).sector > -1) Then
                    
                    'Create polygon vertices
                    CreateSubSectorCeiling ss, polygon()
                    
                    'Create vertexbuffer
                    Set SubSectorCeilings(ss) = CreateVertexBuffer(polygon(), m_subsectors(ss).numvertices)
                    
                    'Create texture reference
                    If (i_SectorCeilings(m_subsectors(ss).sector) Is Nothing) Then
                         If allflats.Exists(sectors(m_subsectors(ss).sector).tceiling) Then
                              Set i_SectorCeilings(m_subsectors(ss).sector) = allflats(sectors(m_subsectors(ss).sector).tceiling).D3DTexture
                         Else
                              Set i_SectorCeilings(m_subsectors(ss).sector) = tex_unknown
                         End If
                    End If
               End If
               
               'Created
               d_SubSectorCeilings(ss) = True
          End If
     Next ssi
     
     'Go for all visible sidedefs
     For ssi = 0 To (r_numsidedefs - 1)
          
          'Get the sidedef index
          sd = r_sidedefs(ssi)
          
          'Check if not already created
          If (d_SidedefUpper(sd) = False) Then
               
               'Make polygon for upper
               If CreateSidedefUpper(sd, sdpolygon()) Then
                    
                    'Create vertex buffer
                    Set SidedefUpper(sd) = CreateVertexBuffer(sdpolygon(), 4)
               End If
               
               'Create texture reference
               If (i_SidedefUpper(sd) Is Nothing) Then
                    If IsTextureName(sidedefs(sd).upper) Then
                         If alltextures.Exists(sidedefs(sd).upper) Then
                              Set i_SidedefUpper(sd) = alltextures(sidedefs(sd).upper).D3DTexture
                         Else
                              Set i_SidedefUpper(sd) = tex_unknown
                         End If
                    Else
                         Set i_SidedefUpper(sd) = tex_missing
                    End If
               End If
               
               'Created
               d_SidedefUpper(sd) = True
          End If
          
          'Check if not already created
          If (d_SidedefMiddle(sd) = False) Then
               
               'Make polygon for middle
               If CreateSidedefMiddle(sd, sdpolygon()) Then
                    
                    'Create vertex buffer
                    Set SidedefMiddle(sd) = CreateVertexBuffer(sdpolygon(), 4)
               End If
               
               'Create texture reference
               If (i_SidedefMiddle(sd) Is Nothing) Then
                    If IsTextureName(sidedefs(sd).middle) Then
                         If alltextures.Exists(sidedefs(sd).middle) Then
                              Set i_SidedefMiddle(sd) = alltextures(sidedefs(sd).middle).D3DTexture
                         Else
                              Set i_SidedefMiddle(sd) = tex_unknown
                         End If
                    Else
                         Set i_SidedefMiddle(sd) = tex_missing
                    End If
               End If
               
               'Created
               d_SidedefMiddle(sd) = True
          End If
          
          'Check if not already created
          If (d_SidedefLower(sd) = False) Then
               
               'Make polygon for lower
               If CreateSidedefLower(sd, sdpolygon()) Then
                    
                    'Create vertex buffer
                    Set SidedefLower(sd) = CreateVertexBuffer(sdpolygon(), 4)
               End If
               
               'Create texture reference
               If (i_SidedefLower(sd) Is Nothing) Then
                    If IsTextureName(sidedefs(sd).lower) Then
                         If alltextures.Exists(sidedefs(sd).lower) Then
                              Set i_SidedefLower(sd) = alltextures(sidedefs(sd).lower).D3DTexture
                         Else
                              Set i_SidedefLower(sd) = tex_unknown
                         End If
                    Else
                         Set i_SidedefLower(sd) = tex_missing
                    End If
               End If
               
               'Created
               d_SidedefLower(sd) = True
          End If
     Next ssi
End Sub

Private Sub PasteCeilingTexture(ByVal s As Long)
     Dim Texture As clsImage
     Dim ss As Long
     
     'Anything to paste?
     If (Len(Trim$(CopiedFlat)) > 0) Then
          
          'Make undo
          CreateUndo "change ceiling texture", UGRP_CEILINGTEXTURECHANGE, s, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Make it so
          sectors(s).tceiling = CopiedFlat
          
          'Check if texture is known
          If allflats.Exists(sectors(s).tceiling) Then
               
               'Get texture object
               Set Texture = allflats(sectors(s).tceiling)
               
               'Show message
               ShowMainText "Ceiling texture pasted:  " & Texture.Name & "  " & Texture.width & "x" & Texture.height
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'Show message
               ShowMainText "Ceiling texture pasted"
          End If
          
          'Go for all subsectors
          For ss = 0 To (numsubsectors - 1)
               
               'Check if subsector is part of this sector
               If (m_subsectors(ss).sector = s) Then
                    
                    'Remove vertexbuffer so it will be recreated
                    d_SubSectorCeilings(ss) = False
                    Set SubSectorCeilings(ss) = Nothing
               End If
          Next ss
          Set i_SectorCeilings(s) = Nothing
     End If
End Sub

Private Sub PasteFloorTexture(ByVal s As Long)
     Dim Texture As clsImage
     Dim ss As Long
     
     'Anything to paste?
     If (Len(Trim$(CopiedFlat)) > 0) Then
          
          'Make undo
          CreateUndo "change floor texture", UGRP_FLOORTEXTURECHANGE, s, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Make it so
          sectors(s).tfloor = CopiedFlat
          
          'Check if texture is known
          If allflats.Exists(sectors(s).tfloor) Then
               
               'Get texture object
               Set Texture = allflats(sectors(s).tfloor)
               
               'Show message
               ShowMainText "Floor texture pasted:  " & Texture.Name & "  " & Texture.width & "x" & Texture.height
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'Show message
               ShowMainText "Floor texture pasted"
          End If
          
          'Go for all subsectors
          For ss = 0 To (numsubsectors - 1)
               
               'Check if subsector is part of this sector
               If (m_subsectors(ss).sector = s) Then
                    
                    'Remove vertexbuffer so it will be recreated
                    d_SubSectorFloors(ss) = False
                    Set SubSectorFloors(ss) = Nothing
               End If
          Next ss
          Set i_SectorFloors(s) = Nothing
     End If
End Sub

Private Sub PasteThing(ByVal th As Long)
     Dim Texture As clsImage
     Dim ss As Long
     Dim oldthing As MAPTHING
     
     'Anything to paste?
     If (CopiedThing.thing <> 0) Then
          
          'Make undo
          CreateUndo "paste thing", UGRP_NONE, 0, True
          
          'Map changed
          mapchanged = True
          
          'Make it so
          oldthing = things(th)
          things(th) = CopiedThing
          
          'Keep some properties
          things(th).sector = oldthing.sector
          things(th).x = oldthing.x
          things(th).y = oldthing.y
          things(th).Z = oldthing.Z
          
          'Show message
          ShowMainText "Thing pasted:  " & GetThingTypeDesc(CopiedThing.thing) & " (" & CopiedThing.thing & ")"
     End If
End Sub


Private Sub InsertThing(ByRef Hotspot As D3DVECTOR)
     Dim t As Long
     
     'Anything to paste?
     If (CopiedThing.thing <> 0) Then LastThing = CopiedThing
     
     'Make undo
     CreateUndo "insert thing", UGRP_NONE, 0, True
     
     'Map changed
     mapchanged = True
     
     'Destroy pointers
     DestroyStructurePointers
     
     'Make thing here
     t = CreateThing
     things(t) = LastThing
     
     'Give pointers to the DLL
     SetStructurePointers m_vertices(0), linedefs(0), VarPtr(sidedefs(0)), m_segs(0), VarPtr(sectors(0)), m_subsectors(0), things(0), m_nodes(0), numnodes, numsectors, numsubsectors, numthings
     
     'Set some properties
     With things(t)
          .selected = 0
          .x = Hotspot.x
          .y = Hotspot.y
     End With
     
     'Determine sector where thing is
     things(t).sector = IntersectSector(things(t).x, -things(t).y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
     
     'Check if we should edit the thing
     If (Config("newthingdialog") = vbChecked) Then
          
          'Edit thing now
          SelectNewThing t
          
          'This is now the last thing
          LastThing = things(t)
     End If
     
     'Show message
     ShowMainText "Thing inserted:  " & GetThingTypeDesc(things(t).thing) & " (" & things(t).thing & ")"
End Sub



Private Sub PasteLowerTexture(ByVal sd As Long)
     Dim Texture As clsImage
     
     'Anything to paste?
     If (Len(Trim$(CopiedTexture)) > 0) Then
          
          'Make undo
          CreateUndo "change lower texture", UGRP_LOWERTEXTURECHANGE, sd, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Make it so
          sidedefs(sd).lower = CopiedTexture
          
          'Check if texture is known
          If alltextures.Exists(sidedefs(sd).lower) Then
               
               'Get texture object
               Set Texture = alltextures(sidedefs(sd).lower)
               
               'Show message
               ShowMainText "Lower texture pasted:  " & Texture.Name & "  " & Texture.width & "x" & Texture.height
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'Show message
               ShowMainText "Lower texture pasted"
          End If
          
          'Remove vertexbuffer so it will be recreated
          d_SidedefLower(sd) = False
          Set SidedefLower(sd) = Nothing
          Set i_SidedefLower(sd) = Nothing
     End If
End Sub

Private Sub PasteMiddleTexture(ByVal sd As Long)
     Dim Texture As clsImage
     
     'Anything to paste?
     If (Len(Trim$(CopiedTexture)) > 0) Then
          
          'Make undo
          CreateUndo "change middle texture", UGRP_MIDDLETEXTURECHANGE, sd, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Make it so
          sidedefs(sd).middle = CopiedTexture
          
          'Check if texture is known
          If alltextures.Exists(sidedefs(sd).middle) Then
               
               'Get texture object
               Set Texture = alltextures(sidedefs(sd).middle)
               
               'Show message
               ShowMainText "Middle texture pasted:  " & Texture.Name & "  " & Texture.width & "x" & Texture.height
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'Show message
               ShowMainText "Middle texture pasted"
          End If
          
          'Remove vertexbuffer so it will be recreated
          d_SidedefMiddle(sd) = False
          Set SidedefMiddle(sd) = Nothing
          Set i_SidedefMiddle(sd) = Nothing
     End If
End Sub

Private Sub PasteSectorProperties(ByVal sector As Long)
     Dim ld As Long
     Dim ss As Long
     Dim RemoveThis As Long
     
     'Make undo
     CreateUndo "paste sector properties", , , True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Make it so
     With sectors(sector)
          .Brightness = CopiedSector.Brightness
          .hceiling = CopiedSector.hceiling
          .hfloor = CopiedSector.hfloor
          .special = CopiedSector.special
          .tag = CopiedSector.tag
          .tceiling = CopiedSector.tceiling
          .tfloor = CopiedSector.tfloor
     End With
     
     'Show message
     ShowMainText "Pasted sector properties"
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Dont assume remove
          RemoveThis = False
          
          'Check side 1
          If (linedefs(ld).s1 > -1) Then
               
               'Check if sidedef refers to this sector
               If (sidedefs(linedefs(ld).s1).sector = sector) Then RemoveThis = True
          End If
          
          'Check side 2
          If (linedefs(ld).s2 > -1) Then
               
               'Check if sidedef refers to this sector
               If (sidedefs(linedefs(ld).s2).sector = sector) Then RemoveThis = True
          End If
          
          'Check if should be removed
          If RemoveThis Then
               
               'Remove vertexbuffers so they will be recreated
               d_SidedefLower(linedefs(ld).s1) = False
               d_SidedefMiddle(linedefs(ld).s1) = False
               d_SidedefUpper(linedefs(ld).s1) = False
               d_SidedefLower(linedefs(ld).s2) = False
               d_SidedefMiddle(linedefs(ld).s2) = False
               d_SidedefUpper(linedefs(ld).s2) = False
               Set SidedefLower(linedefs(ld).s1) = Nothing
               Set SidedefMiddle(linedefs(ld).s1) = Nothing
               Set SidedefUpper(linedefs(ld).s1) = Nothing
               Set SidedefLower(linedefs(ld).s2) = Nothing
               Set SidedefMiddle(linedefs(ld).s2) = Nothing
               Set SidedefUpper(linedefs(ld).s2) = Nothing
               Set i_SidedefLower(linedefs(ld).s1) = Nothing
               Set i_SidedefMiddle(linedefs(ld).s1) = Nothing
               Set i_SidedefUpper(linedefs(ld).s1) = Nothing
               Set i_SidedefLower(linedefs(ld).s2) = Nothing
               Set i_SidedefMiddle(linedefs(ld).s2) = Nothing
               Set i_SidedefUpper(linedefs(ld).s2) = Nothing
          End If
     Next ld
     
     'Go for all subsectors
     For ss = 0 To (numsubsectors - 1)
          
          'Check if subsector is part of this sector
          If (m_subsectors(ss).sector = sector) Then
               
               'Remove vertexbuffer so it will be recreated
               d_SubSectorCeilings(ss) = False
               d_SubSectorFloors(ss) = False
               Set SubSectorCeilings(ss) = Nothing
               Set SubSectorFloors(ss) = Nothing
          End If
     Next ss
     Set i_SectorCeilings(sector) = Nothing
     Set i_SectorFloors(sector) = Nothing
End Sub

Private Sub PasteSidedefProperties(ByVal sd As Long)
     
     'Make undo
     CreateUndo "paste sidedef properties", , , True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Make it so
     With sidedefs(sd)
          .lower = CopiedSidedef.lower
          .middle = CopiedSidedef.middle
          .tx = CopiedSidedef.tx
          .ty = CopiedSidedef.ty
          .upper = CopiedSidedef.upper
     End With
     
     'Show message
     ShowMainText "Pasted sidedef properties"
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefUpper(sd) = False
     d_SidedefMiddle(sd) = False
     d_SidedefLower(sd) = False
     Set SidedefUpper(sd) = Nothing
     Set SidedefMiddle(sd) = Nothing
     Set SidedefLower(sd) = Nothing
     Set i_SidedefUpper(sd) = Nothing
     Set i_SidedefMiddle(sd) = Nothing
     Set i_SidedefLower(sd) = Nothing
End Sub

Private Sub PasteSidedefOffsets(ByVal sd As Long)
     
     'Make undo
     CreateUndo "paste sidedef offsets", , , True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Make it so
     With sidedefs(sd)
          .tx = CopiedX
          .ty = CopiedY
     End With
     
     'Show message
     ShowMainText "Pasted offsets:  " & CopiedX & ", " & CopiedY
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefUpper(sd) = False
     d_SidedefMiddle(sd) = False
     d_SidedefLower(sd) = False
     Set SidedefUpper(sd) = Nothing
     Set SidedefMiddle(sd) = Nothing
     Set SidedefLower(sd) = Nothing
     Set i_SidedefUpper(sd) = Nothing
     Set i_SidedefMiddle(sd) = Nothing
     Set i_SidedefLower(sd) = Nothing
End Sub


Private Sub PasteUpperTexture(ByVal sd As Long)
     Dim Texture As clsImage
     
     'Anything to paste?
     If (Len(Trim$(CopiedTexture)) > 0) Then
          
          'Make undo
          CreateUndo "change upper texture", UGRP_UPPERTEXTURECHANGE, sd, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Make it so
          sidedefs(sd).upper = CopiedTexture
          
          'Check if texture is known
          If alltextures.Exists(sidedefs(sd).upper) Then
               
               'Get texture object
               Set Texture = alltextures(sidedefs(sd).upper)
               
               'Show message
               ShowMainText "Upper texture pasted:  " & Texture.Name & "  " & Texture.width & "x" & Texture.height
               
               'Clean up
               Set Texture = Nothing
          Else
               
               'Show message
               ShowMainText "Upper texture pasted"
          End If
          
          'Remove vertexbuffer so it will be recreated
          d_SidedefUpper(sd) = False
          Set SidedefUpper(sd) = Nothing
          Set i_SidedefUpper(sd) = Nothing
     End If
End Sub

Public Function PickAimedObject(ByRef Index As Long, ByRef Hotspot As D3DVECTOR) As Long
     Dim m_LookAt As D3DVECTOR
     Dim m_Position As D3DVECTOR
     
     'This will return the type of the aimed object or 0 if nothing aimed at
     'Also sets Index to the object's index for its type
     
     'Make lookat vertex
     m_LookAt.x = Position.x + sIn(HAngle) * Cos(VAngle) * (c_videoviewdistance * MAP_RENDER_SCALE)
     m_LookAt.y = Position.y + Cos(HAngle) * Cos(VAngle) * (c_videoviewdistance * MAP_RENDER_SCALE)
     m_LookAt.Z = Position.Z + sIn(VAngle) * (c_videoviewdistance * MAP_RENDER_SCALE)
     
     'Make map pixel coordinates for lookat
     m_LookAt.x = m_LookAt.x * INV_MAP_RENDER_SCALE
     m_LookAt.y = -m_LookAt.y * INV_MAP_RENDER_SCALE
     m_LookAt.Z = m_LookAt.Z * INV_MAP_RENDER_SCALE
     
     'Make map pixel coordinates for position
     m_Position.x = Position.x * INV_MAP_RENDER_SCALE
     m_Position.y = -Position.y * INV_MAP_RENDER_SCALE
     m_Position.Z = Position.Z * INV_MAP_RENDER_SCALE
     
     'Check if there is anything to test
     If (r_numsidedefs > 0) And (r_numsubsectors > 0) Then
          
          'Do the ray intersection tests
          PickAimedObject = PickObject(vertexes(0), linedefs(0), VarPtr(sidedefs(0)), _
                                       VarPtr(sectors(0)), m_subsectors(0), things(0), _
                                       r_sidedefs(0), r_numsidedefs, numlinedefs, r_subsectors(0), _
                                       r_numsubsectors, r_things(0), r_numthings, m_Position, _
                                       m_LookAt, Hotspot, Index)
     Else
          
          'Nothing to try
          Index = -1
          PickAimedObject = OBJ_NOTHING
     End If
End Function

Public Sub PollMouse()
     On Local Error Resume Next
     Dim DIDData(1 To 20) As DIDEVICEOBJECTDATA
     Dim numitems As Long
     Dim c As Long, r As Long
     Dim MousePoint As POINT
     Dim i As Long
     Dim LastIndex As Long
     
     'Get data, if it fails try to acquire the mouse again
     'numitems = DIMouse.GetDeviceData(DIDData, DIGDD_DEFAULT)
     numitems = 20
     DIMouse.GetDeviceData Len(DIDData(1)), DIDData(1), numitems, 0
     If Err.number Then Exit Sub
     
     'Check how we should process data
     If TextureSelecting Then
          
          'Process data
          For i = 1 To numitems
               Select Case DIDData(i).lOfs
                    
                    Case DIMOFS_X, DIMOFS_Y  'Any movement
                         
                         'Check if in windowed mode
                         If (Val(Config("windowedvideo"))) Then
                              
                              'Get mouse coords from form
                              MousePoint.x = frmMain.LastMouseX
                              MousePoint.y = frmMain.LastMouseY
                         Else
                              
                              'Get mouse coords
                              GetCursorPos MousePoint
                         End If
                         
                         'Check if moved
                         If (Abs(MousePoint.x - TLastX) > 1) Or (Abs(MousePoint.y - TLastY) > 1) Then
                              
                              'Limit to texture area
                              If (MousePoint.y >= (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT))) Then MousePoint.y = (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT)) - 1
                              
                              'Calculate the row and col
                              c = Int((MousePoint.x / VideoParams.BackBufferWidth) * TEXTURE_COLS)
                              r = Int((MousePoint.y / (VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT))) * TEXTURE_ROWS)
                              
                              'Change selection index
                              LastIndex = TextureSelectedIndex
                              TextureSelectedIndex = (TextureRowOffset + r) * TEXTURE_COLS + c
                              If (TextureSelectedIndex > (curnumitems - 1)) Then TextureSelectedIndex = curnumitems - 1
                              If (TextureSelectedIndex < 0) Then TextureSelectedIndex = 0
                              If (LastIndex <> TextureSelectedIndex) Then
                                   SelectedName = curitemnames(TextureSelectedIndex)
                                   CreateSelectedTextureText
                              End If
                              
                              'Update coords
                              TLastX = MousePoint.x
                              TLastY = MousePoint.y
                         End If
                         
                    Case DIMOFS_Z  'Scrollwheel
                         
                         'Check if not handled by main handler
                         If (Val(Config("windowedvideo")) = 0) Then
                              
                              'Handle scrollwheel
                              If (DIDData(i).lData > 0) Then
                                   KeydownTextureSelect MOUSE_SCROLL_UP
                              Else
                                   KeydownTextureSelect MOUSE_SCROLL_DOWN
                              End If
                         End If
                         
               End Select
          Next i
          
     Else
          
          'Process data
          For i = 1 To numitems
               Select Case DIDData(i).lOfs
                    
                    Case DIMOFS_X
                         HAngle = HAngle - (DIDData(i).lData * c_mousespeed / 10000)
                         
                         'Set time for updating info panel
                         'InfoUpdateTime = CurrentTime + INFO_UPDATEDELAY
                         
                    Case DIMOFS_Y
                         
                         'Check if using invertex Y
                         If (c_invertmousey = vbChecked) Then
                              VAngle = VAngle + (DIDData(i).lData * c_mousespeed / 10000)
                         Else
                              VAngle = VAngle - (DIDData(i).lData * c_mousespeed / 10000)
                         End If
                         
                         'Limit the Y look
                         If (VAngle > 1.5) Then VAngle = 1.5
                         If (VAngle < -1.5) Then VAngle = -1.5
                         
                         'Set time for updating info panel
                         'InfoUpdateTime = CurrentTime + INFO_UPDATEDELAY
                         
                    Case DIMOFS_Z  'Scrollwheel
                         
                         'Check if not handled by main handler
                         If (Val(Config("windowedvideo")) = 0) Then
                              
                              'Handle scrollwheel
                              If (DIDData(i).lData > 0) Then
                                   Keydown3D MOUSE_SCROLL_UP Or (CurrentShiftMask * (2 ^ 16))
                              Else
                                   Keydown3D MOUSE_SCROLL_DOWN Or (CurrentShiftMask * (2 ^ 16))
                              End If
                         End If
               End Select
          Next i
     End If
End Sub

Public Sub FreeMouse()
     On Error Resume Next
     
     'Show hourglass cursor
     While ShowCursor(True) < 0: Wend
     
     'Stop polling mouse events
     If Not DIMouse Is Nothing Then
          DIMouse.Unacquire
     End If
     
     'Free the cursor movement
     ClipCursor ByVal 0
     
     'Disregard errors
     Err.Clear
End Sub



Private Sub PositionCamera()
     Dim LookAt As D3DVECTOR
     
     'Make lookat vector
     LookAt.x = Position.x + sIn(HAngle) * Cos(VAngle) * c_videoviewdistance
     LookAt.y = Position.y + Cos(HAngle) * Cos(VAngle) * c_videoviewdistance
     LookAt.Z = Position.Z + sIn(VAngle) * c_videoviewdistance
     
     'Make projection matrix
     MatrixLookAtLH matrixView, Position, LookAt, Vector3D(0, 0, 1)
     
     'When in windowed mode, show coordinates
     If (Val(Config("windowedvideo"))) Then
          
          'Check if we should update the coords
          If (InfoCoordsUpdateTime < CurrentTime) Then
               
               'Update cursor position in statusbar
               frmMain.stbStatus.Panels("mousex").Text = "X " & CLng(Position.x * INV_MAP_RENDER_SCALE)
               frmMain.stbStatus.Panels("mousey").Text = "Y " & -CLng(Position.y * INV_MAP_RENDER_SCALE)
               
               'Change update time
               InfoCoordsUpdateTime = CurrentTime + INFO_COORDS_UPDATEDELAY
          End If
     End If
End Sub

Public Function PrepareStructures(ByVal File As clsWAD) As Boolean
     On Error GoTo errorhandler
     Dim lumpindex As Long
     Dim s As Long
     
     'Get the VERTEXES lump
     lumpindex = FindLumpIndex(File, 1, "VERTEXES")
     If (lumpindex = 0) Then Err.Raise 1, , "Could not find required lump VERTEXES!"
     
     'Load the vertices
     LoadVertices File.FileBuffer, File.LumpAddress(lumpindex), File.LumpSize(lumpindex) \ 4
     
     'Get the SEGS lump
     lumpindex = FindLumpIndex(File, 1, "SEGS")
     If (lumpindex = 0) Then Err.Raise 2, , "Could not find required lump SEGS!"
     
     'Load the segs
     numsegs = File.LumpSize(lumpindex) \ 12
     LoadSegs File.FileBuffer, File.LumpAddress(lumpindex), numsegs
     
     'Get the SSECTORS lump
     lumpindex = FindLumpIndex(File, 1, "SSECTORS")
     If (lumpindex = 0) Then Err.Raise 3, , "Could not find required lump SSECTORS!"
     
     'Load the ssectors
     numsubsectors = File.LumpSize(lumpindex) \ 4
     LoadSSectors File.FileBuffer, File.LumpAddress(lumpindex), File.LumpSize(lumpindex) \ 4
     
     'Get the NODES lump
     lumpindex = FindLumpIndex(File, 1, "NODES")
     If (lumpindex = 0) Then Err.Raise 2, , "Could not find required lump NODES!"
     
     'Load the nodes
     numnodes = File.LumpSize(lumpindex) \ 28
     LoadNodes File.FileBuffer, File.LumpAddress(lumpindex), numnodes, m_nodes()
     
     
     'Give pointers to the DLL
     SetStructurePointers m_vertices(0), linedefs(0), VarPtr(sidedefs(0)), m_segs(0), VarPtr(sectors(0)), m_subsectors(0), things(0), m_nodes(0), numnodes, numsectors, numsubsectors, numthings
     
     'Make needed references
     CreateSSectorReferences
     
     'Triangulate ssectors and recalculate ssector boundaries
     PrepareAllSSectors
     
     'Find sectors where Things are
     SetAllThingSectors things(0), numthings, vertexes(0), linedefs(0), numlinedefs, VarPtr(sidedefs(0))
     
          
          
     
     'Make databases
     ReDim SubSectorFloors(0 To numsubsectors - 1)
     ReDim SubSectorCeilings(0 To numsubsectors - 1)
     ReDim SidedefUpper(-1 To numsidedefs - 1)
     ReDim SidedefMiddle(-1 To numsidedefs - 1)
     ReDim SidedefLower(-1 To numsidedefs - 1)
     ReDim d_SubSectorFloors(0 To numsubsectors - 1)
     ReDim d_SubSectorCeilings(0 To numsubsectors - 1)
     ReDim d_SidedefUpper(-1 To numsidedefs - 1)
     ReDim d_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim d_SidedefLower(-1 To numsidedefs - 1)
     ReDim i_SectorFloors(0 To numsectors - 1)
     ReDim i_SectorCeilings(0 To numsectors - 1)
     ReDim i_SidedefUpper(-1 To numsidedefs - 1)
     ReDim i_SidedefMiddle(-1 To numsidedefs - 1)
     ReDim i_SidedefLower(-1 To numsidedefs - 1)
     
     'Done here
     PrepareStructures = True
     Exit Function
     
errorhandler:
     
     'Show error
     MsgBox "Error " & Err.number & " during loading of structures: " & Err.Description, vbCritical
End Function

Private Sub RenderCrosshair()
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 0
     
     'Set texture
     D3DD.SetTexture 0, tex_crosshair
     
     'Set the data
     D3DD.SetStreamSource 0, r_crosshair, 0, TLVERTEXSTRIDE
     
     'Set vertex format
     D3DD.SetFVF TLVERTEXFVF
     
     'Render the crosshair
     D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
End Sub

Private Sub RenderInfoPanel()
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 0
     
     'Set texture
     D3DD.SetTexture 0, Nothing
     
     'Set the data
     D3DD.SetStreamSource 0, r_infopanel, 0, TLVERTEXSTRIDE
     
     'Set vertex format
     D3DD.SetFVF TLVERTEXFVF
     
     'Render the crosshair
     D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
End Sub


Public Sub RenderMouse()
     Dim MousePoly(3) As TLVERTEX
     Dim MousePoint As POINT
     Dim CrosshairWidth As Long
     Dim CrosshairHeight As Long
     
     'Determine crosshair size
     CrosshairWidth = VideoParams.BackBufferWidth / 25 * (CSng(CrosshairInfo.width) / 32)
     CrosshairHeight = (VideoParams.BackBufferWidth / 25) * (CSng(CrosshairInfo.height) / 32)
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo"))) Then
          
          'Get mouse coords from form
          MousePoint.x = frmMain.LastMouseX
          MousePoint.y = frmMain.LastMouseY
     Else
          
          'Get mouse coords
          GetCursorPos MousePoint
     End If
     
     'Create Polgon
     With MousePoly(0)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = MousePoint.x - CrosshairWidth * 0.5
          .sy = MousePoint.y - CrosshairHeight * 0.5
          .tu = 0
          .tv = 0
     End With
     
     With MousePoly(1)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = MousePoint.x - CrosshairWidth * 0.5
          .sy = MousePoint.y + CrosshairHeight * 0.5
          .tu = 0
          .tv = 1
     End With
     
     With MousePoly(2)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = MousePoint.x + CrosshairWidth * 0.5
          .sy = MousePoint.y - CrosshairHeight * 0.5
          .tu = 1
          .tv = 0
     End With
     
     With MousePoly(3)
          .Color = D3DColorMake(1, 1, 1, 1)
          .rhw = 1
          .sx = MousePoint.x + CrosshairWidth * 0.5
          .sy = MousePoint.y + CrosshairHeight * 0.5
          .tu = 1
          .tv = 1
     End With
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     
     'Set texture
     D3DD.SetTexture 0, tex_crosshair
     
     'Set vertex format
     D3DD.SetFVF TLVERTEXFVF
     
     'Render the crosshair
     D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VarPtr(MousePoly(0)), TLVERTEXSTRIDE
End Sub

Public Sub RenderSelection()
     Dim SelectionPoly(3) As TLVERTEX
     Dim cw As Single, ch As Single
     Dim sx As Long, sy As Long
     
     'Check if a selection is made
     If (TextureSelectedIndex >= 0) Then
          
          'Determine the visible index of selection
          sy = TextureSelectedIndex \ TEXTURE_COLS
          sx = TextureSelectedIndex - sy * TEXTURE_COLS
          sy = sy - TextureRowOffset
          
          'Check if visible
          If (sy >= 0) And (sy < TEXTURE_ROWS) Then
               
               'Calculate cell width and height (without spacing)
               cw = VideoParams.BackBufferWidth / TEXTURE_COLS
               ch = VideoParams.BackBufferHeight * (1 - TEXTURE_TEXTHEIGHT) / TEXTURE_ROWS
               
               'Create Polgon
               With SelectionPoly(0)
                    .Color = D3DColorMake(0.2, 0.6, 1, 1)
                    .rhw = 1
                    .sx = cw * sx
                    .sy = ch * sy
                    .tu = 0
                    .tv = 0
               End With
               
               With SelectionPoly(1)
                    .Color = D3DColorMake(0.2, 0.6, 1, 1)
                    .rhw = 1
                    .sx = cw * sx
                    .sy = ch * sy + ch
                    .tu = 0
                    .tv = 1
               End With
               
               With SelectionPoly(2)
                    .Color = D3DColorMake(0.2, 0.6, 1, 1)
                    .rhw = 1
                    .sx = cw * sx + cw
                    .sy = ch * sy
                    .tu = 1
                    .tv = 0
               End With
               
               With SelectionPoly(3)
                    .Color = D3DColorMake(0.2, 0.6, 1, 1)
                    .rhw = 1
                    .sx = cw * sx + cw
                    .sy = ch * sy + ch
                    .tu = 1
                    .tv = 1
               End With
               
               'Set texture
               D3DD.SetTexture 0, Nothing
               
               'Set vertex format
               D3DD.SetFVF TLVERTEXFVF
               
               'Render the crosshair
               D3DD.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VarPtr(SelectionPoly(0)), TLVERTEXSTRIDE
          End If
     End If
End Sub

Private Sub RenderSidedefsLower()
     Dim ssi As Long
     Dim sd As Long
     
     'Go for all visible sidedefs
     '(front to back to reduce overdraw)
     For ssi = 0 To (r_numsidedefs - 1)
          
          'Get the sidedef
          sd = r_sidedefs(ssi)
          
          'Check if its valid
          If Not (SidedefLower(sd) Is Nothing) Then
               
               'Apply the sector brightness
               If (sidedefs(sd).sector > -1) And (FullBrightness = False) Then
                    D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(sidedefs(sd).sector).Brightness)
                    D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(sidedefs(sd).sector).Brightness)
               End If
               
               'Set texture
               D3DD.SetTexture 0, i_SidedefLower(sd)
               
               'Set brightness
               If IsTextureName(sidedefs(sd).lower) = False Then D3DD.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
               
               'Set the data
               D3DD.SetStreamSource 0, SidedefLower(sd), 0, VERTEXSTRIDE
               
               'Set vertex format
               D3DD.SetFVF VERTEXFVF
               
               'Render the floor
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
          End If
     Next ssi
End Sub

Private Sub RenderSidedefsMiddle()
     Dim ssi As Long
     Dim sd As Long
     
     'Go for all visible sidedefs
     '(front to back to reduce overdraw)
     For ssi = 0 To (r_numsidedefs - 1)
          
          'Get the sidedef
          sd = r_sidedefs(ssi)
          
          'Check if its valid
          If Not (SidedefMiddle(sd) Is Nothing) Then
               
               'Check if not a window
               If (linedefs(sidedefs(sd).linedef).s2 = -1) Then
                    
                    'Apply the sector brightness
                    If (sidedefs(sd).sector > -1) And (FullBrightness = False) Then
                         D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(sidedefs(sd).sector).Brightness)
                         D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(sidedefs(sd).sector).Brightness)
                    End If
                    
                    'Set texture
                    D3DD.SetTexture 0, i_SidedefMiddle(sd)
                    
                    'Set brightness
                    If IsTextureName(sidedefs(sd).middle) = False Then D3DD.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
                    
                    'Set the data
                    D3DD.SetStreamSource 0, SidedefMiddle(sd), 0, VERTEXSTRIDE
                    
                    'Set vertex format
                    D3DD.SetFVF VERTEXFVF
                    
                    'Render the floor
                    D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
               End If
          End If
     Next ssi
End Sub

Private Sub RenderSidedefsWindows()
     Dim ssi As Long
     Dim sd As Long
     
     'Go for all visible sidedefs
     '(back to front to correct transparency)
     For ssi = (r_numsidedefs - 1) To 0 Step -1
          
          'Get the sidedef
          sd = r_sidedefs(ssi)
          
          'Check if its valid
          If Not (SidedefMiddle(sd) Is Nothing) Then
               
               'Check if this is a window
               If (linedefs(sidedefs(sd).linedef).s2 > -1) Then
                    
                    'Apply the sector brightness
                    If (sidedefs(sd).sector > -1) And (FullBrightness = False) Then
                         D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(sidedefs(sd).sector).Brightness)
                         D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(sidedefs(sd).sector).Brightness)
                    End If
                    
                    'Set texture
                    D3DD.SetTexture 0, i_SidedefMiddle(sd)
                    
                    'Set brightness
                    If IsTextureName(sidedefs(sd).middle) = False Then D3DD.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
                    
                    'Set the data
                    D3DD.SetStreamSource 0, SidedefMiddle(sd), 0, VERTEXSTRIDE
                    
                    'Set vertex format
                    D3DD.SetFVF VERTEXFVF
                    
                    'Render the floor
                    D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
               End If
          End If
     Next ssi
End Sub


Private Sub RenderSidedefsUpper()
     Dim ssi As Long
     Dim sd As Long
     
     'Go for all visible sidedefs
     '(front to back to reduce overdraw)
     For ssi = 0 To (r_numsidedefs - 1)
          
          'Get the sidedef
          sd = r_sidedefs(ssi)
          
          'Check if its valid
          If Not (SidedefUpper(sd) Is Nothing) Then
               
               'Apply the sector brightness
               If (sidedefs(sd).sector > -1) And (FullBrightness = False) Then
                    D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(sidedefs(sd).sector).Brightness)
                    D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(sidedefs(sd).sector).Brightness)
               End If
               
               'Set texture
               D3DD.SetTexture 0, i_SidedefUpper(sd)
               
               'Set brightness
               If IsTextureName(sidedefs(sd).upper) = False Then D3DD.SetRenderState D3DRS_AMBIENT, &HFFFFFFFF
               
               'Set the data
               D3DD.SetStreamSource 0, SidedefUpper(sd), 0, VERTEXSTRIDE
               
               'Set vertex format
               D3DD.SetFVF VERTEXFVF
               
               'Render the floor
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
          End If
     Next ssi
End Sub

Private Sub RenderSubSectorCeilings()
     Dim ssi As Long
     Dim ss As Long
     Dim mapscalez As Long
     
     'Calculate the Z in map scale
     mapscalez = Position.Z * INV_MAP_RENDER_SCALE
     
     'Go for all visible subsectors
     '(front to back to reduce overdraw)
     For ssi = 0 To (r_numsubsectors - 1)
          
          'Get subsector index
          ss = r_subsectors(ssi)
          
          'Check if its valid
          If Not (SubSectorCeilings(ss) Is Nothing) Then
               
               'Check if eyes below ceiling
               If (mapscalez < sectors(m_subsectors(ss).sector).hceiling) Then
                    
                    'Apply the sector brightness
                    If (FullBrightness = False) Then
                         D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(m_subsectors(ss).sector).Brightness)
                         D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(m_subsectors(ss).sector).Brightness)
                    End If
                    
                    'Set floor texture
                    D3DD.SetTexture 0, i_SectorCeilings(m_subsectors(ss).sector)
                    
                    'Set the floor data
                    D3DD.SetStreamSource 0, SubSectorCeilings(ss), 0, VERTEXSTRIDE
                    
                    'Set vertex format
                    D3DD.SetFVF VERTEXFVF
                    
                    'Render the floor
                    D3DD.DrawPrimitive D3DPT_TRIANGLEFAN, 0, (m_subsectors(ss).numvertices - 2)
               End If
          End If
     Next ssi
End Sub

Private Sub RenderSubSectorFloors()
     Dim ssi As Long
     Dim ss As Long
     Dim mapscalez As Long
     
     'Calculate the Z in map scale
     mapscalez = Position.Z * INV_MAP_RENDER_SCALE
     
     'Go for all visible subsectors
     '(front to back to reduce overdraw)
     For ssi = 0 To (r_numsubsectors - 1)
          
          'Get subsector index
          ss = r_subsectors(ssi)
          
          'Check if its valid
          If Not (SubSectorFloors(ss) Is Nothing) Then
               
               'Check if eyes above floor
               If (mapscalez > sectors(m_subsectors(ss).sector).hfloor) Then
                    
                    'Apply the sector brightness
                    If (FullBrightness = False) Then
                         D3DD.SetRenderState D3DRS_AMBIENT, t_brightness(sectors(m_subsectors(ss).sector).Brightness)
                         D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(m_subsectors(ss).sector).Brightness)
                    End If
                    
                    'Set floor texture
                    D3DD.SetTexture 0, i_SectorFloors(m_subsectors(ss).sector)
                    
                    'Set the floor data
                    D3DD.SetStreamSource 0, SubSectorFloors(ss), 0, VERTEXSTRIDE
                    
                    'Set vertex format
                    D3DD.SetFVF VERTEXFVF
                    
                    'Render the floor
                    D3DD.DrawPrimitive D3DPT_TRIANGLEFAN, 0, (m_subsectors(ss).numvertices - 2)
               End If
          End If
     Next ssi
End Sub

Private Sub RenderThings()
     Dim ti As Long
     Dim t As Long
     Dim sectorbright As Long
     Dim thingcolor As D3DCOLORVALUE
     Dim sectorcolor As D3DCOLORVALUE
     Dim matrixRotate As D3DMATRIX
     Dim matrixScale As D3DMATRIX
     Dim matrixTranslate As D3DMATRIX
     Dim matrixBox As D3DMATRIX
     Dim matrixSprite As D3DMATRIX
     Dim matrixArrow As D3DMATRIX
     Dim matrixTexture As D3DMATRIX
     Dim thingz As Long
     Dim spriteoffsetz As Long
     Dim Div255 As Single
     Dim Sprite As clsImage
     Dim vDir As D3DVECTOR
     Dim ceilheight As Single
     Dim floorheight As Single
     
     'Set vertex format
     D3DD.SetFVF VERTEXFVF
     
     'Dividers
     Div255 = 1 / 255
     
     'Go for all visible things
     '(back to front for correct transparency)
     For ti = (r_numthings - 1) To 0 Step -1
          
          'Get thing index
          t = r_things(ti)
          
          'Create full bright category color
          thingcolor.r = CSng(ScreenPalette(things(t).Color).rgbRed) * Div255
          thingcolor.g = CSng(ScreenPalette(things(t).Color).rgbGreen) * Div255
          thingcolor.b = CSng(ScreenPalette(things(t).Color).rgbBlue) * Div255
          
          'Check if thing is inside a sector
          If (things(t).sector >= 0) Then
               
               'Get heights
               ceilheight = sectors(things(t).sector).hceiling
               floorheight = sectors(things(t).sector).hfloor
               
               'Check if sprite hangs from ceiling
               If (things(t).hangs) Then
                    
                    'Hanging from ceiling
                    thingz = ceilheight - things(t).height
                    If (things(t).Z > 0) Then thingz = thingz - things(t).Z
                    
                    'Check if below floor
                    If (thingz <= floorheight) Then
                         
                         'Put against floor
                         thingz = floorheight
                    End If
               Else
                    
                    'Standing on floor
                    thingz = floorheight
                    If (things(t).Z > 0) Then thingz = thingz + things(t).Z
                    
                    'Check if above ceiling
                    If (thingz + things(t).height >= ceilheight) Then
                         
                         'Put against ceiling
                         thingz = ceilheight - things(t).height
                    End If
               End If
               
               'Check if we must adjust color for sector brightness
               If (FullBrightness = False) Then
                    
                    'Make sector brightness color
                    sectorbright = t_brightness(sectors(things(t).sector).Brightness)
                    sectorcolor.r = CSng((sectorbright And &HFF0000) / (2 ^ 16)) * Div255 + 0.1
                    sectorcolor.g = CSng((sectorbright And &HFF00&) / (2 ^ 8)) * Div255 + 0.1
                    sectorcolor.b = CSng(sectorbright And &HFF&) * Div255 + 0.1
                    If (sectorcolor.r > 1) Then sectorcolor.r = 1
                    If (sectorcolor.g > 1) Then sectorcolor.g = 1
                    If (sectorcolor.b > 1) Then sectorcolor.b = 1
                    
                    'Multiply color with sector brightness
                    ColorModulate thingcolor, thingcolor, sectorcolor
                    
                    'Use the fog as well
                    D3DD.SetRenderState D3DRS_FOGDENSITY, t_fogness(sectors(things(t).sector).Brightness)
               Else
                    
                    'Default sector brightness
                    sectorcolor.r = 1
                    sectorcolor.g = 1
                    sectorcolor.b = 1
               End If
          Else
               
               'Show thing at 0 height
               thingz = 0
               
               'Default sector brightness
               sectorcolor.r = 1
               sectorcolor.g = 1
               sectorcolor.b = 1
          End If
          
          'Create scale
          MatrixScaling matrixScale, CSng(things(t).size * 2) - 0.4, CSng(things(t).size * 2) - 0.4, CSng(things(t).height) - 0.4
          
          'Create translation
          MatrixTranslation matrixTranslate, things(t).x * MAP_RENDER_SCALE, -things(t).y * MAP_RENDER_SCALE, CSng(thingz) * MAP_RENDER_SCALE + 0.002
          
          'Create rotation (thing angle)
          MatrixRotationZ matrixRotate, CSng(270 - things(t).angle) * PiDivMul
          
          'Multiply matrices
          MatrixIdentity matrixBox
          MatrixMultiply matrixBox, matrixBox, matrixScale
          If (things(t).image <> TI_DOT) Then MatrixMultiply matrixArrow, matrixBox, matrixRotate
          MatrixMultiply matrixBox, matrixBox, matrixTranslate
          If (things(t).image <> TI_DOT) Then MatrixMultiply matrixArrow, matrixArrow, matrixTranslate
          
          'Do we want to see an arrow?
          If (things(t).image <> TI_DOT) Then
               
               'Turn off alpha testing and depth buffer writing
               D3DD.SetRenderState D3DRS_ZWRITEENABLE, 0
               D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 0
               
               'Bilinear texture filtering
               SetTextureFilters True
               
               'Apply color
               D3DD.SetRenderState D3DRS_AMBIENT, D3DColorMake(thingcolor.r, thingcolor.g, thingcolor.b, 1)
               
               'Apply matrix
               D3DD.SetTransform D3DTS_WORLD, matrixArrow
               
               'Set texture
               D3DD.SetTexture 0, tex_thingarrow
               
               'Set the data
               D3DD.SetStreamSource 0, r_thingarrow, 0, VERTEXSTRIDE
               
               'Render the thing arrow
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
               
               'Normal texture filtering
               SetTextureFilters False
          End If
          
          'Get thing sprite
          Set Sprite = GetSpriteForThingType(things(t).thing)
          
          'Check if thing has a sprite
          If Not (Sprite Is Nothing) Then
               
               'Apply color
               D3DD.SetRenderState D3DRS_AMBIENT, D3DColorMake(sectorcolor.r, sectorcolor.g, sectorcolor.b, 1)
               
               'Sprite image as texture
               'If (GetThingTypeSpriteName(things(t).thing) = "BFUGA0") Then Stop
               If (Sprite.D3DTexture Is Nothing) Then Sprite.LoadD3DTexture True
               D3DD.SetTexture 0, Sprite.D3DTexture
               
               'Create texture coordinates matrix
               MatrixScaling matrixTexture, Sprite.d3dscalewidth, Sprite.d3dscaleheight, 0
               
               'Create scale
               MatrixScaling matrixScale, Sprite.width, 1, Sprite.height
               
               'Create rotation (billboarding)
               MatrixRotationZ matrixRotate, pi - HAngle
               
               'Check if sprite hangs from ceiling
               If (things(t).hangs) Then
                    
                    'Render sprite hanging
                    spriteoffsetz = things(t).height - Sprite.height
               Else
                    
                    'Render sprite on floor
                    spriteoffsetz = (Sprite.ScaleY - Sprite.height) + 2
               End If
               
               'Create translation
               MatrixTranslation matrixTranslate, things(t).x * MAP_RENDER_SCALE, -things(t).y * MAP_RENDER_SCALE, (thingz + spriteoffsetz) * MAP_RENDER_SCALE
               
               'Multiply matrices
               MatrixIdentity matrixSprite
               MatrixMultiply matrixSprite, matrixSprite, matrixScale
               MatrixMultiply matrixSprite, matrixSprite, matrixRotate
               MatrixMultiply matrixSprite, matrixSprite, matrixTranslate
               
               'Apply matrices
               D3DD.SetTransform D3DTS_WORLD, matrixSprite
               D3DD.SetTransform D3DTS_TEXTURE0, matrixTexture
               
               'Set the data
               D3DD.SetStreamSource 0, r_thingsprite, 0, VERTEXSTRIDE
               
               'Turn on alpha test and depth buffer writing
               D3DD.SetRenderState D3DRS_ZWRITEENABLE, 1
               D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 1
               
               'Set texture stages
               D3DD.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_BORDER
               D3DD.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_BORDER
               D3DD.SetTextureStageState 0, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_COUNT2
               
               'Render the thing box
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
               
               'Disable texture transformation
               D3DD.SetTextureStageState 0, D3DTSS_TEXTURETRANSFORMFLAGS, D3DTTFF_DISABLE
          End If
          
          'Turn off alpha testing and depth buffer writing
          D3DD.SetRenderState D3DRS_ZWRITEENABLE, 0
          D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 0
          
          'Repeat textures
          D3DD.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
          D3DD.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
          
          'Apply color
          D3DD.SetRenderState D3DRS_AMBIENT, D3DColorMake(thingcolor.r, thingcolor.g, thingcolor.b, 1)
          
          'Apply matrix
          D3DD.SetTransform D3DTS_WORLD, matrixBox
          
          'Set texture
          D3DD.SetTexture 0, tex_thingbox
          
          'Set the data
          D3DD.SetStreamSource 0, r_thingboxvb, 0, VERTEXSTRIDE
          
          'Render the thing box
          D3DD.DrawPrimitive D3DPT_TRIANGLELIST, 0, 12
          
          'Set texture
          D3DD.SetTexture 0, Nothing
          
          'Set the data
          D3DD.SetStreamSource 0, r_thingboxlines, 0, VERTEXSTRIDE
          
          'Render the thing box lines
          D3DD.DrawPrimitive D3DPT_LINELIST, 0, 12
     Next ti
     
     'No more indices
     'D3DD.SetIndices Nothing
     
     'Restore matrix
     D3DD.SetTransform D3DTS_WORLD, matrixWorld
End Sub


Private Sub RenderTexts()
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 0
     
     'Set texture
     D3DD.SetTexture 0, tex_font
     
     'Check if there is main text to display
     If Not (r_maintext Is Nothing) Then
          
          'Set the data
          D3DD.SetStreamSource 0, r_maintext, 0, TLVERTEXSTRIDE
          
          'Set vertex format
          D3DD.SetFVF TLVERTEXFVF
          
          'Render the text
          D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, r_nummaintextfaces
     End If
     
     'Check if there is sub text to display
     If Not (r_subtext Is Nothing) Then
          
          'Set the data
          D3DD.SetStreamSource 0, r_subtext, 0, TLVERTEXSTRIDE
          
          'Set vertex format
          D3DD.SetFVF TLVERTEXFVF
          
          'Render the text
          D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, r_numsubtextfaces
     End If
End Sub

Private Sub RenderInfoTexts()
     Dim i As Long
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 0
     
     'Set texture
     D3DD.SetTexture 0, tex_font
     
     'Set vertex format
     D3DD.SetFVF TLVERTEXFVF
     
     'Go for all info panel texts
     For i = 0 To 9
          
          'Check if there is text to display
          If (Not (r_infotexts(i) Is Nothing)) And (r_numinfotextfaces(i) > 0) Then
               
               'Set the data
               D3DD.SetStreamSource 0, r_infotexts(i), 0, TLVERTEXSTRIDE
               
               'Render the text
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, r_numinfotextfaces(i)
          End If
     Next i
End Sub

Private Sub RenderTexturePreviews()
     Dim i As Long
     Dim numitems As Long
     
     'Calculate number of items
     numitems = TEXTURE_COLS * TEXTURE_ROWS
     
     'Go for all items
     For i = 0 To (numitems - 1)
          
          'Check if not nothing
          If Not (r_texpoly(i) Is Nothing) Then
               
               'Set texture
               If (r_texclass(i) Is Nothing) Then
                    D3DD.SetTexture 0, Nothing
               Else
                    D3DD.SetTexture 0, r_texclass(i).D3DTexture
               End If
               
               'Set the data
               D3DD.SetStreamSource 0, r_texpoly(i), 0, TLVERTEXSTRIDE
               
               'Set vertex format
               D3DD.SetFVF TLVERTEXFVF
               
               'Render the crosshair
               D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2
          End If
     Next i
End Sub

Private Sub RenderTextureTexts()
     
     'Render transparent
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 0
     
     'Set texture
     D3DD.SetTexture 0, tex_font
     
     'Check if there is description text to display
     If Not (r_texdesc Is Nothing) Then
          
          'Set the data
          D3DD.SetStreamSource 0, r_texdesc, 0, TLVERTEXSTRIDE
          
          'Set vertex format
          D3DD.SetFVF TLVERTEXFVF
          
          'Render the text
          D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, r_numtexdescfaces
     End If
     
     'Check if there is selection text to display
     If (Not (r_texdesc Is Nothing)) And (r_numtexnamefaces > 0) Then
          
          'Set the data
          D3DD.SetStreamSource 0, r_texname, 0, TLVERTEXSTRIDE
          
          'Set vertex format
          D3DD.SetFVF TLVERTEXFVF
          
          'Render the text
          D3DD.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, r_numtexnamefaces
     End If
End Sub

Public Sub Run3DMode()
     On Error GoTo Leave3DMode
     Dim ErrNumber As Long
     Dim ErrDesc As String
     
     'Precache when preferred
     If (Val(Config("directxprecache")) = vbChecked) Then DirectXPrecache
     
     'Start text
     ShowMainText STATUP_TITLE, STATUP_SUBTITLE
     
     'Start redraw timer
     frmMain.tmr3DRedraw.Enabled = True
     
     'Done
     Exit Sub
     
Leave3DMode:
     
     'Check if not quit nicely
     If (Running3D = True) Or (Err.number <> 0) Then
          
          'Keep error
          ErrNumber = Err.number
          ErrDesc = Err.Description
          
          'Clean up directx mode
          Running3D = False
          TextureSelecting = False
          CleanUp3DMode
          
          'Display error if not device lost error
          If (ErrNumber <> -2005530520) Then MsgBox "Error " & ErrNumber & " in Run3DMode: " & ErrDesc, vbCritical
     End If
End Sub

Public Sub RunSingleFrame(Optional ByVal ProcessMap As Long = True, Optional ByVal RenderMap As Long = True)
     Dim FOV As Long
     Dim nthings As Long
     
     'Position the camera
     PositionCamera
     
     'Check if processing is to be done
     If ProcessMap Then
          
          'Pretend the FOV > 180 when viewing up and down
          'this will make everything behind you visible as well
          If (Abs(VAngle) > 1) Then FOV = 360 Else FOV = c_videofov
          
          'No things when no things shown
          If (ShowThings) Then nthings = MAX_VISIBLE_THINGS Else nthings = 0
          
          'Walk the BSP tree
          ProcessBSP r_subsectors(0), MAX_VISIBLE_SSECTORS, _
                     r_sidedefs(0), MAX_VISIBLE_SIDEDEFS, r_numsubsectors, r_numsidedefs, _
                     Position.x * INV_MAP_RENDER_SCALE, -Position.y * INV_MAP_RENDER_SCALE, _
                     Position.Z * INV_MAP_RENDER_SCALE, HAngle + pi * 0.5, FOV, _
                     c_videoviewdistance, r_things(0), r_numthings, nthings
          
          'Check if differences should be discarded
          If Config("vertexbuffercache") = 0 Then CleanUpDiscards
          
          'Create vertexbuffers for ceilings, floors and walls
          MakeVertexBuffers
          
          'Processed once
          HasProcessed = True
     Else
          
          'Check if processed before
          If (HasProcessed = False) Then Exit Sub
     End If
     
     
     '===== Start scene
     D3DD.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Val(Config("palette")("CLR_BACKGROUND")), 1, 0
     'D3DD.Clear 0, ByVal 0, D3DCLEAR_ZBUFFER, 0, 1, 0
     D3DD.BeginScene
     
     'Apply Matrices
     D3DD.SetTransform D3DTS_PROJECTION, matrixProject
     D3DD.SetTransform D3DTS_VIEW, matrixView
     D3DD.SetTransform D3DTS_WORLD, matrixWorld
     
     'Beginning settings
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
     D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 0
     D3DD.SetRenderState D3DRS_LIGHTING, 1
     D3DD.SetRenderState D3DRS_AMBIENT, D3DColorMake(1, 1, 1, 1)
     D3DD.SetRenderState D3DRS_ZWRITEENABLE, 1
     D3DD.SetSamplerState 0, D3DSAMP_ADDRESSU, D3DTADDRESS_WRAP
     D3DD.SetSamplerState 0, D3DSAMP_ADDRESSV, D3DTADDRESS_WRAP
     'D3DD.SetRenderState D3DRS_FOGDENSITY, 0
     
     'Check if map rendering is to be done
     If RenderMap Then
          
          'Texture filtering as configured
          SetTextureFilters False
          
          'Floor, ceilings and walls dont need culling
          'They are not rendered when you cant see them anyway
          D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
          
          'Render all visible subsector floors
          RenderSubSectorFloors
          
          'Render all visible subsector ceilings
          RenderSubSectorCeilings
          
          'Render all visible sidedef uppers
          RenderSidedefsUpper
          
          'Render all visible sidedef lowers
          RenderSidedefsLower
          
          'Render all visible sidedef middles
          'This does not include windows (transparent middle textures)
          RenderSidedefsMiddle
          
          'Enable transparency
          D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
          D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 1
          
          'Render all visible windows
          RenderSidedefsWindows
          
          'Render things?
          If (ShowThings) Then
               
               'Things need clockwise backface culling
               D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
               
               'Render all things and windows
               RenderThings
          End If
     End If
     
     'Everything else needs no backface culling
     D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
     
     'Bilinear texture filtering
     SetTextureFilters True
     
     'Render the texts
     RenderTexts
     
     'Check if panel must be shown
     If (ShowInfo) Then
          
          'Render info panel
          RenderInfoPanel
          
          'Render the texts
          RenderInfoTexts
     End If
     
     'Normal texture filtering
     SetTextureFilters False
     
     'Render the crosshair
     RenderCrosshair
     
     
     '===== End scene
     D3DD.EndScene
     
     '===== Present scene
     D3DD.Present
End Sub

Private Function RunTextureSelect(ByVal CurrentTexture As String, ByVal UseFlats As Boolean) As String
     On Error GoTo Leave3DMode
     Dim ErrNumber As Long
     Dim ErrDesc As String
     Dim TextRect As SRECT
     Dim MousePoint As POINT
     Dim LastCursorUpdate As Long
     
     'Get mouse coords
     GetCursorPos MousePoint
     
     'Keep coords
     TLastX = MousePoint.x
     TLastY = MousePoint.y
     
     'Determine area
     With TextRect
          .left = 0
          .top = 0.9
          .right = 0.6
          .bottom = 1
     End With
     
     'Make the text
     Set r_texdesc = VertexBufferFromText(TEXTURE_DESC, TextRect, ALIGN_RIGHT, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, TEXT_SIZE)
     r_numtexdescfaces = Len(TEXTURE_DESC) * 4 - 2
     
     'Initiate defaults
     InitTextureSelect CurrentTexture, UseFlats
     
     'Current texture
     SelectedName = CurrentTexture
     CreateSelectedTextureText
     
     'We are now selecting a texture/flat
     TextureSelecting = True
     ThingSelecting = False
     
     'Initiate the textures field
     CreateTexturePreviews
     
     Do
          'Calculate time
          CurrentTime = timeExactTime
          FrameTime = CurrentTime - LastTime
          LastTime = CurrentTime
          
          'Poll the mouse
          PollMouse
          
          'Mouse events can do anything, also terminating 3d mode
          If Not Running3D Then Exit Do
          
          'Check for cursor update
          If ((LastCursorUpdate + CURSOR_FLASH_INTERVAL) < CurrentTime) Then
               
               'Change the cursor
               ShowTextCursor = Not ShowTextCursor
               
               'Recreate the text buffer
               CreateSelectedTextureText
               
               'Keep the time
               LastCursorUpdate = CurrentTime
          End If
          
          '===== Start scene
          D3DD.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Val(Config("palette")("CLR_BACKGROUND")), 1, 0
          D3DD.BeginScene
          
          'Apply Matrices
          D3DD.SetTransform D3DTS_PROJECTION, matrixProject
          D3DD.SetTransform D3DTS_VIEW, matrixView
          D3DD.SetTransform D3DTS_WORLD, matrixWorld
          
          'Beginning settings
          D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 0
          D3DD.SetRenderState D3DRS_LIGHTING, 0
          D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
          
          'Texture filtering as configured
          SetTextureFilters False
          
          'Render selection background
          RenderSelection
          
          'Render texture previews
          RenderTexturePreviews
          
          'Bilinear texture filtering
          SetTextureFilters True
          
          'Render texts
          CreateSelectedTextureText
          RenderTextureTexts
          
          'Texture filtering as configured
          SetTextureFilters False
          
          'Render the mouse
          RenderMouse
          
          
          '===== End scene
          D3DD.EndScene
          
          '===== Present scene
          D3DD.Present
          
          'Process messages
          DoEvents
          
          'Delay frames
          If (DelayVideoFrames) Then Sleep 50 Else Sleep 10
          
          'Next fame input will be done again
          IgnoreInput = False
          
     'Continue until 'dialog' closed
     Loop While TextureSelecting And Running3D
     
     
Leave3DMode:
     
     'Check if 3D Mode was not terminated
     If (Running3D = True) Then
          
          'Check if not quit nicely
          If (TextureSelecting = True) Or (Err.number <> 0) Then
               
               'Keep error
               ErrNumber = Err.number
               ErrDesc = Err.Description
               
               'Clean up directx mode
               Running3D = False
               TextureSelecting = False
               CleanUp3DMode
               
               'Display error if not device lost error
               If (ErrNumber <> -2005530520) Then MsgBox "Error " & ErrNumber & " in RunTextureSelect: " & ErrDesc, vbCritical
               
               'Yes, cancel this
               TextureSelectCancelled = True
          End If
          
          'Check if cancelled
          If TextureSelectCancelled Then
               
               'Keep original texture
               RunTextureSelect = CurrentTexture
          Else
               
               'Check if we should get complete texture name
               If (Val(Config("autocompletetex")) <> 0) And (TextureSelectedIndex >= 0) Then
                    
                    'Return new texture
                    RunTextureSelect = curitemnames(TextureSelectedIndex)
               Else
                    
                    'Use typed name
                    RunTextureSelect = SelectedName
               End If
          End If
     Else
          
          'Clear errors
          Err.Clear
          
          'Keep original texture
          RunTextureSelect = CurrentTexture
     End If
     
     'Clean up arrays
     Erase itemnames()
     Erase useditemnames()
     Erase curitemnames()
     Set collection = Nothing
End Function

Private Function RunThingSelect(ByVal CurrentThingType As Long) As Long
     On Error GoTo Leave3DMode
     Dim ErrNumber As Long
     Dim ErrDesc As String
     Dim TextRect As SRECT
     Dim MousePoint As POINT
     Dim LastCursorUpdate As Long
     
     'Get mouse coords
     GetCursorPos MousePoint
     
     'Keep coords
     TLastX = MousePoint.x
     TLastY = MousePoint.y
     
     'Determine area
     With TextRect
          .left = 0
          .top = 0.9
          .right = 0.37
          .bottom = 1
     End With
     
     'Make the text
     Set r_texdesc = VertexBufferFromText(THING_DESC, TextRect, ALIGN_RIGHT, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, TEXT_SIZE)
     r_numtexdescfaces = Len(THING_DESC) * 4 - 2
     
     'Initiate defaults
     InitThingSelect CurrentThingType
     
     'Current texture
     SelectedName = CStr(CurrentThingType)
     CreateSelectedThingText
     
     'We are now selecting a texture/flat/thing
     TextureSelecting = True
     ThingSelecting = True
     
     'Initiate the things field
     CreateTexturePreviews
     
     Do
          'Calculate time
          CurrentTime = timeExactTime
          FrameTime = CurrentTime - LastTime
          LastTime = CurrentTime
          
          'Poll the mouse
          PollMouse
          
          'Mouse events can do anything, also terminating 3d mode
          If Not Running3D Then Exit Do
          
          'Check for cursor update
          If ((LastCursorUpdate + CURSOR_FLASH_INTERVAL) < CurrentTime) Then
               
               'Change the cursor
               ShowTextCursor = Not ShowTextCursor
               
               'Recreate the text buffer
               CreateSelectedTextureText
               
               'Keep the time
               LastCursorUpdate = CurrentTime
          End If
          
          '===== Start scene
          D3DD.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Val(Config("palette")("CLR_BACKGROUND")), 1, 0
          D3DD.BeginScene
          
          'Apply Matrices
          D3DD.SetTransform D3DTS_PROJECTION, matrixProject
          D3DD.SetTransform D3DTS_VIEW, matrixView
          D3DD.SetTransform D3DTS_WORLD, matrixWorld
          
          'Beginning settings
          D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
          D3DD.SetRenderState D3DRS_ALPHATESTENABLE, 1
          D3DD.SetRenderState D3DRS_LIGHTING, 0
          D3DD.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
          
          'Texture filtering as configured
          SetTextureFilters False
          
          'Render selection background
          RenderSelection
          
          'Render texture previews
          RenderTexturePreviews
          
          'Bilinear texture filtering
          SetTextureFilters True
          
          'Render texts
          CreateSelectedThingText
          RenderTextureTexts
          
          'Texture filtering as configured
          SetTextureFilters False
          
          'Render the mouse
          RenderMouse
          
          
          '===== End scene
          D3DD.EndScene
          
          '===== Present scene
          D3DD.Present
          
          'Process messages
          DoEvents
          
          'Delay frames
          If (DelayVideoFrames) Then Sleep 50 Else Sleep 10
          
          'Next fame input will be done again
          IgnoreInput = False
          
     'Continue until 'dialog' closed
     Loop While TextureSelecting And Running3D
     
     
Leave3DMode:
     
     'Check if 3D Mode was not terminated
     If (Running3D = True) Then
          
          'Check if not quit nicely
          If (TextureSelecting = True) Or (Err.number <> 0) Then
               
               'Keep error
               ErrNumber = Err.number
               ErrDesc = Err.Description
               
               'Clean up directx mode
               Running3D = False
               TextureSelecting = False
               CleanUp3DMode
               
               'Display error if not device lost error
               If (ErrNumber <> -2005530520) Then MsgBox "Error " & ErrNumber & " in RunTextureSelect: " & ErrDesc, vbCritical
               
               'Yes, cancel this
               TextureSelectCancelled = True
          End If
          
          'Check if cancelled
          If TextureSelectCancelled Then
               
               'Keep original thing
               RunThingSelect = CurrentThingType
          Else
               
               'Use selected thing
               RunThingSelect = Val(SelectedName)
          End If
     Else
          
          'Clear errors
          Err.Clear
          
          'Keep original thing
          RunThingSelect = CurrentThingType
     End If
     
     'Clean up arrays
     Erase itemnames()
     Erase useditemnames()
     Erase curitemnames()
     Set collection = Nothing
End Function


Private Sub SelectCeilingTexture(ByVal s As Long)
     On Error Resume Next
     Dim Texture As clsImage
     Dim texturename As String
     Dim ss As Long
     
     'Make undo
     CreateUndo "change ceiling texture", UGRP_CEILINGTEXTURECHANGE, s, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
          
          'Free the mouse
          FreeMouse
          
          'Select texture with standard dialog
          texturename = SelectFlat(sectors(s).tceiling, frmMain)
          
          'Recapture the mouse
          CaptureMouse
     Else
          
          'Select texture with rendered dialog
          texturename = UCase$(RunTextureSelect(sectors(s).tceiling, True))
     End If
     
     'Leave immediately when 3D Mode terminated
     If (Running3D = False) Then Exit Sub
     
     'Apply texture
     sectors(s).tceiling = texturename
     
     'Check if the texture is known
     If flats.Exists(texturename) Then
          
          'Get texture object
          Set Texture = flats(texturename)
          
          'Show message
          ShowMainText "Ceiling texture:  " & texturename & "  " & Texture.width & "x" & Texture.height
          
          'Clean up
          Set Texture = Nothing
     Else
          
          'Show message
          ShowMainText "Ceiling texture:  " & texturename
     End If
     
     'Go for all subsectors
     For ss = 0 To (numsubsectors - 1)
          
          'Check if subsector is part of this sector
          If (m_subsectors(ss).sector = s) Then
               
               'Remove vertexbuffer so it will be recreated
               d_SubSectorCeilings(ss) = False
               Set SubSectorCeilings(ss) = Nothing
          End If
     Next ss
     Set i_SectorCeilings(s) = Nothing
End Sub

Private Sub SelectFloorTexture(ByVal s As Long)
     On Error Resume Next
     Dim Texture As clsImage
     Dim texturename As String
     Dim ss As Long
     
     'Make undo
     CreateUndo "change floor texture", UGRP_FLOORTEXTURECHANGE, s, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
          
          'Free the mouse
          FreeMouse
          
          'Select texture with standard dialog
          texturename = SelectFlat(sectors(s).tfloor, frmMain)
          
          'Recapture the mouse
          CaptureMouse
     Else
          
          'Select texture with rendered dialog
          texturename = UCase$(RunTextureSelect(sectors(s).tfloor, True))
     End If
     
     'Leave immediately when 3D Mode terminated
     If (Running3D = False) Then Exit Sub
     
     'Apply texture
     sectors(s).tfloor = texturename
     
     'Check if the texture is known
     If flats.Exists(texturename) Then
          
          'Get texture object
          Set Texture = flats(texturename)
          
          'Show message
          ShowMainText "Floor texture:  " & texturename & "  " & Texture.width & "x" & Texture.height
          
          'Clean up
          Set Texture = Nothing
     Else
          
          'Show message
          ShowMainText "Floor texture:  " & texturename
     End If
     
     'Go for all subsectors
     For ss = 0 To (numsubsectors - 1)
          
          'Check if subsector is part of this sector
          If (m_subsectors(ss).sector = s) Then
               
               'Remove vertexbuffer so it will be recreated
               d_SubSectorFloors(ss) = False
               Set SubSectorFloors(ss) = Nothing
          End If
     Next ss
     Set i_SectorFloors(s) = Nothing
End Sub

Private Sub SelectLowerTexture(ByVal sd As Long)
     On Error Resume Next
     Dim Texture As clsImage
     Dim texturename As String
     
     'Make undo
     CreateUndo "change lower texture", UGRP_LOWERTEXTURECHANGE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
          
          'Free the mouse
          FreeMouse
          
          'Select texture with standard dialog
          texturename = SelectTexture(sidedefs(sd).lower, frmMain)
          
          'Recapture the mouse
          CaptureMouse
     Else
          
          'Select texture with rendered dialog
          texturename = UCase$(RunTextureSelect(sidedefs(sd).lower, False))
     End If
     
     'Leave immediately when 3D Mode terminated
     If (Running3D = False) Then Exit Sub
     
     'Apply texture
     sidedefs(sd).lower = texturename
     
     'Check if the texture is known
     If textures.Exists(texturename) Then
          
          'Get texture object
          Set Texture = textures(texturename)
          
          'Show message
          ShowMainText "Lower texture:  " & texturename & "  " & Texture.width & "x" & Texture.height
          
          'Clean up
          Set Texture = Nothing
     Else
          
          'Show message
          ShowMainText "Lower texture:  " & texturename
     End If
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefLower(sd) = False
     Set SidedefLower(sd) = Nothing
     Set i_SidedefLower(sd) = Nothing
End Sub

Private Sub SelectMiddleTexture(ByVal sd As Long)
     On Error Resume Next
     Dim Texture As clsImage
     Dim texturename As String
     
     'Make undo
     CreateUndo "change middle texture", UGRP_MIDDLETEXTURECHANGE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
          
          'Free the mouse
          FreeMouse
          
          'Select texture with standard dialog
          texturename = SelectTexture(sidedefs(sd).middle, frmMain)
          
          'Recapture the mouse
          CaptureMouse
     Else
          
          'Select texture with rendered dialog
          texturename = UCase$(RunTextureSelect(sidedefs(sd).middle, False))
     End If
     
     'Leave immediately when 3D Mode terminated
     If (Running3D = False) Then Exit Sub
     
     'Apply texture
     sidedefs(sd).middle = texturename
     
     'Check if the texture is known
     If textures.Exists(texturename) Then
          
          'Get texture object
          Set Texture = textures(texturename)
          
          'Show message
          ShowMainText "Middle texture:  " & texturename & "  " & Texture.width & "x" & Texture.height
          
          'Clean up
          Set Texture = Nothing
     Else
          
          'Show message
          ShowMainText "Middle texture:  " & texturename
     End If
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefMiddle(sd) = False
     Set SidedefMiddle(sd) = Nothing
     Set i_SidedefMiddle(sd) = Nothing
End Sub

Private Sub SelectNewThing(ByVal th As Long)
     On Error Resume Next
     Dim thingtype As Long
     
'     'Check if this thing is editable in 3D Mode
'     'It must have a sprite
'     If (TestSpriteForThingType(things(th).thing)) Then
          
          'Make undo
          CreateUndo "change thing", UGRP_THINGCHANGE, th, True
          
          'Map changed
          mapchanged = True
          
          'Check if in windowed mode
          If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
               
               'Free the mouse
               FreeMouse
               
               'Select thing with standard dialog
               thingtype = SelectThing(things(th).thing, frmMain)
               
               'Recapture the mouse
               CaptureMouse
          Else
               
               'Select thing with rendered dialog
               thingtype = RunThingSelect(things(th).thing)
          End If
          
          'Leave immediately when 3D Mode terminated
          If (Running3D = False) Then Exit Sub
          
          'Apply thing
          things(th).thing = thingtype
          
          'Show message
          ShowMainText "Thing:  " & GetThingTypeDesc(thingtype) & " (" & CStr(thingtype) & ")"
          
          'Update the thing
          UpdateThingCategory th
          UpdateThingImageColor th
          UpdateThingSize th
'     Else
'
'          'Impossible!
'          Beep
'     End If
End Sub


Private Sub SelectUpperTexture(ByVal sd As Long)
     On Error Resume Next
     Dim Texture As clsImage
     Dim texturename As String
     
     'Make undo
     CreateUndo "change upper texture", UGRP_UPPERTEXTURECHANGE, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Check if in windowed mode
     If (Val(Config("windowedvideo")) <> 0) And (Val(Config("standardtexturebrowse")) <> 0) Then
          
          'Free the mouse
          FreeMouse
          
          'Select texture with standard dialog
          texturename = SelectTexture(sidedefs(sd).upper, frmMain)
          
          'Recapture the mouse
          CaptureMouse
     Else
          
          'Select texture with rendered dialog
          texturename = UCase$(RunTextureSelect(sidedefs(sd).upper, False))
     End If
     
     'Leave immediately when 3D Mode terminated
     If (Running3D = False) Then Exit Sub
     
     'Apply texture
     sidedefs(sd).upper = texturename
     
     'Check if the texture is known
     If textures.Exists(texturename) Then
          
          'Get texture object
          Set Texture = textures(texturename)
          
          'Show message
          ShowMainText "Upper texture:  " & texturename & "  " & Texture.width & "x" & Texture.height
          
          'Clean up
          Set Texture = Nothing
     Else
          
          'Show message
          ShowMainText "Upper texture:  " & texturename
     End If
     
     'Remove vertexbuffer so it will be recreated
     d_SidedefUpper(sd) = False
     Set SidedefUpper(sd) = Nothing
     Set i_SidedefUpper(sd) = Nothing
End Sub

Public Sub Stop3DMode()
     
     'Clean up the 3D Mode
     CleanUp3DMode
     
     'Leave when form is unloaded
     If (IsLoaded(frmMain) = False) Then Exit Sub
     
     'Enable editing
     frmMain.picMap.Enabled = True
     
     'Restore window
     If (Val(Config("windowedvideo")) = 0) Then
          frmMain.WindowState = PreviousWindowstate
          frmMain.Form_Resize
          frmMain.Refresh
     End If
     
     'Show status dialog
     frmMain.SetFocus
     frmMain.Refresh
     
     'Check if a map is still open
     If (mapfilename <> "") Then
          
          'Check if the position thing is within bounds
          If (PositionThing >= 0) And (PositionThing < numthings) Then
               
               'Check if the position thing is correct
               If (things(PositionThing).thing = mapconfig("start3dmode")) Then
                    
                    'Update the 3D start thing
                    things(PositionThing).x = Position.x / MAP_RENDER_SCALE
                    things(PositionThing).y = -Position.y / MAP_RENDER_SCALE
                    
                    'Wrap angle
                    While ((HAngle - pi * 0.5) >= pi * 2): HAngle = HAngle - pi * 2: Wend
                    While ((HAngle - pi * 0.5) < 0): HAngle = HAngle + pi * 2: Wend
                    
                    'Apply angle to thing
                    things(PositionThing).angle = HAngle * PiDiv - 90
                    UpdateThingImageColor PositionThing
               End If
          End If
     End If
     
     'Deselect all
     RemoveSelection False
     
     'Hide panels
     frmMain.HideLinedefInfo
     frmMain.HideSectorInfo
     frmMain.HideThingInfo
     frmMain.HideVertexInfo
     
     'Initialize the map screen
     TerminateMapRenderer
     InitializeMapRenderer frmMain.picMap
     
     'Set the viewport
     ChangeView ViewLeft, ViewTop, ViewZoom
     
     'Unload status dialog
     Screen.MousePointer = vbNormal
     Unload frmStatus
     Set frmStatus = Nothing
End Sub

Private Sub UpdateInfoPanel()
     Dim Obj As Long
     Dim ObjType As ENUM_OBJECTTYPES
     Dim ObjSpot As D3DVECTOR
     
     'Get the targeted object
     ObjType = PickAimedObject(Obj, ObjSpot)
     
     'Show info for each different object
     Select Case ObjType
          Case OBJ_SIDEDEFUPPER
               ShowInfoText 0, "Linedef: " & sidedefs(Obj).linedef
               If (linedefs(sidedefs(Obj).linedef).s1 = Obj) Then ShowInfoText 1, "Side: Front" Else ShowInfoText 1, "Side: Back"
               ShowInfoText 2, "Part: Upper"
               ShowInfoText 3, "Length: " & LinedefLength(sidedefs(Obj).linedef)
               ShowInfoText 4, "Sector: " & sidedefs(Obj).sector
               ShowInfoText 5, "Front Height: " & LinedefFrontHeight(sidedefs(Obj).linedef)
               ShowInfoText 6, "Back Height: " & LinedefBackHeight(sidedefs(Obj).linedef)
               ShowInfoText 7, "Part Height: " & SidedefUpperHeight(Obj)
               ShowInfoText 8, "Texture: " & sidedefs(Obj).upper
               ShowInfoText 9, "Offset: " & sidedefs(Obj).tx & ", " & sidedefs(Obj).ty
               
          Case OBJ_SIDEDEFMIDDLE
               ShowInfoText 0, "Linedef: " & sidedefs(Obj).linedef
               If (linedefs(sidedefs(Obj).linedef).s1 = Obj) Then ShowInfoText 1, "Side: Front" Else ShowInfoText 1, "Side: Back"
               ShowInfoText 2, "Part: Middle"
               ShowInfoText 3, "Length: " & LinedefLength(sidedefs(Obj).linedef)
               ShowInfoText 4, "Sector: " & sidedefs(Obj).sector
               ShowInfoText 5, "Front Height: " & LinedefFrontHeight(sidedefs(Obj).linedef)
               ShowInfoText 6, "Back Height: " & LinedefBackHeight(sidedefs(Obj).linedef)
               ShowInfoText 7, "Part Height: " & SidedefMiddleHeight(Obj)
               ShowInfoText 8, "Texture: " & sidedefs(Obj).middle
               ShowInfoText 9, "Offset: " & sidedefs(Obj).tx & ", " & sidedefs(Obj).ty
               
          Case OBJ_SIDEDEFLOWER
               ShowInfoText 0, "Linedef: " & sidedefs(Obj).linedef
               If (linedefs(sidedefs(Obj).linedef).s1 = Obj) Then ShowInfoText 1, "Side: Front" Else ShowInfoText 1, "Side: Back"
               ShowInfoText 2, "Part: Lower"
               ShowInfoText 3, "Length: " & LinedefLength(sidedefs(Obj).linedef)
               ShowInfoText 4, "Sector: " & sidedefs(Obj).sector
               ShowInfoText 5, "Front Height: " & LinedefFrontHeight(sidedefs(Obj).linedef)
               ShowInfoText 6, "Back Height: " & LinedefBackHeight(sidedefs(Obj).linedef)
               ShowInfoText 7, "Part Height: " & SidedefLowerHeight(Obj)
               ShowInfoText 8, "Texture: " & sidedefs(Obj).lower
               ShowInfoText 9, "Offset: " & sidedefs(Obj).tx & ", " & sidedefs(Obj).ty
               
          Case OBJ_SECTORCEILING
               ShowInfoText 0, "Sector: " & Obj
               ShowInfoText 1, "Part: Ceiling"
               ShowInfoText 2, "Height: " & sectors(Obj).hceiling
               ShowInfoText 3, "Texture: " & sectors(Obj).tceiling
               ShowInfoText 4, "Brightness: " & sectors(Obj).Brightness
               ShowInfoText 5, ""
               ShowInfoText 6, ""
               ShowInfoText 7, ""
               ShowInfoText 8, ""
               ShowInfoText 9, ""
               
          Case OBJ_SECTORFLOOR
               ShowInfoText 0, "Sector: " & Obj
               ShowInfoText 1, "Part: Floor"
               ShowInfoText 2, "Height: " & sectors(Obj).hfloor
               ShowInfoText 3, "Texture: " & sectors(Obj).tfloor
               ShowInfoText 4, "Brightness: " & sectors(Obj).Brightness
               ShowInfoText 5, ""
               ShowInfoText 6, ""
               ShowInfoText 7, ""
               ShowInfoText 8, ""
               ShowInfoText 9, ""
               
     End Select
End Sub

Private Sub ShowInfoText(ByVal line As Long, ByVal Text As String)
     Const TextSize As Single = 2
     Dim TextRect As SRECT
     
     'Check if setting
     If (LenB(Text) > 0) Then
          
          'Determine column
          If (line < 5) Then
               
               'Determine area
               With TextRect
                    .left = 0.1 '0.04
                    .top = 0.62 + 0.071 * line
                    .right = 0.5
                    .bottom = 0.5
               End With
          Else
               
               'Determine area
               With TextRect
                    .left = 0.52
                    .top = 0.62 + 0.071 * (line - 5)
                    .right = 0.96
                    .bottom = 0.96
               End With
          End If
          
          'Set the text
          Set r_infotexts(line) = VertexBufferFromText(Text, TextRect, ALIGN_LEFT, ALIGN_TOP, INFO_C1, INFO_C2, INFO_C3, INFO_C4, TextSize)
          r_numinfotextfaces(line) = Len(Text) * 4 - 2
     Else
          
          'Erase
          Set r_infotexts(line) = Nothing
          r_numinfotextfaces(line) = 0
     End If
End Sub

Public Sub ShowMainText(ByVal Text As String, Optional ByVal SubText As String)
     Const TextSize As Single = 2
     Const SubTextSize As Single = 1
     Dim TextRect As SRECT
     
     'Determine area
     With TextRect
          .left = 0
          .top = 0.86
          .right = 1
          .bottom = 1
     End With
     
     'Set the text
     Set r_maintext = VertexBufferFromText(Text, TextRect, ALIGN_CENTER, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, TextSize)
     r_nummaintextfaces = Len(Text) * 4 - 2
     
     'Check if setting sub text as well
     If (SubText <> "") Then
          
          'Determine area
          With TextRect
               .left = 0
               .top = 0.94
               .right = 1
               .bottom = 1
          End With
          
          'Set the text
          Set r_subtext = VertexBufferFromText(SubText, TextRect, ALIGN_CENTER, ALIGN_MIDDLE, TEXT_C1, TEXT_C2, TEXT_C3, TEXT_C4, SubTextSize)
          r_numsubtextfaces = Len(SubText) * 4 - 2
     Else
          
          'Clear sub text
          Set r_subtext = Nothing
     End If
     
     'Set timeout
     TextRemoveTime = GetTickCount + TEXT_SHOWTIME
End Sub

Public Function Start3DMode() As Boolean
     On Error GoTo Error3DMode
     Dim ErrNum As Long, ErrDesc As String
     Dim MapRect As RECT
     Dim Material As D3DMATERIAL9
     Dim aspect As Single
     Dim line As Long
     
     'Copy settings
     c_belowceiling = Val(Config("belowceiling"))
     c_mixresource = Val(mapconfig("mixtexturesflats"))
     c_movespeed = Val(Config("movespeed"))
     c_videowidth = Val(Config("videowidth"))
     c_videoheight = Val(Config("videoheight"))
     c_videoviewdistance = Val(Config("videoviewdistance"))
     c_videofov = Val(Config("videofov"))
     c_invertmousey = Val(Config("invertmousey"))
     c_mousespeed = Val(Config("mousespeed"))
     
     line = 1
     
     line = 4
     
     'Normal processing
     DelayVideoFrames = False
     
     'No movement
     Key3DForward = False
     Key3DBackward = False
     Key3DStrafeLeft = False
     Key3DStrafeRight = False
     
     'No discards yet
     r_numdiscards = 0
     r_numprevsidedefs = 0
     r_numprevsubsectors = 0
     
     line = 5
     
     'Terminate 2D renderer
     TerminateMapRenderer
     Set frmMain.picMap.Picture = Nothing
     
     'Start DirectX
     StartDirectX
     
     line = 6
     
     'Check if gamma or brightness are adjusted
     If (Config("videogamma") <> 0) Or (Config("videobrightness") <> 0) Then
          
          'Make gamma correction
          If (Config("videogamma") > 0) Then
               
               'Positive gamma
               CreateGammaCorrection (Config("videogamma") / 10), Config("videobrightness")
          Else
               
               'Negative gamma
               CreateGammaCorrection ((100 + Config("videogamma")) / 100), Config("videobrightness")
          End If
     End If
     
     line = 7
     
     'Initialize matrices
     MatrixIdentity matrixProject
     MatrixIdentity matrixView
     MatrixIdentity matrixWorld
     
     line = 8
     
     'Determine aspect
     If (Val(Config("videoaspect")) <> 0) Then
          
          'Use fixed aspect
          aspect = VIDEO_FIXED_ASPECT
     Else
          
          'Use resolution aspect
          aspect = VideoParams.BackBufferHeight / VideoParams.BackBufferWidth
     End If
     
     line = 9
     
     'Make projection matrix
     MatrixPerspectiveFovLH matrixProject, Config("videofov") / 57.29577951, 1 / aspect, 0.01, Config("videoviewdistance") * MAP_RENDER_SCALE
     
     line = 10
     
     'Create material
     With Material
          .Ambient.r = 1
          .Ambient.g = 1
          .Ambient.b = 1
          .Ambient.a = 1
          .diffuse.r = 1
          .diffuse.g = 1
          .diffuse.b = 1
          .diffuse.a = 1
     End With
     
     line = 11
     
     'Set material
     D3DD.SetMaterial Material
     
     line = 12
     
     'Set default renderstates
     D3DD.SetRenderState D3DRS_ALPHABLENDENABLE, 1
     D3DD.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATER
     D3DD.SetRenderState D3DRS_ALPHAREF, &HFFFF0000
     D3DD.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
     D3DD.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
     D3DD.SetRenderState D3DRS_CLIPPING, 1
     D3DD.SetRenderState D3DRS_DITHERENABLE, 1
     D3DD.SetRenderState D3DRS_LIGHTING, 1
     D3DD.SetRenderState D3DRS_SHADEMODE, 2 'D3DSHADE_GOURAUD
     D3DD.SetRenderState D3DRS_SPECULARENABLE, 0
     D3DD.SetRenderState D3DRS_ZENABLE, 1
     D3DD.SetRenderState D3DRS_ZWRITEENABLE, 1
     D3DD.SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
     D3DD.SetSamplerState 0, D3DSAMP_BORDERCOLOR, D3DColorMake(0, 0, 0, 0)
     
     line = 13
     
     'Check if fog should be set
     If (Val(Config("showfog")) = vbChecked) Then
          D3DD.SetRenderState D3DRS_FOGCOLOR, Val(Config("palette")("CLR_BACKGROUND"))
          D3DD.SetRenderState D3DRS_FOGTABLEMODE, 2 'D3DFOG_EXP2
          D3DD.SetRenderState D3DRS_FOGSTART, CVL(MKS(1))
          D3DD.SetRenderState D3DRS_FOGEND, CVL(MKS(20))
          D3DD.SetRenderState D3DRS_RANGEFOGENABLE, 0
          D3DD.SetRenderState D3DRS_FOGENABLE, 1
     Else
          D3DD.SetRenderState D3DRS_FOGENABLE, 0
     End If
     
     line = 15
     
     'Disable menus
     frmMain.mnuEdit.Enabled = False
     frmMain.mnuFile.Enabled = False
     frmMain.mnuHelp.Enabled = False
     frmMain.mnuLines.Enabled = False
     frmMain.mnuPrefabs.Enabled = False
     frmMain.mnuScripts.Enabled = False
     frmMain.mnuSectors.Enabled = False
     frmMain.mnuThings.Enabled = False
     frmMain.mnuTools.Enabled = False
     frmMain.mnuVertices.Enabled = False
     
     line = 16
     
     'Create lighting tables
     MakeLightingTables
     
     line = 17
     
     'Make the crosshair
     MakeCrosshair
     
     line = 18
     
     'Make the text font
     MakeTextFont
     
     line = 19
     
     'Make some extra textures
     MakeExtraTextures
     
     line = 20
     
     'Make the info panel
     MakeInfoPanel
     
     line = 21
     
     'Make thing resources
     MakeThingResources
     
     line = 22
     
     'Capture the mouse
     CaptureMouse
     
     'Success
     Running3D = True
     Start3DMode = True
     Exit Function
     
Error3DMode:
     
     'Keep error message
     ErrDesc = Err.Description
     ErrNum = Err.number
     
     'Terminate
     CleanUp3DMode
     
     'Restore window
     frmMain.WindowState = PreviousWindowstate
     frmMain.Form_Resize
     frmMain.Refresh
     
     'Show error
     MsgBox "Error " & ErrNum & " while initializing 3D mode: " & ErrDesc, vbCritical ' & " at section " & line, vbCritical
End Function

Private Sub SwitchSidedefPolygon(ByRef SidedefPoly() As VERTEX)
     Dim SwitchVertex As VERTEX
     
     'Switch vertices
     SwitchVertex.x = SidedefPoly(0).x
     SwitchVertex.y = SidedefPoly(0).y
     SwitchVertex.Z = SidedefPoly(0).Z
     
     SidedefPoly(0).x = SidedefPoly(2).x
     SidedefPoly(0).y = SidedefPoly(2).y
     SidedefPoly(0).Z = SidedefPoly(2).Z
     
     SidedefPoly(2).x = SwitchVertex.x
     SidedefPoly(2).y = SwitchVertex.y
     SidedefPoly(2).Z = SwitchVertex.Z
     
     SwitchVertex.x = SidedefPoly(1).x
     SwitchVertex.y = SidedefPoly(1).y
     SwitchVertex.Z = SidedefPoly(1).Z
     
     SidedefPoly(1).x = SidedefPoly(3).x
     SidedefPoly(1).y = SidedefPoly(3).y
     SidedefPoly(1).Z = SidedefPoly(3).Z
     
     SidedefPoly(3).x = SwitchVertex.x
     SidedefPoly(3).y = SwitchVertex.y
     SidedefPoly(3).Z = SwitchVertex.Z
End Sub

Public Function TestStructures(ByVal File As clsWAD) As Boolean
     On Error GoTo errorhandler
     Dim lumpindex As Long
     
     'Get the VERTEXES lump
     lumpindex = FindLumpIndex(File, 1, "VERTEXES")
     If (lumpindex = 0) Then Exit Function
     If (File.LumpSize(lumpindex) <= 0) Then Exit Function
     
     'Get the SEGS lump
     lumpindex = FindLumpIndex(File, 1, "SEGS")
     If (lumpindex = 0) Then Exit Function
     If (File.LumpSize(lumpindex) <= 0) Then Exit Function
     
     'Get the SSECTORS lump
     lumpindex = FindLumpIndex(File, 1, "SSECTORS")
     If (lumpindex = 0) Then Exit Function
     If (File.LumpSize(lumpindex) <= 0) Then Exit Function
     
     'Get the NODES lump
     lumpindex = FindLumpIndex(File, 1, "NODES")
     If (lumpindex = 0) Then Exit Function
     If (File.LumpSize(lumpindex) <= 0) Then Exit Function
     
     'Done here
     TestStructures = True
     Exit Function
     
errorhandler:
     
     'Show error
     MsgBox "Error " & Err.number & " during structures testing: " & Err.Description, vbCritical
End Function

Private Sub ToggleLowerUnpegged(ByVal sd As Long)
     Dim ld As Long
     Dim sd2 As Long
     
     'Make undo
     CreateUndo "texture alignment", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Get the other side too
     If (linedefs(ld).s1 = sd) Then sd2 = linedefs(ld).s2 Else sd2 = linedefs(ld).s1
     
     'Check if the flag is set
     If (linedefs(ld).Flags And LDF_LOWERUNPEGGED) = LDF_LOWERUNPEGGED Then
          
          'Remove the flag
          linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_LOWERUNPEGGED
     Else
          
          'Add the flag
          linedefs(ld).Flags = linedefs(ld).Flags Or LDF_LOWERUNPEGGED
     End If
     
     'Show message
     ShowMainText "Lower unpegged:  " & OnOff((linedefs(ld).Flags And LDF_LOWERUNPEGGED))
     
     'Remove vertexbuffers so the sides will be recreated
     d_SidedefLower(sd) = False
     d_SidedefMiddle(sd) = False
     d_SidedefUpper(sd) = False
     d_SidedefLower(sd2) = False
     d_SidedefMiddle(sd2) = False
     d_SidedefUpper(sd2) = False
     Set SidedefLower(sd) = Nothing
     Set SidedefMiddle(sd) = Nothing
     Set SidedefUpper(sd) = Nothing
     Set SidedefLower(sd2) = Nothing
     Set SidedefMiddle(sd2) = Nothing
     Set SidedefUpper(sd2) = Nothing
End Sub

Private Sub ToggleMiddleTexture(ByVal sd As Long)
     Dim ld As Long
     Dim sd2 As Long
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Verify that the sidedef is on a doublesided linedef
     If ((linedefs(ld).s1 = sd) And (linedefs(ld).s2 > -1)) Or _
        ((linedefs(ld).s2 = sd) And (linedefs(ld).s1 > -1)) Then
          
          'Make undo
          CreateUndo "toggle middle texture", UGRP_TOGGLEMIDDLETEXTURE, sd, True
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Get the other side too
          If (linedefs(ld).s1 = sd) Then sd2 = linedefs(ld).s2 Else sd2 = linedefs(ld).s1
          
          'Check if a texture must be removed
          If IsTextureName(sidedefs(sd).middle) Then
               
               'Remove the middle texture
               sidedefs(sd).middle = "-"
               sidedefs(sd2).middle = "-"
               
               'Show message
               ShowMainText "Middle textures removed"
          Else
               
               'Check if no default texture set
               If (Trim$(Config("defaulttexture")("middle")) = "-") Or _
                  (Trim$(Config("defaulttexture")("middle")) = "") Then
                    
                    'Add default middle texture
                    sidedefs(sd).middle = alltextures.Keys(0)
                    sidedefs(sd2).middle = alltextures.Keys(0)
               Else
                    
                    'Check if no valid default texture set
                    If (textures.Exists(Trim$(Config("defaulttexture")("middle")))) Then
                         
                         'Add specified default texture
                         sidedefs(sd).middle = Trim$(UCase$(Config("defaulttexture")("middle")))
                         sidedefs(sd2).middle = Trim$(UCase$(Config("defaulttexture")("middle")))
                    Else
                         
                         'Add default middle texture
                         sidedefs(sd).middle = alltextures.Keys(0)
                         sidedefs(sd2).middle = alltextures.Keys(0)
                    End If
               End If
               
               'Show message
               ShowMainText "Middle textures added"
          End If
          
          'Remove vertexbuffer so it will be recreated
          d_SidedefMiddle(sd) = False
          d_SidedefMiddle(sd2) = False
          Set SidedefMiddle(sd) = Nothing
          Set SidedefMiddle(sd2) = Nothing
          Set i_SidedefMiddle(sd) = Nothing
          Set i_SidedefMiddle(sd2) = Nothing
     End If
End Sub

Private Sub ToggleUpperUnpegged(ByVal sd As Long)
     Dim ld As Long
     Dim sd2 As Long
     
     'Make undo
     CreateUndo "texture alignment", UGRP_TEXTUREALIGNMENT, sd, True
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Get the linedef
     ld = sidedefs(sd).linedef
     
     'Get the other side too
     If (linedefs(ld).s1 = sd) Then sd2 = linedefs(ld).s2 Else sd2 = linedefs(ld).s1
     
     'Check if the flag is set
     If (linedefs(ld).Flags And LDF_UPPERUNPEGGED) = LDF_UPPERUNPEGGED Then
          
          'Remove the flag
          linedefs(ld).Flags = linedefs(ld).Flags And Not LDF_UPPERUNPEGGED
     Else
          
          'Add the flag
          linedefs(ld).Flags = linedefs(ld).Flags Or LDF_UPPERUNPEGGED
     End If
     
     'Show message
     ShowMainText "Upper unpegged:  " & OnOff((linedefs(ld).Flags And LDF_UPPERUNPEGGED))
     
     'Remove vertexbuffers so the sides will be recreated
     d_SidedefLower(sd) = False
     d_SidedefMiddle(sd) = False
     d_SidedefUpper(sd) = False
     d_SidedefLower(sd2) = False
     d_SidedefMiddle(sd2) = False
     d_SidedefUpper(sd2) = False
     Set SidedefLower(sd) = Nothing
     Set SidedefMiddle(sd) = Nothing
     Set SidedefUpper(sd) = Nothing
     Set SidedefLower(sd2) = Nothing
     Set SidedefMiddle(sd2) = Nothing
     Set SidedefUpper(sd2) = Nothing
End Sub

Public Sub UpdateLinesSectorsInfo()
     Dim Obj As Long
     Dim ObjType As ENUM_OBJECTTYPES
     Dim ObjSpot As D3DVECTOR
     
     'Get the targeted object
     ObjType = PickAimedObject(Obj, ObjSpot)
     
     'Check if changed
     If (LastInfoObject <> Obj) Or (LastInfoObjectType <> ObjType) Then
          
          'Check if previous type was different
          If (LastInfoObjectType <> ObjType) Then
               
               'What object type was that?
               Select Case LastInfoObjectType
                    
                    'Lines
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER
                         
                         'Hide lines info
                         frmMain.HideLinedefInfo
                         
                    'Sectors
                    Case OBJ_SECTORCEILING, OBJ_SECTORFLOOR
                         
                         'Hide sectors info
                         frmMain.HideSectorInfo
                         
                    'Things
                    Case OBJ_THING
                         
                         'Hide sectors info
                         frmMain.HideThingInfo
                         
               End Select
          End If
          
          'Check if anything
          If (ObjType <> OBJ_NOTHING) Then
               
               'Show info for each different object
               Select Case ObjType
                    
                    'Lines
                    Case OBJ_SIDEDEFUPPER, OBJ_SIDEDEFMIDDLE, OBJ_SIDEDEFLOWER
                         
                         'Show lines info
                         frmMain.ShowLinedefInfo sidedefs(Obj).linedef
                         
                    'Sectors
                    Case OBJ_SECTORCEILING, OBJ_SECTORFLOOR
                         
                         'Show sectors info
                         frmMain.ShowSectorInfo Obj
                         
                    'Thing
                    Case OBJ_THING
                         
                         'Show sectors info
                         frmMain.ShowThingInfo Obj
                         
               End Select
          End If
          
          'Keep object info
          LastInfoObject = Obj
          LastInfoObjectType = ObjType
     End If
End Sub

Public Function VertexBufferFromText( _
                ByRef Text As String, ByRef Position As SRECT, ByVal hAlign As ENUM_HALIGN, _
                ByVal vAlign As ENUM_VALIGN, ByVal c_lt As Long, ByVal c_rt As Long, _
                ByVal c_lb As Long, ByVal c_rb As Long, ByVal CharScale As Single) As Direct3DVertexBuffer9
     
     'Dim TextVertex() As VertexFlat
     Dim TextVertex(1 To TEXT_MAXCHARS * 4) As TLVERTEX
     Dim TextVertices As Long
     
     'Reserve memory for vertices
     TextVertices = Len(Text) * 4
     
     'Create text in memory
     CreateText UCase$(Text), Position, hAlign, vAlign, c_lt, c_rt, c_lb, c_rb, CharScale, TextVertex(1), VideoParams.BackBufferWidth, VideoParams.BackBufferHeight
     
     'Create vertex buffer
     Set VertexBufferFromText = D3DD.CreateVertexBuffer(TextVertices * TLVERTEXSTRIDE, D3DUSAGE_WRITEONLY Or D3DUSAGE_DYNAMIC, TLVERTEXFVF, D3DPOOL_DEFAULT)
     
     'Load vertices into buffer
     VertexBufferFromText.SetData 0, TextVertices * TLVERTEXSTRIDE, VarPtr(TextVertex(1)), 0
End Function

     
     


