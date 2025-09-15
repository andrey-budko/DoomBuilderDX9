Attribute VB_Name = "modErrorCheck"
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


'Ways to solve problems
Public Enum ENUM_ERRORSOLVEFUNCTIONS
     ESF_NONE
     ESF_ERASEUPPERTEXTURE         '1=sidedef
     ESF_ERASEMIDDLETEXTURE        '1=sidedef
     ESF_ERASELOWERTEXTURE         '1=sidedef
     ESF_FLIPSIDEDEFS              '1=linedef
     ESF_FLAGTWOSIDED              '1=linedef
     ESF_UNFLAGTWOSIDED            '1=linedef
     ESF_MERGELINES                '1=linedef 2=linedef
     ESF_DEFAULTLOWERTEXTURE       '1=sidedef
     ESF_DEFAULTMIDDLETEXTURE      '1=sidedef
     ESF_DEFAULTUPPERTEXTURE       '1=sidedef
     ESF_MERGEVERTICES             '1=vertex 2=vertex
     ESF_DELETELINEDEF             '1=linedef
     ESF_DELETETHING               '1=thing
End Enum

'Type for error
Public Type FOUNDERROR
     Title As String
     Description As String
     critical As Boolean
     viewtype As ENUM_EDITMODE
     viewindex As Long
     solvetype As ENUM_ERRORSOLVEFUNCTIONS
     solveindex1 As Long
     solveindex2 As Long
End Type


'Array for found errors
Public FoundErrors() As FOUNDERROR
Public NumFoundErrors As Long

'Settings
Public IgnoreWarningsOption As Integer
Public InvalidTexturesOption As Integer
Public LineErrorsOption As Integer
Public MissingTexturesOption As Integer
Public PlayerStartsOption As Integer
Public UnclosedSectorsOption As Integer
Public VertexErrorsOption As Integer
Public ZeroLengthLinesOption As Integer
Public ThingErrorsOption As Integer


'API Declarations
Public Declare Function TestStuckedThing Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal numlines As Long, ByVal x As Long, ByVal y As Long, ByVal radius As Long) As Long
Public Declare Function OverlappingUnselectedVertex Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByVal sourcevertex As Long) As Long
Public Declare Function OverlappingUnselectedLinedef Lib "builder.dll" (ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal sourceline As Long) As Long
Public Declare Function TestUnclosedSector Lib "builder.dll" (ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByVal numlinedefs As Long, ByVal sector As Long) As Long


Private Sub AddFoundError(ByRef newerror As FOUNDERROR)
     
     'Allocate more memory?
     If (UBound(FoundErrors) = NumFoundErrors) Then
          
          'Allocate new array
          ReDim Preserve FoundErrors(NumFoundErrors + 10)
     End If
     
     'Add to array
     FoundErrors(NumFoundErrors) = newerror
     NumFoundErrors = NumFoundErrors + 1
End Sub


Public Sub ClearFoundErrors()
     
     'Erase array
     ReDim FoundErrors(0)
     NumFoundErrors = 0
End Sub


Private Sub DoClosedSectorsCheck()
     Dim ld As Long
     Dim sc As Long
     Dim fe As FOUNDERROR
     
     'Go for all sectors
     For sc = 0 To numsectors - 1
          
          'Check this sector
          ld = TestUnclosedSector(vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, sc)
          
          'DEBUG
          If (ld > -1) Then
               
               'Create error report
               fe.Title = "Sector " & sc & " is not closed"
               fe.Description = "The lines (sidedefs) that make up this sector do not entirely enclose the sector. This leak may cause visual artifacts and/or unexpected behaviour. The opening in the sector is shown near the highlighted vertex when you select this error."
               fe.critical = True
               fe.viewtype = EM_VERTICES
               fe.viewindex = ld
               fe.solvetype = ESF_NONE
               AddFoundError fe
          End If
     Next sc
     
     'Reset linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
End Sub

Public Function DoErrorChecks() As Boolean
     
     'Initialize array
     ClearFoundErrors
     
     'Check for player starts
     If (PlayerStartsOption = vbChecked) Then DoPlayerStartsCheck
     
     'Check for vertex errors
     If (VertexErrorsOption = vbChecked) Then DoVertexChecks
     
     'Check for line errors
     If (LineErrorsOption = vbChecked) Then DoLineChecks
     
     'Check for sector errors
     If (UnclosedSectorsOption = vbChecked) Then DoClosedSectorsCheck
     
     'Check for zero-length lines
     If (ZeroLengthLinesOption = vbChecked) Then DoZeroLengthLinesCheck
     
     'Check for invalid textures
     If (InvalidTexturesOption = vbChecked) Then DoInvalidTexturesCheck
     
     'Check for missing textures
     If (MissingTexturesOption = vbChecked) Then DoMissingTexturesCheck
     
     'Check for thing errors
     If (ThingErrorsOption = vbChecked) Then DoThingChecks
     
     'Errors found?
     DoErrorChecks = (NumFoundErrors > 0)
End Function


Private Sub DoMissingTexturesCheck()
     Dim ld As Long
     Dim fe As FOUNDERROR
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if linedef has a front sidedef
          If (linedefs(ld).s1 > -1) Then
               
               'Check if upper texture is required
               If RequiresS1Upper(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Upper) Then
                    
                    'Missing upper texture
                    fe.Title = "Linedef " & ld & " requires an upper texture on its front side"
                    fe.Description = "There is no upper texture on the front side of this linedef. If you do not add a texture here, the ceiling of the sector at the back of this line will leak through this part of the wall. Use the Fix button to add the current default build texture here."
                    fe.critical = False
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTUPPERTEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
               
               'Check if middle texture is required
               If RequiresS1Middle(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Middle) Then
                    
                    'Missing middle texture
                    fe.Title = "Linedef " & ld & " requires a middle texture on its front side"
                    fe.Description = "There is no middle texture on the front side of this linedef. The middle texture is required on this line because the line is single-sided and will cause a Hall of Mirrors problem when it has no texture. Use the Fix button to add the current default build texture here."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTMIDDLETEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
               
               'Check if lower texture is required
               If RequiresS1Lower(ld) And Not IsTextureName(sidedefs(linedefs(ld).s1).Lower) Then
                    
                    'Missing lower texture
                    fe.Title = "Linedef " & ld & " requires a lower texture on its front side"
                    fe.Description = "There is no lower texture on the front side of this linedef. If you do not add a texture here, the floor of the sector at the back of this line will leak through this part of the wall. Use the Fix button to add the current default build texture here."
                    fe.critical = False
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTLOWERTEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
          End If
          
          'Check if linedef has a back sidedef
          If (linedefs(ld).s2 > -1) Then
               
               'Check if upper texture is required
               If RequiresS2Upper(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Upper) Then
                    
                    'Missing upper texture
                    fe.Title = "Linedef " & ld & " requires an upper texture on its back side"
                    fe.Description = "There is no upper texture on the back side of this linedef. If you do not add a texture here, the ceiling of the sector at the front of this line will leak through this part of the wall. Use the Fix button to add the current default build texture here."
                    fe.critical = False
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTUPPERTEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
               
               'Check if middle texture is required
               If RequiresS2Middle(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Middle) Then
                    
                    'Missing middle texture
                    fe.Title = "Linedef " & ld & " requires a middle texture on its back side"
                    fe.Description = "There is no middle texture on the back side of this linedef. The middle texture is required on this line because the line is single-sided and will cause a Hall of Mirrors problem when it has no texture. Use the Fix button to add the current default build texture here."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTMIDDLETEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
               
               'Check if lower texture is required
               If RequiresS2Lower(ld) And Not IsTextureName(sidedefs(linedefs(ld).s2).Lower) Then
                    
                    'Missing lower texture
                    fe.Title = "Linedef " & ld & " requires a lower texture on its back side"
                    fe.Description = "There is no lower texture on the back side of this linedef. If you do not add a texture here, the floor of the sector at the front of this line will leak through this part of the wall. Use the Fix button to add the current default build texture here."
                    fe.critical = False
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_DEFAULTLOWERTEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
          End If
     Next ld
End Sub

Private Sub DoInvalidTexturesCheck()
     Dim ld As Long
     Dim fe As FOUNDERROR
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if linedef has a front sidedef
          If (linedefs(ld).s1 > -1) Then
               
               'Check if upper texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s1).Upper)) = 0) Then
                    
                    'Missing upper texture
                    fe.Title = "Linedef " & ld & " has an invalid upper texture on its front side"
                    fe.Description = "The upper texture on the front side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASEUPPERTEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
               
               'Check if middle texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s1).Middle)) = 0) Then
                    
                    'Missing middle texture
                    fe.Title = "Linedef " & ld & " has an invalid middle texture on its front side"
                    fe.Description = "The middle texture on the front side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASEMIDDLETEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
               
               'Check if lower texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s1).Lower)) = 0) Then
                    
                    'Missing lower texture
                    fe.Title = "Linedef " & ld & " has an invalid lower texture on its front side"
                    fe.Description = "The lower texture on the front side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASELOWERTEXTURE
                    fe.solveindex1 = linedefs(ld).s1
                    AddFoundError fe
               End If
          End If
          
          'Check if linedef has a back sidedef
          If (linedefs(ld).s2 > -1) Then
               
               'Check if upper texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s2).Upper)) = 0) Then
                    
                    'Missing upper texture
                    fe.Title = "Linedef " & ld & " has an invalid upper texture on its back side"
                    fe.Description = "The upper texture on the back side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASEUPPERTEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
               
               'Check if middle texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s2).Middle)) = 0) Then
                    
                    'Missing middle texture
                    fe.Title = "Linedef " & ld & " has an invalid middle texture on its back side"
                    fe.Description = "The middle texture on the back side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASEMIDDLETEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
               
               'Check if lower texture is invalid
               If (LenB(Trim$(sidedefs(linedefs(ld).s2).Lower)) = 0) Then
                    
                    'Missing lower texture
                    fe.Title = "Linedef " & ld & " has an invalid lower texture on its back side"
                    fe.Description = "The lower texture on the back side of this linedef is empty. " & _
                                     "If you do not intend to add a texture here, use the dash " & _
                                     "symbol - to indicate no texture. Click the Fix button below " & _
                                     "to solve this problem by replacing it with a dash symbol."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_ERASELOWERTEXTURE
                    fe.solveindex1 = linedefs(ld).s2
                    AddFoundError fe
               End If
          End If
     Next ld
End Sub


Private Sub DoPlayerStartsCheck()
     Dim t As Long
     Dim Player1Starts As Long
     Dim Player2Starts As Long
     Dim Player3Starts As Long
     Dim Player4Starts As Long
     Dim MultiStarts As Long
     Dim fe As FOUNDERROR
     
     'A must must have at least the 4 player starts and
     'at least 4 multiplayer starts.
     
     'Go for all Things
     For t = 0 To numthings - 1
          
          'Check what this is
          Select Case things(t).thing
               Case 1: Player1Starts = Player1Starts + 1
               Case 2: Player2Starts = Player2Starts + 1
               Case 3: Player3Starts = Player3Starts + 1
               Case 4: Player4Starts = Player4Starts + 1
               Case 11: MultiStarts = MultiStarts + 1
          End Select
     Next t
     
     'Now create error report
     
     'No player 1 starts found?
     If (Player1Starts < 1) Then
          
          'No player 1 starts!
          fe.Title = "No player 1 start found"
          fe.Description = "You need at least 1 player 1 start Thing to be able to player your map in singleplayer mode and some engines also require this Thing in other modes.  Fix it yourself."
          fe.critical = True
          fe.viewtype = 0
          fe.viewindex = 0
          fe.solvetype = ESF_NONE
          AddFoundError fe
     End If
     
     'No player 2 starts found?
     If (Player2Starts < 1) Then
          
          'No player 2 starts!
          fe.Title = "No player 2 start found"
          fe.Description = "You need at least 1 player 2 start Thing to be able to player your map in cooperative mode. However, if you do not add a player 2 start Thing, you can still play your map in singleplayer and deathmatch modes."
          fe.critical = False
          fe.viewtype = 0
          fe.viewindex = 0
          fe.solvetype = ESF_NONE
          AddFoundError fe
     End If
     
     'No player 3 starts found?
     If (Player3Starts < 1) Then
          
          'No player 3 starts!
          fe.Title = "No player 3 start found"
          fe.Description = "You need at least 1 player 3 start Thing to be able to player your map in cooperative mode. However, if you do not add a player 3 start Thing, you can still play your map in singleplayer and deathmatch modes."
          fe.critical = False
          fe.viewtype = 0
          fe.viewindex = 0
          fe.solvetype = ESF_NONE
          AddFoundError fe
     End If
     
     'No player 4 starts found?
     If (Player4Starts < 1) Then
          
          'No player 4 starts!
          fe.Title = "No player 4 start found"
          fe.Description = "You need at least 1 player 4 start Thing to be able to player your map in cooperative mode. However, if you do not add a player 4 start Thing, you can still play your map in singleplayer and deathmatch modes."
          fe.critical = False
          fe.viewtype = 0
          fe.viewindex = 0
          fe.solvetype = ESF_NONE
          AddFoundError fe
     End If
     
     'Not enough multiplayer starts found?
     If (MultiStarts < 4) Then
          
          'Not enough multiplayer starts!
          fe.Title = "Not enough multiplayer starts found"
          fe.Description = "You need at least 4 multiplayer start Things to be able to player your map in deathmatch mode. If you attempt to play deathmatch with less than 4 player starts, not everyone may be able to connect."
          fe.critical = False
          fe.viewtype = 0
          fe.viewindex = 0
          fe.solvetype = ESF_NONE
          AddFoundError fe
     End If
End Sub


Private Sub DoVertexChecks()
     Dim v As Long
     Dim ov As Long
     Dim fe As FOUNDERROR
     
     'Go for all vertices
     'This is done backwards so that the highest vertex number
     'is taken for the error preview (this ensures its selection will be visible)
     For v = numvertexes - 1 To 0 Step -1
          
          Do
               'Find overlapping vertex
               ov = OverlappingUnselectedVertex(vertexes(0), numvertexes, v)
               
               'Found anything?
               If (ov > -1) Then
                    
                    'Create error report
                    fe.Title = "Vertex " & v & " overlaps with vertex " & ov
                    fe.Description = "Vertex " & v & " is at the same coordinates as vertex " & ov & ". This may cause problems to the nodebuilder and Doom engine. You can move the vertices manually, or click the Fix button below to merge the vertices now."
                    fe.critical = True
                    fe.viewtype = EM_VERTICES
                    fe.viewindex = v
                    fe.solvetype = ESF_MERGEVERTICES
                    fe.solveindex1 = v
                    fe.solveindex2 = ov
                    AddFoundError fe
                    
                    'Select the vertex
                    vertexes(ov).selected = 1
               End If
               
          'Continue until no more found
          Loop Until (ov = -1)
          
          'Done with this vertex, select it
          vertexes(v).selected = 1
     Next v
     
     'Reset vertex selections
     ResetSelections things(0), 0, linedefs(0), 0, vertexes(0), numvertexes, VarPtr(sectors(0)), 0
End Sub

Private Sub DoLineChecks()
     Dim ld As Long
     Dim old As Long
     Dim fe As FOUNDERROR
     
     'Go for all lines
     'This is done backwards so that the highest linedef number
     'is taken for the error preview (this ensures its selection will be visible)
     For ld = numlinedefs - 1 To 0 Step -1
          
          'Check if linedef has no sides
          If (linedefs(ld).s1 = -1) And (linedefs(ld).s2 = -1) Then
               
               'Create error report
               fe.Title = "Linedef " & ld & " has no sides"
               fe.Description = "This linedef has no sides. A line must have at least one (front) side and optionally a back side. You can add a side manually at the line properties dialog."
               fe.critical = True
               fe.viewtype = EM_LINES
               fe.viewindex = ld
               fe.solvetype = ESF_NONE
               AddFoundError fe
               
          'Check if linedef has a back side, but no front side
          ElseIf (linedefs(ld).s1 = -1) And (linedefs(ld).s2 <> -1) Then
               
               'Create error report
               fe.Title = "Linedef " & ld & " only has a back side"
               fe.Description = "This linedef has a back side, but no front side. A line must have at least a front side and optionally a back side. You can add a front side manually or click the Fix button below to flip the sidedefs."
               fe.critical = True
               fe.viewtype = EM_LINES
               fe.viewindex = ld
               fe.solvetype = ESF_FLIPSIDEDEFS
               fe.solveindex1 = ld
               AddFoundError fe
               
          'Check if this line is marked single sided
          ElseIf (linedefs(ld).Flags And LDF_TWOSIDED) = 0 Then
               
               'It may not have a second side
               If (linedefs(ld).s2 <> -1) Then
                    
                    'Create error report
                    fe.Title = "Linedef " & ld & " is singlesided with a back side"
                    fe.Description = "This linedef is marked as singlesided, but has a front and back side. Singlesided lines may only have a front side. You can remove the back side manually or click the Fix button below to mark the line as doublesided."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_FLAGTWOSIDED
                    fe.solveindex1 = ld
                    AddFoundError fe
               End If
               
          'Check if this line is marked double sided
          ElseIf (linedefs(ld).Flags And LDF_TWOSIDED) = LDF_TWOSIDED Then
               
               'It must have a second side
               If (linedefs(ld).s2 = -1) Then
                    
                    'Create error report
                    fe.Title = "Linedef " & ld & " is doublesided without back side"
                    fe.Description = "This linedef is marked as doublesided, but has no back side. Doublesided lines must have a front and a back side. You can add a back side manually or click the Fix button below to mark the line as singlesided."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_UNFLAGTWOSIDED
                    fe.solveindex1 = ld
                    AddFoundError fe
               End If
          End If
          
          
          'Now find all overlapping lines
          Do
               'Find overlapping linedef
               old = OverlappingUnselectedLinedef(linedefs(0), numlinedefs, ld)
               
               'Found anything?
               If (old > -1) Then
                    
                    'Create error report
                    fe.Title = "Linedef " & ld & " overlaps with linedef " & old
                    fe.Description = "Linedef " & ld & " uses the same vertices as linedef " & old & ". This may cause problems to the nodebuilder and Doom engine. You can click the Fix button below to merge the two lines, but they may not merge flawless when their sector references are incorrect."
                    fe.critical = True
                    fe.viewtype = EM_LINES
                    fe.viewindex = ld
                    fe.solvetype = ESF_MERGELINES
                    fe.solveindex1 = ld
                    fe.solveindex2 = old
                    AddFoundError fe
                    
                    'Select the linedef
                    linedefs(old).selected = 1
               End If
               
          'Continue until no more found
          Loop Until (old = -1)
          
          
          'Done with this linedef, select it
          linedefs(ld).selected = 1
          
     Next ld
     
     'Reset linedef selections
     ResetSelections things(0), 0, linedefs(0), numlinedefs, vertexes(0), 0, VarPtr(sectors(0)), 0
End Sub


Private Sub DoThingChecks()
     Dim th As Long
     Dim oth As Long
     Dim othblock As Long
     Dim fe As FOUNDERROR
     Dim thrad As Long
     Dim othrad As Long
     Dim therr As Long
     Dim result As Long
     Dim tr As RECT
     Dim sr As RECT
     Dim flagsmask1 As Long
     Dim flagsmask2 As Long
     
     'Get the flags masks
     flagsmask1 = mapconfig("thingflagsmask1")
     flagsmask2 = mapconfig("thingflagsmask2")
     If (flagsmask2 = 0) Then flagsmask2 = flagsmask1
     
     'Go for all things
     For th = 0 To numthings - 1
          
          'Get thing error level
          therr = GetThingError(things(th).thing)
          
          'No problems so far
          result = 0
          
          'Check if this thing should be checked if its outside
          If (therr > 0) Then
               
               'Check the sector number at thing coordinates
               result = IntersectSector(things(th).x, -things(th).y, vertexes(0), linedefs(0), VarPtr(sidedefs(0)), numlinedefs, 0)
               If (result = -1) Then
                    
                    'Create error report
                    fe.Title = "Thing " & th & " (" & GetThingTypeDesc(things(th).thing) & ") is outside the map"
                    fe.Description = "This thing is outside the map. It is most likely that this thing will not be seen and/or cannot be used. You can move the thing inside the map or click the Fix button to remove this thing."
                    fe.critical = False
                    fe.viewtype = EM_THINGS
                    fe.viewindex = th
                    fe.solvetype = ESF_DELETETHING
                    fe.solveindex1 = th
                    AddFoundError fe
               End If
          End If
          
          'Check if this thing should be checked if its stucked
          If (therr > 1) And (result <> -1) Then
               
               'Radius to use for testing
               thrad = GetThingTestRadius(things(th).size)
               
               'Check if this thing collides with a line
               result = TestStuckedThing(vertexes(0), linedefs(0), numlinedefs, things(th).x, things(th).y, thrad)
               If (result <> -1) Then
                    
                    'Create error report
                    fe.Title = "Thing " & th & " (" & GetThingTypeDesc(things(th).thing) & ") is stucked in linedef " & result
                    fe.Description = "This thing is probably stucked in a linedef. It is most likely that this thing will not be able to move due to the collision. To solve this collision, simply move the thing away from the line."
                    fe.critical = False
                    fe.viewtype = EM_THINGS
                    fe.viewindex = th
                    fe.solvetype = ESF_NONE
                    AddFoundError fe
               End If
               
               'Make rectangle from thing
               tr.left = things(th).x - thrad
               tr.right = things(th).x + thrad
               tr.top = things(th).y - thrad
               tr.bottom = things(th).y + thrad
               
               'Go for all other things
               For oth = 0 To numthings - 1
                    
                    'Not the same thing?
                    If (oth <> th) Then
                         
                         'Check if this thing is blocking
                         othblock = GetThingBlocking(things(oth).thing)
                         
                         'Check if blocking in any way
                         If (othblock > 0) Then
                              
                              'Compare thing flags by masks
                              If (((things(th).Flags And flagsmask1) And (things(oth).Flags And flagsmask1)) <> 0) And _
                                 (((things(th).Flags And flagsmask2) And (things(oth).Flags And flagsmask2)) <> 0) Then
                                   
                                   'TODO: Add support for True-Height checking
                                   
                                   'Radius to use for testing
                                   othrad = GetThingTestRadius(things(oth).size)
                                   
                                   'Make rectangle from other thing
                                   sr.left = things(oth).x - othrad
                                   sr.right = things(oth).x + othrad
                                   sr.top = things(oth).y - othrad
                                   sr.bottom = things(oth).y + othrad
                                   
                                   'Check if rectangles collide
                                   If (sr.left < tr.right) And (sr.right > tr.left) And _
                                      (sr.top < tr.bottom) And (sr.bottom > tr.top) Then
                                        
                                        'Collision!
                                        'Create error report
                                        fe.Title = "Thing " & th & " (" & GetThingTypeDesc(things(th).thing) & ") is stucked in thing " & oth & " (" & GetThingTypeDesc(things(oth).thing) & ")"
                                        fe.Description = "This thing is probably stucked in another thing. It is most likely that these things will not be able to move due to the collision. To solve this collision, simply move the things away from each other."
                                        fe.critical = False
                                        fe.viewtype = EM_THINGS
                                        fe.viewindex = th
                                        fe.solvetype = ESF_NONE
                                        AddFoundError fe
                                   End If
                              End If
                         End If
                    End If
               Next oth
          End If
     Next th
End Sub



Private Sub DoZeroLengthLinesCheck()
     Dim ld As Long
     Dim fe As FOUNDERROR
     
     'Go for all linedefs
     For ld = 0 To numlinedefs - 1
          
          'Check if linedef refers to same vertices
          If (linedefs(ld).v1 = linedefs(ld).v2) Then
               
               'Zero-length by same vertices
               fe.Title = "Linedef " & ld & " references vertex " & linedefs(ld).v1 & " twice"
               fe.Description = "This linedef ends at the same vertices it starts at, thus the linedef is zero-length. You can use the ""Fix Zero-Length Linedefs"" tool from the Tools menu to remove this linedef, or click the Fix button to remove this linedef now."
               fe.critical = True
               fe.viewtype = EM_VERTICES
               fe.viewindex = linedefs(ld).v1
               fe.solvetype = ESF_DELETELINEDEF
               fe.solveindex1 = ld
               AddFoundError fe
          Else
               
               'Check if both vertices are at same location
               If (CLng(vertexes(linedefs(ld).v1).x) = CLng(vertexes(linedefs(ld).v2).x)) And _
                  (CLng(vertexes(linedefs(ld).v1).y) = CLng(vertexes(linedefs(ld).v2).y)) Then
                    
                    'Zero-length by two vertices
                    fe.Title = "Linedef " & ld & " references overlapping vertices " & linedefs(ld).v1 & " and " & linedefs(ld).v2
                    fe.Description = "The vertex where this linedef ends is at the same coordinates as the vertex where this linedef starts at, thus the linedef is zero-length. You can use the ""Fix Zero-Length Linedefs"" tool from the Tools menu to remove this linedef, or click the Fix button to remove this linedef now."
                    fe.critical = True
                    fe.viewtype = EM_VERTICES
                    If (linedefs(ld).v1 > linedefs(ld).v2) Then fe.viewindex = linedefs(ld).v1 Else fe.viewindex = linedefs(ld).v2
                    fe.solvetype = ESF_DELETELINEDEF
                    fe.solveindex1 = ld
                    AddFoundError fe
               End If
          End If
     Next ld
End Sub


Private Function GetThingTestRadius(ByVal radius As Long) As Long
     
     'Radius to use for testing
     'This is tweaked a bit so the error checker will not whine too much with false alarms
     GetThingTestRadius = Int((CSng(radius) * 0.75) + 0.5)
End Function


