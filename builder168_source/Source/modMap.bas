Attribute VB_Name = "modMap"
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


'Memory allocation
'Public Const DECLARE_THINGS As Long = 50
'Public Const DECLARE_LINEDEFS As Long = 100
'Public Const DECLARE_SIDEDEFS As Long = 200
'Public Const DECLARE_VERTICES As Long = 200
'Public Const DECLARE_SECTORS As Long = 50

'Standard linedef flags
Public Enum ENUM_LINEDEFFLAGS
     LDF_IMPASSIBLE = 1
     LDF_BLOCKMONSTER = 2
     LDF_TWOSIDED = 4
     LDF_UPPERUNPEGGED = 8
     LDF_LOWERUNPEGGED = 16
     LDF_SECRET = 32
     LDF_BLOCKSOUND = 64
     LDF_HIDDEN = 128
     LDF_SHOWN = 256
End Enum

'Map formats
Public Enum ENUM_MAPFORMATS
     MFMT_DOOM = 1
     MFMT_HEXEN = 2
End Enum


'NOTE: The following type structures are optimized for internal use only
'The map structures will be formatted correctly when written to and read from file

'THINGS
Public Type MAPTHING
     tag As Long
     x As Long                'X position
     y As Long                'Y position
     Z As Long                'Z position
     angle As Long            'Degrees
     thing As Long
     Flags As Long
     effect As Long
     arg0 As Long
     arg1 As Long
     arg2 As Long
     arg3 As Long
     arg4 As Long
     
     'Optimization variables
     category As Long         'category
     Color As Long            'Color from palette to render thing with
     image As Long            'Image to be rendered with (depends on type and angle)
     size As Long
     height As Long
     hangs As Long
     selected As Long
     argref0 As Long
     argref1 As Long
     argref2 As Long
     argref3 As Long
     argref4 As Long
     sector As Long           '(only used in 3D Mode)
End Type

'LINEDEFS
Public Type MAPLINEDEF
     v1 As Long               'Start vertex
     v2 As Long               'End vertex
     Flags As Long
     effect As Long           'Action to perform on sector...
     tag As Long              '...with this same tag
     arg0 As Long
     arg1 As Long
     arg2 As Long
     arg3 As Long
     arg4 As Long
     s1 As Long               'Right sidedef
     s2 As Long               'Left sidedef (or -1 for singleside lines)
     
     'Optimization variables
     selected As Long
     argref0 As Long
     argref1 As Long
     argref2 As Long
     argref3 As Long
     argref4 As Long
End Type

'SIDEDEFS
Public Type MAPSIDEDEF
     tx As Long               'Texture X offset
     ty As Long               'Texture Y offset
     Upper As String
     Lower As String
     Middle As String
     sector As Long           'Sector to which this side belongs
     
     'Optimization variables
     linedef As Long          'Linedef on which this sidedef is
     MiddleTop As Long        'Top of middle texture (only used in 3D Mode)
     MiddleBottom As Long     'Bottom of middle texture (only used in 3D Mode)
End Type

'VERTICES
Public Type MAPVERTEX
     x As Single
     y As Single
     
     'Optimization variables
     selected As Long
End Type

'SECTORS
Public Type MAPSECTOR
     hfloor As Long           'Floor height
     hceiling As Long         'Ceiling height
     tfloor As String
     tceiling As String
     Brightness As Long       '0-255
     special As Long
     tag As Long
     
     'Optimization variables
     selected As Long
     visible As Long          '(only used in 3D Mode)
End Type


'API Declarations
Public Declare Sub Rereference_Vertices Lib "builder.dll" (ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal oldref As Long, ByVal newref As Long)
Public Declare Sub Rereference_Sidedefs Lib "builder.dll" (ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal oldref As Long, ByVal newref As Long)
Public Declare Sub Rereference_SidedefsLinedef Lib "builder.dll" (ByVal ptr_sidedefs As Long, ByVal numsidedefs As Long, ByVal oldref As Long, ByVal newref As Long)
Public Declare Sub Rereference_Sectors Lib "builder.dll" (ByVal ptr_sidedefs As Long, ByVal numsidedefs As Long, ByVal oldref As Long, ByVal newref As Long)
Public Declare Function CountSectorSidedefs Lib "builder.dll" (ByVal ptr_sidedefs As Long, ByVal numsidedefs As Long, ByVal sector As Long) As Long
Public Declare Function CountVertexLinedefs Lib "builder.dll" (ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal VERTEX As Long) As Long
Public Declare Sub ExportWavefrontObj Lib "builder.dll" (ByVal filepathname As String, ByRef vertices As MAPVERTEX, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByVal ptr_sectors As Long, ByRef things As MAPTHING, ByVal numvertices As Long, ByVal numlinedefs As Long, ByVal numsidedefs As Long, ByVal numsectors As Long, ByVal numthings As Long)


'Map structure and things
Public things(66000) As MAPTHING
Public linedefs(66000) As MAPLINEDEF
Public sidedefs(66000) As MAPSIDEDEF
Public vertexes(66000) As MAPVERTEX
Public sectors(66000) As MAPSECTOR

'Map counts
Public numthings As Long
Public numlinedefs As Long
Public numsidedefs As Long
Public numvertexes As Long
Public numsectors As Long

'Map preferences and filenames
Public mapchanged As Boolean
Public mapnodeschanged As Boolean
Public maptempfile As String
Public mapfile As String
Public mapfilename As String
Public maplumpname As String
Public mapoldlumpname As String         'When this is set, this lump will be removed on save
Public mapisiwad As Boolean
Public mapgame As String
Public mapsaved As Boolean
Public PositionThing As Long

'Current map
Public MapWAD As clsWAD

'Current extra textures file
Public addwadfile As String
Public addtexdir As String
Public addflatdir As String
Public AddWAD As New clsWAD

'Current IWAD
Public IWAD As New clsWAD

'Temporary WAD
Public TempWAD As clsWAD

'Configuration
Public mapconfig As Dictionary

Private Sub ApplyDecorateThings(ByRef WadFile As clsWAD)
     On Error GoTo ApplyDecorateThingsError
     Dim lumpindex As Long
     Dim DecorateData As String
     Dim ScopeLevel As Long
     Dim IgnoreString As Long
     Dim LastWord As String
     Dim WordsFound As New collection
     Dim itemname As String
     Dim ItemNumber As String
     Dim ItemCategory As String
     Dim ItemWidth As Long
     Dim NewItem As Dictionary
     Dim c As Long
     Dim prevchar As String * 1
     Dim char As String * 1
     
     'Find DECORATE lump index
     lumpindex = FindLumpIndex(WadFile, 1, "DECORATE")
     
     'Found it?
     If (lumpindex > 0) Then
          
          'Get the data
          DecorateData = WadFile.GetLump(lumpindex)
          
          'Remove comments
          RemoveDecorateComments DecorateData
          
          'Replace all newlines with spaces
          DecorateData = Replace$(DecorateData, vbCr, "")
          DecorateData = Replace$(DecorateData, vbLf, " ")
          DecorateData = Replace$(DecorateData, vbTab, " ")
          
          'Go for each char
          For c = 1 To Len(DecorateData)
               
               'Get the char
               prevchar = char
               char = Mid$(DecorateData, c, 1)
               
               'Are we not in a string block?
               If (IgnoreString = 0) Then
                    
                    'Are we not in a scope level to ignore?
                    If (ScopeLevel <= 1) Then
                         
                         'Check if a string starts
                         If (char = """") Then
                              
                              'Now in a string block
                              IgnoreString = True
                         Else
                              
                              'End of word (space)?
                              If (char = " ") Then
                                   
                                   'Do we have a word?
                                   If (LastWord <> "") Then
                                        
                                        'Add to list of found words
                                        WordsFound.Add LastWord
                                        
                                        'In a scope level?
                                        If (ScopeLevel > 0) Then
                                             
                                             'More than 1 word?
                                             If (WordsFound.Count > 1) Then
                                                  
                                                  'Check what the previous word was
                                                  Select Case UCase$(WordsFound(WordsFound.Count - 1))
                                                       
                                                       'DoomEdNum indicates Thing ID
                                                       Case "DOOMEDNUM"
                                                            
                                                            'Get the number
                                                            ItemNumber = Val(WordsFound(WordsFound.Count))
                                                            
                                                            'Clear found words
                                                            Set WordsFound = New collection
                                                       
                                                       'Radius indicates Thing width
                                                       Case "RADIUS"
                                                            
                                                            'Get the width
                                                            ItemWidth = Val(WordsFound(WordsFound.Count))
                                                            
                                                            'Clear found words
                                                            Set WordsFound = New collection
                                                            
                                                       '//$Category indicates Thing category
                                                       Case "//$CATEGORY"
                                                            
                                                            'Get the category
                                                            ItemCategory = WordsFound(WordsFound.Count)
                                                            
                                                            'Clear found words
                                                            Set WordsFound = New collection
                                                  End Select
                                             End If
                                        End If
                                        
                                        'Erase lastword
                                        LastWord = ""
                                   End If
                              Else
                                   
                                   'Add character to word
                                   LastWord = LastWord & char
                              End If
                         End If
                    End If
                    
                    'Opening scope?
                    If (char = "{") Then
                         
                         'Only one word? Then this is the thing name.
                         If (WordsFound.Count = 1) Then
                              
                              'Name found
                              itemname = WordsFound(1)
                              
                         'More than 1 word found?
                         ElseIf (WordsFound.Count > 1) Then
                              
                              'Check the first word
                              Select Case UCase$(WordsFound(1))
                                   
                                   'These have thing name as second word
                                   Case "PICKUP", "BREAKABLE", "PROJECTILE"
                                        
                                        'Get the name
                                        itemname = WordsFound(2)
                                        
                                   'Actor has a more complex definition
                                   Case "ACTOR"
                                        
                                        'Get the name
                                        itemname = WordsFound(2)
                                        
                                        'More than 2 words?
                                        If (WordsFound.Count > 2) Then
                                             
                                             'Third word is : ?
                                             If (WordsFound(3) = ":") Then
                                                  
                                                  'More than four words?
                                                  If (WordsFound.Count > 4) Then
                                                       
                                                       'Fifth word is thing ID
                                                       ItemNumber = Val(WordsFound(5))
                                                  End If
                                             Else
                                                  
                                                  'Third word is thing ID
                                                  ItemNumber = Val(WordsFound(3))
                                             End If
                                        End If
                              End Select
                         End If
                         
                         'Erase lastword
                         LastWord = ""
                         
                         'Clear found words
                         Set WordsFound = New collection
                         
                         'Scope deeper
                         ScopeLevel = ScopeLevel + 1
                         
                    'Closing scope?
                    ElseIf (char = "}") Then
                         
                         'All required parameters found?
                         If (itemname <> "") And (ItemNumber <> "") Then
                              
                              'Give default category name if category not specified
                              If (Trim$(ItemCategory) = "") Then ItemCategory = "DECORATE"
                              
                              'Check if the thing number does not already exist
                              If (GetThingTypeDesc(ItemNumber, "") = "") Then
                                   
                                   'Make a the category if it doesnt exists yet
                                   If (mapconfig("thingtypes").Exists(ItemCategory) = False) Then
                                        
                                        'Make category
                                        Set NewItem = New Dictionary
                                        
                                        'Set the category properties
                                        With NewItem
                                             .Add "color", 8
                                             .Add "arrow", 1
                                             .Add "title", ItemCategory
                                             .Add "width", 0
                                        End With
                                        
                                        'Add category to things
                                        mapconfig("thingtypes").Add ItemCategory, NewItem
                                   End If
                                   
                                   'Check if the thing doesnt exist yet
                                   If (mapconfig("thingtypes")(ItemCategory).Exists(ItemNumber) = False) Then
                                        
                                        'Make the thing object
                                        Set NewItem = New Dictionary
                                        
                                        'Add the thing properties
                                        With NewItem
                                             .Add "title", itemname
                                             .Add "width", ItemWidth
                                        End With
                                        
                                        'Add the thing to category
                                        mapconfig("thingtypes")(ItemCategory).Add ItemNumber, NewItem
                                   End If
                              Else
                                   
                                   'Make error for this DECORATE thing
                                   ErrorLog_Add "WARNING: Thing number " & ItemNumber & " for """ & itemname & """ already used.", False
                                   ErrorLog_Add "Please check your DECORATE lump for errors.", False
                              End If
                         End If
                         
                         'Erase lastword
                         LastWord = ""
                         
                         'Clear settings
                         ItemWidth = 0
                         itemname = ""
                         ItemNumber = ""
                         ItemCategory = ""
                         
                         'Clear found words
                         Set WordsFound = New collection
                         
                         'Scope higher
                         ScopeLevel = ScopeLevel - 1
                    End If
                    
               'In a string block
               Else
                    
                    'Only check for string ending
                    If (char = """") And (prevchar <> "\") Then
                         
                         'End of string block
                         IgnoreString = False
                    End If
               End If
          Next c
     End If
     
     'Leave now
     Exit Sub
     
     
ApplyDecorateThingsError:
     
     'Non-fatal error
     ErrorLog_Add "ERROR " & Err.number & " in ApplyDecorateThings(): " & Err.Description, False
     ErrorLog_Add "Please check your DECORATE lump for errors.", False
End Sub


Public Function ChangeMapOptions(Optional ByVal NewMap As Boolean) As Boolean
     Dim i As Long
     
     'Load the map options dialog
     Load frmMapOptions
     frmMapOptions.Loading = Not NewMap
     
     'Go for al configs
     For i = 0 To (AllGameConfigs.Count - 1)
          
          'Add to list
          frmMapOptions.cmbGameConfig.AddItem AllGameConfigs.Keys(i)
          
          'Check if this one should be selected
          If (StrComp(mapgame, AllGameConfigs.Keys(i), vbTextCompare) = 0) Then frmMapOptions.cmbGameConfig.ListIndex = frmMapOptions.cmbGameConfig.NewIndex
     Next i
     
     'Set the current settings
     With frmMapOptions
          If (.cmbGameConfig.ListIndex < 0) Then .cmbGameConfig.ListIndex = 0
          If Not NewMap Then .txtMapLumpName.Text = maplumpname Else .txtMapLumpName.Text = ""
          .txtWAD.Text = addwadfile
          .txtTexDir = addtexdir
          .txtFlatDir = addflatdir
          .tag = Abs(NewMap)
     End With
     
     'Show the dialog
     frmMapOptions.Loading = False
     frmMapOptions.Show 1, frmMain
     
     'Check the result
     If (frmMapOptions.tag = "1") Then
          
          'Apply the settings
          With frmMapOptions
               maplumpname = .txtMapLumpName.Text
               addwadfile = .txtWAD.Text
               addtexdir = .txtTexDir.Text
               addflatdir = .txtFlatDir.Text
               mapgame = .cmbGameConfig.Text
          End With
          
          'Make full directories
          If (Len(addtexdir) > 0) Then If (right$(addtexdir, 1) <> "\") Then addtexdir = addtexdir & "\"
          If (Len(addflatdir) > 0) Then If (right$(addflatdir, 1) <> "\") Then addflatdir = addflatdir & "\"
          
          'Update caption
          frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
          
          'Return True, settings changed
          ChangeMapOptions = True
     Else
          
          'Return False, nothing changed
          ChangeMapOptions = False
     End If
     
     'Unload the dialog
     Unload frmMapOptions: Set frmMapOptions = Nothing
End Function

Private Sub CopyLumpsByType(ByRef SourceWAD As clsWAD, ByVal SourceHeaderLumpName As String, ByRef TargetWAD As clsWAD, ByVal TargetHeaderLumpName As String, ByVal LumpTypes As ENUM_MAPLUMPTYPES)
     On Local Error GoTo errorhandler
     Dim TargetHeaderIndex As Long
     Dim SourceHeaderIndex As Long
     Dim NextLumpName As String
     Dim lumpindex As Long
     Dim FoundIndex As Long
     Dim MapLumps As Variant
     Dim i As Long
     
     
     'This will copy the map lumps of specified type(s).
     
     
     'Find the map header lump in the TargetWAD
     TargetHeaderIndex = FindLumpIndex(TargetWAD, 1, TargetHeaderLumpName)
     lumpindex = TargetHeaderIndex
     If (TargetHeaderIndex > 0) Then
          
          'Remove that lump
          TargetWAD.DeleteLump lumpindex
          
          'Get next lump name
          NextLumpName = ""
          If (lumpindex <= TargetWAD.LumpCount) Then NextLumpName = TargetWAD.LumpName(lumpindex)
          
          'Continue deleting lumps until no more map-related lumps
          Do Until (GetMapLumpType(NextLumpName) = ML_UNKNOWN)
               
               'Make reliable lumpname
               NextLumpName = Trim$(UCase$(NextLumpName))
               
               'Check if this lump should be copied
               If (GetMapLumpType(NextLumpName) And LumpTypes) > 0 Then
                    
                    'Remove that lump
                    TargetWAD.DeleteLump lumpindex
               Else
                    
                    'Advance to the next lump
                    lumpindex = lumpindex + 1
               End If
               
               'Get next lump name
               NextLumpName = ""
               If (lumpindex <= TargetWAD.LumpCount) Then NextLumpName = TargetWAD.LumpName(lumpindex) Else Exit Do
          Loop
     End If
     
     
     
     'Find the map header lump in the SourceWAD
     SourceHeaderIndex = FindLumpIndex(SourceWAD, 1, SourceHeaderLumpName)
     
     'Check if found
     If (SourceHeaderIndex > 0) Then
          
          'Copy the map header from SourceWAD
          TargetWAD.AddLump SourceWAD.GetLump(SourceHeaderIndex), TargetHeaderLumpName, TargetHeaderIndex
          lumpindex = FindLumpIndex(TargetWAD, 1, TargetHeaderLumpName) + 1
          
          'Go for all lump name as defined by map configuration
          MapLumps = mapconfig("maplumpnames").Keys
          For i = LBound(MapLumps) To UBound(MapLumps)
               
               'Make reliable lumpname
               NextLumpName = Trim$(UCase$(MapLumps(i)))
               
               'Check if this lump should be copied
               If (GetMapLumpType(NextLumpName) And LumpTypes) > 0 Then
                    
                    'Find the lump in the SourceWAD
                    FoundIndex = FindLumpIndex(SourceWAD, SourceHeaderIndex, NextLumpName, UBound(MapLumps) + 2)
                    If (FoundIndex > 0) Then
                         
                         'Copy to TargetWAD
                         TargetWAD.AddLump SourceWAD.GetLump(FoundIndex), NextLumpName, lumpindex
                         lumpindex = lumpindex + 1
                    End If
               End If
          Next i
     End If
     
     'Leave now
     Exit Sub
     
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in CopyLumpsByType(): " & Err.Description, vbCritical
End Sub

Public Function GetWadMapSettings(ByVal Filename As String, ByVal MapName As String) As Dictionary
     Dim ExtPos As Long
     Dim DBSFilename As String
     Dim WadConfigFile As New clsConfiguration
     Dim ConfigStruct As Dictionary
     
     'Make filename for the .dbs file
     ExtPos = InStrRev(Filename, ".")
     DBSFilename = left$(Filename, ExtPos - 1) & ".dbs"
     
     'Check if configuration exists
     If (Dir(DBSFilename) <> "") Then
          
          'Open the configuration
          WadConfigFile.LoadConfiguration DBSFilename
          Set ConfigStruct = WadConfigFile.Root(True)
          Set WadConfigFile = Nothing
          
          'Verify type
          If (ConfigStruct("type") = SETTINGS_CONFIG_TYPE) Then
               
               'Find the structure with map name
               If (ConfigStruct.Exists(UCase$(MapName)) = True) Then
                    
                    'Get structure
                    Set GetWadMapSettings = ConfigStruct(UCase$(MapName))
               End If
          End If
     End If
     
     'Clean up
     Set ConfigStruct = Nothing
End Function

Public Function GetWadSettings(ByVal Filename As String) As Dictionary
     Dim ExtPos As Long
     Dim DBSFilename As String
     Dim WadConfigFile As New clsConfiguration
     Dim ConfigStruct As Dictionary
     
     'Make filename for the .dbs file
     ExtPos = InStrRev(Filename, ".")
     DBSFilename = left$(Filename, ExtPos - 1) & ".dbs"
     
     'Check if configuration exists
     If (Dir(DBSFilename) <> "") Then
          
          'Open the configuration
          WadConfigFile.LoadConfiguration DBSFilename
          Set ConfigStruct = WadConfigFile.Root(True)
          Set WadConfigFile = Nothing
          
          'Verify type
          If (ConfigStruct("type") = SETTINGS_CONFIG_TYPE) Then
               
               'Return structure
               Set GetWadSettings = ConfigStruct
          End If
     End If
     
     'Clean up
     Set ConfigStruct = Nothing
End Function


Public Sub PutCurrentWadMapSettings(ByVal Filename As String)
     Dim ExtPos As Long
     Dim DBSFilename As String
     Dim WadConfigFile As New clsConfiguration
     Dim ConfigStruct As Dictionary
     Dim MapStruct As Dictionary
     
     'Make filename for the .dbs file
     ExtPos = InStrRev(Filename, ".")
     DBSFilename = left$(Filename, ExtPos - 1) & ".dbs"
     
     'Check if configuration exists
     If (Dir(DBSFilename) <> "") Then
          
          'Open the configuration
          WadConfigFile.LoadConfiguration DBSFilename
          Set ConfigStruct = WadConfigFile.Root(True)
     Else
          
          'Create configuration
          WadConfigFile.NewConfiguration
          Set ConfigStruct = WadConfigFile.Root(True)
          
          'Set type
          ConfigStruct("type") = SETTINGS_CONFIG_TYPE
     End If
     
     'Verify type
     If (ConfigStruct("type") = SETTINGS_CONFIG_TYPE) Then
          
          'Set the current configuration
          ConfigStruct("config") = mapconfig("game")
          
          'Find the structure with current map name
          If (ConfigStruct.Exists(UCase$(maplumpname)) = True) Then
               
               'Get structure
               Set MapStruct = ConfigStruct(UCase$(maplumpname))
          Else
               
               'Create structure
               Set MapStruct = New Dictionary
               ConfigStruct.Add UCase$(maplumpname), MapStruct
          End If
          
          'Apply settings
          MapStruct("addwad") = addwadfile
          MapStruct("addtexdir") = addtexdir
          MapStruct("addflatdir") = addflatdir
          
          'Save configuration
          WadConfigFile.SaveConfiguration DBSFilename
     End If
     
     'Clean up
     Set ConfigStruct = Nothing
     Set MapStruct = Nothing
End Sub


Public Sub RemoveDecorateComments(ByRef DecorateData As String)
     Dim IgnoreComment As Boolean
     Dim CommentPos1 As Long
     Dim CommentPos2 As Long
     
     'Replace tabs with spaces
     DecorateData = Replace$(DecorateData, vbTab, " ")
     
     'Remove line comments
     CommentPos1 = 1
     Do
          'Find line comment
          CommentPos1 = InStr(CommentPos1, DecorateData, "//", vbBinaryCompare)
          
          'Found?
          If (CommentPos1 > 0) Then
               
               'Does a $ follow the comment?
               If (Mid$(DecorateData, CommentPos1 + 2, 1) = "$") Then
                    
                    'This must be kept
                    'Advance position
                    CommentPos1 = CommentPos1 + 2
               Else
                    
                    'Find end of line
                    CommentPos2 = InStr(CommentPos1, DecorateData, vbLf, vbBinaryCompare)
                    
                    'Found?
                    If (CommentPos2 > 0) Then
                         
                         'Remove the comment
                         DecorateData = left$(DecorateData, CommentPos1 - 1) & right$(DecorateData, Len(DecorateData) - (CommentPos2 - 2))
                    Else
                         
                         'Remove all up to the end
                         DecorateData = left$(DecorateData, CommentPos1 - 1)
                    End If
               End If
               
               'More may come
               IgnoreComment = True
          Else
               
               'No more line comments
               IgnoreComment = False
          End If
     Loop While IgnoreComment
     
     'Now remove block comments
     CommentPos1 = 1
     Do
          'Find next comment start
          CommentPos1 = InStr(CommentPos1, DecorateData, "/*", vbBinaryCompare)
          
          'Found?
          If (CommentPos1 > 0) Then
               
               'Yes, comment ignoring
               IgnoreComment = True
               
               'Find the end of the comment
               CommentPos2 = InStr(CommentPos1, DecorateData, "*/", vbBinaryCompare)
               
               'Found?
               If (CommentPos2 > 0) Then
                    
                    'Remove the comment
                    DecorateData = left$(DecorateData, CommentPos1 - 1) & right$(DecorateData, Len(DecorateData) - (CommentPos2 + 1))
               Else
                    
                    'Remove all up to the end
                    DecorateData = left$(DecorateData, CommentPos1 - 1)
               End If
          Else
               
               'No luck
               IgnoreComment = False
          End If
     Loop While IgnoreComment
End Sub

Private Function RemoveLumpsByType(ByRef TargetWAD As clsWAD, ByVal TargetHeaderLumpName As String, ByVal LumpTypes As ENUM_MAPLUMPTYPES) As Long
     On Local Error GoTo errorhandler
     Dim TargetHeaderIndex As Long
     Dim SourceHeaderIndex As Long
     Dim NextLumpName As String
     Dim lumpindex As Long
     Dim FoundIndex As Long
     Dim MapLumps As Variant
     Dim i As Long
     
     
     'This will remove the map lumps of specified type(s).
     'Returns the index of the removed header lump or 0 when not found.
     
     
     'Find the map header lump in the TargetWAD
     TargetHeaderIndex = FindLumpIndex(TargetWAD, 1, TargetHeaderLumpName)
     lumpindex = TargetHeaderIndex + 1
     If (TargetHeaderIndex > 0) Then
          
          'Get next lump name
          NextLumpName = ""
          If (lumpindex <= TargetWAD.LumpCount) Then NextLumpName = TargetWAD.LumpName(lumpindex)
          
          'Continue deleting lumps until no more map-related lumps
          Do Until (GetMapLumpType(NextLumpName) = ML_UNKNOWN)
               
               'Make reliable lumpname
               NextLumpName = Trim$(UCase$(NextLumpName))
               
               'Check if this lump should be deleted
               If (GetMapLumpType(NextLumpName) And LumpTypes) > 0 Then
                    
                    'Remove that lump
                    TargetWAD.DeleteLump lumpindex
               Else
                    
                    'Advance to the next lump
                    lumpindex = lumpindex + 1
               End If
               
               'Get next lump name
               NextLumpName = ""
               If (lumpindex <= TargetWAD.LumpCount) Then NextLumpName = TargetWAD.LumpName(lumpindex) Else Exit Do
          Loop
     End If
     
     'Return header index
     RemoveLumpsByType = TargetHeaderIndex
     
     'Leave now
     Exit Function
     
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in RemoveLumpsByType(): " & Err.Description, vbCritical
End Function

Public Function CreateLinedef() As Long
     
     'Check if more memory needs to be allocated
     'If (numlinedefs = UBound(linedefs) - 1) Then
     '
     '     'Allocate memory
     '     ReDim Preserve linedefs(0 To numlinedefs + DECLARE_LINEDEFS)
     'End If
     
     'Increase number of linedefs
     numlinedefs = numlinedefs + 1
     
     'Return the last linedef index
     CreateLinedef = numlinedefs - 1
End Function

Public Sub CreateOptimizations()
     Dim th As Long      'Thing
     Dim ld As Long      'Linedef
     Dim sd As Long      'Sidedef
     Dim xl As Long, yl As Long
     Dim sdref() As Boolean
     Dim SidedefCompression As Long
     Dim SidedefsRemoved As Long
     Dim SectorsRemoved As Long
     
     
     '
     ' === CHECK FOR INVALID VERTEX REFERENCES
     ' === CHECK FOR INVALID SIDEDEF REFERENCES
     ' === COPY DOUBLE REFERENCED SIDEDEFS
     ' === REFERENCE ALL SIDEDEFS BACK TO THEIR LINEDEFS
     '
     
     'Make array for keeping track of references
     If (numsidedefs > 0) Then ReDim sdref(0 To (numsidedefs - 1)) Else ReDim sdref(0 To 0)
     
     'Go for all linedefs
     ld = numlinedefs - 1
     Do While ld >= 0
          
          'Verify vertex 1
          If (linedefs(ld).v1 < 0) Or (linedefs(ld).v1 >= numvertexes) Then
               
               'Report error
               ErrorLog_Add "ERROR: Linedef " & ld & " references invalid vertex " & linedefs(ld).v1 & ". Linedef has been removed.", True
               
               'Remove the linedef
               RemoveLinedef ld, True, False, True
               
          'Verify vertex 2
          ElseIf (linedefs(ld).v2 < 0) Or (linedefs(ld).v2 >= numvertexes) Then
               
               'Report error
               ErrorLog_Add "ERROR: Linedef " & ld & " references invalid vertex " & linedefs(ld).v2 & ". Linedef has been removed.", True
               
               'Remove the linedef
               RemoveLinedef ld, True, False, True
               
          'Verify sidedef 1
          ElseIf (linedefs(ld).s1 < -1) Or (linedefs(ld).s1 >= numsidedefs) Then
               
               'Report error
               ErrorLog_Add "ERROR: Linedef " & ld & " references invalid sidedef1 " & linedefs(ld).s1 & ". Linedef has been removed.", True
               
               'Remove the linedef
               RemoveLinedef ld, True, False, True
               
          'Verify sidedef 2
          ElseIf (linedefs(ld).s2 < -1) Or (linedefs(ld).s2 >= numsidedefs) Then
               
               'Report error
               ErrorLog_Add "ERROR: Linedef " & ld & " references invalid sidedef2 " & linedefs(ld).s2 & ". Linedef has been removed.", True
               
               'Remove the linedef
               RemoveLinedef ld, True, False, True
          End If
          
          'Verify length of line
          xl = vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x
          yl = vertexes(linedefs(ld).v2).y - vertexes(linedefs(ld).v1).y
          If (CLng(Sqr(xl * xl + yl * yl)) = 0) Then ErrorLog_Add "WARNING: Linedef " & ld & " has length zero", False
          
          'Check if the line has an action which we know
          If (mapconfig("linedeftypes").Exists(CStr(linedefs(ld).effect)) = True) Then
               
               'Set the marking references on the linedef
               With linedefs(ld)
                    .argref0 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark1")
                    .argref1 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark2")
                    .argref2 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark3")
                    .argref3 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark4")
                    .argref4 = mapconfig("linedeftypes")(CStr(linedefs(ld).effect))("mark5")
               End With
          End If
          
          'Check if first sidedef set
          If (linedefs(ld).s1 > -1) Then
               
               'Check if sidedef is already referenced
               If (sdref(linedefs(ld).s1)) Then
                    
                    'Create new sidedef
                    sd = CreateSidedef
                    
                    'Fill it with the same info
                    With sidedefs(linedefs(ld).s1)
                         sidedefs(sd).Lower = .Lower
                         sidedefs(sd).Middle = .Middle
                         sidedefs(sd).Upper = .Upper
                         sidedefs(sd).sector = .sector
                         sidedefs(sd).tx = .tx
                         sidedefs(sd).ty = .ty
                    End With
                    
                    'Reference linedef to new sidedef
                    linedefs(ld).s1 = sd
                    
                    'Sidedef compression has been solved
                    SidedefCompression = SidedefCompression + 1
                    
                    'Correct the references aray
                    ReDim Preserve sdref(0 To (numsidedefs - 1))
                    sdref(numsidedefs - 1) = True
               Else
                    
                    'This sidedef is referenced
                    sdref(linedefs(ld).s1) = True
               End If
               
               'Backlink the sidedef to this linedef
               sidedefs(linedefs(ld).s1).linedef = ld
          End If
          
          'Check if second sidedef set
          If (linedefs(ld).s2 > -1) Then
               
               'Check if sidedef is already referenced
               If (sdref(linedefs(ld).s2)) Then
                    
                    'Create new sidedef
                    sd = CreateSidedef
                    
                    'Fill it with the same info
                    With sidedefs(linedefs(ld).s2)
                         sidedefs(sd).Lower = .Lower
                         sidedefs(sd).Middle = .Middle
                         sidedefs(sd).Upper = .Upper
                         sidedefs(sd).sector = .sector
                         sidedefs(sd).tx = .tx
                         sidedefs(sd).ty = .ty
                    End With
                    
                    'Reference linedef to new sidedef
                    linedefs(ld).s2 = sd
                    
                    'Sidedef compression has been solved
                    SidedefCompression = SidedefCompression + 1
                    
                    'Correct the references aray
                    ReDim Preserve sdref(0 To (numsidedefs - 1))
                    sdref(numsidedefs - 1) = True
               Else
                    
                    'This sidedef is referenced
                    sdref(linedefs(ld).s2) = True
               End If
               
               'Backlink the sidedef to this linedef
               sidedefs(linedefs(ld).s2).linedef = ld
          End If
          
          'Next lindedef
          ld = ld - 1
     Loop
     
     'Add note for sidedef compression if there was any
     If SidedefCompression Then ErrorLog_Add "NOTE: " & SidedefCompression & " Sidedefs have been copied as result of sidedefs decompression.", False
     
     
     '
     ' === REMOVE UNREFERENCED SIDEDEFS
     ' === REPLACE ZERO TEXTURE NAMES WITH -
     '
     
     'Go for all sidedefs
     For sd = (numsidedefs - 1) To 0 Step -1
          
          'Check texture names and replace with "-" as needed
          If (LenB(Trim$(sidedefs(sd).Lower)) = 0) Then sidedefs(sd).Lower = "-"
          If (LenB(Trim$(sidedefs(sd).Middle)) = 0) Then sidedefs(sd).Middle = "-"
          If (LenB(Trim$(sidedefs(sd).Upper)) = 0) Then sidedefs(sd).Upper = "-"
          
          'Check if unreferenced
          If (sdref(sd) = False) Then
               
               'Remove this sidedef
               RemoveSidedef sd, True, True, False
               SidedefsRemoved = SidedefsRemoved + 1
          End If
     Next sd
     
     'Add note for sidedef removal if any
     If SidedefsRemoved Then ErrorLog_Add "NOTE: " & SidedefsRemoved & " Unused sidedefs have been removed.", False
     
     
     '
     ' === REMOVE UNREFERENCED SECTORS
     ' === REPLACE ZERO FLAT NAMES WITH -
     '
     
     'Go for all sectors
     For sd = (numsectors - 1) To 0 Step -1
          
          'Check flat names and replace with "-" as needed
          If (LenB(Trim$(sectors(sd).tfloor)) = 0) Then sectors(sd).tfloor = "-"
          If (LenB(Trim$(sectors(sd).tceiling)) = 0) Then sectors(sd).tceiling = "-"
          
          'Check if unreferenced
          If (CountSectorSidedefs(VarPtr(sidedefs(0)), numsidedefs, sd) = 0) Then
               
               'Remove this sector
               RemoveSector sd, False
               SectorsRemoved = SectorsRemoved + 1
          End If
     Next sd
     
     'Add note for sector removal if any
     If SectorsRemoved Then ErrorLog_Add "NOTE: " & SectorsRemoved & " Unused sectors have been removed.", False
     
     
     '
     ' === CHECK FOR INVALID SECTOR REFERENCES
     '
     
     'Go for all sidedefs
     sd = numsidedefs - 1
     Do While sd >= 0
          
          'Verify its sectors
          If (sidedefs(sd).sector < 0) Or (sidedefs(sd).sector >= numsectors) Then
               
               'Report error
               ErrorLog_Add "ERROR: Sidedef " & sd & " references invalid sector " & sidedefs(sd).sector & ". Sidedef has been removed.", True
               
               'Remove the sidedef
               RemoveSidedef sd, True, False, False
               
               'Map changed
               mapchanged = True
               mapnodeschanged = True
          End If
          
          'Next sidedef
          sd = sd - 1
     Loop
     
     
     '
     ' === SET ALL THING'S COLORS AND IMAGES
     '
     
     'Go for all things
     PositionThing = -1
     For th = 0 To (numthings - 1)
          
          'Check if the line has an action which we know
          If (mapconfig("linedeftypes").Exists(CStr(things(th).effect)) = True) Then
               
               'Set the marking references on the linedef
               With things(th)
                    .argref0 = mapconfig("linedeftypes")(CStr(things(th).effect))("mark1")
                    .argref1 = mapconfig("linedeftypes")(CStr(things(th).effect))("mark2")
                    .argref2 = mapconfig("linedeftypes")(CStr(things(th).effect))("mark3")
                    .argref3 = mapconfig("linedeftypes")(CStr(things(th).effect))("mark4")
                    .argref4 = mapconfig("linedeftypes")(CStr(things(th).effect))("mark5")
               End With
          End If
          
          'Update thing image, color and size
          UpdateThingImageColor th
          UpdateThingSize th
          UpdateThingCategory th
          
          'Check if this is the 3D start position
          If (things(th).thing = mapconfig("start3dmode")) Then ApplyPositionFromThing th
     Next th
End Sub

Public Function CreateSector() As Long
     
     'Check if more memory needs to be allocated
     'If (numsectors = UBound(sectors) - 1) Then
     '
     '     'Allocate memory
     '     ReDim Preserve sectors(0 To numsectors + DECLARE_SECTORS)
     'End If
     
     'Increase number of sectors
     numsectors = numsectors + 1
     
     'Return the last sector index
     CreateSector = numsectors - 1
End Function

Public Function CreateSidedef() As Long
     
     'Check if more memory needs to be allocated
     'If (numsidedefs = UBound(sidedefs) - 1) Then
     '
     '     'Allocate memory
     '     ReDim Preserve sidedefs(0 To numsidedefs + DECLARE_SIDEDEFS)
     'End If
     
     'Increase number of sidedefs
     numsidedefs = numsidedefs + 1
     
     'Return the last sidedef index
     CreateSidedef = numsidedefs - 1
End Function

Public Function CreateThing() As Long
     
     'Check if more memory needs to be allocated
     'If (numthings = UBound(things) - 1) Then
     '
     '     'Allocate memory
     '     ReDim Preserve things(0 To numthings + DECLARE_THINGS)
     'End If
     
     'Increase number of numthings
     numthings = numthings + 1
     
     'Return the last thing index
     CreateThing = numthings - 1
End Function

Public Function CreateVertex() As Long
     
     'Check if more memory needs to be allocated
     'If (numvertexes = UBound(vertexes) - 1) Then
     '
     '     'Allocate memory
     '     ReDim Preserve vertexes(0 To numvertexes + DECLARE_VERTICES)
     'End If
     
     'Increase number of vertices
     numvertexes = numvertexes + 1
     
     'Return the last vertex index
     CreateVertex = numvertexes - 1
End Function

Public Sub LoadMapConfiguration(ByVal Gameconfig As String)
     On Error GoTo ConfigError
     Dim cfg As New clsConfiguration
     Dim c As Long, t As Long
     Dim ThingCats As Variant
     Dim ThingItems As Variant
     Dim LinedefIDs As Variant
     Dim NewObj As Dictionary
     Dim ThisObj As Dictionary
     Dim newthings As New Dictionary
     Dim thingtypes As Dictionary
     Dim OldMousePointer As Long
     Dim catwidth As Long
     Dim catarrow As Long
     Dim caterror As Long
     Dim catblock As Long
     Dim catheight As Long
     Dim cathangs As Long
     
     'Change mousepointer
     OldMousePointer = Screen.MousePointer
     Screen.MousePointer = vbHourglass
     
     'Load the game config
     cfg.LoadConfiguration GetGameConfigFile(Gameconfig)
     
     'Get the dictionary object
     Set mapconfig = cfg.Root(True)
     
     'Clean up
     Set cfg = Nothing
     
     
     'We now will expand the configuration for things that have
     'properties set for the entire category.
     
     'Also make a new collection which contains all things without
     'the category collections. This is for faster lookups.
     
     'Go for all thing categories
     Set thingtypes = mapconfig("thingtypes")
     ThingCats = mapconfig("thingtypes").Keys
     For c = LBound(ThingCats) To UBound(ThingCats)
          
          'Get properties specified at category level
          If thingtypes(ThingCats(c)).Exists("width") Then catwidth = thingtypes(ThingCats(c))("width") Else catwidth = 0
          If thingtypes(ThingCats(c)).Exists("arrow") Then catarrow = thingtypes(ThingCats(c))("arrow") Else catarrow = 0
          If thingtypes(ThingCats(c)).Exists("error") Then caterror = thingtypes(ThingCats(c))("error") Else caterror = 1
          If thingtypes(ThingCats(c)).Exists("blocking") Then catblock = thingtypes(ThingCats(c))("blocking") Else catblock = 0
          If thingtypes(ThingCats(c)).Exists("height") Then catheight = thingtypes(ThingCats(c))("height") Else catheight = 0
          If thingtypes(ThingCats(c)).Exists("hangs") Then cathangs = thingtypes(ThingCats(c))("hangs") Else cathangs = 0
          
          'Go for all items in the category
          ThingItems = thingtypes(ThingCats(c)).Keys
          For t = LBound(ThingItems) To UBound(ThingItems)
               
               'Check if this is an item
               If IsNumeric(ThingItems(t)) Then
                    
                    'Check if not already of object type
                    If IsObject(thingtypes(ThingCats(c))(ThingItems(t))) = False Then
                         
                         'Create new object
                         Set NewObj = New Dictionary
                         
                         'Add properties to object
                         NewObj.Add "title", thingtypes(ThingCats(c))(ThingItems(t))
                         NewObj.Add "width", catwidth
                         NewObj.Add "arrow", catarrow
                         NewObj.Add "error", caterror
                         NewObj.Add "blocking", catblock
                         NewObj.Add "height", catheight
                         NewObj.Add "hangs", cathangs
                         NewObj.Add "category", ThingCats(c)
                         
                         'Replace the thing in category
                         Set thingtypes(ThingCats(c)).Item(ThingItems(t)) = NewObj
                         
                         'Add to new things collection
                         newthings.Add ThingItems(t), NewObj
                         
                         'Clean up
                         Set NewObj = Nothing
                    Else
                         
                         'Get the object
                         Set ThisObj = thingtypes(ThingCats(c))(ThingItems(t))
                         
                         'Add properties to it which it has not specified
                         If (ThisObj.Exists("width") = False) Then ThisObj.Add "width", catwidth
                         If (ThisObj.Exists("arrow") = False) Then ThisObj.Add "arrow", catarrow
                         If (ThisObj.Exists("error") = False) Then ThisObj.Add "error", caterror
                         If (ThisObj.Exists("blocking") = False) Then ThisObj.Add "blocking", catblock
                         If (ThisObj.Exists("height") = False) Then ThisObj.Add "height", catheight
                         If (ThisObj.Exists("hangs") = False) Then ThisObj.Add "hangs", cathangs
                         ThisObj.Add "category", ThingCats(c)
                         
                         'Add to new things collection
                         newthings.Add ThingItems(t), ThisObj
                    End If
               End If
          Next t
     Next c
     
     'Add the new things to configuration
     mapconfig.Add "__things", newthings
     
     
     'Expand the linedeftypes structure
     
     'Go for all linedefs
     LinedefIDs = mapconfig("linedeftypes").Keys
     For c = LBound(LinedefIDs) To UBound(LinedefIDs)
          
          'Check if this linedef is a string
          If VarType(mapconfig("linedeftypes")(LinedefIDs(c))) = vbString Then
               
               'Make a structure from this
               
               'Create new object
               Set NewObj = New Dictionary
               
               'Add the title
               NewObj.Add "title", mapconfig("linedeftypes")(LinedefIDs(c))
               
               'Remove old linedef type
               mapconfig("linedeftypes").Remove LinedefIDs(c)
               
               'Add new linedef type
               mapconfig("linedeftypes").Add LinedefIDs(c), NewObj
               
               'Clean up
               Set NewObj = Nothing
          End If
          
          'Add marks if needed
          If (mapconfig("linedeftypes")(LinedefIDs(c)).Exists("mark1") = False) Then mapconfig("linedeftypes")(LinedefIDs(c)).Add "mark1", 0
          If (mapconfig("linedeftypes")(LinedefIDs(c)).Exists("mark2") = False) Then mapconfig("linedeftypes")(LinedefIDs(c)).Add "mark2", 0
          If (mapconfig("linedeftypes")(LinedefIDs(c)).Exists("mark3") = False) Then mapconfig("linedeftypes")(LinedefIDs(c)).Add "mark3", 0
          If (mapconfig("linedeftypes")(LinedefIDs(c)).Exists("mark4") = False) Then mapconfig("linedeftypes")(LinedefIDs(c)).Add "mark4", 0
          If (mapconfig("linedeftypes")(LinedefIDs(c)).Exists("mark5") = False) Then mapconfig("linedeftypes")(LinedefIDs(c)).Add "mark5", 0
     Next c
     
     'Reset mousepointer
     Screen.MousePointer = OldMousePointer
     
     'Leave now
     Exit Sub
     
     
ConfigError:
     
     MsgBox "Error " & Err.number & " in LoadMapConfiguration: " & Err.Description & vbLf & vbLf & "Game configuration: " & GetGameConfigFile(Gameconfig), vbCritical
     Terminate
End Sub

Public Function MapBuild(ByVal StatusShown As Boolean, ByVal ExportBuild As Boolean) As Boolean
     Dim BuildWAD As New clsWAD
     Dim BuildFile As String
     Dim ResultFile As String
     Dim lumpindex As Long
     Dim Lumpnames As Variant
     Dim ThisLumpName As String
     Dim ThisLumpType As ENUM_MAPLUMPTYPES
     Dim Parameters As String
     Dim i As Long
     
     'Check if we should load the status dialog
     If (StatusShown = False) Then
          
          'Show status dialog
          frmStatus.Show 0, frmMain
          frmMain.SetFocus
          frmMain.Refresh
     End If
     
     'Set status
     DisplayStatus "Building nodes..."
     
     'Presume no problems
     MapBuild = True
     
     'Write memory to TempWAD
     StoreMapLumps
     
     'Check what nodebuild settings to use
     If ExportBuild Then Parameters = Config("buildexportparams") Else Parameters = Config("buildparams")
     
     'Check if %T or %F is missing
     If (InStr(1, Parameters, "%F", vbTextCompare) = 0) Or _
        (InStr(1, Parameters, "%T", vbTextCompare) = 0) Then
          
          'Create a single temporary wad file to build in
          BuildFile = MakeTempFile(True)
          ResultFile = BuildFile
     Else
          
          'Create two temporary wad files to build in
          BuildFile = MakeTempFile(True)
          ResultFile = MakeTempFile(False)
     End If
     
     'Open build file
     Kill BuildFile
     BuildWAD.NewFile BuildFile, False
     
     'Copy all needed lumps
     CopyLumpsByType TempWAD, "MAP01", BuildWAD, "MAP01", ML_REQUIRED
     
     'Write changes and close wad file
     BuildWAD.WriteChanges
     BuildWAD.CloseFile
     
     'Focus to main window. This is a workaround from some
     'driver issues with microsoft mouse scrollwheel
     AppActivate frmMain.Caption
     frmMain.SetFocus
     DisplayStatus "Building nodes..."
     
     
     'Check what nodebuild settings to use
     If ExportBuild Then
          
          'Replace placeholders in parameters
          Parameters = Replace$(Parameters, "%F", BuildFile, , , vbTextCompare)
          Parameters = Replace$(Parameters, "%T", ResultFile, , , vbTextCompare)
          
          'Let the nodebuilder do its work
          If (Execute(Config("buildexportexec"), Parameters, SW_HIDE, True) = False) Then MsgBox "Warning: Could not run the nodebuilder! Check your configuration!", vbExclamation
     Else
          
          'Replace placeholders in parameters
          Parameters = Replace$(Parameters, "%F", BuildFile, , , vbTextCompare)
          Parameters = Replace$(Parameters, "%T", ResultFile, , , vbTextCompare)
          
          'Let the nodebuilder do its work
          If (Execute(Config("buildexec"), Parameters, SW_HIDE, True) = False) Then MsgBox "Warning: Could not run the nodebuilder! Check your configuration!", vbExclamation
     End If
     
     
     'Refresh main window
     frmMain.Show
     frmMain.Refresh
     
     'Set status
     DisplayStatus "Waiting for Nodebuilder to complete..."
     
     'Wait for result file
     If WaitForSingleFile(ResultFile, 3000, 0) Then
          
          'Open the result file to copy from
          On Error Resume Next
          BuildWAD.OpenFile ResultFile, False
          
          'Check for errors
          If (Err.number = 0) Then
               
               'Go for all defined lump names
               Lumpnames = mapconfig("maplumpnames").Keys
               For i = LBound(Lumpnames) To UBound(Lumpnames)
                    
                    'Get lump name
                    ThisLumpName = Trim$(UCase$(Lumpnames(i)))
                    
                    'Get lump type
                    ThisLumpType = GetMapLumpType(ThisLumpName)
                    
                    'Check if this is a lump from nodebuilder
                    'If (CLng(mapconfig("maplumpnames")(ThisLumpName)) And ML_NODEBUILD) = ML_NODEBUILD Then
                    If (ThisLumpType And ML_NODEBUILD) = ML_NODEBUILD Then
                         
                         'Check if lump is missing
                         lumpindex = FindLumpIndex(BuildWAD, 1, ThisLumpName)
                         If (lumpindex = 0) Then
                              
                              'Nodebuilder did not build the required nodes!
                              MapBuild = False
                              Exit For
                              
                         'Check if its missing content
                         ElseIf (BuildWAD.LumpSize(lumpindex) <= 0) And ((ThisLumpType And ML_EMPTYALLOWED) = 0) Then
                              
                              'Nodebuilder did not build the required nodes!
                              MapBuild = False
                              Exit For
                         End If
                    End If
               Next i
          Else
               
               'Nodes did not build
               MapBuild = False
          End If
          On Error GoTo 0
     Else
          
          'File was not written!
          MapBuild = False
     End If
     
     'Check if all required lumps do exist
     If (MapBuild) Then
          
          'Copy nodebuilder lumps to TempWad
          CopyLumpsByType BuildWAD, "MAP01", TempWAD, "MAP01", ML_NODEBUILD
          
          'Write changes to temporary file
          TempWAD.WriteChanges
          
          'Reload map lumps (they may have been modified by the nodebuilder)
          ReadMapLumps
          
          'Create data structure optimizations
          DisplayStatus "Optimizing data structures..."
          CreateOptimizations
          
          'Remove unused vertices
          DisplayStatus "Removing unused vertices..."
          RemoveUnusedVertices
          
          'Deselect and redraw map
          RemoveSelection True
          
          'Nodes rebuilt
          mapnodeschanged = False
     Else
          
          'Nodes were not built
          'Remove old nodebuilder lumps from TempWAD
          RemoveLumpsByType TempWAD, "MAP01", ML_NODEBUILD
          
          'Write memory to TempWAD so that these required lumps still exist
          StoreMapLumps
     End If
     
     
     'Close and remove the build file
     On Local Error Resume Next
     BuildWAD.CloseFile
     Kill BuildFile
     If (Dir(ResultFile) <> "") Then Kill ResultFile
     On Local Error GoTo 0
     
     'Check if we should unload the status dialog
     If (StatusShown = False) Then Unload frmStatus: Set frmStatus = Nothing
End Function

Private Sub MapCompressSidedefs()
     Dim cs As Long      'Current sidedef
     Dim ts As Long      'Sidedef being tested
     Dim Same As Long
     Dim t As Single
     Dim p As Single
     
     'Show status
     DisplayStatus "Compressing sidedefs...   0%"
     t = Timer
     
     'Go for all sidedefs until no more
     Do While cs < numsidedefs
          
          'Begin at this position.
          'Previous sidedefs have already been examined
          ts = cs + 1
          
          'Go for all sidedefs from this point to find equal ones
          Do While ts < numsidedefs
               
               'Assume not the same
               Same = False
               
               'Test if these are the same
               If (sidedefs(cs).sector = sidedefs(ts).sector) Then
                If (sidedefs(cs).Middle = sidedefs(ts).Middle) Then
                 If (sidedefs(cs).Lower = sidedefs(ts).Lower) Then
                  If (sidedefs(cs).Upper = sidedefs(ts).Upper) Then
                   If (sidedefs(cs).tx = sidedefs(ts).tx) Then
                    If (sidedefs(cs).ty = sidedefs(ts).ty) Then Same = True
                   End If
                  End If
                 End If
                End If
               End If
               
               'Check if they are same
               If (Same) Then
                    
                    'All linedefs that refer to the tested sidedef must
                    'now be rereferenced to the current sidedef.
                    Rereference_Sidedefs linedefs(0), numlinedefs, ts, cs
                    
                    'Remove this sidedef
                    RemoveSidedef ts, False, False, False
               Else
                    
                    'Move on to test next sidedef
                    ts = ts + 1
               End If
               
               'Update status?
               If (t + 0.5 < Timer) Then
                    
                    'Calculate percent
                    p = CSng(cs / numsidedefs) * 100
                    
                    'Update status
                    DisplayStatus "Compressing sidedefs...   " & CLng(Sqr(p * 100)) & "%"
                    t = Timer
               End If
          Loop
          
          'Next sidedef
          cs = cs + 1
     Loop
End Sub

Public Function MapLoad(ByVal Filename As String, ByRef FileWAD As clsWAD, ByVal LumpName As String, ByVal ShowStatusDialog As Boolean) As Boolean
     On Error GoTo MapLoadError
     Dim FileBuffer As Integer
     Dim MapLumpIndex As Long
     Dim SubLumpIndex As Long
     Dim MapRect As RECT
     Dim WidthZoom As Single
     Dim HeightZoom As Single
     Dim i As Long
     
     'Check if we should add current path
     If (InStr(Filename, "\") = 0) And _
        (InStr(Filename, "/") = 0) And _
        (InStr(Filename, ":") = 0) Then
          
          'Add current path to filename
          If (right$(CurDir, 1) = "\") Then
               Filename = CurDir & Filename
          Else
               Filename = CurDir & "\" & Filename
          End If
     End If
     
     'Set the default map lump name
     maplumpname = UCase$(UnPadded(LumpName))
     
     'Set the map file name
     'Do NOT set mapfile yet
     mapfilename = Dir(Filename)
     frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
     mapchanged = False
     mapnodeschanged = False
     mapsaved = True
     mapisiwad = FileWAD.IWAD
     Set MapWAD = FileWAD
     
     'No more selections
     Set selected = New Dictionary
     numselected = 0
     Set dragselected = New Dictionary
     dragnumselected = 0
     Erase changedlines()
     ReDim changedlines(0)
     numchangedlines = 0
     
     'Load map configuration
     DisplayStatus "Loading configuration..."
     LoadMapConfiguration mapgame
     
     'Check if an IWAD is configured
     If (Trim$(GetCurrentIWADFile) = "") Then
          
          'Ask the user to configure an IWAD
          If (MsgBox("You do not have the IWAD configured for this configuration yet." & vbLf & "Would you like to browse for the correct IWAD now?", vbQuestion Or vbYesNo) = vbYes) Then
               
               'Configure IWAD
               frmMain.ShowConfiguration 3
          End If
          
     'Check if IWAD can be found
     ElseIf (Dir(GetCurrentIWADFile) = "") Then
          
          'Ask the user to configure an IWAD
          If (MsgBox("The IWAD configured for this configuration cannot be found!" & vbLf & "Would you like to browse for the correct IWAD now?", vbQuestion Or vbYesNo) = vbYes) Then
               
               'Configure IWAD
               frmMain.ShowConfiguration 3
          End If
     End If
     
     'Now set mapfile
     mapfile = Filename
     
     'Open additional wads
     OpenIWADFile
     OpenADDWADFile
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Check if we should show the status dialog
     If (ShowStatusDialog) Then
          
          'Show status dialog
          frmStatus.Show 0, frmMain
          frmMain.SetFocus
          frmMain.Refresh
     End If
     
     'Load the error log
     ErrorLog_Load
     
     
     'Set status
     DisplayStatus "Loading DECORATE things..."
     
     'Load additional items from DECORATE
     ApplyDecorateThings IWAD
     ApplyDecorateThings AddWAD
     ApplyDecorateThings MapWAD
     
     'Set status
     DisplayStatus "Loading map data..."
     
     'Get file buffer
     FileBuffer = MapWAD.FileBuffer
     
     'Get lump index of map
     For i = 1 To MapWAD.LumpCount
          If (MapWAD.LumpName(i) = LumpName) Then MapLumpIndex = i: Exit For
     Next i
     
     
     'Make a temporary file for the extra lumps
     Set TempWAD = New clsWAD
     maptempfile = MakeTempFile(False)
     TempWAD.NewFile maptempfile, False
     
     'Copy all lumps to the TempWad
     CopyLumpsByType MapWAD, LumpName, TempWAD, "MAP01", ML_REQUIRED Or ML_RESPECTED Or ML_NODEBUILD Or ML_CUSTOM
     
     'Load the map lumps to memory
     ReadMapLumps
     
     
     'Precache resources
     MapLoadResources
     
     
     'Create data structure optimizations
     DisplayStatus "Optimizing data structures..."
     CreateOptimizations
     
     'Remove unused vertices
     DisplayStatus "Removing unused vertices..."
     RemoveUnusedVertices
     
     'Set the defaults
     gridsizex = Config("defaultgrid")
     gridsizey = Config("defaultgrid")
     gridx = 0: gridy = 0
     snapmode = Config("defaultsnap")
     stitchmode = Config("defaultstitch")
     filterthings = False
     filtersettings.category = -1
     filtersettings.filtermode = 0
     filtersettings.Flags = 0
     
     'Ensure correct textures to build with
     CorrectDefaultTextures
     
     'Initialize undo/redo
     DisplayStatus "Initializing Undo/Redo..."
     InitializeUndoRedo
     
     
     'Initialize map screen renderer
     DisplayStatus "Initializing Renderer..."
     InitializeMapRenderer frmMain.picMap
     
     
     'Calculating map rect
     MapRect = CalculateMapRect
     
     'Center map in view
     CenterViewAt MapRect, True
     
     
     'Default new thing
     With LastThing
          .thing = Config("defaultthing")
          .angle = 0
          .arg0 = 0
          .arg1 = 0
          .arg2 = 0
          .arg3 = 0
          .arg4 = 0
          .Color = 0
          .image = 0
          .effect = 0
          .Flags = mapconfig("defaulthingflags")
          .selected = 0
          .tag = 0
          .Z = 0
     End With
     
     
     'Enable map editing controls
     EnableMapEditing
     
     'Update scripts menu
     UpdateScriptLumpsMenu
     
     'Select the current editing mode, this will also draw map
     frmMain.itmEditMode_Click CInt(EM_LINES)
     
     'Remove from list if the map is in recent list
     i = GetRecentFileIndex(Filename)
     If (i > 0) Then RemoveRecentFile i
     
     'Add to recent list
     AddRecentFile Filename
     
     'Update menu with list
     UpdateRecentFilesMenu
     
     'Update status bar
     UpdateStatusBar
     
     'Unload status dialog
     If (ShowStatusDialog) Then Unload frmStatus: Set frmStatus = Nothing
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
     
     'Show the errors and warnings dialog
     ErrorLog_DisplayAndFlush
     
     'Map is loaded
     MapLoad = True
     Exit Function
     
     
MapLoadError:
     
     'Show error message
     MsgBox "Error " & Err.number & " while loading map: " & Err.Description, vbCritical
     
     'Unload map
     MapUnload
     
     'Unload error log
     ErrorLog_Flush
     
     'Unload dialog
     If (ShowStatusDialog) Then Unload frmStatus: Set frmStatus = Nothing
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
End Function

Public Sub MapLoadResources()
     
     'Load sprites and optionally precache
     DisplayStatus "Loading sprites..."
     
     'Unload sprites
     CleanUpSpriteImages
     
     'Initialze sprites
     InitializeSprites
     
     'Load texture names and optionally precache
     DisplayStatus "Loading texture names..."
     
     'Unload old list
     UnloadAllTextures
     
     'Unload old list
     UnloadAllFlats
     
     'Load new textures and check for errors
     If (Not LoadAllTextures) Then Err.Raise 0, , "Unknown error in LoadAllTextures()"
     
     'Check if we should precache
     If (Config("textureprecache")) Then
          
          'Precache them
          DisplayStatus "Precaching texture resources..."
          PrecacheTextures
     End If
     
     
     'Load flat names and optionally precache
     DisplayStatus "Loading flat names..."
     
     'Load new textures and check for errors
     If (Not LoadAllFlats) Then Err.Raise 0, , "Unknown error in LoadAllFlats()"
     
     'Check if we should precache
     If (Config("textureprecache")) Then
          
          'Precache them
          DisplayStatus "Precaching flat resources..."
          PrecacheFlats
     End If
End Sub

Public Function MapNew(ByVal LumpName As String, ByVal ShowStatusDialog As Boolean, ByVal ShowMapOptions As Boolean) As Boolean
     On Error GoTo MapLoadError
     Dim MapRect As RECT
     Dim WidthZoom As Single
     Dim HeightZoom As Single
     Dim i As Long
     
     'Set the default map lump name
     maplumpname = UCase$(UnPadded(LumpName))
     
     'Show map options
     If (ShowMapOptions) Then If (Not ChangeMapOptions(True)) Then Exit Function
     
     'Set the map file name
     mapfile = MakeTempFile(False)
     mapfilename = "untitled.wad"
     frmMain.Caption = App.Title & " - " & mapfilename & " (" & maplumpname & ")"
     mapchanged = False
     mapnodeschanged = False
     mapsaved = False
     mapisiwad = False
     
     
     'Load map configuration
     DisplayStatus "Loading configuration..."
     LoadMapConfiguration mapgame
     
     'Open additional wads
     OpenIWADFile
     OpenADDWADFile
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Check if we should show the status dialog
     If (ShowStatusDialog) Then
          
          'Show status dialog
          frmStatus.Show 0, frmMain
          frmMain.SetFocus
          frmMain.Refresh
     End If
     
     'Load the error log
     ErrorLog_Load
     
     
     
     'Display status
     DisplayStatus "Creating temporary file..."
     
     'Make temporary file
     Set MapWAD = New clsWAD
     MapWAD.NewFile mapfile, mapisiwad
     
     'Add map header
     MapWAD.AddLump "", maplumpname
     
     'Make a temporary file for the extra lumps
     Set TempWAD = New clsWAD
     maptempfile = MakeTempFile(False)
     TempWAD.NewFile maptempfile, False
     
     'Create all required lumps in the TempWad
     CompleteMapLumps TempWAD, "MAP01"
     
     
     
     'Display status
     DisplayStatus "Allocating memory for structures..."
     
     'Allocate new memory
     'ReDim things(0 To DECLARE_THINGS)
     'ReDim linedefs(0 To DECLARE_LINEDEFS)
     'ReDim sidedefs(0 To DECLARE_SIDEDEFS)
     'ReDim vertexes(0 To DECLARE_VERTICES)
     'ReDim sectors(0 To DECLARE_SECTORS)
     
     
     
     'Load resources and optionally precache
     MapLoadResources
     
'     DisplayStatus "Loading textures..."
'     UnloadAllTextures
'     UnloadAllFlats
'     If (Not LoadAllTextures) Then Err.Raise 0, , "Unknown error in LoadAllTextures()"
'     If (Config("textureprecache")) Then
'          DisplayStatus "Precaching textures..."
'          PrecacheTextures
'     End If
'
'     'Load flat names and optionally precache
'     DisplayStatus "Loading flats..."
'     If (Not LoadAllFlats) Then Err.Raise 0, , "Unknown error in LoadAllFlats()"
'     If (Config("textureprecache")) Then
'          DisplayStatus "Precaching flats..."
'          PrecacheFlats
'     End If
     
     'Create data structure optimizations
     DisplayStatus "Optimizing data structures..."
     CreateOptimizations
     
     'Set the defaults
     gridsizex = Config("defaultgrid")
     gridsizey = Config("defaultgrid")
     gridx = 0: gridy = 0
     snapmode = Config("defaultsnap")
     stitchmode = Config("defaultstitch")
     filterthings = False
     filtersettings.category = -1
     filtersettings.filtermode = 0
     filtersettings.Flags = 0
     
     
     'Initialize undo/redo
     DisplayStatus "Initializing Undo/Redo..."
     InitializeUndoRedo
     
     
     'Initialize map screen renderer
     DisplayStatus "Initializing Renderer..."
     InitializeMapRenderer frmMain.picMap
     
     
     'Default viewport
     ChangeView -100, 100, 1
     
     'Default new thing
     With LastThing
          .thing = Config("defaultthing")
          .angle = 0
          .arg0 = 0
          .arg1 = 0
          .arg2 = 0
          .arg3 = 0
          .arg4 = 0
          .Color = 0
          .image = 0
          .effect = 0
          .Flags = mapconfig("defaulthingflags")
          .selected = 0
          .tag = 0
          .Z = 0
     End With
     
     'Enable map editing controls
     EnableMapEditing
     
     'Update scripts menu
     UpdateScriptLumpsMenu
     
     'Select the current editing mode, this will also draw map
     frmMain.itmEditMode_Click CInt(EM_LINES)
     
     'Update status bar
     UpdateStatusBar
     
     'Unload status dialog
     If (ShowStatusDialog) Then Unload frmStatus: Set frmStatus = Nothing
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
     
     'Show the errors and warnings dialog
     ErrorLog_DisplayAndFlush
     
     'Map is loaded
     MapNew = True
     Exit Function
     
     
MapLoadError:
     
     'Show error message
     If (Err.number <> -1) Then MsgBox "Error " & Err.number & " while creating map: " & Err.Description, vbCritical
     
     'Unload map
     MapUnload
     
     'Unload error log
     ErrorLog_Flush
     
     'Unload dialog
     If (ShowStatusDialog) Then Unload frmStatus: Set frmStatus = Nothing
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
End Function

Public Function MapSave(ByVal Filename As String, ByVal SaveMode As ENUM_SAVEMODES, Optional ByVal CompressSidedefs As Boolean = False) As Boolean
     On Error GoTo MapSaveError
     Dim ExistingIndex As Long
     Dim NextLumpName As String
     Dim BuildWAD As New clsWAD
     Dim BuildFile As String
     Dim DoNodebuilder As Boolean
     Dim FoundIndex As Long
     Dim MapLumps As Variant
     Dim MapDeleting As Boolean
     Dim i As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Show status dialog
     frmStatus.Show 0, frmMain
     frmStatus.Refresh
     frmMain.SetFocus
     frmMain.Refresh
     
     'Disable editing
     frmMain.picMap.Enabled = False
     
     'Presume no problems
     MapSave = True
     
     'Close IWAD
     IWAD.CloseFile
     
     'Close additional files
     AddWAD.CloseFile
     MapWAD.CloseFile
     
     'Restore original map configuration if its not for this map
     If (mapconfig("game") <> mapgame) Then LoadMapConfiguration mapgame
     
     'Show status
     DisplayStatus "Creating backup of previous file..."
     
     'Check if we should make a backup
     If (SaveMode = SM_SAVE) Or (SaveMode = SM_SAVEAS) Or _
        (SaveMode = SM_SAVEINTO) Or (SaveMode = SM_EXPORT) Then
          
          'Check if preferred to make backups
          If (Config("savebackup")) Then
               
               'Kill oldest backup
               If (Dir(Filename & ".backup3") <> "") Then Kill Filename & ".backup3"
               
               'Move previous backup
               If (Dir(Filename & ".backup2") <> "") Then Name Filename & ".backup2" As Filename & ".backup3"
               
               'Move previous backup
               If (Dir(Filename & ".backup1") <> "") Then Name Filename & ".backup1" As Filename & ".backup2"
               
               'Copy the file with .bak extension
               If (Dir(Filename) <> "") Then FileCopy Filename, Filename & ".backup1"
          End If
     End If
     
     'Check if we should remove existing target file
     If (SaveMode = SM_SAVEAS) Or (SaveMode = SM_TEST) Then
          
          'Check if this file exists
          If (Dir(Filename) <> "") Then
               
               'Remove the file
               If (Dir(Filename) <> "") Then Kill Filename
          End If
     End If
     
     'Check if the source exists
     If (Dir(mapfile) <> "") Then
          
          'Check if we should rebuild the original file to the temp build file
          'so that all required textures and other lumps are with the map
          If (SaveMode = SM_SAVE) Or (SaveMode = SM_SAVEAS) Or (SaveMode = SM_EXPORT) Then
               
               'Show status
               DisplayStatus "Rebuilding WAD file resources..."
               
               'Check if the two are the same
               If (StrComp(Trim$(mapfile), Trim$(Filename), vbTextCompare) = 0) Then
                    
                    'Copy the original file to a temp file
                    mapfile = MakeTempFile(False)
                    FileCopy Filename, mapfile
                    
                    'Open temp file
                    MapWAD.OpenFile mapfile, True
                    
                    'Reset the filename
                    mapfile = Filename
               Else
                    
                    'Open target file
                    MapWAD.OpenFile mapfile, True
               End If
               
               'Open new file
               BuildWAD.NewFile Filename, mapisiwad
               
               'Start copying all resources
               For i = 1 To MapWAD.LumpCount
                    
                    'Check if we should ignore this set of lumps
                    If (mapoldlumpname <> "") And (StrComp(left$(MapWAD.LumpName(i), Len(mapoldlumpname)), mapoldlumpname, vbTextCompare) = 0) Then
                         
                         'Dont copy this
                         'Indicate not to copy anymore until
                         'another lump found that does not belong to map
                         MapDeleting = True
                         
                    'Check if this is not the map lump name
                    ElseIf StrComp(left$(MapWAD.LumpName(i), Len(maplumpname)), maplumpname, vbTextCompare) = 0 Then
                         
                         'Dont copy this
                         'Indicate not to copy anymore until
                         'another lump found that does not belong to map
                         MapDeleting = True
                    Else
                         
                         'Check if ignoring map lumps
                         If MapDeleting Then
                              
                              'Check if we should stop ignoring here
                              If GetMapLumpType(MapWAD.LumpName(i)) = ML_UNKNOWN Then MapDeleting = False
                         End If
                         
                         'Copy the lump directly if not supposed to be ignored
                         If Not MapDeleting Then BuildWAD.AddLump MapWAD.GetLump(i), MapWAD.LumpName(i)
                    End If
               Next i
               
               'Write table and header
               BuildWAD.WriteChanges
               
               'Close the files
               MapWAD.CloseFile
               BuildWAD.CloseFile
               
               'No more old map lump name
               mapoldlumpname = ""
               
          'Check if we should make a fast copy of the original file
          ElseIf (SaveMode = SM_TEST) Then
               
               'Show status
               DisplayStatus "Copying WAD file resources..."
               
               'Copy file
               FileCopy mapfile, Filename
          End If
     End If
     
     
     'Check if sidedefs must be compressed
     If (CompressSidedefs) Then
          
          'Show status
          DisplayStatus "Compressing sidedefs..."
          
          'Compress sidedefs
          MapCompressSidedefs
     End If
     
     
     'Save the script if the script editor is open
     If (ScriptEditor) Then frmScript.Save
     
     'Write memory to TempWAD
     StoreMapLumps
     
     
     'Possible preferences for nodebuilding:
     ' 0 = Always rebuild nodes
     ' 2 = Ask to rebuild nodes
     ' 4 = Never rebuild nodes
     
     'Check if user allows nodebuilding
     If (Config("buildnodes") <> 4) Or (SaveMode = SM_EXPORT) Then
          
          'Check if we should ask
          If ((Config("buildnodes") = 2) And (SaveMode <> SM_EXPORT)) Then
               
               'Ask to rebuild nodes and keep the result
               DoNodebuilder = (MsgBox("Do you want to build the nodes for your map now?", vbQuestion Or vbYesNo) = vbYes)
          Else
               
               'Build the nodes
               DoNodebuilder = True
          End If
          
          'Build the nodes if user chose to do so
          If DoNodebuilder Then
               
               'Build the nodes and check if failed
               If MapBuild(True, (SaveMode = SM_EXPORT)) = False Then
                    
                    'Nodebuilder failed!
                    MsgBox "The nodebuilder did not build the required structures." & vbLf & "Please check your map for errors or select a different nodebuilder." & vbLf & vbLf & "Your map will be saved without the nodes.", vbCritical
                    
                    'Nodes were not built
                    'MapSave = False
               End If
          End If
     End If
     
     
     'Check if the target exists
     If (Dir(Filename) <> "") Then
          
          'Open the target WAD file
          MapWAD.OpenFile Filename, False
     Else
          
          'Make new target WAD file
          MapWAD.NewFile Filename, mapisiwad
     End If
     
     
     'Add new map lumps
     DisplayStatus "Writing new map resources..."
     
     'Copy and replace resource from TempWAD to MapWAD
     CopyLumpsByType TempWAD, "MAP01", MapWAD, maplumpname, ML_REQUIRED Or ML_RESPECTED Or ML_NODEBUILD Or ML_CUSTOM
     
     
     'Finish off
     DisplayStatus "Closing files..."
     
     'Write changes and close wad file
     MapWAD.WriteChanges
     MapWAD.CloseFile
     
     'Check if we were only exporting or testing
     If (SaveMode = SM_EXPORT) Or (SaveMode = SM_TEST) Then
          
          'Re-open the original wad as read only
          MapWAD.OpenFile mapfile, True
     Else
          
          'Re-open the new wad as read only
          MapWAD.OpenFile Filename, True
     End If
     
     'Open additional wads
     OpenIWADFile
     OpenADDWADFile
     
     'Check if file should be on recent list
     If (SaveMode = SM_SAVEAS) Or (SaveMode = SM_SAVEINTO) Then
          
          'Remove from list if the map is in recent list
          i = GetRecentFileIndex(Filename)
          If (i > 0) Then RemoveRecentFile i
          
          'Add to recent list
          AddRecentFile Filename
          
          'Update menu with list
          UpdateRecentFilesMenu
     End If
     
     'Store wad map info if preferred
     If (Val(Config("storeeditinginfo"))) Then PutCurrentWadMapSettings Filename
     
     'Unload status dialog
     Unload frmStatus: Set frmStatus = Nothing
     ErrorLog_DisplayAndFlush
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
     
     'Enable editing
     frmMain.picMap.Enabled = True
     
     'Reselect current editing mode
     frmMain.itmEditMode_Click CInt(mode)
     
     'Map is saved
     Exit Function
     
     
MapSaveError:
     
     'Show error message
     MsgBox "Error " & Err.number & " while saving map: " & Err.Description, vbCritical
     
     'Unload dialog
     Unload frmStatus: Set frmStatus = Nothing
     
     'Problems!
     MapSave = False
     
     'Enable editing
     frmMain.picMap.Enabled = True
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
End Function

Private Sub StoreMapLumps()
     On Local Error GoTo errorhandler
     Dim HeaderIndex As Long
     Dim NextLumpName As String
     Dim lumpindex As Long
     Dim Tempfile As String
     Dim FileBuffer As Integer
     Dim ShortValue As Integer
     Dim StringValue As String * 8
     Dim ByteValue As Byte
     Dim i As Long
     
     
     'This will store the map lumps VERTEXES, LINEDEFS, SIDEDEFS
     'SECTORS and THINGS from memory into the TempWAD.
     'These will be added/replaced after the map header.
     
     
     'Find the map header lump in the TempWAD
     HeaderIndex = FindLumpIndex(TempWAD, 1, "MAP01")
     lumpindex = HeaderIndex + 1
     If (HeaderIndex > 0) Then
          
          'Get next lump name
          NextLumpName = ""
          If (lumpindex <= TempWAD.LumpCount) Then NextLumpName = TempWAD.LumpName(lumpindex)
          
          'Continue deleting lumps until no more map-related lumps
          Do Until (GetMapLumpType(NextLumpName) = ML_UNKNOWN)
               
               'Make reliable lumpname
               NextLumpName = Trim$(UCase$(NextLumpName))
               
               'Check if this is a map lump
               If (NextLumpName = "VERTEXES") Or (NextLumpName = "LINEDEFS") Or _
                  (NextLumpName = "SIDEDEFS") Or (NextLumpName = "SECTORS") Or _
                  (NextLumpName = "THINGS") Then
                    
                    'Remove that lump
                    TempWAD.DeleteLump lumpindex
               Else
                    
                    'Advance to the next lump
                    lumpindex = lumpindex + 1
               End If
               
               'Get next lump name
               NextLumpName = ""
               If (lumpindex <= TempWAD.LumpCount) Then NextLumpName = TempWAD.LumpName(lumpindex) Else Exit Do
          Loop
     End If
     
     
     
     '========= WRITE THINGS
     Tempfile = MakeTempFile
     FileBuffer = FreeFile
     Open Tempfile For Binary As #FileBuffer
     
     'Check if writing in Doom format
     If (mapconfig("mapformat") = 1) Then
          
          'Go for all things to write
          For i = 0 To (numthings - 1)
               ShortValue = things(i).x: Put #FileBuffer, , ShortValue
               ShortValue = things(i).y: Put #FileBuffer, , ShortValue
               ShortValue = things(i).angle: Put #FileBuffer, , ShortValue
               ShortValue = LtoI(things(i).thing): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(things(i).Flags): Put #FileBuffer, , ShortValue
          Next i
          
     'Check if writing in Hexen format
     ElseIf (mapconfig("mapformat") = 2) Then
          
          'Go for all things to write
          For i = 0 To (numthings - 1)
               ShortValue = LtoI(things(i).tag): Put #FileBuffer, , ShortValue
               ShortValue = things(i).x: Put #FileBuffer, , ShortValue
               ShortValue = things(i).y: Put #FileBuffer, , ShortValue
               ShortValue = things(i).Z: Put #FileBuffer, , ShortValue
               ShortValue = things(i).angle: Put #FileBuffer, , ShortValue
               ShortValue = LtoI(things(i).thing): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(things(i).Flags): Put #FileBuffer, , ShortValue
               ByteValue = things(i).effect: Put #FileBuffer, , ByteValue
               ByteValue = things(i).arg0: Put #FileBuffer, , ByteValue
               ByteValue = things(i).arg1: Put #FileBuffer, , ByteValue
               ByteValue = things(i).arg2: Put #FileBuffer, , ByteValue
               ByteValue = things(i).arg3: Put #FileBuffer, , ByteValue
               ByteValue = things(i).arg4: Put #FileBuffer, , ByteValue
          Next i
     End If
     
     'Close temporary file
     Close #FileBuffer
     
     'Import the lump
     TempWAD.ImportLump Tempfile, "THINGS", HeaderIndex + 1
     
     'Clean up the temporary file
     Kill Tempfile
     
     
     
     '========= WRITE LINEDEFS
     Tempfile = MakeTempFile
     FileBuffer = FreeFile
     Open Tempfile For Binary As #FileBuffer
     
     'Check if writing in Doom format
     If (mapconfig("mapformat") = 1) Then
          
          'Go for all lines to write
          For i = 0 To (numlinedefs - 1)
               ShortValue = LtoI(linedefs(i).v1): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(linedefs(i).v2): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(linedefs(i).Flags): Put #FileBuffer, , ShortValue
               ShortValue = linedefs(i).effect: Put #FileBuffer, , ShortValue
               ShortValue = LtoI(linedefs(i).tag): Put #FileBuffer, , ShortValue
               
               'Fix signs
               If (linedefs(i).s1 = -1) Then ShortValue = -1 Else ShortValue = LtoI(linedefs(i).s1)
               Put #FileBuffer, , ShortValue
               
               'Fix signs
               If (linedefs(i).s2 = -1) Then ShortValue = -1 Else ShortValue = LtoI(linedefs(i).s2)
               Put #FileBuffer, , ShortValue
          Next i
          
     'Check if writing in Hexen format
     ElseIf (mapconfig("mapformat") = 2) Then
          
          'Go for all lines to write
          For i = 0 To (numlinedefs - 1)
               ShortValue = LtoI(linedefs(i).v1): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(linedefs(i).v2): Put #FileBuffer, , ShortValue
               ShortValue = LtoI(linedefs(i).Flags): Put #FileBuffer, , ShortValue
               ByteValue = CVB(MKL(linedefs(i).effect)): Put #FileBuffer, , ByteValue
               ByteValue = linedefs(i).arg0: Put #FileBuffer, , ByteValue
               ByteValue = linedefs(i).arg1: Put #FileBuffer, , ByteValue
               ByteValue = linedefs(i).arg2: Put #FileBuffer, , ByteValue
               ByteValue = linedefs(i).arg3: Put #FileBuffer, , ByteValue
               ByteValue = linedefs(i).arg4: Put #FileBuffer, , ByteValue
               
               'Fix signs
               If (linedefs(i).s1 = -1) Then ShortValue = -1 Else ShortValue = LtoI(linedefs(i).s1)
               Put #FileBuffer, , ShortValue
               
               'Fix signs
               If (linedefs(i).s2 = -1) Then ShortValue = -1 Else ShortValue = LtoI(linedefs(i).s2)
               Put #FileBuffer, , ShortValue
          Next i
     End If
     
     'Close temporary file
     Close #FileBuffer
     
     'Import the lump
     TempWAD.ImportLump Tempfile, "LINEDEFS", HeaderIndex + 2
     
     'Clean up the temporary file
     Kill Tempfile
     
     
     
     '========= WRITE SIDEDEFS
     Tempfile = MakeTempFile
     FileBuffer = FreeFile
     Open Tempfile For Binary As #FileBuffer
     
     'Go for all sides to write
     For i = 0 To (numsidedefs - 1)
          ShortValue = sidedefs(i).tx: Put #FileBuffer, , ShortValue
          ShortValue = sidedefs(i).ty: Put #FileBuffer, , ShortValue
          StringValue = Padded(sidedefs(i).Upper, 8): Put #FileBuffer, , StringValue
          StringValue = Padded(sidedefs(i).Lower, 8): Put #FileBuffer, , StringValue
          StringValue = Padded(sidedefs(i).Middle, 8): Put #FileBuffer, , StringValue
          ShortValue = LtoI(sidedefs(i).sector): Put #FileBuffer, , ShortValue
     Next i
     
     'Close temporary file
     Close #FileBuffer
     
     'Import the lump
     TempWAD.ImportLump Tempfile, "SIDEDEFS", HeaderIndex + 3
     
     'Clean up the temporary file
     Kill Tempfile
     
     
     
     '========= WRITE VERTEXES
     Tempfile = MakeTempFile
     FileBuffer = FreeFile
     Open Tempfile For Binary As #FileBuffer
     
     'Go for all vertices to write
     For i = 0 To (numvertexes - 1)
          ShortValue = vertexes(i).x: Put #FileBuffer, , ShortValue
          ShortValue = vertexes(i).y: Put #FileBuffer, , ShortValue
     Next i
     
     'Close temporary file
     Close #FileBuffer
     
     'Import the lump
     TempWAD.ImportLump Tempfile, "VERTEXES", HeaderIndex + 4
     
     'Clean up the temporary file
     Kill Tempfile
     
     
     
     '========= WRITE SECTORS
     Tempfile = MakeTempFile
     FileBuffer = FreeFile
     Open Tempfile For Binary As #FileBuffer
     
     'Go for all sectors to write
     For i = 0 To (numsectors - 1)
          ShortValue = sectors(i).hfloor: Put #FileBuffer, , ShortValue
          ShortValue = sectors(i).hceiling: Put #FileBuffer, , ShortValue
          StringValue = Padded(sectors(i).tfloor, 8): Put #FileBuffer, , StringValue
          StringValue = Padded(sectors(i).tceiling, 8): Put #FileBuffer, , StringValue
          ShortValue = sectors(i).Brightness: Put #FileBuffer, , ShortValue
          ShortValue = LtoI(sectors(i).special): Put #FileBuffer, , ShortValue
          ShortValue = LtoI(sectors(i).tag): Put #FileBuffer, , ShortValue
     Next i
     
     'Close temporary file
     Close #FileBuffer
     
     'Import the lump
     TempWAD.ImportLump Tempfile, "SECTORS", HeaderIndex + 5
     
     'Clean up the temporary file
     Kill Tempfile
     
     'Leave now
     Exit Sub
     
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in StoreMapLumps(): " & Err.Description, vbCritical
End Sub

Private Sub ReadMapLumps()
     On Local Error GoTo errorhandler
     Dim HeaderIndex As Long
     Dim lumpindex As Long
     Dim FileBuffer As Integer
     Dim ShortValue As Integer
     Dim StringValue As String * 8
     Dim ByteValue As Byte
     Dim i As Long
     
     
     'This will read the map lumps VERTEXES, LINEDEFS, SIDEDEFS
     'SECTORS and THINGS from the TempWAD into memory.
     
     
     'Find the map header lump in the TempWAD
     HeaderIndex = FindLumpIndex(TempWAD, 1, "MAP01")
     
     'Get the buffer
     FileBuffer = TempWAD.FileBuffer
     
     
     
     '========= READ THINGS
     lumpindex = FindLumpIndex(TempWAD, HeaderIndex, "THINGS")
     If (lumpindex = 0) Then Err.Raise vbObjectError + 1, , "Could not find required lump THINGS!"
     
     'Check if reading in Doom format
     If (mapconfig("mapformat") = 1) Then
          
          'Calculate number of things
          numthings = TempWAD.LumpSize(lumpindex) \ 10
          
          'Allocate memory for things
          'ReDim things(0 To numthings + DECLARE_THINGS)
          
          'Go for all things to load
          Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
          For i = 0 To (numthings - 1)
               Get #FileBuffer, , ShortValue: things(i).x = ShortValue
               Get #FileBuffer, , ShortValue: things(i).y = ShortValue
               Get #FileBuffer, , ShortValue: things(i).angle = ShortValue
               Get #FileBuffer, , ShortValue: things(i).thing = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: things(i).Flags = ItoL(ShortValue)
          Next i
          
     'Check if reading in Hexen format
     ElseIf (mapconfig("mapformat") = 2) Then
          
          'Calculate number of things
          numthings = TempWAD.LumpSize(lumpindex) \ 20
          
          'Allocate memory for things
          'ReDim things(0 To numthings + DECLARE_THINGS)
          
          'Go for all things to load
          Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
          For i = 0 To (numthings - 1)
               Get #FileBuffer, , ShortValue: things(i).tag = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: things(i).x = ShortValue
               Get #FileBuffer, , ShortValue: things(i).y = ShortValue
               Get #FileBuffer, , ShortValue: things(i).Z = ShortValue
               Get #FileBuffer, , ShortValue: things(i).angle = ShortValue
               Get #FileBuffer, , ShortValue: things(i).thing = ShortValue
               Get #FileBuffer, , ShortValue: things(i).Flags = ItoL(ShortValue)
               Get #FileBuffer, , ByteValue: things(i).effect = ByteValue
               Get #FileBuffer, , ByteValue: things(i).arg0 = ByteValue
               Get #FileBuffer, , ByteValue: things(i).arg1 = ByteValue
               Get #FileBuffer, , ByteValue: things(i).arg2 = ByteValue
               Get #FileBuffer, , ByteValue: things(i).arg3 = ByteValue
               Get #FileBuffer, , ByteValue: things(i).arg4 = ByteValue
          Next i
     End If
     
     
     
     '========= READ LINEDEFS
     lumpindex = FindLumpIndex(TempWAD, HeaderIndex, "LINEDEFS")
     If (lumpindex = 0) Then Err.Raise vbObjectError + 2, , "Could not find required lump LINEDEFS!"
     
     'Check if reading in Doom format
     If (mapconfig("mapformat") = 1) Then
          
          'Calculate number of linedefs
          numlinedefs = TempWAD.LumpSize(lumpindex) \ 14
          
          'Allocate memory for linedefs
          'ReDim linedefs(0 To numlinedefs + DECLARE_LINEDEFS)
          
          'Go for all linedefs to load
          Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
          For i = 0 To (numlinedefs - 1)
               Get #FileBuffer, , ShortValue: linedefs(i).v1 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).v2 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).Flags = ShortValue
               Get #FileBuffer, , ShortValue: linedefs(i).effect = ShortValue
               Get #FileBuffer, , ShortValue: linedefs(i).tag = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).s1 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).s2 = ItoL(ShortValue)
               
               'Fix signs
               If (linedefs(i).s1 = 65535) Then linedefs(i).s1 = -1
               If (linedefs(i).s2 = 65535) Then linedefs(i).s2 = -1
          Next i
          
     'Check if reading in Hexen format
     ElseIf (mapconfig("mapformat") = 2) Then
          
          'Calculate number of linedefs
          numlinedefs = TempWAD.LumpSize(lumpindex) \ 16
          
          'Allocate memory for linedefs
          'ReDim linedefs(0 To numlinedefs + DECLARE_LINEDEFS)
          
          'Go for all linedefs to load
          Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
          For i = 0 To (numlinedefs - 1)
               Get #FileBuffer, , ShortValue: linedefs(i).v1 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).v2 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).Flags = ShortValue
               Get #FileBuffer, , ByteValue: linedefs(i).effect = ByteValue
               Get #FileBuffer, , ByteValue: linedefs(i).arg0 = ByteValue
               Get #FileBuffer, , ByteValue: linedefs(i).arg1 = ByteValue
               Get #FileBuffer, , ByteValue: linedefs(i).arg2 = ByteValue
               Get #FileBuffer, , ByteValue: linedefs(i).arg3 = ByteValue
               Get #FileBuffer, , ByteValue: linedefs(i).arg4 = ByteValue
               Get #FileBuffer, , ShortValue: linedefs(i).s1 = ItoL(ShortValue)
               Get #FileBuffer, , ShortValue: linedefs(i).s2 = ItoL(ShortValue)
               
               'Fix signs
               If (linedefs(i).s1 = 65535) Then linedefs(i).s1 = -1
               If (linedefs(i).s2 = 65535) Then linedefs(i).s2 = -1
          Next i
     End If
     
     
     
     '========= READ SIDEDEFS
     lumpindex = FindLumpIndex(TempWAD, HeaderIndex, "SIDEDEFS")
     If (lumpindex = 0) Then Err.Raise vbObjectError + 3, , "Could not find required lump SIDEDEFS!"
     
     'Calculate number of sidedefs
     numsidedefs = TempWAD.LumpSize(lumpindex) \ 30
     
     'Allocate memory for sidedefs
     'ReDim sidedefs(0 To numsidedefs + DECLARE_SIDEDEFS)
     
     'Go for all sidedefs to load
     Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
     For i = 0 To (numsidedefs - 1)
          Get #FileBuffer, , ShortValue: sidedefs(i).tx = ShortValue
          Get #FileBuffer, , ShortValue: sidedefs(i).ty = ShortValue
          Get #FileBuffer, , StringValue: sidedefs(i).Upper = UCase$(Trim$(UnPadded(StringValue)))
          Get #FileBuffer, , StringValue: sidedefs(i).Lower = UCase$(Trim$(UnPadded(StringValue)))
          Get #FileBuffer, , StringValue: sidedefs(i).Middle = UCase$(Trim$(UnPadded(StringValue)))
          Get #FileBuffer, , ShortValue: sidedefs(i).sector = ItoL(ShortValue)
     Next i
     
     
     
     '========= READ VERTEXES
     lumpindex = FindLumpIndex(TempWAD, HeaderIndex, "VERTEXES")
     If (lumpindex = 0) Then Err.Raise vbObjectError + 4, , "Could not find required lump VERTEXES!"
     
     'Calculate number of vertexes
     numvertexes = TempWAD.LumpSize(lumpindex) \ 4
     
     'Allocate memory for vertexes
     'ReDim vertexes(0 To numvertexes + DECLARE_VERTICES)
     
     'Go for all vertexes to load
     Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
     For i = 0 To (numvertexes - 1)
          Get #FileBuffer, , ShortValue: vertexes(i).x = ShortValue
          Get #FileBuffer, , ShortValue: vertexes(i).y = ShortValue
     Next i
     
     
     
     '========= READ SECTORS
     lumpindex = FindLumpIndex(TempWAD, HeaderIndex, "SECTORS")
     If (lumpindex = 0) Then Err.Raise vbObjectError + 5, , "Could not find required lump SECTORS!"
     
     'Calculate number of sectors
     numsectors = TempWAD.LumpSize(lumpindex) \ 26
     
     'Allocate memory for sectors
     'ReDim sectors(0 To numsectors + DECLARE_SECTORS)
     
     'Go for all sectors to load
     Seek #FileBuffer, TempWAD.LumpAddress(lumpindex) + 1
     For i = 0 To (numsectors - 1)
          Get #FileBuffer, , ShortValue: sectors(i).hfloor = ShortValue
          Get #FileBuffer, , ShortValue: sectors(i).hceiling = ShortValue
          Get #FileBuffer, , StringValue: sectors(i).tfloor = UCase$(Trim$(UnPadded(StringValue)))
          Get #FileBuffer, , StringValue: sectors(i).tceiling = UCase$(Trim$(UnPadded(StringValue)))
          Get #FileBuffer, , ShortValue: sectors(i).Brightness = ShortValue
          Get #FileBuffer, , ShortValue: sectors(i).special = ShortValue
          Get #FileBuffer, , ShortValue: sectors(i).tag = ItoL(ShortValue)
     Next i
     
     'Leave now
     Exit Sub
     
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in ReadMapLumps(): " & Err.Description, vbCritical
End Sub


Private Sub CompleteMapLumps(ByRef TargetWAD As clsWAD, ByVal TargetHeaderLumpName As String)
     On Local Error GoTo errorhandler
     Dim HeaderIndex As Long
     Dim lumpindex As Long
     Dim FoundIndex As Long
     Dim MapLumps As Variant
     Dim NextLumpName As String
     Dim i As Long
     
     
     'This will create the map lumps defined as 'required' in
     'the TargetWAD if they do not exist yet.
     'They will be inserted after the given header, in order as defined.
     
     
     'Find the map header lump in the TargetWAD
     HeaderIndex = FindLumpIndex(TargetWAD, 1, TargetHeaderLumpName)
     
     'When it is not found, create it
     If (lumpindex = 0) Then
          TargetWAD.AddLump "", TargetHeaderLumpName
          HeaderIndex = TargetWAD.LumpCount
     End If
     
     'Start adding after the header
     lumpindex = HeaderIndex + 1
     
     'Go for all lump names as defined by map configuration
     MapLumps = mapconfig("maplumpnames").Keys
     For i = LBound(MapLumps) To UBound(MapLumps)
          
          'Make reliable lumpname
          NextLumpName = Trim$(UCase$(MapLumps(i)))
          
          'Check if this lump is required
          If (GetMapLumpType(NextLumpName) And ML_REQUIRED) = ML_REQUIRED Then
               
               'Find the lump in the target wad (but within range!)
               FoundIndex = FindLumpIndex(TargetWAD, HeaderIndex, NextLumpName, UBound(MapLumps) + 2)
               
               'Check if it is missing
               If (FoundIndex = 0) Then
                    
                    'Create the lump
                    TargetWAD.AddLump "", NextLumpName, lumpindex
               End If
               
               'Next index
               lumpindex = lumpindex + 1
          End If
     Next i
     
     'Leave now
     Exit Sub
     
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in CompleteMapLumps(): " & Err.Description, vbCritical
End Sub

Public Function MapUnload() As Boolean
     
     'Unload the script editor if its loaded
     If (ScriptEditor) Then Unload frmScript
     
     'Stop 3D Mode if still running
     If (Running3D) Then
          
          'Stop 3D Mode
          Stop3DMode
          Running3D = False
     End If
     
     'Check if the map changed
     If (mapchanged) Then
          
          'Ask for saving changes
          Select Case MsgBox("Do you want to save changes to " & mapfilename & " (" & maplumpname & ")?", vbQuestion Or vbYesNoCancel)
               
               'YES: Save changes now
               Case vbYes: frmMain.itmFileSaveMap_Click
               
               'CANCEL: Leave immediately
               Case vbCancel: Exit Function
          End Select
     End If
     
     'Close files
     If Not (IWAD Is Nothing) Then IWAD.CloseFile
     If Not (AddWAD Is Nothing) Then AddWAD.CloseFile
     If Not (MapWAD Is Nothing) Then MapWAD.CloseFile
     If Not (TempWAD Is Nothing) Then TempWAD.CloseFile
     
     'Remove temporary file
     If ((Dir(maptempfile) <> "") And (maptempfile <> "")) Then Kill maptempfile
     
     'No more map loaded
     mapchanged = False
     mapnodeschanged = False
     Set mapconfig = Nothing
     mapfile = ""
     mapfilename = ""
     addwadfile = ""
     frmMain.Caption = App.Title
     Set MapWAD = Nothing
     Set TempWAD = Nothing
     
     'No more selections
     Set selected = New Dictionary
     numselected = 0
     Set dragselected = New Dictionary
     dragnumselected = 0
     Erase changedlines()
     ReDim changedlines(0)
     numchangedlines = 0
     
     'Deallocate memory
     Erase things, vertexes, linedefs, sidedefs, sectors
     numthings = 0
     numvertexes = 0
     numlinedefs = 0
     numsidedefs = 0
     numsectors = 0
     
     'Unload sprites
     CleanUpSpriteImages
     
     'Unload textures
     UnloadAllTextures
     
     'Unload flats
     UnloadAllFlats
     
     'Disable controls
     DisableMapEditing
     
     'Terminate Undo/Redo
     TerminateUndoRedo
     
     'Terminate renderer
     TerminateMapRenderer
     
     'Restore original background color
     frmMain.picMap.BackColor = vbApplicationWorkspace
     
     'Clean up temporary files
     CleanUpTemporaries
     
     'Clear screen
     Set frmMain.picMap.Picture = Nothing
     frmMain.picMap.Refresh
     
     'Map unloaded
     MapUnload = True
End Function

Public Sub RemoveLinedef(ByVal LinedefIndex As Long, Optional ByVal RemoveAllSidedefs As Boolean = True, Optional ByVal RemoveAllUnusedVertices As Boolean = True, Optional ByVal RemoveAllUnusedSectors As Boolean = True)
     Dim lastlinedef As Long
     Dim i As Long
     
     'Boundary check
     If (LinedefIndex < 0) Or (LinedefIndex >= numlinedefs) Then Exit Sub
     
     'Check if we should delete all sidedefs
     If RemoveAllSidedefs Then
          
          'Remove both sidedefs of the line
          'Do not modify its double-sided flag, the line will go bye-bye anyway
          If (linedefs(LinedefIndex).s1 > -1) And (linedefs(LinedefIndex).s1 < numsidedefs) Then RemoveSidedef linedefs(LinedefIndex).s1, False, RemoveAllUnusedSectors, False
          If (linedefs(LinedefIndex).s2 > -1) And (linedefs(LinedefIndex).s2 < numsidedefs) Then RemoveSidedef linedefs(LinedefIndex).s2, False, RemoveAllUnusedSectors, False
     End If
     
     'Check if we should remove unused vertices
     If RemoveAllUnusedVertices Then
          
          'Remove this vertex when no lines are referring to it
          If (CountVertexLinedefs(linedefs(0), numlinedefs, linedefs(LinedefIndex).v1) = 1) Then RemoveVertex linedefs(LinedefIndex).v1
          If (CountVertexLinedefs(linedefs(0), numlinedefs, linedefs(LinedefIndex).v2) = 1) Then RemoveVertex linedefs(LinedefIndex).v2
     End If
     
     'Check if this is not the last linedef
     If (numlinedefs > 1) Then
          
          'Calculate the last linedef index
          lastlinedef = numlinedefs - 1
          
          'Remove from selection if in there
          If (selectedtype = EM_LINES) And (selected.Exists(CStr(LinedefIndex))) Then
               selected.Remove CStr(LinedefIndex)
               numselected = selected.Count
          End If
          
          'Replace the linedef with the last one
          linedefs(LinedefIndex) = linedefs(lastlinedef)
          
          'Re-reference sidedefs to the moved linedef
          Rereference_SidedefsLinedef VarPtr(sidedefs(0)), numsidedefs, lastlinedef, LinedefIndex
          
          'Update changed lines
          For i = LBound(changedlines) To UBound(changedlines)
               
               'Check if same as moved linedef
               If (changedlines(i) = lastlinedef) Then changedlines(i) = LinedefIndex
          Next i
     End If
     
     'Decrease number of linedefs
     numlinedefs = numlinedefs - 1
End Sub

Public Sub RemoveSector(ByVal SectorIndex As Long, Optional ByVal RemoveAllSidedefs As Boolean = True)
     Dim lastsector As Long
     Dim sd As Long
     
     'Boundary check
     If (SectorIndex < 0) Or (SectorIndex >= numsectors) Then Exit Sub
     
     'Check if we should delete all sidedefs
     If RemoveAllSidedefs Then
          
          'Go for all sidedefs to delete
          Do While (sd < numsidedefs)
               
               'Check if this sidedef refers to this sector
               If (sidedefs(sd).sector = SectorIndex) Then
                    
                    'Remove the sidedef
                    RemoveSidedef sd, True, False
               Else
                    
                    'Next sidedef
                    sd = sd + 1
               End If
          Loop
     End If
     
     'Check if this is not the last sector
     If (numsectors > 1) Then
          
          'Calculate the last sector index
          lastsector = numsectors - 1
          
          'Remove from selection if in there
          If (selectedtype = EM_SECTORS) And (selected.Exists(CStr(SectorIndex))) Then
               selected.Remove CStr(SectorIndex)
               numselected = selected.Count
          End If
          
          'Replace the sector with the last one
          sectors(SectorIndex) = sectors(lastsector)
          
          'Re-reference sidedefs to the moved sector
          Rereference_Sectors VarPtr(sidedefs(0)), numsidedefs, lastsector, SectorIndex
     End If
     
     'Decrease number of sectors
     numsectors = numsectors - 1
End Sub

Public Sub RemoveSidedef(ByVal SidedefIndex As Long, Optional ByVal UnsetDoublesided As Boolean = True, Optional ByVal RemoveUnusedSector As Boolean = True, Optional ByVal RemoveUnusedLinedefs As Boolean = True)
     Dim lastsidedef As Long
     Dim ld As Long, s As Long
     Dim os As Long
     
     'Boundary check
     If (SidedefIndex < 0) Or (SidedefIndex >= numsidedefs) Then Exit Sub
     
     'Get the sidedef on the other side of the line
     If (linedefs(sidedefs(SidedefIndex).linedef).s1 = SidedefIndex) Then
          os = linedefs(sidedefs(SidedefIndex).linedef).s2
     Else
          os = linedefs(sidedefs(SidedefIndex).linedef).s1
     End If
     
     'Check if we should correct linedef properties
     If UnsetDoublesided Then
          
          'Remove double-sided flag from linedef
          linedefs(sidedefs(SidedefIndex).linedef).Flags = linedefs(sidedefs(SidedefIndex).linedef).Flags And Not LDF_TWOSIDED
          
          'Add impassable flag
          linedefs(sidedefs(SidedefIndex).linedef).Flags = linedefs(sidedefs(SidedefIndex).linedef).Flags Or LDF_IMPASSIBLE
          
          'Check if there is another side
          If (os > -1) Then
               
               'Check if its missing middle texture
               If (IsTextureName(sidedefs(os).Middle) = False) Then
                    
                    'Check if it has an upper texture
                    If (IsTextureName(sidedefs(os).Upper)) Then
                         
                         'Apply upper to middle
                         sidedefs(os).Middle = sidedefs(os).Upper
                         
                    'Otherwise check if it has a lower texture
                    ElseIf (IsTextureName(sidedefs(os).Lower)) Then
                         
                         'Apply upper to middle
                         sidedefs(os).Middle = sidedefs(os).Lower
                         
                    'And otherwise use the default
                    Else
                         
                         'Apply default to middle
                         sidedefs(os).Middle = Config("defaulttexture")("middle")
                    End If
               End If
               
               'Remove upper and lower textures
               sidedefs(os).Upper = "-"
               sidedefs(os).Lower = "-"
          End If
     End If
     
     'Reference linedef to sidedef -1 (none)
     If (linedefs(sidedefs(SidedefIndex).linedef).s1 = SidedefIndex) Then linedefs(sidedefs(SidedefIndex).linedef).s1 = -1
     If (linedefs(sidedefs(SidedefIndex).linedef).s2 = SidedefIndex) Then linedefs(sidedefs(SidedefIndex).linedef).s2 = -1
     
     'Keep linedef and sector for later use
     ld = sidedefs(SidedefIndex).linedef
     s = sidedefs(SidedefIndex).sector
     
     'Check if this is not the last sidedef
     If (numsidedefs > 1) Then
          
          'Calculate the last sidedef index
          lastsidedef = numsidedefs - 1
          
          'Replace the sidedef with the last one
          sidedefs(SidedefIndex) = sidedefs(lastsidedef)
          
          'Re-reference linedefs to the moved sidedef
          Rereference_Sidedefs linedefs(0), numlinedefs, lastsidedef, SidedefIndex
     End If
     
     'Decrease number of sidedefs
     numsidedefs = numsidedefs - 1
     
     'Check if we should correct linedef properties
     'if it has only a sidedef 2 left over
     If UnsetDoublesided Then
          
          'Check if it has only sidedef 2
          If (linedefs(ld).s1 = -1) And (linedefs(ld).s2 > -1) Then
               
               'Flip the linedef
               FlipLinedefVertices ld
               
               'Flip sidedefs (because they flipped with the linedef's vertices)
               FlipLinedefSidedefs ld
          End If
     End If
     
     'Check if we should remove the linedef when it
     'has no more attached sidedefs
     If RemoveUnusedLinedefs Then
          
          'Check if the linedef is bald
          If (linedefs(ld).s1 = -1) And (linedefs(ld).s2 = -1) Then
               
               'Remove this linedef
               RemoveLinedef ld, False
          End If
     End If
     
     'Check if we should remove the referenced sector if
     'this is the last sidedef referring to it
     If RemoveUnusedSector Then
          
          'Check if this is the last sidedef referring to its sector
          If (CountSectorSidedefs(VarPtr(sidedefs(0)), numsidedefs, s) = 0) Then
               
               'Remove this sector
               RemoveSector s, False
          End If
     End If
End Sub

Public Sub RemoveThing(ByVal ThingIndex As Long)
     Dim LastTh As Long
     
     'Boundary check
     If (ThingIndex < 0) Or (ThingIndex >= numthings) Then Exit Sub
     
     'Check if this is not the last thing
     If (numthings > 1) Then
          
          'Calculate the last thing index
          LastTh = numthings - 1
          
          'Remove from selection if in there
          If (selectedtype = EM_THINGS) And (selected.Exists(CStr(ThingIndex))) Then
               selected.Remove CStr(ThingIndex)
               numselected = selected.Count
          End If
          
          'Replace the thing with the last one
          things(ThingIndex) = things(LastTh)
          
          'Rereference
          If (PositionThing = LastTh) Then PositionThing = ThingIndex
     End If
     
     'Decrease number of things
     numthings = numthings - 1
End Sub

Public Sub RemoveVertex(ByVal VertexIndex As Long)
     Dim lastvertex As Long
     
     'NOTE: This function does NOT remove any linedefs that refer to this vertex!
     
     'Boundary check
     If (VertexIndex < 0) Or (VertexIndex >= numvertexes) Then Exit Sub
     
     'Check if this is not the last vertex
     If (numvertexes > 1) Then
          
          'Calculate the last vertex index
          lastvertex = numvertexes - 1
          
          'Remove from selection if in there
          If (selectedtype = EM_VERTICES) And (selected.Exists(CStr(VertexIndex))) Then
               selected.Remove CStr(VertexIndex)
               numselected = selected.Count
          End If
          
          'Replace the vertex with the last one
          vertexes(VertexIndex) = vertexes(lastvertex)
          
          'Re-reference linedefs to the moved vertex
          Rereference_Vertices linedefs(0), numlinedefs, lastvertex, VertexIndex
     End If
     
     'Decrease number of vertices
     numvertexes = numvertexes - 1
End Sub
