Attribute VB_Name = "modTextures"
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


'Constants
Public Const TRANSPARENCY_INDEX As Long = 1
Public Const DEFAULT_INDEX As Long = 16 * 5
Public Const ALTERNATE_INDEX As Long = 0
Public Const BLACK_LIMIT As Long = 5


'Texture and flat flags
Public Enum ENUM_IMAGEFLAGS
     IF_WORLDCOORDS = -1
End Enum


'Texture and flat formats
Public Enum ENUM_IMAGEFORMAT
     TF_UNKNOWN          'Unknown format
     TF_INVALID          'Invalid format
     TF_IMAGE            'Doom Image format  (column list rendered data)
     TF_FLAT             'Doom Flat format   (raw 8-bit pixel data)
     TF_PNG              'Portable Network Graphic
     TF_BITMAP_P8        'Bitmap 8-bit Paletted
     TF_BITMAP_B5G6R5    'Bitmap 16-bit
     TF_BITMAP_B8G8R8    'Bitmap 24-bit
     TF_BITMAP_A8B8G8R8  'Bitmap 32-bit
End Enum


'Texture and flat source files
Public Enum ENUM_IMAGESOURCE
     TS_MAPWAD
     TS_ADDWAD
     TS_IWAD
     TS_TEXDIR
     TS_FLATDIR
End Enum


'API Declarations
Public Declare Sub BuildConversionTable Lib "builder.dll" (ByRef palette As BITMAPRGB, ByVal numcolors As Long)
Public Declare Sub Draw_Image Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)
Public Declare Sub Draw_Flat Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)
Public Declare Sub Draw_BitmapP8 Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)
Public Declare Sub Draw_BitmapB5G6R5 Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)
Public Declare Sub Draw_BitmapB8G8R8 Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)
Public Declare Sub Draw_BitmapA8B8G8R8 Lib "builder.dll" (ByRef Texture As Byte, ByVal tw As Long, ByVal th As Long, ByVal pdata As String, ByVal pdatalen As Long, ByVal px As Long, ByVal py As Long, ByVal pw As Long, ByVal ph As Long, ByVal TRANSPARENCY_INDEX As Byte, ByVal ALTERNATE_INDEX As Byte)


'PLAYPAL
Public playpal(0 To 255) As BITMAPRGB

'PNAMES
Private pnames() As Long
Private pfile() As Long
Private numpnames As Long

'Settings
Private defaulttexturescale As Single

'TEXTUREs
Public textures As Dictionary                'clsTexture objects with TextureName as key
Public alltextures As Dictionary             'clsTexture objects with TextureName as key

Public Sub GetScaledTexturePicture(ByVal texturename As String, ByRef target As image, Optional ByVal NoCaching As Boolean, Optional Required As Boolean)
     Dim Texture As clsImage
     Dim sw As Long, sh As Long
     
     'Check if texture is set
     'If ((TextureName = "-") Or (TextureName = "")) Then
     If (IsTextureName(texturename) = False) Then
          
          'Check if required
          If (Required And (texturename = "-")) Then
               
               'Set to missing texture
               Set target.Picture = frmMain.imgMissingTexture.Picture
          Else
               
               'Set nothing
               Set target.Picture = Nothing
          End If
          
          'Move the box
          target.Move 0, 0, 64, 64
          
     'Check if the texture is known
     ElseIf alltextures.Exists(UCase$(texturename)) Then
          
          'Get the texture object
          Set Texture = alltextures(UCase$(texturename))
          
          'Set the texture
          Set target.Picture = Texture.Picture(NoCaching)
          
          'Move the image box depending on scale
          Texture.GetScale 64, 64, sw, sh, NoCaching
          target.Move (64 - sw) \ 2, (64 - sh) \ 2, sw, sh
     Else
          
          'Return the Unknown Texture image
          Set target.Picture = frmMain.imgUnknownTexture.Picture
          
          'Move the box
          target.Move 0, 0, 64, 64
     End If
End Sub

Public Sub GetScaledTexturePictureEx(ByVal texturename As String, ByRef target As image, ByVal BoxWidth As Long, ByVal BoxHeight As Long, Optional ByVal NoCaching As Boolean, Optional Required As Boolean)
     Dim Texture As clsImage
     Dim sw As Long, sh As Long
     
     'Check if texture is set
     'If ((TextureName = "-") Or (TextureName = "")) Then
     If (IsTextureName(texturename) = False) Then
          
          'Check if required
          If (Required And (texturename = "-")) Then
               
               'Set to missing texture
               Set target.Picture = frmMain.imgMissingTexture.Picture
          Else
               
               'Set nothing
               Set target.Picture = Nothing
          End If
          
          'Move the box
          target.Move (BoxWidth - 64) \ 2, (BoxHeight - 64) \ 2, 64, 64
          
     'Check if the texture is known
     ElseIf alltextures.Exists(UCase$(texturename)) Then
          
          'Get the texture object
          Set Texture = alltextures(UCase$(texturename))
          
          'Set the texture
          Set target.Picture = Texture.Picture(NoCaching)
          
          'Move the image box depending on scale
          Texture.GetScale BoxWidth, BoxHeight, sw, sh, NoCaching
          target.Move (BoxWidth - sw) \ 2, (BoxHeight - sh) \ 2, sw, sh
     Else
          
          'Return the Unknown Texture image
          Set target.Picture = frmMain.imgUnknownTexture.Picture
          
          'Move the box
          target.Move (BoxWidth - 64) \ 2, (BoxHeight - 64) \ 2, 64, 64
     End If
End Sub

Public Function GetTextureFileData(ByRef LumpName As String) As String
     Dim Filename As String
     Dim filedata As String
     Dim filebuf As Integer
     
     'Make the full file/pathname
     Filename = addtexdir & LumpName & ".*"
     Filename = Dir(Filename)
     
     'Found anything?
     If (Filename <> "") Then
          
          'Read the data
          filebuf = FreeFile
          Open addtexdir & Filename For Binary Access Read Lock Write As filebuf
          filedata = Space$(LOF(filebuf))
          Get #filebuf, , filedata
          Close #filebuf
          
          'Return data
          GetTextureFileData = filedata
     End If
End Function

Public Function LoadAllTextures() As Boolean
     Dim Locations As Variant
     Dim StartIndex As Long, EndIndex As Long
     Dim i As Long
     
     'Load the PLAYPAL
     If (Not LoadPlaypal(MapWAD)) Then
          If (Not LoadPlaypal(AddWAD)) Then
               If (Not LoadPlaypal(IWAD)) Then
                    
                    'Could not load any PLAYPAL lump
                    ErrorLog_Add "WARNING: Could not find required lump PLAYPAL", False
               End If
          End If
     End If
     
     'Determine default scale
     defaulttexturescale = Val(mapconfig("defaulttexturescale"))
     If (defaulttexturescale < 0.00001) Then defaulttexturescale = 1
     
     'Create textures collections
     Set textures = New Dictionary
     Set alltextures = New Dictionary
     
     'Load standard textures from IWAD
     LoadPNames IWAD
     LoadTextureSet IWAD, "TEXTURE1", True
     LoadTextureSet IWAD, "TEXTURE2", False
     
     'Load standard textures from AddWAD
     LoadPNames AddWAD
     LoadTextureSet AddWAD, "TEXTURE1", True
     LoadTextureSet AddWAD, "TEXTURE2", False
     
     'Load standard textures from MapWAD
     LoadPNames MapWAD
     LoadTextureSet MapWAD, "TEXTURE1", True
     LoadTextureSet MapWAD, "TEXTURE2", False
     
     'Check if textures could be found
     'If (textures.Count = 0) Then
     '
     '     'Could not load any TEXTURE lump
     '     ErrorLog_Add "WARNING: Could not find required lumps TEXTURE1 or TEXTURE2", False
     'End If
     
     
     'Go for all defined texture source locations
     Locations = mapconfig("textures").Items
     For i = LBound(Locations) To UBound(Locations)
          
          'Load textures from IWAD
          StartIndex = FindLumpIndex(IWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(IWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadTextureRange IWAD, TS_IWAD, StartIndex, EndIndex
          
          'Load textures from AddWAD
          StartIndex = FindLumpIndex(AddWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(AddWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadTextureRange AddWAD, TS_ADDWAD, StartIndex, EndIndex
          
          'Load textures from MapWAD
          StartIndex = FindLumpIndex(MapWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(MapWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadTextureRange MapWAD, TS_MAPWAD, StartIndex, EndIndex
     Next i
     
     
     'Load textures from specified directory
     If (addtexdir <> "") Then LoadTextureDirectory addtexdir, TS_TEXDIR
     
     
     'Clean up
     Erase Locations
     Erase pnames, pfile
     numpnames = 0
     
     'Sort textures
     SortTextures
     
     'No problems
     LoadAllTextures = True
End Function

Private Sub LoadTextureRange(ByRef WadFile As clsWAD, ByVal FileSource As ENUM_IMAGESOURCE, ByVal StartIndex As Long, ByVal EndIndex As Long)
     Dim Texture As clsImage
     Dim i As Long, f As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim LumpName As String
     Dim ListTexture As Boolean
     
     'Check if not closed
     If (WadFile.Filename = "") Then Exit Sub
     
     'Get the filter lists from config
     RequiredList = mapconfig("texturesfilter").Items
     LimitedList = mapconfig("notexturesfilter").Items
     
     'Go for all lumps between start and end
     For i = StartIndex To EndIndex
          
          'Check if not empty
          If (WadFile.LumpSize(i) > 0) Then
               
               'Get lump name
               LumpName = UCase$(Trim$(WadFile.LumpName(i)))
               
               'Continue if name is valid
               If (LumpName <> "") Then
                    
                    'Create new texture
                    Set Texture = New clsImage
                    
                    'Set the properties
                    With Texture
                         .Name = LumpName
                         .width = 0
                         .height = 0
                         .ScaleX = defaulttexturescale
                         .ScaleY = defaulttexturescale
                         .FlatCandidate = True
                         .AddPatch 0, 0, 0, 0, i, FileSource, TF_UNKNOWN
                    End With
                    
                    'Remove if already added before (overwrite)
                    If (textures.Exists(LumpName)) Then textures.Remove LumpName
                    If (alltextures.Exists(LumpName)) Then alltextures.Remove LumpName
                    
                    'Store the texture info
                    alltextures.Add LumpName, Texture
                    
                    'Go by each required filter
                    ListTexture = False
                    For f = LBound(RequiredList) To UBound(RequiredList)
                         If (LumpName Like RequiredList(f)) Then ListTexture = True: Exit For
                    Next f
                    
                    'Go by each limited filter
                    For f = LBound(LimitedList) To UBound(LimitedList)
                         If (LumpName Like LimitedList(f)) Then ListTexture = False: Exit For
                    Next f
                    
                    'Add texture to listing if not filtered out
                    If ListTexture Then textures.Add LumpName, Texture
                    
                    'Clean up references
                    Set Texture = Nothing
               End If
          End If
     Next i
End Sub

Private Sub LoadTextureDirectory(ByVal directory As String, ByVal FileSource As ENUM_IMAGESOURCE)
     Dim Texture As clsImage
     Dim i As Long, f As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim Filename As String
     Dim LumpName As String
     Dim ListTexture As Boolean
     Dim ext As String
     
     'Get the filter lists from config
     RequiredList = mapconfig("texturesfilter").Items
     LimitedList = mapconfig("notexturesfilter").Items
     
     'Find first file
     Filename = Dir(directory & "*.*")
     
     'Continue until end of directory
     While (Filename <> "")
          
          'Get file extension
          ext = LCase$(right$(Filename, 3))
          
          'Check if this is a known extension
          If (ext = "bmp") Or (ext = "png") Then
               
               'Determine texture name
               LumpName = UCase$(Mid$(Filename, 1, Len(Filename) - 4))
               If (Len(LumpName) > 8) Then LumpName = left$(LumpName, 8)
               
               'Continue if name is valid
               If (LumpName <> "") Then
                    
                    'Create new texture
                    Set Texture = New clsImage
                    
                    'Set the properties
                    With Texture
                         .Name = LumpName
                         .width = 0
                         .height = 0
                         .ScaleX = defaulttexturescale
                         .ScaleY = defaulttexturescale
                         .FlatCandidate = True
                         .AddPatch 0, 0, 0, 0, i, FileSource, TF_UNKNOWN
                    End With
                    
                    'Remove if already added before (overwrite)
                    If (textures.Exists(LumpName)) Then textures.Remove LumpName
                    If (alltextures.Exists(LumpName)) Then alltextures.Remove LumpName
                    
                    'Store the texture info
                    alltextures.Add LumpName, Texture
                    
                    'Go by each required filter
                    ListTexture = False
                    For f = LBound(RequiredList) To UBound(RequiredList)
                         If (LumpName Like RequiredList(f)) Then ListTexture = True: Exit For
                    Next f
                    
                    'Go by each limited filter
                    For f = LBound(LimitedList) To UBound(LimitedList)
                         If (LumpName Like LimitedList(f)) Then ListTexture = False: Exit For
                    Next f
                    
                    'Add texture to listing if not filtered out
                    If ListTexture Then textures.Add LumpName, Texture
                    
                    'Clean up references
                    Set Texture = Nothing
               End If
          End If
          
          'Find next
          Filename = Dir()
     Wend
End Sub



Private Function LoadPlaypal(ByRef WadFile As clsWAD) As Boolean
     Dim FileBuffer As Integer
     Dim playpalindex As Long
     Dim Color As Byte
     Dim i As Long
     
     'Check if not closed
     If (WadFile.Filename = "") Then Exit Function
     
     'Get the WadFile filebuffer
     FileBuffer = WadFile.FileBuffer
     
     'Find the PLAYPAL lump
     playpalindex = FindLumpIndex(WadFile, 1, "PLAYPAL")
     If (playpalindex = 0) Then Exit Function
     
     'Seek to its address
     Seek #FileBuffer, WadFile.LumpAddress(playpalindex) + 1
     
     'Read all palette entries
     For i = 0 To 255
          
          'Read red color
          Get #FileBuffer, , Color
          playpal(i).rgbRed = Color
          
          'Read green color
          Get #FileBuffer, , Color
          playpal(i).rgbGreen = Color
          
          'Read blue color
          Get #FileBuffer, , Color
          playpal(i).rgbBlue = Color
          
          'Check if the color is black
          If (playpal(i).rgbRed < BLACK_LIMIT) And _
             (playpal(i).rgbGreen < BLACK_LIMIT) And _
             (playpal(i).rgbBlue < BLACK_LIMIT) Then
               
               'Make the color just not really black
               'we use black for transparency (reliable with DirectX)
               playpal(i).rgbRed = BLACK_LIMIT
               playpal(i).rgbGreen = BLACK_LIMIT
               playpal(i).rgbBlue = BLACK_LIMIT
          End If
     Next i
     
     'Copy the array
     'CopyMemory playpal_wb(0), playpal(0), 256 * 4
     
     'Make the tranparency index exactly black
     playpal(TRANSPARENCY_INDEX).rgbRed = 0
     playpal(TRANSPARENCY_INDEX).rgbGreen = 0
     playpal(TRANSPARENCY_INDEX).rgbBlue = 0
     
     'Make the tranparency index the windows color
     'playpal_wb(TRANSPARENCY_INDEX) = LongToBITMAPRGB(LongToBGRLong(GetSysColor(WCOLOR_APPWORKSPACE)))
     
     'Make the color conversion table
     BuildConversionTable playpal(0), 256
     
     
     'No problems
     LoadPlaypal = True
End Function

Private Function LoadPNames(ByRef WadFile As clsWAD) As Boolean
     Dim FileBuffer As Integer
     Dim pnamesindex As Long
     Dim pnamestrings() As String * 8
     Dim LookupIWAD As New Dictionary
     Dim LookupMapWAD As New Dictionary
     Dim LookupAddWAD As New Dictionary
     Dim i As Long
     Dim nstr As String
     
     'Check if not closed
     If (WadFile.Filename = "") Then Exit Function
     
     'Get the IWAD filebuffer
     FileBuffer = WadFile.FileBuffer
     
     'Find the PNAMES lump
     pnamesindex = FindLumpIndex(WadFile, 1, "PNAMES")
     If (pnamesindex = 0) Then Exit Function
     
     'Seek to its address
     Seek #FileBuffer, WadFile.LumpAddress(pnamesindex) + 1
     
     'Read the number of PNAMES
     Get #FileBuffer, , numpnames
     
     'Allocate memory for PNAMES
     ReDim pnames(0 To numpnames - 1)
     ReDim pfile(0 To numpnames - 1)
     ReDim pnamestrings(0 To numpnames - 1)
     
     'Red the PNAMES from file
     Get #FileBuffer, , pnamestrings
     
     'Make lookup dictionaries
     For i = 1 To IWAD.LumpCount
          nstr = UCase$(IWAD.LumpName(i))
          If (LookupIWAD.Exists(nstr) = False) Then LookupIWAD.Add nstr, i
     Next i
     For i = 1 To MapWAD.LumpCount
          nstr = UCase$(MapWAD.LumpName(i))
          If (LookupMapWAD.Exists(nstr) = False) Then LookupMapWAD.Add nstr, i
     Next i
     For i = 1 To AddWAD.LumpCount
          nstr = UCase$(AddWAD.LumpName(i))
          If (LookupAddWAD.Exists(nstr) = False) Then LookupAddWAD.Add nstr, i
     Next i
     
     'Go for all string names to find the lump indices
     For i = 0 To (numpnames - 1)
          
          'Get the pname string
          nstr = UCase$(UnPadded(pnamestrings(i)))
          
          'Find lump in MapWAD
          If LookupMapWAD.Exists(nstr) Then
               
               'Set the lookup table values
               pfile(i) = TS_MAPWAD
               pnames(i) = LookupMapWAD(nstr)
          Else
               
               'Find lump in AddWAD
               If LookupAddWAD.Exists(nstr) Then
                    
                    'Set the lookup table values
                    pfile(i) = TS_ADDWAD
                    pnames(i) = LookupAddWAD(nstr)
               Else
                    
                    'Find lump in IWAD
                    If LookupIWAD.Exists(nstr) Then
                         
                         'Set the lookup table values
                         pfile(i) = TS_IWAD
                         pnames(i) = LookupIWAD(nstr)
                    Else
                         
                         'Lump could not be found
                         ErrorLog_Add "WARNING: Could not find the required lump for the patch " & UnPadded(pnamestrings(i)), False
                         
                         'This indicates the pname was not found
                         pfile(i) = -1
                         pnames(i) = 0
                    End If
               End If
          End If
     Next i
     
     'No problems
     LoadPNames = True
End Function

Private Function LoadTextureSet(ByRef WadFile As clsWAD, ByVal LumpName As String, ByVal DiscardFirst As Boolean) As Boolean
     Dim NewTexture As clsImage
     Dim FileBuffer As Integer
     Dim texturesindex As Long
     Dim numtex As Long
     Dim numpatch As Long
     Dim ReadString As String * 8
     Dim ReadShort As Integer
     Dim ReadByte As Byte
     Dim i As Long, p As Long, f As Long
     Dim px As Integer, py As Integer
     Dim pi As Integer, pd As Integer, pc As Integer
     Dim AddTexture As Boolean
     Dim AddedPatches As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim ListTexture As Boolean
     Dim StrifePatch As Boolean
     
     'Check if not closed
     If (WadFile.Filename = "") Then Exit Function
     
     'Get the WadFile filebuffer
     FileBuffer = WadFile.FileBuffer
     
     'Find the PNAMES lump
     texturesindex = FindLumpIndex(WadFile, 1, LumpName)
     If (texturesindex = 0) Then Exit Function
     
     'Get the filter lists from config
     RequiredList = mapconfig("texturesfilter").Items
     LimitedList = mapconfig("notexturesfilter").Items
     
     'Seek to its address
     Seek #FileBuffer, WadFile.LumpAddress(texturesindex) + 1
     
     'Get the number of textures
     Get #FileBuffer, , numtex
     
     'Skip the offset bytes, we'll read it as a sequence
     Seek #FileBuffer, WadFile.LumpAddress(texturesindex) + 1 + 4 + 4 * numtex
     
     'Go for all texture definitions in file
     For i = 1 To numtex
          
          'Create new texture object
          Set NewTexture = New clsImage
          
          'Assume Doom format
          StrifePatch = False
          
          'Read data from file
          Get #FileBuffer, , ReadString: NewTexture.Name = UCase$(Trim$(UnPadded(ReadString)))
          Get #FileBuffer, , ReadShort: NewTexture.Flags = ItoL(ReadShort)
          Get #FileBuffer, , ReadByte: If (ReadByte > 0) Then NewTexture.ScaleX = CSng(ReadByte) / CSng(8) Else NewTexture.ScaleX = defaulttexturescale
          Get #FileBuffer, , ReadByte: If (ReadByte > 0) Then NewTexture.ScaleY = CSng(ReadByte) / CSng(8) Else NewTexture.ScaleY = defaulttexturescale
          Get #FileBuffer, , ReadShort: NewTexture.width = ReadShort
          Get #FileBuffer, , ReadShort: NewTexture.height = ReadShort
          Get #FileBuffer, , ReadShort: If (ReadShort <> 0) Then StrifePatch = True
          If Not StrifePatch Then
               Get #FileBuffer, , ReadShort
               Get #FileBuffer, , ReadShort: numpatch = ReadShort
          Else
               numpatch = ReadShort
          End If
          
          'Check if texture is valid
          If (NewTexture.width > 0) And (NewTexture.height > 0) And (numpatch > 0) Then
               
               'Go for all patch references in file
               AddTexture = True
               AddedPatches = 0
               For p = 1 To numpatch
                    
                    'Read data from file
                    Get #FileBuffer, , px
                    Get #FileBuffer, , py
                    Get #FileBuffer, , pi
                    If Not StrifePatch Then
                         Get #FileBuffer, , pd
                         Get #FileBuffer, , pc
                    End If
                    
                    'Check if patch number if valid
                    If (pi > -1) And (pi < numpnames) Then
                         
                         'Check if the required patch can be found
                         If (pnames(pi) > 0) And (pfile(pi) > -1) Then
                              
                              'Add patch to texture object
                              NewTexture.AddPatch px, py, NewTexture.width, NewTexture.height, pnames(pi), pfile(pi), TF_UNKNOWN
                              AddedPatches = AddedPatches + 1
                         Else
                              
                              'Patches are missing
                              'AddTexture = False
                              'Exit For
                         End If
                    Else
                         
                         'Patches are invalid
                         'AddTexture = False
                         'Exit For
                    End If
               Next p
               
               'Check if the first texture should be discarded
               If (DiscardFirst = False) Or (i > 1) Then
                    
                    'Check if texture should be added
                    If (AddTexture = True) And (AddedPatches > 0) Then
                         
                         'Remove if already added before (overwrite)
                         If (textures.Exists(NewTexture.Name)) Then textures.Remove NewTexture.Name
                         If (alltextures.Exists(NewTexture.Name)) Then alltextures.Remove NewTexture.Name
                         
                         'Add texture object to dictionary
                         alltextures.Add NewTexture.Name, NewTexture
                         
                         'Go by each required filter
                         ListTexture = False
                         For f = LBound(RequiredList) To UBound(RequiredList)
                              If (NewTexture.Name Like RequiredList(f)) Then ListTexture = True: Exit For
                         Next f
                         
                         'Go by each limited filter
                         For f = LBound(LimitedList) To UBound(LimitedList)
                              If (NewTexture.Name Like LimitedList(f)) Then ListTexture = False: Exit For
                         Next f
                         
                         'Add texture to listing if not filtered out
                         If ListTexture Then textures.Add NewTexture.Name, NewTexture
                    Else
                         
                         'Show warning
                         'ErrorLog_Add "WARNING: Texture patches for " & NewTexture.Name & " could not be found.", False
                    End If
               End If
          Else
               
               'Discard the rest of texture data
               For p = 1 To numpatch
                    
                    'Read data from file
                    Get #FileBuffer, , px
                    Get #FileBuffer, , py
                    Get #FileBuffer, , pi
                    If Not StrifePatch Then
                         Get #FileBuffer, , pd
                         Get #FileBuffer, , pc
                    End If
               Next p
               
               'Show warning
               ErrorLog_Add "WARNING: Texture patches for " & NewTexture.Name & " contain invalid header information.", False
          End If
          
          'Destroy texture object reference
          Set NewTexture = Nothing
     Next i
     
     'No problems
     LoadTextureSet = True
End Function

Public Sub PrecacheTextures()
     Dim i As Long
     Dim TextureKeys As Variant
     Dim Texture As clsImage
     
     'Go for all textures
     TextureKeys = alltextures.Keys
     For i = LBound(TextureKeys) To UBound(TextureKeys)
          
          'Get the texture object
          Set Texture = alltextures(TextureKeys(i))
          
          'Load the texture
          If (Texture.IsLoaded = False) Then Texture.LoadImage
          
          'Clean up
          Set Texture = Nothing
     Next i
End Sub

Public Sub SortTextures()
     Set textures = SortDictionary(textures)
End Sub

Public Sub UnloadAllTextures()
     
     'Clear textures
     Set textures = Nothing
     Set alltextures = Nothing
End Sub

Public Sub UnloadDirect3DTextures()
     Dim i As Long
     Dim TextureKeys As Variant
     Dim Texture As clsImage
     
     'Go for all textures
     TextureKeys = alltextures.Keys
     For i = LBound(TextureKeys) To UBound(TextureKeys)
          
          'Get the texture object
          Set Texture = alltextures(TextureKeys(i))
          
          'Unload the direct3d texture
          Set Texture.D3DTexture = Nothing
          
          'Clean up
          Set Texture = Nothing
     Next i
End Sub
