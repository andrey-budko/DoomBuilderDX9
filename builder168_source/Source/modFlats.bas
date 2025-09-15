Attribute VB_Name = "modFlats"
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


'FLATS
Public IWADfstart As Long
Public IWADfend As Long
Public WADFILEfstart As Long
Public WADFILEfend As Long
Public ADDWADfstart As Long
Public ADDWADfend As Long


'API Declarations
Public Declare Sub Flip_Flat Lib "builder.dll" (ByRef flat As Byte, ByVal tw As Long, ByVal th As Long)


Public flats As Dictionary         'clsTexture objects with TextureName as key
Public allflats As Dictionary      'clsTexture objects with TextureName as key


Public Sub GetScaledFlatPicture(ByVal FlatName As String, ByRef target As image, Optional ByVal NoCaching As Boolean)
     On Local Error Resume Next
     Dim flat As clsImage
     Dim sw As Long, sh As Long
     
     'Check if texture is set
     If (LenB(FlatName) = 0) Then
          
          'Set nothing
          Set target.Picture = Nothing
     
     'Check if the flat is known
     ElseIf allflats.Exists(UCase$(FlatName)) Then
          
          'Get the flat object
          Set flat = allflats(UCase$(FlatName))
          
          'Set the texture
          Set target.Picture = flat.Picture(NoCaching)
          
          'Move the image box depending on scale
          flat.GetScale 64, 64, sw, sh, NoCaching
          target.Move (64 - sw) \ 2, (64 - sh) \ 2, sw, sh
     Else
          
          'Return the Unknown Texture image
          Set target.Picture = frmMain.imgUnknownFlat.Picture
          
          'Move the box
          target.Move 0, 0, 64, 64
     End If
End Sub

Public Function LoadAllFlats() As Boolean
     Dim Locations As Variant
     Dim StartIndex As Long, EndIndex As Long
     Dim i As Long
     
     'Create flats dictionary
     If (Val(mapconfig("mixtexturesflats")) = vbChecked) Then
          Set flats = textures
          Set allflats = alltextures
     Else
          Set flats = New Dictionary
          Set allflats = New Dictionary
     End If
     
     'Flat lumps search starts after first F_START and ends at F_END
     'WADFILEfstart = FindLumpIndex(MapWAD, 1, "F_START") + 1
     'If (WADFILEfstart = 1) Then WADFILEfstart = FindLumpIndex(MapWAD, 1, "FF_START") + 1
     'WADFILEfend = FindLumpIndex(MapWAD, 1, "F_END")
     'If (WADFILEfend = 0) Then WADFILEfend = FindLumpIndex(MapWAD, 1, "FF_END")
     'If (WADFILEfend = 0) Then WADFILEfend = MapWAD.LumpCount
     
     'ADDWADfstart = FindLumpIndex(AddWAD, 1, "F_START") + 1
     'If (ADDWADfstart = 1) Then ADDWADfstart = FindLumpIndex(AddWAD, 1, "FF_START") + 1
     'ADDWADfend = FindLumpIndex(AddWAD, 1, "F_END")
     'If (ADDWADfend = 0) Then ADDWADfend = FindLumpIndex(AddWAD, 1, "FF_END")
     'If (ADDWADfend = 0) Then ADDWADfend = AddWAD.LumpCount
     
     'IWADfstart = FindLumpIndex(IWAD, 1, "F_START") + 1
     'If (IWADfstart = 1) Then IWADfstart = FindLumpIndex(IWAD, 1, "FF_START") + 1
     'IWADfend = FindLumpIndex(IWAD, 1, "F_END")
     'If (IWADfend = 0) Then IWADfend = FindLumpIndex(IWAD, 1, "FF_END")
     'If (IWADfend = 0) Then IWADfend = IWAD.LumpCount
     
     'Add IWAD flats
     'If (IWADfstart > 1) Then LoadFlats IWAD, TS_IWAD, IWADfstart, IWADfend
     
     'Add ADDWAD flats
     'If (ADDWADfstart > 1) Then LoadFlats AddWAD, TS_ADDWAD, ADDWADfstart, ADDWADfend
     
     'Add WAD File flats
     'If (WADFILEfstart > 1) Then LoadFlats MapWAD, TS_MAPWAD, WADFILEfstart, WADFILEfend
     
     
     'Go for all defined flats source locations
     Locations = mapconfig("flats").Items
     For i = LBound(Locations) To UBound(Locations)
          
          'Load flats from IWAD
          StartIndex = FindLumpIndex(IWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(IWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadFlatRange IWAD, TS_IWAD, StartIndex, EndIndex
          
          'Load flats from AddWAD
          StartIndex = FindLumpIndex(AddWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(AddWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadFlatRange AddWAD, TS_ADDWAD, StartIndex, EndIndex
          
          'Load flats from MapWAD
          StartIndex = FindLumpIndex(MapWAD, 1, Locations(i)("start"))
          EndIndex = FindLumpIndex(MapWAD, 1, Locations(i)("end"))
          If (StartIndex <= EndIndex) And (StartIndex > 0) Then LoadFlatRange MapWAD, TS_MAPWAD, StartIndex, EndIndex
     Next i
     
     
     'Load flats from specified directory
     If (addflatdir <> "") Then LoadFlatDirectory addflatdir, TS_FLATDIR
     
     
     'Sort flats
     SortFlats
     
     'Sorting created a new flats dictionary,
     'so we most reference it with textures
     'when using mixed resources
     If (Val(mapconfig("mixtexturesflats")) = vbChecked) Then Set textures = flats
     
     'No problems
     LoadAllFlats = True
End Function

Private Sub LoadFlatDirectory(ByVal directory As String, ByVal FileSource As ENUM_IMAGESOURCE)
     Dim flat As clsImage
     Dim i As Long, f As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim filename As String
     Dim lumpname As String
     Dim ListFlat As Boolean
     Dim ext As String
     
     'Get the filter lists from config
     RequiredList = mapconfig("flatsfilter").Items
     LimitedList = mapconfig("noflatsfilter").Items
     
     'Find first file
     filename = Dir(directory & "*.*")
     
     'Continue until end of directory
     While (filename <> "")
          
          'Get file extension
          ext = LCase$(right$(filename, 3))
          
          'Check if this is a known extension
          If (ext = "bmp") Or (ext = "png") Then
               
               'Determine texture name
               lumpname = UCase$(Mid$(filename, 1, Len(filename) - 4))
               If (Len(lumpname) > 8) Then lumpname = left$(lumpname, 8)
               
               'Continue if name is valid
               If (lumpname <> "") Then
                    
                    'Create new flat
                    Set flat = New clsImage
                    
                    'Set the properties
                    With flat
                         .Name = lumpname
                         .width = 0
                         .height = 0
                         .ScaleX = 1
                         .ScaleY = 1
                         .FlatCandidate = True
                         .AddPatch 0, 0, 0, 0, i, FileSource, TF_UNKNOWN
                    End With
                    
                    'Remove if already added before (overwrite)
                    If (flats.Exists(lumpname)) Then flats.Remove lumpname
                    If (allflats.Exists(lumpname)) Then allflats.Remove lumpname
                    
                    'Store the flat info
                    allflats.Add lumpname, flat
                    
                    'Go by each required filter
                    ListFlat = False
                    For f = LBound(RequiredList) To UBound(RequiredList)
                         If (lumpname Like RequiredList(f)) Then ListFlat = True: Exit For
                    Next f
                    
                    'Go by each limited filter
                    For f = LBound(LimitedList) To UBound(LimitedList)
                         If (lumpname Like LimitedList(f)) Then ListFlat = False: Exit For
                    Next f
                    
                    'Add flat to listing if not filtered out
                    If ListFlat Then flats.Add lumpname, flat
                    
                    'Clean up references
                    Set flat = Nothing
               End If
          End If
          
          'Find next
          filename = Dir()
     Wend
End Sub




Public Function GetFlatFileData(ByRef lumpname As String) As String
     Dim filename As String
     Dim filedata As String
     Dim filebuf As Integer
     
     'Make the full file/pathname
     filename = addflatdir & lumpname & ".*"
     filename = Dir(filename)
     
     'Found anything?
     If (filename <> "") Then
          
          'Read the data
          filebuf = FreeFile
          Open addflatdir & filename For Binary Access Read Lock Write As filebuf
          filedata = Space$(LOF(filebuf))
          Get #filebuf, , filedata
          Close #filebuf
          
          'Return data
          GetFlatFileData = filedata
     End If
End Function


Private Sub LoadFlats(ByRef WadFile As clsWAD, ByVal FileSource As ENUM_IMAGESOURCE, ByVal F_START As Long, ByVal F_END As Long)
     Dim flat As clsImage
     Dim i As Long, f As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim ListFlat As Boolean
     Dim lumpname As String
     
     'Check if not closed
     If (WadFile.filename = "") Then Exit Sub
     
     'Get the filter lists from config
     RequiredList = mapconfig("flatsfilter").Items
     LimitedList = mapconfig("noflatsfilter").Items
     
     'Go for all lumps between F_START and F_END
     For i = F_START To F_END
          
          'Check if this flat has correct size
          'If (WadFile.LumpSize(i) = 64 * 64) Then
          
          'Check if not empty
          If (WadFile.LumpSize(i) > 0) Then
               
               'Get lump name
               lumpname = UCase$(Trim$(WadFile.lumpname(i)))
               
               'Continue if name is valid
               If (lumpname <> "") Then
                         
                    'Create new flat
                    Set flat = New clsImage
                    
                    'Check if this flat is square
                    If (Sqr(WadFile.LumpSize(i)) = Int(Sqr(WadFile.LumpSize(i)))) Then
                         
                         'Set the properties
                         With flat
                              .Name = lumpname
                              .width = Sqr(WadFile.LumpSize(i))
                              .height = Sqr(WadFile.LumpSize(i))
                              .ScaleX = 1
                              .ScaleY = 1
                              .FlatCandidate = True
                              .AddPatch 0, 0, .width, .height, i, FileSource, TF_UNKNOWN
                         End With
                         
                    'Check if this flat is larger than 4096
                    ElseIf (WadFile.LumpSize(i) > 4096) Then
                         
                         'Set the properties
                         With flat
                              .Name = lumpname
                              .width = 64
                              .height = 64
                              .ScaleX = 1
                              .ScaleY = 1
                              .FlatCandidate = True
                              .AddPatch 0, 0, .width, .height, i, FileSource, TF_UNKNOWN
                         End With
                         
                    'Invalid flat, but add it anyway (it will be black when invalid data)
                    Else
                         
                         'Set the properties
                         With flat
                              .Name = lumpname
                              .width = 0
                              .height = 0
                              .ScaleX = 1
                              .ScaleY = 1
                              .FlatCandidate = True
                              .AddPatch 0, 0, .width, .height, i, FileSource, TF_UNKNOWN
                         End With
                    End If
                    
                    'Remove if already added before (overwrite)
                    If (flats.Exists(lumpname)) Then flats.Remove lumpname
                    If (allflats.Exists(lumpname)) Then allflats.Remove lumpname
                    
                    'Store the flat info
                    allflats.Add lumpname, flat
                    
                    'Go by each required filter
                    ListFlat = False
                    For f = LBound(RequiredList) To UBound(RequiredList)
                         If (lumpname Like RequiredList(f)) Then ListFlat = True: Exit For
                    Next f
                    
                    'Go by each limited filter
                    For f = LBound(LimitedList) To UBound(LimitedList)
                         If (lumpname Like LimitedList(f)) Then ListFlat = False: Exit For
                    Next f
                    
                    'Add flat to listing if not filtered out
                    If ListFlat Then flats.Add lumpname, flat
                    
                    'Clean up references
                    Set flat = Nothing
               End If
          End If
          
          'End If
     Next i
End Sub

Private Sub LoadFlatRange(ByRef WadFile As clsWAD, ByVal FileSource As ENUM_IMAGESOURCE, ByVal StartIndex As Long, ByVal EndIndex As Long)
     Dim flat As clsImage
     Dim i As Long, f As Long
     Dim RequiredList As Variant
     Dim LimitedList As Variant
     Dim ListFlat As Boolean
     Dim lumpname As String
     
     'Check if not closed
     If (WadFile.filename = "") Then Exit Sub
     
     'Get the filter lists from config
     RequiredList = mapconfig("flatsfilter").Items
     LimitedList = mapconfig("noflatsfilter").Items
     
     'Go for all lumps between StartIndex and EndIndex
     For i = StartIndex To EndIndex
          
          'Check if not empty
          If (WadFile.LumpSize(i) > 0) Then
               
               'Get lump name
               lumpname = UCase$(Trim$(WadFile.lumpname(i)))
               
               'Continue if name is valid
               If (lumpname <> "") Then
                    
                    'Create new flat
                    Set flat = New clsImage
                    
                    'Set the properties
                    With flat
                         .Name = lumpname
                         .width = 0
                         .height = 0
                         .ScaleX = 1
                         .ScaleY = 1
                         .FlatCandidate = True
                         .AddPatch 0, 0, 0, 0, i, FileSource, TF_UNKNOWN
                    End With
                    
                    'Remove if already added before (overwrite)
                    If (flats.Exists(lumpname)) Then flats.Remove lumpname
                    If (allflats.Exists(lumpname)) Then allflats.Remove lumpname
                    
                    'Store the flat info
                    allflats.Add lumpname, flat
                    
                    'Go by each required filter
                    ListFlat = False
                    For f = LBound(RequiredList) To UBound(RequiredList)
                         If (lumpname Like RequiredList(f)) Then ListFlat = True: Exit For
                    Next f
                    
                    'Go by each limited filter
                    For f = LBound(LimitedList) To UBound(LimitedList)
                         If (lumpname Like LimitedList(f)) Then ListFlat = False: Exit For
                    Next f
                    
                    'Add flat to listing if not filtered out
                    If ListFlat Then flats.Add lumpname, flat
                    
                    'Clean up references
                    Set flat = Nothing
               End If
          End If
     Next i
End Sub


Public Sub PrecacheFlats()
     Dim i As Long
     Dim FlatKeys As Variant
     Dim flat As clsImage
     
     'Go for all flats
     FlatKeys = allflats.Keys
     For i = LBound(FlatKeys) To UBound(FlatKeys)
          
          'Get the flat object
          Set flat = allflats(FlatKeys(i))
          
          'Load the flat
          If (flat.IsLoaded = False) Then flat.LoadImage
          
          'Clean up
          Set flat = Nothing
     Next i
End Sub

Public Sub SortFlats()
     Set flats = SortDictionary(flats)
End Sub

Public Sub UnloadAllFlats()
     
     'Destroy dictionary
     Set flats = Nothing
     Set allflats = Nothing
End Sub

Public Sub UnloadDirect3DFlats()
     Dim i As Long
     Dim FlatKeys As Variant
     Dim flat As clsImage
     
     'Go for all flats
     FlatKeys = allflats.Keys
     For i = LBound(FlatKeys) To UBound(FlatKeys)
          
          'Get the flat object
          Set flat = allflats(FlatKeys(i))
          
          'Unload the Direct3D texture
          Set flat.D3DTexture = Nothing
          
          'Clean up
          Set flat = Nothing
     Next i
End Sub
