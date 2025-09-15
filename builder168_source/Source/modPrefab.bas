Attribute VB_Name = "modPrefab"
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


'Prefab insertion modes
Public Enum ENUM_PREFABINCLUDEMODE
     PIM_VERTICES
     PIM_STRUCTURE
     PIM_THINGS
End Enum

Public ClipboardFile As String
Public LastPrefab As String

Public PrefabFloorHeight As Long, PrefabCeilHeight As Long
Public PrefabAdjustHeights As Boolean

Public Sub ClipboardCleanup()
     
     'Check if the clipboard refers to my file
     If (StrComp(ClipboardGetDescriptor, ClipboardFile, vbTextCompare) = 0) Then
          
          'Clear the clipboard when file doenst exists
          If (Dir(ClipboardFile) = "") Then Clipboard.SetText "", vbCFText
     Else
          
          'Kill the file if exists
          If (Dir(ClipboardFile) <> "") Then Kill ClipboardFile
     End If
End Sub

Public Function ClipboardGetDescriptor() As String
     On Local Error Resume Next
     Dim Descriptor As String
     
     'Get the descriptor
     Descriptor = Clipboard.GetText(vbCFText)
     
     'Get filename from descriptor
     ClipboardGetDescriptor = Mid$(Descriptor, Len("DOOMBUILDER:") + 1)
End Function

Public Sub ClipboardSetDescriptor()
     Dim Descriptor As String
     
     'Add signature to filename
     Descriptor = "DOOMBUILDER:" & ClipboardFile
     
     'Set on clipboard
     Clipboard.Clear
     Clipboard.SetText Descriptor, vbCFText
End Sub

Public Sub InitializeClipboard()
     
     'Make clipboard temporary file name
     ClipboardFile = TempPath & "dbclpbrd.tmp"
     
End Sub

Public Function InsertPrefab(ByVal Filename As String, ByVal X As Long, ByVal Y As Long, ByVal IncludeMode As ENUM_PREFABINCLUDEMODE) As Long
     Dim c_vertexes() As MAPVERTEX
     Dim c_linedefs() As MAPLINEDEF
     Dim c_sidedefs() As MAPSIDEDEF
     Dim c_sectors() As MAPSECTOR
     Dim c_things() As MAPTHING
     Dim arrayvertexes() As Long
     Dim arraysidedefs() As Long
     Dim arraysectors() As Long
     
     Dim cvertexes As Long
     Dim clinedefs As Long
     Dim csidedefs As Long
     Dim csectors As Long
     Dim cthings As Long
     
     Dim ox As Long, oy As Long
     Dim FileBuffer As Integer
     Dim i As Long, n As Long
     Dim s As Long, v As Long
     Dim th As Long, sd As Long
     Dim ld As Long
     Dim fstr As String * 8
     
     'Set hourglass mousepointer
     Screen.MousePointer = vbArrowHourglass
     
     'Remove any existing selection
     RemoveSelection False
     
     'Open the data file
     FileBuffer = FreeFile
     Open Filename For Binary As #FileBuffer
     
     'Read the count numbers
     Get #FileBuffer, , ox
     Get #FileBuffer, , oy
     Get #FileBuffer, , cvertexes
     Get #FileBuffer, , csectors
     Get #FileBuffer, , csidedefs
     Get #FileBuffer, , clinedefs
     Get #FileBuffer, , cthings
     
     'Reserve memory
     If (cthings > 0) Then ReDim c_things(0 To (cthings - 1))
     If (cvertexes > 0) Then ReDim c_vertexes(0 To (cvertexes - 1))
     If (clinedefs > 0) Then ReDim c_linedefs(0 To (clinedefs - 1))
     If (csidedefs > 0) Then ReDim c_sidedefs(0 To (csidedefs - 1))
     If (csectors > 0) Then ReDim c_sectors(0 To (csectors - 1))
     
     'Allocate memory for re-reference arrays
     If (cvertexes > 0) Then ReDim arrayvertexes(0 To (cvertexes - 1))
     If (csectors > 0) Then ReDim arraysectors(0 To (csectors - 1))
     If (csidedefs > 0) Then ReDim arraysidedefs(0 To (csidedefs - 1))
     
     'Read vertices
     For v = 0 To cvertexes - 1
          Get #FileBuffer, , c_vertexes(v).X
          Get #FileBuffer, , c_vertexes(v).Y
          Get #FileBuffer, , c_vertexes(v).selected
     Next v
     
     'Read sectors
     For s = 0 To csectors - 1
          Get #FileBuffer, , c_sectors(s).hfloor
          Get #FileBuffer, , c_sectors(s).hceiling
          Get #FileBuffer, , fstr: c_sectors(s).tfloor = UnPadded(fstr)
          Get #FileBuffer, , fstr: c_sectors(s).tceiling = UnPadded(fstr)
          Get #FileBuffer, , c_sectors(s).Brightness
          Get #FileBuffer, , c_sectors(s).special
          Get #FileBuffer, , c_sectors(s).tag
          Get #FileBuffer, , c_sectors(s).selected
     Next s
     
     'Read sidedefs
     For sd = 0 To csidedefs - 1
          Get #FileBuffer, , c_sidedefs(sd).tx
          Get #FileBuffer, , c_sidedefs(sd).ty
          Get #FileBuffer, , fstr: c_sidedefs(sd).Upper = UnPadded(fstr)
          Get #FileBuffer, , fstr: c_sidedefs(sd).Lower = UnPadded(fstr)
          Get #FileBuffer, , fstr: c_sidedefs(sd).Middle = UnPadded(fstr)
          Get #FileBuffer, , c_sidedefs(sd).sector
          Get #FileBuffer, , c_sidedefs(sd).linedef
     Next sd
     
     'Read linedefs
     For ld = 0 To clinedefs - 1
          Get #FileBuffer, , c_linedefs(ld).v1
          Get #FileBuffer, , c_linedefs(ld).v2
          Get #FileBuffer, , c_linedefs(ld).Flags
          Get #FileBuffer, , c_linedefs(ld).effect
          Get #FileBuffer, , c_linedefs(ld).tag
          Get #FileBuffer, , c_linedefs(ld).arg0
          Get #FileBuffer, , c_linedefs(ld).arg1
          Get #FileBuffer, , c_linedefs(ld).arg2
          Get #FileBuffer, , c_linedefs(ld).arg3
          Get #FileBuffer, , c_linedefs(ld).arg4
          Get #FileBuffer, , c_linedefs(ld).s1
          Get #FileBuffer, , c_linedefs(ld).s2
          Get #FileBuffer, , c_linedefs(ld).selected
          Get #FileBuffer, , c_linedefs(ld).argref0
          Get #FileBuffer, , c_linedefs(ld).argref1
          Get #FileBuffer, , c_linedefs(ld).argref2
          Get #FileBuffer, , c_linedefs(ld).argref3
          Get #FileBuffer, , c_linedefs(ld).argref4
     Next ld
     
     'Read things
     For th = 0 To cthings - 1
          Get #FileBuffer, , c_things(th).tag
          Get #FileBuffer, , c_things(th).X
          Get #FileBuffer, , c_things(th).Y
          Get #FileBuffer, , c_things(th).Z
          Get #FileBuffer, , c_things(th).angle
          Get #FileBuffer, , c_things(th).thing
          Get #FileBuffer, , c_things(th).Flags
          Get #FileBuffer, , c_things(th).effect
          Get #FileBuffer, , c_things(th).arg0
          Get #FileBuffer, , c_things(th).arg1
          Get #FileBuffer, , c_things(th).arg2
          Get #FileBuffer, , c_things(th).arg3
          Get #FileBuffer, , c_things(th).arg4
          Get #FileBuffer, , c_things(th).category
          Get #FileBuffer, , c_things(th).Color
          Get #FileBuffer, , c_things(th).image
          Get #FileBuffer, , c_things(th).size
          Get #FileBuffer, , c_things(th).selected
          Get #FileBuffer, , c_things(th).argref0
          Get #FileBuffer, , c_things(th).argref1
          Get #FileBuffer, , c_things(th).argref2
          Get #FileBuffer, , c_things(th).argref3
          Get #FileBuffer, , c_things(th).argref4
     Next th
     
     'Heights information available?
     If (EOF(FileBuffer) = False) Then
          
          'Read floor/ceiling height adjustment
          Get #FileBuffer, , PrefabFloorHeight
          Get #FileBuffer, , PrefabCeilHeight
          PrefabAdjustHeights = True
     Else
          
          'Use absolute heights
          PrefabAdjustHeights = False
     End If
     
     'Close file
     Close #FileBuffer
     
     
     'Check if there are any vertices
     If (cvertexes > 0) And ((IncludeMode = PIM_VERTICES) Or (IncludeMode = PIM_STRUCTURE)) Then
          
          'Go for all new vertices
          For i = 0 To (cvertexes - 1)
               
               'Create vertex on map
               n = CreateVertex
               
               'Set the re-reference array item for this vertex
               arrayvertexes(i) = n
               
               'Set vertex properties
               With vertexes(n)
                    .selected = 1
                    .X = c_vertexes(i).X - ox + X
                    .Y = c_vertexes(i).Y - oy - Y
               End With
               
               'Add to selection
               selected.Add CStr(n), n
               numselected = selected.Count
               
               'Count this
               InsertPrefab = InsertPrefab + 1
          Next i
     End If
     
     'Check if there are any sectors
     If (csectors > 0) And (IncludeMode = PIM_STRUCTURE) Then
          
          'Go for all new sectors
          For i = 0 To (csectors - 1)
               
               'Create sector on map
               n = CreateSector
               
               'Set the re-reference array item for this sector
               arraysectors(i) = n
               
               'Set sector properties
               sectors(n) = c_sectors(i)
               sectors(n).selected = 0
               
               'Check if we should erase tag and actions
               If (Config("copytagpaste") = vbUnchecked) Then
                    With sectors(n)
                         .special = 0
                         .tag = 0
                    End With
               End If
          Next i
     End If
     
     'Check if there are any sidedefs
     If (csidedefs > 0) And (IncludeMode = PIM_STRUCTURE) Then
          
          'Go for all new sidedefs
          For i = 0 To (csidedefs - 1)
               
               'Create sidedef on map
               n = CreateSidedef
               
               'Set the re-reference array item for this sidedef
               arraysidedefs(i) = n
               
               'Set sidedef properties
               sidedefs(n) = c_sidedefs(i)
               If (sidedefs(n).sector > -1) Then sidedefs(n).sector = arraysectors(c_sidedefs(i).sector)
          Next i
     End If
     
     'Check if there are any linedefs
     If (clinedefs > 0) And (IncludeMode = PIM_STRUCTURE) Then
          
          'Go for all new linedefs
          For i = 0 To (clinedefs - 1)
               
               'Create linedef on map
               n = CreateLinedef
               
               'Set linedef properties
               linedefs(n) = c_linedefs(i)
               With linedefs(n)
                    .selected = 0
                    If (.s1 > -1) Then .s1 = arraysidedefs(.s1)
                    If (.s2 > -1) Then .s2 = arraysidedefs(.s2)
                    .v1 = arrayvertexes(.v1)
                    .v2 = arrayvertexes(.v2)
                    .selected = 1
               End With
               
               'Check if we should erase tag and actions
               If (Config("copytagpaste") = vbUnchecked) Then
                    With linedefs(n)
                         .arg0 = 0
                         .arg1 = 0
                         .arg2 = 0
                         .arg3 = 0
                         .arg4 = 0
                         .argref0 = 0
                         .argref1 = 0
                         .argref2 = 0
                         .argref3 = 0
                         .argref4 = 0
                         .effect = 0
                         .tag = 0
                    End With
               End If
               
'               'If there is not second sidedef, make one referring to parent
'               If (linedefs(n).s2 = -1) Then
'
'                    'Make sidedef
'                    linedefs(n).s2 = CreateSidedef
'
'                    'Set the sidedef properties
'                    With sidedefs(linedefs(n).s2)
'                         .Upper = "-"
'                         .Middle = "-"
'                         .Lower = "-"
'                         .sector = -1
'                         .tx = 0
'                         .ty = 0
'                    End With
'
'                    'Set doublesided and remove impassable
'                    linedefs(n).flags = linedefs(n).flags And Not LDF_IMPASSIBLE
'                    linedefs(n).flags = linedefs(n).flags Or LDF_TWOSIDED
'               End If
               
               'Reference sidedefs back to this linedef
               If (linedefs(n).s1 > -1) Then sidedefs(linedefs(n).s1).linedef = n
               If (linedefs(n).s2 > -1) Then sidedefs(linedefs(n).s2).linedef = n
          Next i
     End If
     
     'Check if there are any things
     If (cthings > 0) And (IncludeMode = PIM_THINGS) Then
          
          'Go for all new things
          For i = 0 To (cthings - 1)
               
               'Create thing on map
               n = CreateThing
               
               'Set thing properties
               things(n) = c_things(i)
               With things(n)
                    .selected = 1
                    .X = c_things(i).X - ox + X
                    .Y = c_things(i).Y - oy - Y
               End With
               
               'Check if we should erase tag and actions
               If (Config("copytagpaste") = vbUnchecked) Then
                    With things(n)
                         .arg0 = 0
                         .arg1 = 0
                         .arg2 = 0
                         .arg3 = 0
                         .arg4 = 0
                         .argref0 = 0
                         .argref1 = 0
                         .argref2 = 0
                         .argref3 = 0
                         .argref4 = 0
                         .effect = 0
                         .tag = 0
                    End With
               End If
               
               'Check if we should snap it to grid
               If ((cthings = 1) And (snapmode = True)) Then
                    
                    'Snap thing to grid
                    things(n).X = SnappedToGridX(things(n).X)
                    things(n).Y = SnappedToGridY(things(n).Y)
               End If
               
               'Add to selection
               selected.Add CStr(n), n
               numselected = selected.Count
               
               'Count this
               InsertPrefab = InsertPrefab + 1
          Next i
     End If
     
     
     'Set normal mousepointer
     Screen.MousePointer = vbDefault
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
End Function

Public Function PasteAvailable() As Boolean
     On Error Resume Next
     Dim Filename As String
     Dim DirFilename As String
     
     'Check if a descriptor is available
     If (Clipboard.GetFormat(vbCFText) = True) Then
          
          'Get filename from clipboard
          Filename = ClipboardGetDescriptor
          DirFilename = Dir(Filename)
          
          'Check if the file still exists
          If ((Trim$(Filename) = "") Or (DirFilename = "")) Then
               
               'Cant paste, file is gone
               PasteAvailable = False
          Else
               
               'Paste is possible if no errors occurred
               PasteAvailable = (Err.number = 0)
          End If
     Else
          
          'Cant paste, nothing on clipboard
          PasteAvailable = False
     End If
     
     'Clean up error
     Err.Clear
End Function

Public Sub SavePrefabSelection(ByVal Filename As String)
     Dim FileBuffer As Integer
     Dim c_vertexes() As MAPVERTEX
     Dim c_linedefs() As MAPLINEDEF
     Dim c_sidedefs() As MAPSIDEDEF
     Dim c_sectors() As MAPSECTOR
     Dim c_things() As MAPTHING
     Dim ref_vertexes() As Long         'Contains the original index for this index
     Dim ref_linedefs() As Long         'Contains the original index for this index
     Dim ref_sidedefs() As Long         'Contains the original index for this index
     Dim ref_sectors() As Long          'Contains the original index for this index
     Dim sc_sectors() As Long           'Count of sidedefs copied that refer to this sector
     Dim st_sectors() As Long           'Total of sidedefs that refer to this sector
     Dim n_vertexes As Long
     Dim n_linedefs As Long
     Dim n_sidedefs As Long
     Dim n_sectors As Long
     Dim n_things As Long
     Dim seladd As Long
     Dim selrect As RECT
     Dim ox As Long, oy As Long
     Dim s As Long, sd As Long
     Dim v As Long, ld As Long
     Dim th As Long
     Dim fstr As String * 8
     Dim OuterFloorHeight As Long, OuterCeilHeight As Long
     Dim InnerFloorHeight As Long, InnerCeilHeight As Long
     Const StartFloorheight As Long = 2147483640
     Const StartCeilheight As Long = -2147483640
     
     'Set hourglass mousepointer
     Screen.MousePointer = vbArrowHourglass
     
     'Start with very high/low floor/ceiling
     OuterFloorHeight = StartFloorheight
     OuterCeilHeight = StartCeilheight
     InnerFloorHeight = StartFloorheight
     InnerCeilHeight = StartCeilheight
     
     'Reserve enough memory
     If numthings Then ReDim c_things(0 To numthings)
     If numvertexes Then ReDim c_vertexes(0 To numvertexes)
     If numlinedefs Then ReDim c_linedefs(0 To numlinedefs)
     If numsidedefs Then ReDim c_sidedefs(0 To numsidedefs)
     If numsectors Then ReDim c_sectors(0 To numsectors)
     If numvertexes Then ReDim ref_vertexes(0 To numvertexes)
     If numlinedefs Then ReDim ref_linedefs(0 To numlinedefs)
     If numsidedefs Then ReDim ref_sidedefs(0 To numsidedefs)
     If numsectors Then ReDim ref_sectors(0 To numsectors)
     If numsectors Then ReDim sc_sectors(0 To numsectors)
     If numsectors Then ReDim st_sectors(0 To numsectors)
     
     'Erase references
     For v = 0 To numvertexes: ref_vertexes(v) = -1: Next v
     For v = 0 To numlinedefs: ref_linedefs(v) = -1: Next v
     For v = 0 To numsidedefs: ref_sidedefs(v) = -1: Next v
     For v = 0 To numsectors: ref_sectors(v) = -1: Next v
     
     'Go for all things
     For th = 0 To (numthings - 1)
          
          'Check if selected
          If things(th).selected Then
               
               'Add the thing
               c_things(n_things) = things(th)
               
               'Apply thing to rect
               If seladd Then
                    If (c_things(n_things).X < selrect.left) Then selrect.left = c_things(n_things).X
                    If (c_things(n_things).X > selrect.right) Then selrect.right = c_things(n_things).X
                    If (c_things(n_things).Y < selrect.top) Then selrect.top = c_things(n_things).Y
                    If (c_things(n_things).Y > selrect.bottom) Then selrect.bottom = c_things(n_things).Y
               Else
                    selrect.left = c_things(n_things).X
                    selrect.right = c_things(n_things).X
                    selrect.top = c_things(n_things).Y
                    selrect.bottom = c_things(n_things).Y
                    seladd = True
               End If
               
               'Count the thing
               n_things = n_things + 1
          End If
     Next th
     
     
     'Go for all vertices
     For v = 0 To (numvertexes - 1)
          
          'Check if selected
          If vertexes(v).selected Then
               
               'Add the vertex
               c_vertexes(n_vertexes) = vertexes(v)
               ref_vertexes(v) = n_vertexes
               
               'Apply vertex to rect
               If seladd Then
                    If (c_vertexes(n_vertexes).X < selrect.left) Then selrect.left = c_vertexes(n_vertexes).X
                    If (c_vertexes(n_vertexes).X > selrect.right) Then selrect.right = c_vertexes(n_vertexes).X
                    If (c_vertexes(n_vertexes).Y < selrect.top) Then selrect.top = c_vertexes(n_vertexes).Y
                    If (c_vertexes(n_vertexes).Y > selrect.bottom) Then selrect.bottom = c_vertexes(n_vertexes).Y
               Else
                    selrect.left = c_vertexes(n_vertexes).X
                    selrect.right = c_vertexes(n_vertexes).X
                    selrect.top = c_vertexes(n_vertexes).Y
                    selrect.bottom = c_vertexes(n_vertexes).Y
                    seladd = True
               End If
               
               'Count the vertex
               n_vertexes = n_vertexes + 1
          End If
     Next v
     
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if selected
          If linedefs(ld).selected Then
               
               'This linedef has 2 vertices, 1 or 2 sidedefs and 1 or 2 sectors.
               'They must be copied, counted and rereferred.
               
               'Add the linedef
               c_linedefs(n_linedefs) = linedefs(ld)
               ref_linedefs(ld) = n_linedefs
               
               
               'Check if the first vertex is not yet copied
               If (ref_vertexes(c_linedefs(n_linedefs).v1) = -1) Then
                    
                    'Add the vertex
                    v = n_vertexes
                    c_vertexes(v) = vertexes(c_linedefs(n_linedefs).v1)
                    ref_vertexes(c_linedefs(n_linedefs).v1) = v
                    
                    'Apply vertex to rect
                    If seladd Then
                         If (c_vertexes(v).X < selrect.left) Then selrect.left = c_vertexes(v).X
                         If (c_vertexes(v).X > selrect.right) Then selrect.right = c_vertexes(v).X
                         If (c_vertexes(v).Y < selrect.top) Then selrect.top = c_vertexes(v).Y
                         If (c_vertexes(v).Y > selrect.bottom) Then selrect.bottom = c_vertexes(v).Y
                    Else
                         selrect.left = c_vertexes(v).X
                         selrect.right = c_vertexes(v).X
                         selrect.top = c_vertexes(v).Y
                         selrect.bottom = c_vertexes(v).Y
                         seladd = True
                    End If
                    
                    'Count the vertex
                    n_vertexes = n_vertexes + 1
               End If
               
               'Refer the linedef to this vertex
               c_linedefs(n_linedefs).v1 = ref_vertexes(c_linedefs(n_linedefs).v1)
               
               
               'Check if the second vertex is not yet copied
               If (ref_vertexes(c_linedefs(n_linedefs).v2) = -1) Then
                    
                    'Add the vertex
                    v = n_vertexes
                    c_vertexes(v) = vertexes(c_linedefs(n_linedefs).v2)
                    ref_vertexes(c_linedefs(n_linedefs).v2) = v
                    
                    'Apply vertex to rect
                    If seladd Then
                         If (c_vertexes(v).X < selrect.left) Then selrect.left = c_vertexes(v).X
                         If (c_vertexes(v).X > selrect.right) Then selrect.right = c_vertexes(v).X
                         If (c_vertexes(v).Y < selrect.top) Then selrect.top = c_vertexes(v).Y
                         If (c_vertexes(v).Y > selrect.bottom) Then selrect.bottom = c_vertexes(v).Y
                    Else
                         selrect.left = c_vertexes(v).X
                         selrect.right = c_vertexes(v).X
                         selrect.top = c_vertexes(v).Y
                         selrect.bottom = c_vertexes(v).Y
                         seladd = True
                    End If
                    
                    'Count the vertex
                    n_vertexes = n_vertexes + 1
               End If
               
               'Refer the linedef to this vertex
               c_linedefs(n_linedefs).v2 = ref_vertexes(c_linedefs(n_linedefs).v2)
               
               
               'Check if line has a front sidedef
               If (c_linedefs(n_linedefs).s1 <> -1) Then
                    
                    'Check if the first sidedef is not yet copied
                    If (ref_sidedefs(c_linedefs(n_linedefs).s1) = -1) Then
                         
                         'Add the sidedef
                         sd = n_sidedefs
                         c_sidedefs(sd) = sidedefs(c_linedefs(n_linedefs).s1)
                         ref_sidedefs(c_linedefs(n_linedefs).s1) = sd
                         
                         'Check if the sector is not yet copied
                         If (ref_sectors(c_sidedefs(sd).sector) = -1) Then
                              
                              'Add the sector
                              s = n_sectors
                              c_sectors(s) = sectors(c_sidedefs(sd).sector)
                              ref_sectors(c_sidedefs(sd).sector) = s
                              
                              'Count the sidedefs the original sector has
                              st_sectors(s) = CountSectorSidedefs(VarPtr(sidedefs(0)), numsidedefs, c_sidedefs(sd).sector)
                              
                              'Count the sector
                              n_sectors = n_sectors + 1
                         End If
                         
                         'Refer the sidedef to this sector
                         c_sidedefs(sd).sector = ref_sectors(c_sidedefs(sd).sector)
                         
                         'Count the reference
                         sc_sectors(c_sidedefs(sd).sector) = sc_sectors(c_sidedefs(sd).sector) + 1
                         
                         'Refer the sidedef to this linedef
                         c_sidedefs(sd).linedef = n_linedefs
                         
                         'Count the sidedefs
                         n_sidedefs = n_sidedefs + 1
                    End If
                    
                    'Refer the linedef to this sidedef
                    c_linedefs(n_linedefs).s1 = ref_sidedefs(c_linedefs(n_linedefs).s1)
               End If
               
               
               'Check if line has a back sidedef
               If (c_linedefs(n_linedefs).s2 <> -1) Then
                    
                    'Check if the second sidedef is not yet copied
                    If (ref_sidedefs(c_linedefs(n_linedefs).s2) = -1) Then
                         
                         'Add the sidedef
                         sd = n_sidedefs
                         c_sidedefs(sd) = sidedefs(c_linedefs(n_linedefs).s2)
                         ref_sidedefs(c_linedefs(n_linedefs).s2) = sd
                         
                         'Check if the sector is not yet copied
                         If (ref_sectors(c_sidedefs(sd).sector) = -1) Then
                              
                              'Add the sector
                              s = n_sectors
                              c_sectors(s) = sectors(c_sidedefs(sd).sector)
                              ref_sectors(c_sidedefs(sd).sector) = s
                              
                              'Count the sidedefs the original sector has
                              st_sectors(s) = CountSectorSidedefs(VarPtr(sidedefs(0)), numsidedefs, c_sidedefs(sd).sector)
                              
                              'Count the sector
                              n_sectors = n_sectors + 1
                         End If
                         
                         'Refer the sidedef to this sector
                         c_sidedefs(sd).sector = ref_sectors(c_sidedefs(sd).sector)
                         
                         'Count the reference
                         sc_sectors(c_sidedefs(sd).sector) = sc_sectors(c_sidedefs(sd).sector) + 1
                         
                         'Refer the sidedef to this linedef
                         c_sidedefs(sd).linedef = n_linedefs
                         
                         'Count the sidedefs
                         n_sidedefs = n_sidedefs + 1
                    End If
                    
                    'Refer the linedef to this sidedef
                    c_linedefs(n_linedefs).s2 = ref_sidedefs(c_linedefs(n_linedefs).s2)
               End If
               
               
               'Count the linedef
               n_linedefs = n_linedefs + 1
          End If
     Next ld
     
     
     'Go for all copied sectors
     s = (n_sectors - 1)
     Do While (s >= 0)
          
          'Check if sidedefs are missing for this sector
          If (sc_sectors(s) < st_sectors(s)) Then
               
               'Keep the lowest floor height and highest ceiling height
               If (c_sectors(s).hfloor < OuterFloorHeight) Then OuterFloorHeight = c_sectors(s).hfloor
               If (c_sectors(s).hceiling > OuterCeilHeight) Then OuterCeilHeight = c_sectors(s).hceiling
               
               'Refer all sidedefs referring to it to sector -1
               '(to indicate to use parent sector)
               Rereference_Sectors VarPtr(c_sidedefs(0)), n_sidedefs, s, -1
               
               'Take the last sector and move it here
               c_sectors(s) = c_sectors(n_sectors - 1)
               
               'Update the count lists
               sc_sectors(s) = sc_sectors(n_sectors - 1)
               st_sectors(s) = st_sectors(n_sectors - 1)
               
               'Rerefer all sidedefs with moved sector
               Rereference_Sectors VarPtr(c_sidedefs(0)), n_sidedefs, (n_sectors - 1), s
               
               'Thats one sector less
               n_sectors = n_sectors - 1
          Else
               
               'Keep the lowest floor height and highest ceiling height
               If (c_sectors(s).hfloor < InnerFloorHeight) Then InnerFloorHeight = c_sectors(s).hfloor
               If (c_sectors(s).hceiling > InnerCeilHeight) Then InnerCeilHeight = c_sectors(s).hceiling
          End If
          
          'Next sector
          s = s - 1
     Loop
     
     
     'Crop memory to used entries
     If n_things Then ReDim Preserve c_things(0 To (n_things - 1))
     If n_vertexes Then ReDim Preserve c_vertexes(0 To (n_vertexes - 1))
     If n_linedefs Then ReDim Preserve c_linedefs(0 To (n_linedefs - 1))
     If n_sidedefs Then ReDim Preserve c_sidedefs(0 To (n_sidedefs - 1))
     If n_sectors Then ReDim Preserve c_sectors(0 To (n_sectors - 1))
     
     'Calculate offsets
     ox = selrect.left + (selrect.right - selrect.left) \ 2
     oy = selrect.top + (selrect.bottom - selrect.top) \ 2
     
     
     'Open new data file
     FileBuffer = FreeFile
     Open Filename For Binary As #FileBuffer
     
     'Output count numbers
     Put #FileBuffer, , ox                   'Offset X
     Put #FileBuffer, , oy                   'Offset Y
     Put #FileBuffer, , n_vertexes
     Put #FileBuffer, , n_sectors
     Put #FileBuffer, , n_sidedefs
     Put #FileBuffer, , n_linedefs
     Put #FileBuffer, , n_things
     
     'Output vertices
     For v = 0 To n_vertexes - 1
          Put #FileBuffer, , c_vertexes(v).X
          Put #FileBuffer, , c_vertexes(v).Y
          Put #FileBuffer, , c_vertexes(v).selected
     Next v
     
     'Output sectors
     For s = 0 To n_sectors - 1
          Put #FileBuffer, , c_sectors(s).hfloor
          Put #FileBuffer, , c_sectors(s).hceiling
          fstr = Padded(c_sectors(s).tfloor, 8): Put #FileBuffer, , fstr
          fstr = Padded(c_sectors(s).tceiling, 8): Put #FileBuffer, , fstr
          Put #FileBuffer, , c_sectors(s).Brightness
          Put #FileBuffer, , c_sectors(s).special
          Put #FileBuffer, , c_sectors(s).tag
          Put #FileBuffer, , c_sectors(s).selected
     Next s
     
     'Output sidedefs
     For sd = 0 To n_sidedefs - 1
          Put #FileBuffer, , c_sidedefs(sd).tx
          Put #FileBuffer, , c_sidedefs(sd).ty
          fstr = Padded(c_sidedefs(sd).Upper, 8): Put #FileBuffer, , fstr
          fstr = Padded(c_sidedefs(sd).Lower, 8): Put #FileBuffer, , fstr
          fstr = Padded(c_sidedefs(sd).Middle, 8): Put #FileBuffer, , fstr
          Put #FileBuffer, , c_sidedefs(sd).sector
          Put #FileBuffer, , c_sidedefs(sd).linedef
     Next sd
     
     'Output linedefs
     For ld = 0 To n_linedefs - 1
          Put #FileBuffer, , c_linedefs(ld).v1
          Put #FileBuffer, , c_linedefs(ld).v2
          Put #FileBuffer, , c_linedefs(ld).Flags
          Put #FileBuffer, , c_linedefs(ld).effect
          Put #FileBuffer, , c_linedefs(ld).tag
          Put #FileBuffer, , c_linedefs(ld).arg0
          Put #FileBuffer, , c_linedefs(ld).arg1
          Put #FileBuffer, , c_linedefs(ld).arg2
          Put #FileBuffer, , c_linedefs(ld).arg3
          Put #FileBuffer, , c_linedefs(ld).arg4
          Put #FileBuffer, , c_linedefs(ld).s1
          Put #FileBuffer, , c_linedefs(ld).s2
          Put #FileBuffer, , c_linedefs(ld).selected
          Put #FileBuffer, , c_linedefs(ld).argref0
          Put #FileBuffer, , c_linedefs(ld).argref1
          Put #FileBuffer, , c_linedefs(ld).argref2
          Put #FileBuffer, , c_linedefs(ld).argref3
          Put #FileBuffer, , c_linedefs(ld).argref4
     Next ld
     
     'Output things
     For th = 0 To n_things - 1
          Put #FileBuffer, , c_things(th).tag
          Put #FileBuffer, , c_things(th).X
          Put #FileBuffer, , c_things(th).Y
          Put #FileBuffer, , c_things(th).Z
          Put #FileBuffer, , c_things(th).angle
          Put #FileBuffer, , c_things(th).thing
          Put #FileBuffer, , c_things(th).Flags
          Put #FileBuffer, , c_things(th).effect
          Put #FileBuffer, , c_things(th).arg0
          Put #FileBuffer, , c_things(th).arg1
          Put #FileBuffer, , c_things(th).arg2
          Put #FileBuffer, , c_things(th).arg3
          Put #FileBuffer, , c_things(th).arg4
          Put #FileBuffer, , c_things(th).category
          Put #FileBuffer, , c_things(th).Color
          Put #FileBuffer, , c_things(th).image
          Put #FileBuffer, , c_things(th).size
          Put #FileBuffer, , c_things(th).selected
          Put #FileBuffer, , c_things(th).argref0
          Put #FileBuffer, , c_things(th).argref1
          Put #FileBuffer, , c_things(th).argref2
          Put #FileBuffer, , c_things(th).argref3
          Put #FileBuffer, , c_things(th).argref4
     Next th
     
     'Floor and ceiling height offsets
     If (OuterFloorHeight = StartFloorheight) Then OuterFloorHeight = InnerFloorHeight
     If (OuterCeilHeight = StartCeilheight) Then OuterCeilHeight = InnerCeilHeight
     Put #FileBuffer, , OuterFloorHeight
     Put #FileBuffer, , OuterCeilHeight
     
     'Close the file
     Close #FileBuffer
     
     'Set normal mousepointer
     Screen.MousePointer = vbDefault
End Sub
