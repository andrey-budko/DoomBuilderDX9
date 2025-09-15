Attribute VB_Name = "modUndoRedo"
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


'Max undo/redo levels
Public Const UNDOLIMIT As Long = 100


'Undo/Redo Descriptors
Public Type UNDOREDODESC
     Description As String
     mapchanged As Boolean
     mapnodeschanged As Boolean
     Allow3DUndo As Boolean
     pointer As Long
End Type


'Undo Groupings
Public Enum ENUM_UNDOGROUP
     UGRP_NONE
     UGRP_FLOORTEXTURECHANGE
     UGRP_CEILINGTEXTURECHANGE
     UGRP_UPPERTEXTURECHANGE
     UGRP_LOWERTEXTURECHANGE
     UGRP_MIDDLETEXTURECHANGE
     UGRP_TEXTUREALIGNMENT
     UGRP_BRIGHNESSCHANGE
     UGRP_CEILINGHEIGHTCHANGE
     UGRP_FLOORHEIGHTCHANGE
     UGRP_UPPERTEXTUREDELETE
     UGRP_LOWERTEXTUREDELETE
     UGRP_MIDDLETEXTUREDELETE
     UGRP_TOGGLEMIDDLETEXTURE
     UGRP_THINGHEIGHTCHANGE
     UGRP_THINGANGLECHANGE
     UGRP_THINGCHANGE
End Enum


'API Declarations
Private Declare Sub UndoRedo_Init Lib "builder.dll" ()
Private Declare Sub UndoRedo_Term Lib "builder.dll" ()
Private Declare Function UndoRedo_Put Lib "builder.dll" (ByRef things As MAPTHING, ByVal numthings As Long, ByRef linedefs As MAPLINEDEF, ByVal numlinedefs As Long, ByVal ptr_sidedefs As Long, ByVal numsidedefs As Long, ByRef vertices As MAPVERTEX, ByVal numvertices As Long, ByVal ptr_sectors As Long, ByVal numsectors As Long) As Long
Private Declare Sub UndoRedo_GetSizes Lib "builder.dll" (ByVal Index As Long, ByRef numthings As Long, ByRef numlinedefs As Long, ByRef numsidedefs As Long, ByRef numvertices As Long, ByRef numsectors As Long)
Private Declare Sub UndoRedo_GetImages Lib "builder.dll" (ByVal Index As Long, ByRef things As MAPTHING, ByRef linedefs As MAPLINEDEF, ByVal ptr_sidedefs As Long, ByRef vertices As MAPVERTEX, ByVal ptr_sectors As Long)
Private Declare Sub UndoRedo_Delete Lib "builder.dll" (ByVal Index As Long)

'Grouping
Private LastUndoGroup As ENUM_UNDOGROUP
Private LastUndoGroupIndex As Long

'Buffers
Private UndoBuffer(0 To (UNDOLIMIT - 1)) As UNDOREDODESC
Private RedoBuffer(0 To (UNDOLIMIT - 1)) As UNDOREDODESC
Public Undos As Long
Public Redos As Long

Public Function AllowThis3DRedo() As Boolean
     
     'Return Allow3DUndo
     If (Redos > 0) Then AllowThis3DRedo = RedoBuffer(0).Allow3DUndo
End Function

Public Function AllowThis3DUndo() As Boolean
     
     'Return Allow3DUndo
     If (Undos > 0) Then AllowThis3DUndo = UndoBuffer(0).Allow3DUndo
End Function

Public Sub CreateUndo(ByVal Description As String, Optional ByVal Group As ENUM_UNDOGROUP = UGRP_NONE, Optional ByVal GroupIndex As Long, Optional Allow3DUndo As Boolean)
     Dim OldMousePointer As Long
     
     'Change mousepointer
     OldMousePointer = Screen.MousePointer
     If (Screen.MousePointer = vbDefault) Then Screen.MousePointer = vbArrowHourglass
     
     'Do not make an undo when the group and
     'group indexmatch previous grouping
     If (Group = UGRP_NONE) Or _
        (Group <> LastUndoGroup) Or _
        (GroupIndex <> LastUndoGroupIndex) Then
          
          'Make the undo
          PushUndo Description, Allow3DUndo
          
          'Keep the grouping info
          LastUndoGroup = Group
          LastUndoGroupIndex = GroupIndex
          
          'Clear the redo's
          ResetRedos
          
          'Update the menu
          frmMain.itmEditUndo.Enabled = True
          frmMain.itmEditUndo.Caption = "&Undo " & Description
          frmMain.itmEditRedo.Enabled = False
          frmMain.itmEditRedo.Caption = "&Redo"
          frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
          frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
          frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
          frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
          frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
          frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
     End If
     
     'Change mousepointer back
     Screen.MousePointer = OldMousePointer
End Sub

Public Sub InitializeUndoRedo()
     
     'Initialize UndoRedo structures in DLL
     UndoRedo_Init
     
     'Set the defaults
     Undos = 0
     Redos = 0
     
     'Update the menu
     frmMain.itmEditUndo.Enabled = False
     frmMain.itmEditUndo.Caption = "&Undo"
     frmMain.itmEditRedo.Enabled = False
     frmMain.itmEditRedo.Caption = "&Redo"
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Sub PerformRedo(Optional ByVal KeepSelected As Boolean)
     Dim i As Long
     
     'Make an undo option first
     PushUndo RedoBuffer(0).Description, RedoBuffer(0).Allow3DUndo
     
     'Get memory sizes
     UndoRedo_GetSizes RedoBuffer(0).pointer, numthings, numlinedefs, numsidedefs, numvertexes, numsectors
     
     'Reserve memory
     'ReDim things(0 To numthings + DECLARE_THINGS)
     'ReDim linedefs(0 To numlinedefs + DECLARE_LINEDEFS)
     'ReDim sidedefs(0 To numsidedefs + DECLARE_SIDEDEFS)
     'ReDim vertexes(0 To numvertexes + DECLARE_VERTICES)
     'ReDim sectors(0 To numsectors + DECLARE_SECTORS)
     
     'Prepare sidedef strings
     For i = 0 To numsidedefs '+ DECLARE_SIDEDEFS
          With sidedefs(i)
               .Lower = Space$(8)
               .Middle = Space$(8)
               .Upper = Space$(8)
          End With
     Next i
     
     'Prepare sector strings
     For i = 0 To numsectors '+ DECLARE_SECTORS
          With sectors(i)
               .tfloor = Space$(8)
               .tceiling = Space$(8)
          End With
     Next i
     
     'Get map structure images
     UndoRedo_GetImages RedoBuffer(0).pointer, things(0), linedefs(0), VarPtr(sidedefs(0)), vertexes(0), VarPtr(sectors(0))
     UndoRedo_Delete RedoBuffer(0).pointer
     
     'None of this stuff should be selected
     If Not KeepSelected Then ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     
     'Remove spaces in sidedef strings
     For i = 0 To numsidedefs '+ DECLARE_SIDEDEFS
          With sidedefs(i)
               .Lower = Trim$(.Lower)
               .Middle = Trim$(.Middle)
               .Upper = Trim$(.Upper)
          End With
     Next i
     
     'Remove spaces in sector strings
     For i = 0 To numsectors '+ DECLARE_SECTORS
          With sectors(i)
               .tfloor = Trim$(.tfloor)
               .tceiling = Trim$(.tceiling)
          End With
     Next i
     
     'Restore settings
     mapchanged = True
     mapnodeschanged = True
     
     'Move all redo descriptors one down
     For i = 1 To (Redos - 1)
          RedoBuffer(i - 1) = RedoBuffer(i)
     Next i
     
     'Decrease number of redo's
     Redos = Redos - 1
     
     'Reset the grouping info
     LastUndoGroup = UGRP_NONE
     
     'Update the menu
     If (Undos > 0) Then
          frmMain.itmEditUndo.Enabled = True
          frmMain.itmEditUndo.Caption = "&Undo " & UndoBuffer(0).Description
     Else
          frmMain.itmEditUndo.Enabled = False
          frmMain.itmEditUndo.Caption = "&Undo"
     End If
     If (Redos > 0) Then
          frmMain.itmEditRedo.Enabled = True
          frmMain.itmEditRedo.Caption = "&Redo " & RedoBuffer(0).Description
     Else
          frmMain.itmEditRedo.Enabled = False
          frmMain.itmEditRedo.Caption = "&Redo"
     End If
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Sub PerformUndo(Optional ByVal KeepSelected As Boolean)
     Dim i As Long
     
     'Make a redo option first
     PushRedo UndoBuffer(0).Description, UndoBuffer(0).Allow3DUndo
     
     'Get memory sizes
     UndoRedo_GetSizes UndoBuffer(0).pointer, numthings, numlinedefs, numsidedefs, numvertexes, numsectors
     
     'Reserve memory
     'ReDim things(0 To numthings + DECLARE_THINGS)
     'ReDim linedefs(0 To numlinedefs + DECLARE_LINEDEFS)
     'ReDim sidedefs(0 To numsidedefs + DECLARE_SIDEDEFS)
     'ReDim vertexes(0 To numvertexes + DECLARE_VERTICES)
     'ReDim sectors(0 To numsectors + DECLARE_SECTORS)
     
     'Prepare sidedef strings
     For i = 0 To numsidedefs '+ DECLARE_SIDEDEFS
          With sidedefs(i)
               .Lower = Space$(8)
               .Middle = Space$(8)
               .Upper = Space$(8)
          End With
     Next i
     
     'Prepare sector strings
     For i = 0 To numsectors '+ DECLARE_SECTORS
          With sectors(i)
               .tfloor = Space$(8)
               .tceiling = Space$(8)
          End With
     Next i
     
     'Get map structure images
     UndoRedo_GetImages UndoBuffer(0).pointer, things(0), linedefs(0), VarPtr(sidedefs(0)), vertexes(0), VarPtr(sectors(0))
     UndoRedo_Delete UndoBuffer(0).pointer
     
     'None of this stuff should be selected
     If Not KeepSelected Then ResetSelections things(0), numthings, linedefs(0), numlinedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors
     
     'Remove spaces in sidedef strings
     For i = 0 To numsidedefs '+ DECLARE_SIDEDEFS
          With sidedefs(i)
               .Lower = Trim$(.Lower)
               .Middle = Trim$(.Middle)
               .Upper = Trim$(.Upper)
          End With
     Next i
     
     'Remove spaces in sector strings
     For i = 0 To numsectors '+ DECLARE_SECTORS
          With sectors(i)
               .tfloor = Trim$(.tfloor)
               .tceiling = Trim$(.tceiling)
          End With
     Next i
     
     'Restore settings
     mapchanged = True
     mapnodeschanged = True
     
     'Move all undo descriptors one down
     For i = 1 To (Undos - 1)
          UndoBuffer(i - 1) = UndoBuffer(i)
     Next i
     
     'Decrease number of undo's
     Undos = Undos - 1
     
     'Reset the grouping info
     LastUndoGroup = UGRP_NONE
     
     'Update the menu
     If (Undos > 0) Then
          frmMain.itmEditUndo.Enabled = True
          frmMain.itmEditUndo.Caption = "&Undo " & UndoBuffer(0).Description
     Else
          frmMain.itmEditUndo.Enabled = False
          frmMain.itmEditUndo.Caption = "&Undo"
     End If
     If (Redos > 0) Then
          frmMain.itmEditRedo.Enabled = True
          frmMain.itmEditRedo.Caption = "&Redo " & RedoBuffer(0).Description
     Else
          frmMain.itmEditRedo.Enabled = False
          frmMain.itmEditRedo.Caption = "&Redo"
     End If
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Private Sub PushRedo(ByVal Description As String, ByVal Allow3DUndo As Boolean)
     Dim i As Long
     
     'Check if we should remove a level
     If (Redos >= Config("maxundos")) Then
          
          'Remove the last redo level
          UndoRedo_Delete RedoBuffer(Redos - 1).pointer
          
          'Thats one less
          Redos = Redos - 1
     End If
     
     'Move all redo descriptors one up
     For i = Redos To 1 Step -1
          RedoBuffer(i) = RedoBuffer(i - 1)
     Next i
     
     'Set the redo descriptor details
     With RedoBuffer(0)
          .Description = Description
          .mapchanged = mapchanged
          .mapnodeschanged = mapnodeschanged
          .Allow3DUndo = Allow3DUndo
     End With
     
     'Save current map structure and get a pointer
     RedoBuffer(0).pointer = UndoRedo_Put(things(0), numthings, linedefs(0), numlinedefs, VarPtr(sidedefs(0)), numsidedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors)
     
     'One redo has been added
     Redos = Redos + 1
End Sub

Private Sub PushUndo(ByVal Description As String, ByVal Allow3DUndo As Boolean)
     Dim i As Long
     
     'Check if we should remove a level
     If (Undos >= Config("maxundos")) Then
          
          'Remove the last undo level
          UndoRedo_Delete UndoBuffer(Undos - 1).pointer
          
          'Thats one less
          Undos = Undos - 1
     End If
     
     'Move all undo descriptors one up
     For i = Undos To 1 Step -1
          UndoBuffer(i) = UndoBuffer(i - 1)
     Next i
     
     'Set the undo descriptor details
     With UndoBuffer(0)
          .Description = Description
          .mapchanged = mapchanged
          .mapnodeschanged = mapnodeschanged
          .Allow3DUndo = Allow3DUndo
     End With
     
     'Save current map structure and get a pointer
     UndoBuffer(0).pointer = UndoRedo_Put(things(0), numthings, linedefs(0), numlinedefs, VarPtr(sidedefs(0)), numsidedefs, vertexes(0), numvertexes, VarPtr(sectors(0)), numsectors)
     
     'One undo has been added
     Undos = Undos + 1
End Sub

Public Function RedoDescription() As String
     
     'Return next redo description
     If (Redos > 0) Then RedoDescription = RedoBuffer(0).Description
End Function

Public Sub RenameUndo(ByVal NewDescription As String)
     
     'Rename the next undo
     UndoBuffer(0).Description = NewDescription
     
     'Update the menu
     frmMain.itmEditUndo.Enabled = True
     frmMain.itmEditUndo.Caption = "&Undo " & NewDescription
     frmMain.itmEditRedo.Enabled = False
     frmMain.itmEditRedo.Caption = "&Redo"
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Sub ResetRedos()
     Dim i As Long
     
     'Go for all current redos
     For i = 0 To (Redos - 1)
          
          'Remove the redo level
          UndoRedo_Delete RedoBuffer(i).pointer
     Next i
     
     'Erase descriptors
     Erase RedoBuffer
     
     'No more redo's
     Redos = 0
     
     'Update the menu
     frmMain.itmEditRedo.Enabled = False
     frmMain.itmEditRedo.Caption = "&Redo"
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Sub ResetUndos()
     Dim i As Long
     
     'Go for all current undos
     For i = 0 To (Undos - 1)
          
          'Remove the undos level
          UndoRedo_Delete UndoBuffer(i).pointer
     Next i
     
     'Erase descriptors
     Erase UndoBuffer
     
     'No more undos's
     Undos = 0
     
     'Reset the grouping info
     LastUndoGroup = UGRP_NONE
     
     'Update the menu
     frmMain.itmEditUndo.Enabled = False
     frmMain.itmEditUndo.Caption = "&Undo"
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
End Sub

Public Sub TerminateUndoRedo()
     
     'Deallocate memory used by DLL
     UndoRedo_Term
     
     'Erase descriptors
     Erase UndoBuffer, RedoBuffer
     
     'Set the defaults
     Undos = 0
     Redos = 0
     
     'Reset the grouping info
     LastUndoGroup = UGRP_NONE
     
     'Update the menu
     frmMain.itmEditUndo.Enabled = False
     frmMain.itmEditUndo.Caption = "&Undo"
     frmMain.itmEditRedo.Enabled = False
     frmMain.itmEditRedo.Caption = "&Redo"
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Function UndoDescription() As String
     
     'Return next undo description
     If (Undos > 0) Then UndoDescription = UndoBuffer(0).Description
End Function

Public Sub WithdrawRedo()
     Dim i As Long
     
     'Delete this redo from memory
     UndoRedo_Delete RedoBuffer(0).pointer
     
     'Move all redo descriptors one down
     For i = 1 To (Redos - 1)
          RedoBuffer(i - 1) = RedoBuffer(i)
     Next i
     
     'Decrease number of redo's
     Redos = Redos - 1
     
     'Update the menu
     If (Redos > 0) Then
          frmMain.itmEditRedo.Enabled = True
          frmMain.itmEditRedo.Caption = "&Redo " & RedoBuffer(0).Description
     Else
          frmMain.itmEditRedo.Enabled = False
          frmMain.itmEditRedo.Caption = "&Redo"
     End If
     frmMain.tlbToolbar.Buttons("EditRedo").Enabled = frmMain.itmEditRedo.Enabled
     frmMain.tlbToolbar.Buttons("EditRedo").ToolTipText = frmMain.itmEditRedo.Caption
     frmMain.itmEditRedo.Caption = MenuNameForShortcut(frmMain.itmEditRedo.Caption, "editredo")
End Sub

Public Sub WithdrawUndo()
     Dim i As Long
     
     'Delete this undo from memory
     UndoRedo_Delete UndoBuffer(0).pointer
     
     'Move all undo descriptors one down
     For i = 1 To (Undos - 1)
          UndoBuffer(i - 1) = UndoBuffer(i)
     Next i
     
     'Decrease number of undo's
     Undos = Undos - 1
     
     'Reset the grouping info
     LastUndoGroup = UGRP_NONE
     
     'Update the menu
     If (Undos > 0) Then
          frmMain.itmEditUndo.Enabled = True
          frmMain.itmEditUndo.Caption = "&Undo " & UndoBuffer(0).Description
     Else
          frmMain.itmEditUndo.Enabled = False
          frmMain.itmEditUndo.Caption = "&Undo"
     End If
     frmMain.tlbToolbar.Buttons("EditUndo").Enabled = frmMain.itmEditUndo.Enabled
     frmMain.tlbToolbar.Buttons("EditUndo").ToolTipText = frmMain.itmEditUndo.Caption
     frmMain.itmEditUndo.Caption = MenuNameForShortcut(frmMain.itmEditUndo.Caption, "editundo")
End Sub
