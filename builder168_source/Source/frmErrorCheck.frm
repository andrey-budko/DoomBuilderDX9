VERSION 5.00
Begin VB.Form frmErrorCheck 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find map errors"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmErrorCheck.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   190
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix and Recheck"
      Height          =   330
      Left            =   4125
      TabIndex        =   17
      Top             =   6105
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CheckBox chkIgnoreWarnings 
      Caption         =   "Hide warnings"
      Height          =   255
      Left            =   345
      TabIndex        =   7
      Top             =   2430
      Width           =   1860
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   300
      Left            =   60
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   385
      TabIndex        =   13
      Top             =   2895
      Visible         =   0   'False
      Width           =   5805
      Begin VB.Image imgInfo 
         Height          =   240
         Left            =   45
         Picture         =   "frmErrorCheck.frx":000C
         Top             =   15
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmErrorCheck.frx":0596
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0 items have been replaced"
         ForeColor       =   &H80000017&
         Height          =   240
         Left            =   375
         TabIndex        =   14
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   5310
      End
   End
   Begin VB.ListBox lstResults 
      Height          =   1845
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   10
      Top             =   3270
      Visible         =   0   'False
      Width           =   5820
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   4245
      TabIndex        =   9
      Top             =   2385
      Width           =   1500
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Default         =   -1  'True
      Height          =   330
      Left            =   2670
      TabIndex        =   8
      Top             =   2385
      Width           =   1500
   End
   Begin VB.Frame frmChecks 
      Caption         =   " Error Checks "
      Height          =   1560
      Left            =   180
      TabIndex        =   11
      Top             =   705
      Width           =   5565
      Begin VB.CheckBox chkThingErrors 
         Caption         =   "Thing warnings (stucked, outside)"
         Height          =   255
         Left            =   2595
         TabIndex        =   18
         Top             =   1140
         Value           =   1  'Checked
         Width           =   2820
      End
      Begin VB.CheckBox chkVertexErrors 
         Caption         =   "Vertex errors (overlappings)"
         Height          =   255
         Left            =   2595
         TabIndex        =   5
         Top             =   570
         Value           =   1  'Checked
         Width           =   2700
      End
      Begin VB.CheckBox chkZeroLengthLines 
         Caption         =   "Zero-length lines"
         Height          =   255
         Left            =   165
         TabIndex        =   3
         Top             =   1140
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkLineErrors 
         Caption         =   "Line errors (sides, overlappings)"
         Height          =   255
         Left            =   2595
         TabIndex        =   4
         Top             =   285
         Value           =   1  'Checked
         Width           =   2700
      End
      Begin VB.CheckBox chkInvalidTextures 
         Caption         =   "Invalid textures"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   855
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkMissingTextures 
         Caption         =   "Missing textures"
         Height          =   255
         Left            =   165
         TabIndex        =   1
         Top             =   570
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkPlayerStarts 
         Caption         =   "Player start Things"
         Height          =   255
         Left            =   165
         TabIndex        =   0
         Top             =   285
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CheckBox chkUnclosedSectors 
         Caption         =   "Unclosed sectors"
         Height          =   255
         Left            =   2595
         TabIndex        =   6
         Top             =   855
         Value           =   1  'Checked
         Width           =   2700
      End
   End
   Begin VB.Label lblDescription 
      Caption         =   "lblDescription"
      Height          =   810
      Left            =   105
      TabIndex        =   16
      Top             =   5385
      UseMnemonic     =   0   'False
      Width           =   5730
   End
   Begin VB.Label lblDescriptionCaption 
      AutoSize        =   -1  'True
      Caption         =   "Error Description:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   15
      Top             =   5190
      UseMnemonic     =   0   'False
      Width           =   1455
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmErrorCheck.frx":0B20
      Height          =   465
      Left            =   180
      TabIndex        =   12
      Top             =   165
      Width           =   5565
   End
End
Attribute VB_Name = "frmErrorCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private OptionsChanged As Boolean

Private Sub chkInvalidTextures_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkLineErrors_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkMissingTextures_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkPlayerStarts_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkThingErrors_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkUnclosedSectors_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkVertexErrors_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub chkZeroLengthLines_Click()
     
     'Options changed
     OptionsChanged = True
End Sub

Private Sub cmdCheck_Click()
     Dim i As Long
     
     'Busy
     Screen.MousePointer = vbHourglass
     
     'Clear the list
     lstResults.Clear
     lstResults.visible = False
     picWarning.visible = False
     cmdFix.visible = False
     
     'Clear selection
     RemoveSelection True
     
     'Apply settings
     IgnoreWarningsOption = chkIgnoreWarnings.Value
     InvalidTexturesOption = chkInvalidTextures.Value
     LineErrorsOption = chkLineErrors.Value
     MissingTexturesOption = chkMissingTextures.Value
     PlayerStartsOption = chkPlayerStarts.Value
     UnclosedSectorsOption = chkUnclosedSectors.Value
     VertexErrorsOption = chkVertexErrors.Value
     ZeroLengthLinesOption = chkZeroLengthLines.Value
     ThingErrorsOption = chkThingErrors.Value
     
     'Do the error checks
     If DoErrorChecks Then
          
          'Errors found
          
          'Fill the list with the errors
          For i = 0 To NumFoundErrors - 1
               
               'Check if this is an ERROR
               If (FoundErrors(i).critical) Then
                    
                    'Add to list
                    lstResults.AddItem "ERROR: " & FoundErrors(i).Title
                    lstResults.ItemData(lstResults.NewIndex) = i
                    
               'Check if a WARNING should be displayed
               ElseIf (chkIgnoreWarnings.Value = vbUnchecked) Then
                    
                    'Add to list
                    lstResults.AddItem "WARNING: " & FoundErrors(i).Title
                    lstResults.ItemData(lstResults.NewIndex) = i
               End If
          Next i
     End If
     
     'Anything listed?
     If (lstResults.ListCount > 0) Then
          
          'Adjust the height and set number of errors
          height = 462 * Screen.TwipsPerPixelY
          imgWarning.visible = True
          imgInfo.visible = False
          lblTotal.Caption = lstResults.ListCount & " issues have been found"
          lblDescription.Caption = "Please select an error from the list above."
          lstResults.visible = True
          picWarning.visible = True
          
          'Now select the first error
          lstResults.ListIndex = 0
     Else
          
          'No errors found
          'Adjust the height and set number of errors
          height = 242 * Screen.TwipsPerPixelY
          imgWarning.visible = False
          imgInfo.visible = True
          lblTotal.Caption = "No issues have been found"
          lstResults.visible = False
          picWarning.visible = True
     End If
     
     'Options not checked
     OptionsChanged = False
     
     'Done
     Screen.MousePointer = vbNormal
End Sub


Private Sub cmdClose_Click()
     
     'Clear selection
     RemoveSelection True
     
     'Clean up errors
     ClearFoundErrors
     
     'Leave
     Unload Me
End Sub


Private Sub cmdFix_Click()
     Dim errindex As Long
     Dim arg1 As Long
     Dim arg2 As Long
     
     'Check if anything selected
     If (lstResults.ListIndex > -1) Then
          
          'Get the error index
          errindex = lstResults.ItemData(lstResults.ListIndex)
          
          'Get arguments
          arg1 = FoundErrors(errindex).solveindex1
          arg2 = FoundErrors(errindex).solveindex2
          
          'Check how to fix this problem
          Select Case FoundErrors(errindex).solvetype
               
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_ERASEUPPERTEXTURE
                    
                    'Create undo
                    CreateUndo "fixing upper texture", , , True
                    
                    'Erase upper texture
                    sidedefs(arg1).Upper = "-"
               
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_ERASEMIDDLETEXTURE
                    
                    'Create undo
                    CreateUndo "fixing middle texture", , , True
                    
                    'Erase middle texture
                    sidedefs(arg1).Middle = "-"
               
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_ERASELOWERTEXTURE
                    
                    'Create undo
                    CreateUndo "fixing lower texture", , , True
                    
                    'Erase lower texture
                    sidedefs(arg1).Lower = "-"
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_FLIPSIDEDEFS
                    
                    'Create undo
                    CreateUndo "fixing sidedefs"
                    
                    'Flip sidedefs of the given line
                    FlipLinedefSidedefs arg1
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_FLAGTWOSIDED
                    
                    'Create undo
                    CreateUndo "fixing doublesided"
                    
                    'Ensure a twosided flag
                    If (linedefs(arg1).Flags And LDF_TWOSIDED) = 0 Then linedefs(arg1).Flags = linedefs(arg1).Flags Or LDF_TWOSIDED
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_UNFLAGTWOSIDED
                    
                    'Create undo
                    CreateUndo "fixing doublesided"
                    
                    'Remove twosided flag
                    If (linedefs(arg1).Flags And LDF_TWOSIDED) = LDF_TWOSIDED Then linedefs(arg1).Flags = linedefs(arg1).Flags And Not LDF_TWOSIDED
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_MERGELINES
                    
                    'Create undo
                    CreateUndo "merge linedefs"
                    
                    'Merge two lines
                    MergeLinedefs arg1, arg2
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_DEFAULTLOWERTEXTURE
                    
                    'Create undo
                    CreateUndo "fixing lower texture", , , True
                    
                    'Ensure valid textures are used to build with
                    CorrectDefaultTextures
                    
                    'Default lower texture
                    sidedefs(arg1).Lower = Config("defaulttexture")("lower")
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_DEFAULTMIDDLETEXTURE
                    
                    'Create undo
                    CreateUndo "fixing middle texture", , , True
                    
                    'Ensure valid textures are used to build with
                    CorrectDefaultTextures
                    
                    'Default middle texture
                    sidedefs(arg1).Middle = Config("defaulttexture")("middle")
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_DEFAULTUPPERTEXTURE
                    
                    'Create undo
                    CreateUndo "fixing upper texture", , , True
                    
                    'Ensure valid textures are used to build with
                    CorrectDefaultTextures
                    
                    'Default upper texture
                    sidedefs(arg1).Upper = Config("defaulttexture")("upper")
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_MERGEVERTICES
                    
                    'Create undo
                    CreateUndo "merge vertices"
                    
                    'Merge two vertices
                    StitchVertices arg1, arg2
                    
                    'Select the resulting vertex
                    vertexes(arg1).selected = 1
                    
                    'Remove looped linedefs
                    RemoveLoopedLinedefs
                    
                    'Find all changing lines
                    FindChangingLines True, True
                    
                    'Due to auto-stitch, linedefs could be overlapping
                    'Combine these into one now
                    MergeDoubleLinedefs
                    
                    'We dont need these anymore
                    ReDim changedlines(0)
                    numchangedlines = 0
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_DELETELINEDEF
                    
                    'Create undo
                    CreateUndo "linedef delete"
                    
                    'Remove the linedef
                    RemoveLinedef arg1, , , True
                    
               Case ENUM_ERRORSOLVEFUNCTIONS.ESF_DELETETHING
                    
                    'Create undo
                    CreateUndo "thing delete"
                    
                    'Remove the thing
                    RemoveThing arg1
                    
          End Select
          
          'Map changed
          mapchanged = True
          mapnodeschanged = True
          
          'Recheck
          cmdCheck_Click
     End If
End Sub

Private Sub Form_Load()
     
     'Restore settings
     chkIgnoreWarnings.Value = IgnoreWarningsOption
     chkInvalidTextures.Value = InvalidTexturesOption
     chkLineErrors.Value = LineErrorsOption
     chkMissingTextures.Value = MissingTexturesOption
     chkPlayerStarts.Value = PlayerStartsOption
     chkUnclosedSectors.Value = UnclosedSectorsOption
     chkVertexErrors.Value = VertexErrorsOption
     chkZeroLengthLines.Value = ZeroLengthLinesOption
     chkThingErrors.Value = ThingErrorsOption
     
     'Move to left top of parent
     left = frmMain.left + 50 * Screen.TwipsPerPixelX
     top = frmMain.top + 100 * Screen.TwipsPerPixelY
End Sub


Private Sub lstResults_Click()
     Dim errindex As Long
     Dim TargetRect As RECT
     Dim ItemIndex As Long
     
     'Check if anything selected
     If (lstResults.ListIndex > -1) Then
          
          'Get the error index
          errindex = lstResults.ItemData(lstResults.ListIndex)
          
          'Display the error description
          lblDescription.Caption = FoundErrors(errindex).Description
          
          'Display fix button if this problem can be solved by a click on a button
          cmdFix.visible = (FoundErrors(errindex).solvetype <> ESF_NONE)
          
          'Get the item index
          ItemIndex = FoundErrors(errindex).viewindex
          
          'Clear selection
          RemoveSelection False
          
          'Check type
          Select Case FoundErrors(errindex).viewtype
               
               'Vertex
               Case EM_VERTICES
                    
                    'Switch to the correct mode
                    If (mode <> EM_VERTICES) Then frmMain.itmEditMode_Click EM_VERTICES
                    
                    'Select this vertex
                    vertexes(ItemIndex).selected = 1
                    selected.Add CStr(ItemIndex), ItemIndex
                    numselected = 1
                    
                    'Make rect for vertex
                    With TargetRect
                         .left = vertexes(ItemIndex).x
                         .right = vertexes(ItemIndex).x
                         .top = -vertexes(ItemIndex).y
                         .bottom = -vertexes(ItemIndex).y
                    End With
                    
                    'Show it
                    CenterViewAt TargetRect, True, , 0.6
                    
               'Linedef
               Case EM_LINES
                    
                    'Switch to the correct mode
                    If (mode <> EM_LINES) Then frmMain.itmEditMode_Click EM_LINES
                    
                    'Select this linedef
                    linedefs(ItemIndex).selected = 1
                    selected.Add CStr(ItemIndex), ItemIndex
                    numselected = 1
                    
                    'Make rect for linedef
                    TargetRect = CalculateLinedefRect(ItemIndex)
                    
                    'Show it
                    CenterViewAt TargetRect, True, 200, 0.5, 1
                    
               'Sector
               Case EM_SECTORS
                    
                    'Switch to the correct mode
                    If (mode <> EM_SECTORS) Then frmMain.itmEditMode_Click EM_SECTORS
                    
                    'Select this sector
                    SelectSector ItemIndex
                    selected.Add CStr(ItemIndex), ItemIndex
                    numselected = 1
                    
                    'Make rect for sector
                    TargetRect = CalculateSectorRect(ItemIndex)
                    
                    'Show it
                    CenterViewAt TargetRect, True, 200, 0.5, 1
                    
               'Thing
               Case EM_THINGS
                    
                    'Switch to the correct mode
                    If (mode <> EM_THINGS) Then frmMain.itmEditMode_Click EM_THINGS
                    
                    'Select this thing
                    things(ItemIndex).selected = 1
                    selected.Add CStr(ItemIndex), ItemIndex
                    numselected = 1
                    
                    'Make rect for thing
                    With TargetRect
                         .left = things(ItemIndex).x
                         .right = things(ItemIndex).x
                         .top = -things(ItemIndex).y
                         .bottom = -things(ItemIndex).y
                    End With
                    
                    'Show it
                    CenterViewAt TargetRect, True, , 0.6
                    
          End Select
          
          'Render map
          RedrawMap False
          
          'Update status bar
          UpdateStatusBar
     End If
End Sub


