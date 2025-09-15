VERSION 5.00
Begin VB.Form frmMapSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Map"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
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
   Icon            =   "frmMapSelect.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   417
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDetect 
      BorderStyle     =   0  'None
      Height          =   4785
      Left            =   45
      TabIndex        =   14
      Top             =   750
      Width           =   5865
      Begin VB.Label lblDetect 
         Alignment       =   2  'Center
         Caption         =   "Attempting to detect appropriate game configuration, please wait..."
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   1290
         Width           =   5760
      End
   End
   Begin VB.ListBox lstMap 
      Columns         =   6
      Height          =   1410
      IntegralHeight  =   0   'False
      Left            =   105
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   5760
   End
   Begin VB.TextBox txtWAD 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   3630
      Width           =   4275
   End
   Begin VB.CommandButton cmdBrowseFlatDir 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   4440
      TabIndex        =   9
      Top             =   5055
      Width           =   1425
   End
   Begin VB.TextBox txtFlatDir 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   5070
      Width           =   4275
   End
   Begin VB.CommandButton cmdBrowseTexDir 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   4440
      TabIndex        =   7
      Top             =   4335
      Width           =   1425
   End
   Begin VB.TextBox txtTexDir 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   4350
      Width           =   4275
   End
   Begin VB.Timer tmrFindContents 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5430
      Top             =   5700
   End
   Begin VB.CommandButton cmdBrowseWAD 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   4440
      TabIndex        =   5
      Top             =   3615
      Width           =   1425
   End
   Begin VB.ComboBox cmbGameConfig 
      Height          =   330
      IntegralHeight  =   0   'False
      ItemData        =   "frmMapSelect.frx":000C
      Left            =   1260
      List            =   "frmMapSelect.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   810
      Width           =   3090
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   510
      Left            =   60
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   388
      TabIndex        =   10
      Top             =   60
      Width           =   5850
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: These settings control the way your map is loaded and saved. Be sure to configure these correctly."
         ForeColor       =   &H80000017&
         Height          =   450
         Left            =   375
         TabIndex        =   11
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   5340
      End
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmMapSelect.frx":0010
         Top             =   90
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3015
      TabIndex        =   1
      Top             =   5760
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   1230
      TabIndex        =   0
      Top             =   5760
      Width           =   1665
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Additional Flats from directory:"
      Height          =   210
      Left            =   105
      TabIndex        =   18
      Top             =   4830
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Additional Textures from directory:"
      Height          =   210
      Left            =   105
      TabIndex        =   17
      Top             =   4110
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Additional Textures and Flats from WAD file:"
      Height          =   210
      Left            =   105
      TabIndex        =   16
      Top             =   3390
      Width           =   3195
   End
   Begin VB.Label Label1 
      Caption         =   "With the above selected game, the maps shown below were found in the chosen WAD file. Please select the map to load for editing."
      Height          =   435
      Left            =   105
      TabIndex        =   13
      Top             =   1305
      Width           =   5820
   End
   Begin VB.Label lblGameConfig 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Configuration:"
      Height          =   210
      Left            =   105
      TabIndex        =   12
      Top             =   870
      UseMnemonic     =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "frmMapSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private LastSelected As String

Private Sub cmbGameConfig_Change()
     On Local Error Resume Next
     Dim FileWAD As New clsWAD
     
     'If no filename is set on the tag yet, leave
     If (Trim$(tag) = "") Then Exit Sub
     
     'If no items selected yet, leave
     If (cmbGameConfig.ListIndex < 0) Then Exit Sub
     
     'Temporarely load this configuration
     LoadMapConfiguration cmbGameConfig.Text
     
     'Open the file
     FileWAD.OpenFile tag, True
     
     'Change the extra wad file to the default for this config
     txtWAD.Text = mapconfig("texturesfile")
     
     'Refill the list
     UpdateMapsList FileWAD
     
     'Close the file
     FileWAD.CloseFile
End Sub

Private Sub cmbGameConfig_Click()
     cmbGameConfig_Change
End Sub

Private Sub cmbGameConfig_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbGameConfig_Change
End Sub

Private Sub cmdBrowseFlatDir_Click()
     Dim NewFolder As String
     
     'Browse for new file
     NewFolder = SelectFolder(Me.hWnd, "Select additional flats directory")
     
     'Check if not cancelled
     If (Trim$(NewFolder) <> "") Then
          
          'Set the new file in textbox
          txtFlatDir.Text = NewFolder
          txtFlatDir.SelStart = Len(txtFlatDir.Text)
          txtFlatDir.SetFocus
     End If
End Sub

Private Sub cmdBrowseTexDir_Click()
     Dim NewFolder As String
     
     'Browse for new file
     NewFolder = SelectFolder(Me.hWnd, "Select additional textures directory")
     
     'Check if not cancelled
     If (Trim$(NewFolder) <> "") Then
          
          'Set the new file in textbox
          txtTexDir.Text = NewFolder
          txtTexDir.SelStart = Len(txtTexDir.Text)
          txtTexDir.SetFocus
     End If
End Sub

Private Sub cmdBrowseWAD_Click()
     Dim NewFile As String
     
     'Browse for new file
     NewFile = OpenFile(Me.hWnd, "Select Extra WAD File", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|All Files|*.*", "", cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'Set the new file in textbox
          txtWAD.Text = NewFile
          txtWAD.SelStart = Len(txtWAD.Text)
          txtWAD.SetFocus
     End If
End Sub

Private Sub cmdCancel_Click()
     On Local Error Resume Next
     
     'Check if a map was loaded
     If (mapfile <> "") Then
          
          'Reload original configuration
          LoadMapConfiguration mapgame
     End If
     
     'Leave now
     Unload Me
     Set frmMapSelect = Nothing
End Sub

Private Sub cmdOK_Click()
     Dim FileWAD As New clsWAD
     
     'Hide dialog
     Hide
     
     'Unload old map
     If (MapUnload) Then
          
          'Change map configuration
          mapgame = cmbGameConfig.Text
          
          'Change add wad file
          addwadfile = txtWAD.Text
          addtexdir = txtTexDir.Text
          addflatdir = txtFlatDir.Text
          
          'Make full directories
          If (Len(addtexdir) > 0) Then If (right$(addtexdir, 1) <> "\") Then addtexdir = addtexdir & "\"
          If (Len(addflatdir) > 0) Then If (right$(addflatdir, 1) <> "\") Then addflatdir = addflatdir & "\"
          
          'Open the file
          FileWAD.OpenFile tag, True
          
          'Load the map from file
          MapLoad tag, FileWAD, lstMap.List(lstMap.ListIndex), True
     End If
     
     'Unload dialog
     Unload Me
     Set frmMapSelect = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     On Local Error Resume Next
     Dim i As Long
     
     'Go for al configs
     For i = 0 To (AllGameConfigs.Count - 1)
          
          'Add to list
          cmbGameConfig.AddItem AllGameConfigs.Keys(i)
     Next i
End Sub

Private Sub Form_Resize()
     
     'Refresh dialog
     Refresh
     
     'Dialog shows up, Find WAD contents
     tmrFindContents.Enabled = True
End Sub

Private Sub lstMap_Click()
     Dim ConfigStruct As Dictionary
     
     'Anything selected?
     If (lstMap.ListIndex >= 0) Then
          
          'Enabled OK button
          cmdOK.Enabled = True
          
          'Keep last selected map name
          LastSelected = lstMap.List(lstMap.ListIndex)
          
          
          'Get settings for this map from DBS file
          Set ConfigStruct = GetWadMapSettings(tag, lstMap.List(lstMap.ListIndex))
          
          'Check if settings available
          If Not (ConfigStruct Is Nothing) Then
               
               'Apply settings
               If (ConfigStruct.Exists("addwad")) Then txtWAD.Text = ConfigStruct("addwad")
               If (ConfigStruct.Exists("addtexdir")) Then txtTexDir.Text = ConfigStruct("addtexdir")
               If (ConfigStruct.Exists("addflatdir")) Then txtFlatDir.Text = ConfigStruct("addflatdir")
               
               'Clean up
               Set ConfigStruct = Nothing
          End If
     Else
          
          'Not able to open anything when nothing selected
          cmdOK.Enabled = False
     End If
End Sub

Private Sub lstMap_DblClick()
     If (cmdOK.Enabled) Then cmdOK_Click
End Sub

Private Sub lstMap_KeyUp(KeyCode As Integer, Shift As Integer)
     lstMap_Click
End Sub

Private Sub lstMap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     lstMap_Click
End Sub

Private Sub tmrFindContents_Timer()
     On Local Error GoTo BrowseMapError
     Dim ConfigStruct As Dictionary
     Dim FileWAD As New clsWAD
     Dim GameDetermined As Boolean
     Dim i As Long
     
     'Disable timer
     tmrFindContents.Enabled = False
     
     'Select nothing
     cmbGameConfig.ListIndex = -1
     
     'Check if file exists
     If (Dir(tag) <> "") Then
          
          'Get settings for this map from DBS file
          Set ConfigStruct = GetWadSettings(tag)
          
          'Check if settings available
          If Not (ConfigStruct Is Nothing) Then
               
               'Check if game configuration is given
               If (ConfigStruct.Exists("config") = True) Then
                    
                    'Go for all game configurations
                    For i = 0 To (cmbGameConfig.ListCount - 1)
                         
                         'Check if the given configuration exists
                         If (StrComp(cmbGameConfig.List(i), ConfigStruct("config"), vbTextCompare) = 0) Then
                              
                              'Then select it and be done with it
                              cmbGameConfig.ListIndex = i
                              GameDetermined = True
                              
                              'Leave now
                              Exit For
                         End If
                    Next i
               End If
          End If
          
          'Check if no game configuration chosen yet
          If (GameDetermined = False) Then
               
               'Go for all game configurations
               For i = 0 To (cmbGameConfig.ListCount - 1)
                    
                    'Check if an IWAD is configured for this config
                    If (Trim$(GetCurrentIWADFile(cmbGameConfig.List(i))) <> "") Then
                         
                         'Temporarely load this configuration
                         LoadMapConfiguration cmbGameConfig.List(i)
                         
                         'Open the file
                         FileWAD.OpenFile tag, True
                         
                         'Validate the game
                         If ValidateGameWAD(FileWAD) Then
                              
                              'Fill the list
                              UpdateMapsList FileWAD
                              
                              'Check if theres a positive result
                              If (lstMap.ListCount > 0) Then
                                   
                                   'Close the file
                                   FileWAD.CloseFile
                                   
                                   'Select this game
                                   cmbGameConfig.ListIndex = i
                                   
                                   'Leave now
                                   Exit For
                              End If
                         End If
                         
                         'Close the file
                         FileWAD.CloseFile
                    End If
               Next i
          End If
     Else
          
          'Show error
          lblDetect.visible = False
          MsgBox "Cannot open that WAD file, the file does not exist." & vbLf & LCase(tag), vbCritical
          
          'Cancel and close
          cmdCancel_Click
     End If
     
     'Remove loading panel
     fraDetect.visible = False
     cmdCancel.Enabled = True
     
     'Leave
     Exit Sub
     
     
BrowseMapError:
     
     'Show error
     lblDetect.visible = False
     MsgBox "Could not browse the WAD file contents in this file." & vbLf & "Error " & Err.number & ": " & Err.Description, vbCritical
     
     'Cancel and close
     cmdCancel_Click
End Sub

Private Sub UpdateMapsList(ByRef FileWAD As clsWAD)
     Dim RequiredLumps As Long
     Dim VerifiedLumps As Long
     Dim MapLumps As Variant
     Dim l As Long
     Dim nl As Long
     Dim mlt As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Clear list
     lstMap.Clear
     
     'Count the number of required map lumps
     MapLumps = mapconfig("maplumpnames").Items
     For l = LBound(MapLumps) To UBound(MapLumps)
          If (CLng(MapLumps(l)) And ML_REQUIRED) = ML_REQUIRED Then RequiredLumps = RequiredLumps + 1
     Next l
     
     'Go for all lumps
     For l = 1 To FileWAD.LumpCount
          
          'Check if not a lumpname which can be part of a map and not at EOF
          If ((GetMapLumpType(FileWAD.LumpnamePadded(l), False) = ML_UNKNOWN) And (l < FileWAD.LumpCount - 4)) Then
               
               'Check for required lumps
               VerifiedLumps = 0
               nl = 1
               mlt = GetMapLumpType(FileWAD.LumpnamePadded(l + nl), False)
               Do Until (mlt = ML_UNKNOWN)
                    
                    'This lump is verified when its a lump required by the editor
                    If (mlt And ML_REQUIRED) = ML_REQUIRED Then VerifiedLumps = VerifiedLumps + 1
                    
                    'Next lump
                    nl = nl + 1
                    If (l + nl > FileWAD.LumpCount) Then Exit Do
                    mlt = GetMapLumpType(FileWAD.LumpnamePadded(l + nl), False)
               Loop
               
               'Add the name to the list if this is a map that can be loaded
               If (VerifiedLumps >= RequiredLumps) Then lstMap.AddItem FileWAD.LumpName(l)
          End If
     Next l
     
     'Check if a selection was made
     If (LastSelected <> "") Then
          
          'Find select index
          For l = 0 To (lstMap.ListCount - 1)
               
               'Select this item if previously selected
               If (lstMap.List(l) = LastSelected) Then
                    lstMap.ListIndex = l
                    Exit For
               End If
          Next l
     
     'Else check if only 1 map in the wad
     ElseIf (lstMap.ListCount = 1) Then
          
          'Select this
          lstMap.ListIndex = 0
     End If
     
     'Validate lstMap
     lstMap_Click
     
     'Make the sure the OK button is enabled/disabled correctly
     cmdOK.Enabled = (lstMap.ListIndex > -1)
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
End Sub

Private Function ValidateGameWAD(ByRef FileWAD As clsWAD) As Boolean
     Dim Lumpnames As Variant
     Dim OneFound As Long
     Dim i As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Go for all game detect lumpnames
     Lumpnames = mapconfig("gamedetect").Keys
     For i = LBound(Lumpnames) To UBound(Lumpnames)
          
          'Check if this lump may be found
          If (mapconfig("gamedetect")(Lumpnames(i)) = 1) Then
               
               'Check if we can find it
               If (FindLumpIndex(FileWAD, 1, Lumpnames(i)) > 0) Then OneFound = True
               
          'Check if this lump may not be found
          ElseIf (mapconfig("gamedetect")(Lumpnames(i)) = 2) Then
               
               'Check if we can find it
               If (FindLumpIndex(FileWAD, 1, Lumpnames(i)) > 0) Then
                    
                    'This may not be, leave now
                    OneFound = False
                    Exit For
               End If
               
          'Check if this lump must be found
          ElseIf (mapconfig("gamedetect")(Lumpnames(i)) = 3) Then
               
               'Check if we can find it
               If (FindLumpIndex(FileWAD, 1, Lumpnames(i)) = 0) Then
                    
                    'This cant be found, leave now
                    OneFound = False
                    Exit For
               End If
          End If
     Next i
     
     'Reset mousepointer
     Screen.MousePointer = vbDefault
     
     'Return result
     ValidateGameWAD = OneFound
End Function

Private Sub txtFlatDir_GotFocus()
     SelectAllText txtFlatDir
End Sub


Private Sub txtTexDir_GotFocus()
     SelectAllText txtTexDir
End Sub


Private Sub txtWAD_GotFocus()
     SelectAllText txtWAD
End Sub


