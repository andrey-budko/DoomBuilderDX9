VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
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
   Icon            =   "frmFind.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   374
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstResults 
      Height          =   1035
      IntegralHeight  =   0   'False
      Left            =   45
      TabIndex        =   14
      Top             =   2715
      Visible         =   0   'False
      Width           =   5430
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   300
      Left            =   45
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   360
      TabIndex        =   12
      Top             =   2370
      Visible         =   0   'False
      Width           =   5430
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmFind.frx":000C
         Top             =   15
         Width           =   240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "0 items have been replaced"
         ForeColor       =   &H80000017&
         Height          =   240
         Left            =   375
         TabIndex        =   13
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   4980
      End
   End
   Begin VB.CheckBox chkReplaceOnly 
      Caption         =   "Replace only (no select)"
      Height          =   255
      Left            =   1410
      TabIndex        =   6
      Top             =   1905
      Width           =   2085
   End
   Begin VB.CommandButton cmdBrowseReplace 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   5
      Top             =   1485
      Width           =   375
   End
   Begin VB.TextBox txtReplace 
      Height          =   330
      Left            =   1410
      TabIndex        =   4
      Top             =   1485
      Width           =   1545
   End
   Begin VB.CommandButton cmdBrowseFind 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3000
      TabIndex        =   2
      Top             =   660
      Width           =   375
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   1410
      TabIndex        =   1
      Top             =   660
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   330
      Left            =   4185
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   1155
   End
   Begin VB.CheckBox chkWithinSelection 
      Caption         =   "Within current selection"
      Height          =   255
      Left            =   1410
      TabIndex        =   3
      Top             =   1095
      Width           =   2085
   End
   Begin VB.ComboBox cmbFindType 
      Height          =   330
      ItemData        =   "frmFind.frx":0596
      Left            =   1410
      List            =   "frmFind.frx":0598
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   1965
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find"
      Default         =   -1  'True
      Height          =   330
      Left            =   4185
      TabIndex        =   7
      Top             =   210
      Width           =   1155
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Search type:"
      Height          =   210
      Left            =   345
      TabIndex        =   11
      Top             =   270
      UseMnemonic     =   0   'False
      Width           =   930
   End
   Begin VB.Label lblReplace 
      AutoSize        =   -1  'True
      Caption         =   "Replace with:"
      Height          =   210
      Left            =   285
      TabIndex        =   10
      Top             =   1545
      UseMnemonic     =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Find what:"
      Height          =   210
      Left            =   510
      TabIndex        =   9
      Top             =   720
      UseMnemonic     =   0   'False
      Width           =   765
   End
End
Attribute VB_Name = "frmFind"
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


Option Explicit

Private ResultType As ENUM_EDITMODE

Private Sub FillResultsList()
     Dim i As Long
     Dim ResultItems As Variant
     
     'Clear listbox
     lstResults.Clear
     
     'Get result item indices
     ResultItems = selected.Items
     
     'Cant list more than 32766 items in list
     If (UBound(ResultItems) > 32766) Then ReDim Preserve ResultItems(0 To 32766)
     
     'Check how to list the items
     Select Case ResultType
          
          'Vertices
          Case EM_VERTICES
               
               'Go for all items
               For i = LBound(ResultItems) To UBound(ResultItems)
                    
                    'Add description to list
                    lstResults.AddItem "Vertex " & ResultItems(i) & " at " & vertexes(ResultItems(i)).x & ", " & vertexes(ResultItems(i)).y
                    lstResults.ItemData(lstResults.NewIndex) = ResultItems(i)
               Next i
               
          'Linedefs
          Case EM_LINES
               
               'Go for all items
               For i = LBound(ResultItems) To UBound(ResultItems)
                    
                    'Add description to list
                    lstResults.AddItem "Linedef " & ResultItems(i)
                    lstResults.ItemData(lstResults.NewIndex) = ResultItems(i)
               Next i
          
          'Sectors
          Case EM_SECTORS
               
               'Go for all items
               For i = LBound(ResultItems) To UBound(ResultItems)
                    
                    'Add description to list
                    lstResults.AddItem "Sector " & ResultItems(i)
                    lstResults.ItemData(lstResults.NewIndex) = ResultItems(i)
               Next i
          
          'Things
          Case EM_THINGS
               
               'Go for all items
               For i = LBound(ResultItems) To UBound(ResultItems)
                    
                    'Add description to list
                    lstResults.AddItem "Thing " & ResultItems(i) & " (" & GetThingTypeDesc(things(ResultItems(i)).thing) & ")" & " at " & things(ResultItems(i)).x & ", " & things(ResultItems(i)).y
                    lstResults.ItemData(lstResults.NewIndex) = ResultItems(i)
               Next i
     End Select
End Sub

Private Sub cmbFindType_Change()
     
     'Check what search to perform
     Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
          
          Case FR_VERTEXNUMBER
               txtFind.tag = ""
               txtReplace.Enabled = False
               cmdBrowseFind.Enabled = False
          
          Case FR_LINEDEFNUMBER
               txtFind.tag = ""
               txtReplace.Enabled = False
               cmdBrowseFind.Enabled = False
               
          Case FR_LINEDEFACTION
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
               
          Case FR_LINEDEFSECTORTAG, FR_LINEDEFTHINGTAG
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = False
               
          Case FR_LINEDEFTEXTURE
               txtFind.tag = "TEXT"
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
          
          Case FR_SECTORNUMBER
               txtFind.tag = ""
               txtReplace.Enabled = False
               cmdBrowseFind.Enabled = False
               
          Case FR_SECTOREFFECT
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
               
          Case FR_SECTORTAG
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = False
               
          Case FR_SECTORFLAT
               txtFind.tag = "TEXT"
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
          
          Case FR_THINGNUMBER
               txtFind.tag = ""
               txtReplace.Enabled = False
               cmdBrowseFind.Enabled = False
               
          Case FR_THINGACTION
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
               
          Case FR_THINGTAG, FR_THINGSECTORTAG, FR_THINGTHINGTAG
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = False
               
          Case FR_THINGTYPE
               txtFind.tag = ""
               txtReplace.Enabled = True
               cmdBrowseFind.Enabled = True
     End Select
     
     'Set some other properties
     txtFind.Text = ""
     txtReplace.Text = ""
     txtReplace.tag = txtFind.tag
     chkReplaceOnly.Enabled = txtReplace.Enabled
     cmdBrowseReplace.Enabled = cmdBrowseFind.Enabled And txtReplace.Enabled
     If (chkReplaceOnly.Enabled = False) Then chkReplaceOnly.Value = vbUnchecked
     lblReplace.Enabled = txtReplace.Enabled
     'If (txtReplace.Enabled = True) Then lblReplace.ForeColor = vbGrayText Else lblReplace.ForeColor = vbButtonText
     
     'Change button caption
     If (txtReplace.visible And txtReplace.Enabled) Then
          cmdFind.Caption = "Replace"
     Else
          cmdFind.Caption = "Find"
     End If
End Sub

Private Sub cmbFindType_Click()
     
     'Same as changing
     cmbFindType_Change
End Sub

Private Sub cmbFindType_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Same as changing
     cmbFindType_Change
End Sub

Private Sub cmdBrowseFind_Click()
     
     'Open the correct dialog for selecting
     Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
          Case FR_LINEDEFACTION: txtFind.Text = SelectAction(txtFind.Text, Me)
          Case FR_LINEDEFTEXTURE: txtFind.Text = SelectTexture(txtFind.Text, Me)
          Case FR_SECTOREFFECT: txtFind.Text = SelectSectorEffect(txtFind.Text, Me)
          Case FR_SECTORFLAT: txtFind.Text = SelectFlat(txtFind.Text, Me)
          Case FR_THINGACTION: txtFind.Text = SelectAction(txtFind.Text, Me)
          Case FR_THINGTYPE: txtFind.Text = SelectThing(txtFind.Text, Me)
     End Select
End Sub


Private Sub cmdBrowseReplace_Click()
     
     'Open the correct dialog for selecting
     Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
          Case FR_LINEDEFACTION: txtReplace.Text = SelectAction(txtReplace.Text, Me)
          Case FR_LINEDEFTEXTURE: txtReplace.Text = SelectTexture(txtReplace.Text, Me)
          Case FR_SECTOREFFECT: txtReplace.Text = SelectSectorEffect(txtReplace.Text, Me)
          Case FR_SECTORFLAT: txtReplace.Text = SelectFlat(txtReplace.Text, Me)
          Case FR_THINGACTION: txtReplace.Text = SelectAction(txtReplace.Text, Me)
          Case FR_THINGTYPE: txtReplace.Text = SelectThing(txtReplace.Text, Me)
     End Select
End Sub


Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdFind_Click()
     Dim Count As Long
     
     'Ask for replace input when needed
     If (txtReplace.Enabled = True) And (txtReplace.visible = True) And (Trim$(txtReplace.Text) = "") Then MsgBox "Please enter a value to replace the found items with.", vbExclamation: Exit Sub
     
     'Switch to the correct mode
     Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
          
          'Lines
          Case FR_LINEDEFACTION, FR_LINEDEFNUMBER, _
               FR_LINEDEFSECTORTAG, FR_LINEDEFTHINGTAG, FR_LINEDEFTEXTURE
               If (mode <> EM_LINES) Then frmMain.itmEditMode_Click EM_LINES
               
          'Sectors
          Case FR_SECTOREFFECT, FR_SECTORFLAT, _
               FR_SECTORNUMBER, FR_SECTORTAG
               If (mode <> EM_SECTORS) Then frmMain.itmEditMode_Click EM_SECTORS
          
          'Things
          Case FR_THINGACTION, FR_THINGNUMBER, _
               FR_THINGTAG, FR_THINGTHINGTAG, _
               FR_THINGSECTORTAG, FR_THINGTYPE
               If (mode <> EM_THINGS) Then frmMain.itmEditMode_Click EM_THINGS
          
          'Vertices
          Case FR_VERTEXNUMBER
               If (mode <> EM_VERTICES) Then frmMain.itmEditMode_Click EM_VERTICES
               
     End Select
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Create undo when replacing
     If (txtReplace.Text <> "") Then CreateUndo "replace " & LCase$(cmbFindType.Text) & "s"
     
     'Do the find and replace
     Count = FindSelectAndReplace(cmbFindType.ItemData(cmbFindType.ListIndex), _
                                  txtFind.Text, (chkWithinSelection.Value = vbChecked), _
                                  txtReplace.Text, (chkReplaceOnly.Value = vbChecked))
     
     'Keep copy of selection type for the results list
     ResultType = mode
     
     'Fill the list of results
     FillResultsList
     
     'Check if result tooltip not already shown
     If (picWarning.visible = False) Then
          picWarning.top = ScaleHeight + 2
          height = height + (picWarning.height + 5) * Screen.TwipsPerPixelY
          picWarning.visible = True
          'cmdShowResults.top = picWarning.top + 1
          'cmdShowResults.Visible = True
     End If
     
     'Show result
     If (txtReplace.Enabled = True) And (txtReplace.visible = True) Then
          lblWarning.Caption = Count & " items have been replaced."
     Else
          lblWarning.Caption = Count & " items have been found."
     End If
     
     'Check if items were found
     If (Count > 0) Then
          
          'Show results now
          height = height + 150 * Screen.TwipsPerPixelY
          lstResults.top = picWarning.top + picWarning.height + 4
          lstResults.height = ScaleHeight - lstResults.top - 6
          lstResults.visible = True
          'cmdShowResults.Caption = "5"
          lstResults.SetFocus
     Else
          
          'Nothing done, remove the undo when we were replacing
          If (txtReplace.Text <> "") Then WithdrawUndo
     End If
     
     'Redraw map
     RedrawMap
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
     
     'Store last find type
     LastFindType = cmbFindType.ListIndex
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
     
     'Add search options
     cmbFindType.AddItem "Vertex number": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_VERTEXNUMBER
     cmbFindType.AddItem "Linedef number": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_LINEDEFNUMBER
     cmbFindType.AddItem "Linedef action": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_LINEDEFACTION
     cmbFindType.AddItem "Linedef sector tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_LINEDEFSECTORTAG
     cmbFindType.AddItem "Linedef thing tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_LINEDEFTHINGTAG
     cmbFindType.AddItem "Linedef texture": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_LINEDEFTEXTURE
     cmbFindType.AddItem "Sector number": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_SECTORNUMBER
     cmbFindType.AddItem "Sector effect": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_SECTOREFFECT
     cmbFindType.AddItem "Sector tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_SECTORTAG
     cmbFindType.AddItem "Sector flat": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_SECTORFLAT
     cmbFindType.AddItem "Thing number": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGNUMBER
     cmbFindType.AddItem "Thing action": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGACTION
     cmbFindType.AddItem "Thing tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGTAG
     cmbFindType.AddItem "Thing sector tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGSECTORTAG
     cmbFindType.AddItem "Thing thing tag": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGTHINGTAG
     cmbFindType.AddItem "Thing type": cmbFindType.ItemData(cmbFindType.NewIndex) = FR_THINGTYPE
     
     'Select recent used
     cmbFindType.ListIndex = LastFindType
     cmbFindType_Change
     
     'Move to left top of parent
     left = frmMain.left + 50 * Screen.TwipsPerPixelX
     top = frmMain.top + 100 * Screen.TwipsPerPixelY
End Sub


Private Sub lstResults_Click()
     Dim ItemIndex As Long
     Dim TargetRect As RECT
     
     'Item selected?
     If (lstResults.ListIndex > -1) Then
          
          'Get index
          ItemIndex = lstResults.ItemData(lstResults.ListIndex)
          
          'Clear selection
          'RemoveSelection False
          
          'Check type
          Select Case ResultType
               
               'Vertex
               Case EM_VERTICES
                    
                    'Select this vertex only
                    'vertexes(ItemIndex).selected = 1
                    'selected.Add CStr(ItemIndex), ItemIndex
                    'numselected = 1
                    
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
                    
                    'Select this linedef only
                    'linedefs(ItemIndex).selected = 1
                    'selected.Add CStr(ItemIndex), ItemIndex
                    'numselected = 1
                    
                    'Make rect for linedef
                    TargetRect = CalculateLinedefRect(ItemIndex)
                    
                    'Show it
                    CenterViewAt TargetRect, True, 200, 0.5, 1
                    
               'Sector
               Case EM_SECTORS
                    
                    'Select this sector only
                    'SelectSector ItemIndex
                    'selected.Add CStr(ItemIndex), ItemIndex
                    'numselected = 1
                    
                    'Make rect for sector
                    TargetRect = CalculateSectorRect(ItemIndex)
                    
                    'Show it
                    CenterViewAt TargetRect, True, 200, 0.5, 1
                    
               'Thing
               Case EM_THINGS
                    
                    'Select this thing only
                    'things(ItemIndex).selected = 1
                    'selected.Add CStr(ItemIndex), ItemIndex
                    'numselected = 1
                    
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

Private Sub txtFind_GotFocus()
     SelectAllText txtFind
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
     
     'Check if not typing text
     If (txtFind.tag = "") Then
          
          'Check if key is allowed
          If ((KeyAscii <> 8) And (KeyAscii <> 45) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
     End If
     
     'Convert to uppercase
     If (KeyAscii <> 0) Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Use autocomplete?
     If Val(Config("autocompletetypetex")) Then
          
          'Finish texture names?
          Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
               Case FR_LINEDEFTEXTURE: CompleteTextureName KeyCode, Shift, txtFind
               Case FR_SECTORFLAT: CompleteFlatName KeyCode, Shift, txtFind
          End Select
     End If
End Sub

Private Sub txtFind_Validate(Cancel As Boolean)
     
     'Use autocomplete?
     If Val(Config("autocompletetex")) Then
          
          'Finish texture names?
          Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
               Case FR_LINEDEFTEXTURE: txtFind.Text = GetNearestTextureName(txtFind.Text)
               Case FR_SECTORFLAT: txtFind.Text = GetNearestFlatName(txtFind.Text)
          End Select
     End If
End Sub


Private Sub txtReplace_GotFocus()
     SelectAllText txtReplace
End Sub

Private Sub txtReplace_KeyPress(KeyAscii As Integer)
     
     'Check if not typing text
     If (txtReplace.tag = "") Then
          
          'Check if key is allowed
          If ((KeyAscii <> 8) And (KeyAscii <> 45) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
     End If
     
     'Convert to uppercase
     If (KeyAscii <> 0) Then KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub


Private Sub txtReplace_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Use autocomplete?
     If Val(Config("autocompletetypetex")) Then
          
          'Finish texture names?
          Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
               Case FR_LINEDEFTEXTURE: CompleteTextureName KeyCode, Shift, txtReplace
               Case FR_SECTORFLAT: CompleteFlatName KeyCode, Shift, txtReplace
          End Select
     End If
End Sub

Private Sub txtReplace_Validate(Cancel As Boolean)
     
     'Use autocomplete?
     If Val(Config("autocompletetex")) Then
          
          'Finish texture names?
          Select Case cmbFindType.ItemData(cmbFindType.ListIndex)
               Case FR_LINEDEFTEXTURE: txtReplace.Text = GetNearestTextureName(txtReplace.Text)
               Case FR_SECTORFLAT: txtReplace.Text = GetNearestFlatName(txtReplace.Text)
          End Select
     End If
End Sub


