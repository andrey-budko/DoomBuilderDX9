VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLinedefType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Action"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
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
   Icon            =   "frmLinedefType.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   433
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   442
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4845
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3060
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6015
      Width           =   1665
   End
   Begin VB.Frame fraStandard 
      Height          =   5160
      Left            =   285
      TabIndex        =   16
      Top             =   510
      Width           =   6045
      Begin MSComctlLib.TreeView trvTypes 
         Height          =   4005
         Left            =   210
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   7064
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
         FullRowSelect   =   -1  'True
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lstTypes 
         Height          =   4005
         Left            =   210
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   7064
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Description"
            Object.Width           =   6879
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "R = Repeatable"
         Height          =   210
         Left            =   2655
         TabIndex        =   32
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "1 = Once"
         Height          =   210
         Left            =   4485
         TabIndex        =   31
         Top             =   420
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "G = Gunfire"
         Height          =   210
         Left            =   4470
         TabIndex        =   30
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "W = Walk over"
         Height          =   210
         Left            =   2640
         TabIndex        =   29
         Top             =   210
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "H = Hexen action which uses extended arguments and activation"
         Height          =   210
         Left            =   600
         TabIndex        =   19
         Top             =   630
         Width           =   4710
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "D = Door (do not tag)"
         Height          =   210
         Left            =   600
         TabIndex        =   18
         Top             =   420
         Width           =   1545
      End
      Begin VB.Label Label3523246436 
         AutoSize        =   -1  'True
         Caption         =   "S = Switch"
         Height          =   210
         Left            =   600
         TabIndex        =   17
         Top             =   210
         Width           =   825
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   " Options "
      Height          =   4200
      Left            =   285
      TabIndex        =   20
      Top             =   1470
      Visible         =   0   'False
      Width           =   6045
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   870
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1770
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2220
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2670
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   3570
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   0
         Left            =   150
         TabIndex        =   28
         Top             =   480
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   1
         Left            =   150
         TabIndex        =   27
         Top             =   930
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   2
         Left            =   150
         TabIndex        =   26
         Top             =   1380
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   3
         Left            =   150
         TabIndex        =   25
         Top             =   1830
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   4
         Left            =   150
         TabIndex        =   24
         Top             =   2280
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   5
         Left            =   150
         TabIndex        =   23
         Top             =   2730
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   6
         Left            =   150
         TabIndex        =   22
         Top             =   3180
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   7
         Left            =   150
         TabIndex        =   21
         Top             =   3630
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1725
      End
   End
   Begin VB.Frame fraCategorie 
      Caption         =   " Category "
      Height          =   825
      Left            =   285
      TabIndex        =   14
      Top             =   510
      Visible         =   0   'False
      Width           =   6045
      Begin VB.ComboBox cmbCategory 
         Height          =   330
         Left            =   2010
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   2865
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Select Category:"
         Height          =   210
         Left            =   675
         TabIndex        =   15
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1200
      End
   End
   Begin MSComctlLib.TabStrip tbsPanel 
      Height          =   5745
      Left            =   105
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   105
      Width           =   6405
      _ExtentX        =   11298
      _ExtentY        =   10134
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   4207
      TabMinWidth     =   1323
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Standard Linedefs"
            Key             =   "standard"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Generalized Linedefs"
            Key             =   "generalized"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmLinedefType"
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

Private Sub cmbCategory_Change()
     Dim i As Long, k As Long
     Dim Cats As Variant
     Dim Cat As Dictionary
     Dim Opts As Variant
     Dim NextIndex As Long
     Dim Opt As Dictionary
     Dim OptKeys As Variant
     
     'Get category object
     Cats = mapconfig("gen_linedeftypes").Items
     Set Cat = Cats(cmbCategory.ListIndex)
     
     'Go for all items in category
     Opts = Cat.Keys
     For i = 0 To (Cat.Count - 1)
          
          'Check if this item is an option
          If (VarType(Cat(Opts(i))) = vbObject) Then
               
               'Clear combobox
               cmbOption(NextIndex).Clear
               
               'Get the option dictionary
               Set Opt = Cat(Opts(i))
               
               'Fill the combobox
               OptKeys = Opt.Keys
               For k = 0 To (Opt.Count - 1)
                    
                    'Add to combo
                    With cmbOption(NextIndex)
                         .AddItem Opt(OptKeys(k))
                         .ItemData(.NewIndex) = Val(OptKeys(k))
                    End With
               Next k
               
               'Select the first item
               cmbOption(NextIndex).ListIndex = 0
               
               'Set the caption
               lblOption(NextIndex).Caption = StrConv(Opts(i), vbProperCase) & ":"
               
               'Show the combobox
               cmbOption(NextIndex).Enabled = True
               cmbOption(NextIndex).Visible = True
               lblOption(NextIndex).Visible = True
               
               'Change next index
               NextIndex = NextIndex + 1
               If (NextIndex > cmbOption.UBound) Then Exit For
          End If
     Next i
     
     'Remove all other options
     For i = NextIndex To cmbOption.UBound
          
          'Clear and hide the combobox
          cmbOption(i).Clear
          cmbOption(i).Enabled = False
          cmbOption(i).Visible = False
          lblOption(i).Visible = False
     Next i
End Sub

Private Sub cmbCategory_Click()
     cmbCategory_Change
End Sub

Private Sub cmbCategory_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbCategory_Change
End Sub

Private Sub cmdCancel_Click()
     tag = 0
     Hide
End Sub

Private Sub cmdOK_Click()
     tag = 1
     Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Check what key is pressed
     If (KeyCode = vbKeyTab) And (Shift = vbCtrlMask) Then
          
          'Switch to next panel
          If (tbsPanel.SelectedItem.Index = tbsPanel.Tabs.Count) Then
               tbsPanel.Tabs(1).selected = True
          Else
               tbsPanel.Tabs(tbsPanel.SelectedItem.Index + 1).selected = True
          End If
          
          'Focus to panel
          tbsPanel.SetFocus
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     Dim Keys As Variant
     Dim flags As String
     Dim Cat As String
     Dim desc As String
     Dim sp As Long
     Dim i As Long
     Dim li As ListItem
     Dim ni As Node
     
     'Check if showing tree or list
     If (Val(Config("linestree")) = vbUnchecked) Then
          
          'Fill the types list with all linedef types
          Keys = mapconfig("linedeftypes").Keys
          For i = LBound(Keys) To UBound(Keys)
               
               'Add the item to list
               Set li = lstTypes.ListItems.Add(, "L" & CStr(Keys(i)), Space$(5 - Len(CStr(Keys(i)))) & Keys(i))
               
               'Get type description
               desc = mapconfig("linedeftypes")(Keys(i))("title")
               
               'Get the first space position
               sp = InStr(desc, " ")
               
               'Check if we can add with seperate activation type
               If (sp > 0) Then
                    
                    'Add seperated activation type and description
                    li.ListSubItems.Add , , Trim$(left$(desc, sp - 1))
                    li.ListSubItems.Add , , Trim$(Mid$(desc, sp + 1))
               Else
                    
                    'Add only description
                    li.ListSubItems.Add , , ""
                    li.ListSubItems.Add , , desc
               End If
               
               'Clean up
               Set li = Nothing
          Next i
          
          'Sort the list
          lstTypes.SortKey = Abs(Val(Config("linedefssort"))) - 1
          lstTypes.SortOrder = Abs(Val(Config("linedefssort")) < 0)
          lstTypes.Visible = True
     Else
          
          'Fill the types tree with all linedef types
          Keys = mapconfig("linedeftypes").Keys
          For i = LBound(Keys) To UBound(Keys)
               
               'Get description
               desc = mapconfig("linedeftypes")(Keys(i))("title")
               
               'Get the first space position
               sp = InStr(desc, " ")
               
               'Get flags
               flags = left$(desc, sp - 1)
               desc = Mid$(desc, sp + 1)
               
               'Get the next space position
               sp = InStr(desc, " ")
               If (sp = 0) Then sp = Len(desc) + 1
               
               'Get category
               Cat = " " & Trim$(left$(desc, sp - 1))
               
               'Sort normal to top
               If IsNumeric(Keys(i)) And (Val(Keys(i)) = 0) Then
                    desc = "  " & desc
               Else
                    desc = Trim$(desc)
               End If
               
               'Check for category
               If (Trim$(Cat) <> "") Then
                    
                    'Get category item
                    On Local Error Resume Next
                    Set ni = trvTypes.nodes(LCase$(Cat))
                    On Local Error GoTo 0
                    
                    'Check if category does not yet exist
                    If (ni Is Nothing) Then
                         
                         'Make category
                         Set ni = trvTypes.nodes.Add(, , LCase$(Cat), Cat)
                         ni.Sorted = True
                    End If
                    
                    'Add the item to list
                    Set ni = trvTypes.nodes.Add(ni, tvwChild, "L" & CStr(Keys(i)), flags & " " & desc & " (" & CStr(Keys(i)) & ")")
                    ni.tag = CStr(Keys(i))
               Else
                    
                    'Add the item to list
                    Set ni = trvTypes.nodes.Add(, , "L" & CStr(Keys(i)), flags & " " & desc & " (" & CStr(Keys(i)) & ")")
                    ni.tag = CStr(Keys(i))
               End If
               
               'Clean up
               Set ni = Nothing
          Next i
          
          'Show list
          trvTypes.Visible = True
     End If
     
     
     'Check if generalized lindefs are supported
     If (Val(mapconfig("generalizedlinedefs")) <> 0) Then
          
          'Go for each category
          Keys = mapconfig("gen_linedeftypes").Keys
          For i = LBound(Keys) To UBound(Keys)
               
               'Add category title to combo
               cmbCategory.AddItem mapconfig("gen_linedeftypes")(Keys(i))("title")
          Next i
          
          'When there are no standard linedefs, remove the tab
          If (lstTypes.ListItems.Count = 0) And (trvTypes.nodes.Count = 0) Then
               
               'Open other tab instead
               tbsPanel.Tabs("generalized").selected = True
               
               'Remove standard tab
               tbsPanel.Tabs.Remove "standard"
          End If
     Else
          
          'No generalized linedefs
          tbsPanel.Tabs.Remove "generalized"
     End If
End Sub

Private Sub lstTypes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     
     'Check if already sorted by this column
     If lstTypes.SortKey = (ColumnHeader.Index - 1) Then
          
          'Reverse sort
          If lstTypes.SortOrder = lvwAscending Then
               lstTypes.SortOrder = lvwDescending
          Else
               lstTypes.SortOrder = lvwAscending
          End If
     Else
          
          'Change sort key
          lstTypes.SortKey = ColumnHeader.Index - 1
          lstTypes.SortOrder = lvwAscending
          lstTypes.Sorted = True
     End If
     
     'Save sort
     If (lstTypes.SortOrder = lvwAscending) Then
          Config("linedefssort") = (lstTypes.SortKey + 1)
     Else
          Config("linedefssort") = -(lstTypes.SortKey + 1)
     End If
End Sub

Private Sub lstTypes_DblClick()
     
     'Click OK
     cmdOK_Click
End Sub

Private Sub tbsPanel_Click()
     If (tbsPanel.SelectedItem.Index = 1) Then
          fraStandard.Visible = True
          fraCategorie.Visible = False
          fraOptions.Visible = False
     Else
          fraStandard.Visible = False
          fraCategorie.Visible = True
          fraOptions.Visible = True
          If (cmbCategory.ListIndex = -1) And (cmbCategory.ListCount > 0) Then cmbCategory.ListIndex = 0
     End If
End Sub

Private Sub trvTypes_DblClick()
     
     'Check if node can be used
     If (trvTypes.SelectedItem.Children = 0) Then
          
          'Click OK
          cmdOK_Click
     End If
End Sub
