VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSectorType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Sector Type"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5310
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
   Icon            =   "frmSectorType.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   377
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   354
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1710
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5175
      Width           =   1665
   End
   Begin VB.Frame fraStandard 
      Height          =   4215
      Left            =   315
      TabIndex        =   21
      Top             =   600
      Width           =   4695
      Begin MSComctlLib.ListView lstTypes 
         Height          =   3705
         Left            =   180
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Width           =   4365
         _ExtentX        =   7699
         _ExtentY        =   6535
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Num"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   5292
         EndProperty
      End
   End
   Begin VB.Frame fraOptions 
      Caption         =   " Options "
      Height          =   4215
      Left            =   315
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3570
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   6
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2670
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   4
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2220
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1770
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   2
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1320
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   870
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.ComboBox cmbOption 
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   420
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   7
         Left            =   180
         TabIndex        =   20
         Top             =   3630
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   6
         Left            =   180
         TabIndex        =   19
         Top             =   3180
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   5
         Left            =   180
         TabIndex        =   18
         Top             =   2730
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   4
         Left            =   180
         TabIndex        =   17
         Top             =   2280
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   3
         Left            =   180
         TabIndex        =   16
         Top             =   1830
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   2
         Left            =   180
         TabIndex        =   15
         Top             =   1380
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   930
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lblOption 
         Alignment       =   1  'Right Justify
         Caption         =   "Option:"
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   480
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin MSComctlLib.TabStrip tbsPanels 
      Height          =   4845
      Left            =   150
      TabIndex        =   11
      Top             =   150
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   8546
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   3678
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Standard Effects"
            Key             =   "standard"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Generalized Effects"
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
Attribute VB_Name = "frmSectorType"
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
          If (tbsPanels.SelectedItem.Index = tbsPanels.Tabs.Count) Then
               tbsPanels.Tabs(1).selected = True
          Else
               tbsPanels.Tabs(tbsPanels.SelectedItem.Index + 1).selected = True
          End If
          
          'Focus to panel
          tbsPanels.SetFocus
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     Dim Keys As Variant
     Dim Typedesc As String
     Dim i As Long
     Dim li As ListItem
     Dim o As Long
     Dim Opts As Dictionary
     Dim OptKeys As Variant
     
     'Check if using generalized sector effects
     If (Val(mapconfig("generalizedsectors")) <> 0) Then
          
          'Get the options dictionary
          Set Opts = mapconfig("gen_sectortypes")
          OptKeys = Opts.Keys
          
          'Go for all options
          For o = 0 To (Opts.Count - 1)
               
               'Check if combos are available
               If (o <= cmbOption.UBound) Then
                    
                    'Clear box
                    cmbOption(o).Clear
                    
                    'Set the caption
                    lblOption(o).Caption = StrConv(OptKeys(o), vbProperCase) & ":"
                    
                    'Get the key values
                    Keys = Opts(OptKeys(o)).Keys
                    
                    'Go for all items to add them to the combo
                    For i = LBound(Keys) To UBound(Keys)
                         
                         'Add to combo
                         With cmbOption(o)
                              .AddItem Opts(OptKeys(o))(Keys(i))
                              .ItemData(.NewIndex) = Val(Keys(i))
                         End With
                    Next i
                    
                    'Select the first item
                    cmbOption(o).ListIndex = 0
                    
                    'Show option
                    cmbOption(o).Visible = True
                    cmbOption(o).Enabled = True
                    lblOption(o).Visible = True
               End If
          Next o
     Else
          
          'No generalized shit
          tbsPanels.Tabs.Remove "generalized"
     End If
     
     'Fill the types list with all linedef types
     Keys = mapconfig("sectortypes").Keys
     For i = LBound(Keys) To UBound(Keys)
          
          'Add the item to list
          Set li = lstTypes.ListItems.Add(, "L" & CStr(Keys(i)), Space$(5 - Len(CStr(Keys(i)))) & Keys(i))
          
          'Get type description
          Typedesc = mapconfig("sectortypes")(Keys(i))
          
          'Add description
          li.ListSubItems.Add , , Typedesc
          
          'Clean up
          Set li = Nothing
     Next i
     
     'Sort the list
     lstTypes.SortKey = Abs(Val(Config("sectorssort"))) - 1
     lstTypes.SortOrder = Abs(Val(Config("sectorssort")) < 0)
     
     'When there are no standard types but there are generalized types, remove the standard tab
     If (mapconfig("sectortypes").Count = 0) And (Val(mapconfig("generalizedsectors")) <> 0) Then
          
          'Open other tab instead
          tbsPanels.Tabs("generalized").selected = True
          
          'Remove standard tab
          tbsPanels.Tabs.Remove "standard"
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
          Config("sectorssort") = (lstTypes.SortKey + 1)
     Else
          Config("sectorssort") = -(lstTypes.SortKey + 1)
     End If
End Sub

Private Sub lstTypes_DblClick()
     
     'Click OK
     cmdOK_Click
End Sub

Private Sub tbsPanels_Click()
     If (tbsPanels.SelectedItem.Key = "standard") Then
          fraStandard.Visible = True
          fraOptions.Visible = False
     Else
          fraStandard.Visible = False
          fraOptions.Visible = True
     End If
End Sub
