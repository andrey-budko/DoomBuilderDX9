VERSION 5.00
Begin VB.Form frmThingFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Things Filter"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5220
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
   Icon            =   "frmThingFilter.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   348
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCatagory 
      Caption         =   " Category "
      Height          =   855
      Left            =   180
      TabIndex        =   25
      Top             =   1155
      Width           =   4845
      Begin VB.ComboBox cmbCatagory 
         Height          =   330
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   315
         Width           =   2655
      End
      Begin VB.Label lblCatagory 
         AutoSize        =   -1  'True
         Caption         =   "Thing Category:"
         Height          =   210
         Left            =   495
         TabIndex        =   26
         Top             =   360
         Width           =   1140
      End
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
      ScaleWidth      =   338
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   60
      Width           =   5100
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmThingFilter.frx":000C
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Filtering Things may cause confusion, because the filtered Things are not displayed on your screen, but they do exist!"
         ForeColor       =   &H80000017&
         Height          =   450
         Left            =   375
         TabIndex        =   24
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   4710
      End
   End
   Begin VB.CheckBox chkFilterThings 
      Caption         =   "Only show Things according to the settings below"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   765
      Value           =   1  'Checked
      Width           =   4215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2648
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1665
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   908
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5580
      Width           =   1665
   End
   Begin VB.Frame fraFlags 
      Caption         =   " Flags "
      Height          =   3195
      Left            =   180
      TabIndex        =   21
      Top             =   2130
      Width           =   4845
      Begin VB.OptionButton optMode 
         Caption         =   "All flags"
         Height          =   255
         Index           =   1
         Left            =   1785
         TabIndex        =   28
         Top             =   2790
         Width           =   1155
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Exact flags"
         Height          =   255
         Index           =   2
         Left            =   3135
         TabIndex        =   29
         Top             =   2790
         Width           =   1155
      End
      Begin VB.OptionButton optMode 
         Caption         =   "Any flags"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   27
         Top             =   2790
         Width           =   1155
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   15
         Left            =   2550
         TabIndex        =   17
         Tag             =   "0"
         Top             =   2295
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   14
         Left            =   2550
         TabIndex        =   16
         Tag             =   "0"
         Top             =   2010
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   13
         Left            =   2550
         TabIndex        =   15
         Tag             =   "0"
         Top             =   1725
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   12
         Left            =   2550
         TabIndex        =   14
         Tag             =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   11
         Left            =   2550
         TabIndex        =   13
         Tag             =   "0"
         Top             =   1155
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   10
         Left            =   2550
         TabIndex        =   12
         Tag             =   "0"
         Top             =   870
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   9
         Left            =   2550
         TabIndex        =   11
         Tag             =   "0"
         Top             =   585
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   8
         Left            =   2550
         TabIndex        =   10
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   7
         Left            =   390
         TabIndex        =   9
         Tag             =   "0"
         Top             =   2295
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   6
         Left            =   390
         TabIndex        =   8
         Tag             =   "0"
         Top             =   2010
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   5
         Left            =   390
         TabIndex        =   7
         Tag             =   "0"
         Top             =   1725
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   4
         Left            =   390
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1440
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   3
         Left            =   390
         TabIndex        =   5
         Tag             =   "0"
         Top             =   1155
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   2
         Left            =   390
         TabIndex        =   4
         Tag             =   "0"
         Top             =   870
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   1
         Left            =   390
         TabIndex        =   3
         Tag             =   "0"
         Top             =   585
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.CheckBox chkFlag 
         Caption         =   "Thing Flag"
         Height          =   255
         Index           =   0
         Left            =   390
         TabIndex        =   2
         Tag             =   "0"
         Top             =   300
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.TextBox txtRawFlags 
         Height          =   315
         Left            =   2580
         MaxLength       =   6
         TabIndex        =   20
         Text            =   "value"
         Top             =   2895
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "flags value"
         Height          =   195
         Left            =   3585
         TabIndex        =   22
         Top             =   2955
         Visible         =   0   'False
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmThingFilter"
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
Private Sub chkFilterThings_Click()
     Dim f As Long
     
     'Disable controls
     cmbCatagory.Enabled = (chkFilterThings.Value = vbChecked)
     optMode(0).Enabled = (chkFilterThings.Value = vbChecked)
     optMode(1).Enabled = (chkFilterThings.Value = vbChecked)
     optMode(2).Enabled = (chkFilterThings.Value = vbChecked)
     
     'Disable array of flags
     For f = chkFlag.LBound To chkFlag.UBound
          chkFlag(f).Enabled = (chkFilterThings.Value = vbChecked)
     Next f
     
     'Gray labels
     If (chkFilterThings.Value = vbChecked) Then lblCatagory.ForeColor = vbButtonText Else lblCatagory.ForeColor = vbGrayText
End Sub

Private Sub chkFlag_Click(Index As Integer)
     txtRawFlags.Text = ""
     txtRawFlags.Text = ""
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdOK_Click()
     Dim f As Long
     
     'Set filter on or off
     filterthings = (chkFilterThings.Value = vbChecked)
     
     'Set category
     filtersettings.category = cmbCatagory.ListIndex - 1
     
     'Check if raw code set
     If (Trim$(txtRawFlags.Text) = "") Then
          
          'Begin with 0
          filtersettings.flags = 0
          
          'Go for all individual flags
          For f = 0 To 15
               
               'Check if this flag can be set
               If (chkFlag(f).tag <> "0") Then
                    
                    'Check if the flag is marked to be set
                    If (chkFlag(f).Value = vbChecked) Then
                         
                         'Add the flag on the thing
                         filtersettings.flags = filtersettings.flags Or CLng(chkFlag(f).tag)
                    End If
               End If
          Next f
     Else
          
          'Set flags from raw
          On Error Resume Next
          filtersettings.flags = Val(txtRawFlags.Text)
          On Error GoTo 0
     End If
     
     'Set mode
     If (optMode(0).Value) Then filtersettings.filtermode = 0
     If (optMode(1).Value) Then filtersettings.filtermode = 1
     If (optMode(2).Value) Then filtersettings.filtermode = 2
     
     'Redraw map
     RedrawMap
     
     'Leave
     Unload Me
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
     Dim Keys As Variant
     Dim nflag As Long
     Dim i As Long
     Dim Cats As Variant
     
     'Add no category
     cmbCatagory.AddItem "(any category)"
     
     'Go for all categories
     Cats = mapconfig("thingtypes").Items
     For i = LBound(Cats) To UBound(Cats)
          
          'Add category to combobox
          cmbCatagory.AddItem Cats(i)("title")
     Next i
     
     'Set category
     On Error Resume Next
     cmbCatagory.ListIndex = filtersettings.category + 1
     On Error GoTo 0
     
     'Go for all flags
     Keys = mapconfig("thingflags").Keys
     For i = 0 To 15
          
          'Check if this flag is known
          If (mapconfig("thingflags").Exists(CStr(2 ^ i)) = True) Then
               
               'Check if not unset
               If CStr(mapconfig("thingflags")(CStr(2 ^ i))) <> "0" Then
                    
                    'Set the checkbox properties
                    chkFlag(nflag).tag = CStr(2 ^ i)
                    chkFlag(nflag).Visible = True
                    chkFlag(nflag).Caption = mapconfig("thingflags")(CStr(2 ^ i))
                    
                    'Check this flag
                    chkFlag(nflag).Value = Abs((filtersettings.flags And (2 ^ i)) = (2 ^ i))
                    
                    'Next flag
                    nflag = nflag + 1
               Else
                    
                    'Zero tag
                    chkFlag(nflag).tag = "0"
               End If
          Else
               
               'Zero tag
               chkFlag(nflag).tag = "0"
          End If
     Next i
     
     'Set raw value
     txtRawFlags.Text = filtersettings.flags
     
     'Set option
     optMode(filtersettings.filtermode).Value = True
     
     'Filter on or off
     chkFilterThings.Value = Abs(filterthings)
End Sub


