VERSION 5.00
Begin VB.Form frmSector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Sector Selection"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
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
   Icon            =   "frmSector.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   299
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSetDefault 
      Caption         =   "Set as build defaults"
      Height          =   255
      Left            =   225
      TabIndex        =   9
      Top             =   4050
      Width           =   2265
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3795
      TabIndex        =   10
      Top             =   4020
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5445
      TabIndex        =   11
      Top             =   4020
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   " Floor and Ceiling "
      Height          =   1815
      Left            =   75
      TabIndex        =   17
      Top             =   1980
      Width           =   6975
      Begin VB.PictureBox picTCeiling 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   4650
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Ceiling Texture"
         Top             =   255
         Width           =   1020
         Begin VB.Image imgTCeiling 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Ceiling Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.TextBox txtTCeiling 
         Height          =   315
         Left            =   4650
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "-"
         Top             =   1335
         Width           =   1020
      End
      Begin VB.PictureBox picTFloor 
         BackColor       =   &H8000000C&
         CausesValidation=   0   'False
         ClipControls    =   0   'False
         HasDC           =   0   'False
         Height          =   1020
         Left            =   5790
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Floor Texture"
         Top             =   255
         Width           =   1020
         Begin VB.Image imgTFloor 
            Height          =   960
            Left            =   0
            Stretch         =   -1  'True
            ToolTipText     =   "Floor Texture"
            Top             =   0
            Width           =   960
         End
      End
      Begin VB.TextBox txtTFloor 
         Height          =   315
         Left            =   5790
         MaxLength       =   8
         TabIndex        =   8
         Text            =   "-"
         Top             =   1335
         Width           =   1020
      End
      Begin DoomBuilder.ctlValueBox txtHCeiling 
         Height          =   360
         Left            =   1455
         TabIndex        =   5
         Top             =   360
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         MaxLength       =   8
         Min             =   -32767
         SmallChange     =   8
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin DoomBuilder.ctlValueBox txtHFloor 
         Height          =   360
         Left            =   1455
         TabIndex        =   6
         Top             =   810
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
         MaxLength       =   8
         Min             =   -32767
         SmallChange     =   8
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   210
         Left            =   1500
         TabIndex        =   24
         Top             =   1350
         Width           =   90
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector height:"
         Height          =   210
         Left            =   270
         TabIndex        =   23
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Floor height:"
         Height          =   210
         Left            =   390
         TabIndex        =   20
         Top             =   885
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Ceiling height:"
         Height          =   210
         Left            =   285
         TabIndex        =   19
         Top             =   435
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Action "
      Height          =   855
      Left            =   75
      TabIndex        =   14
      Top             =   1035
      Width           =   6975
      Begin VB.CommandButton cmdNextTag 
         Caption         =   "Next Unused"
         Height          =   345
         Left            =   2355
         TabIndex        =   4
         Top             =   300
         Width           =   1335
      End
      Begin DoomBuilder.ctlValueBox txtTag 
         Height          =   360
         Left            =   1455
         TabIndex        =   3
         Top             =   285
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32768
         EmptyAllowed    =   -1  'True
         RelativeAllowed =   -1  'True
         Unsigned        =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector Tag:"
         Height          =   210
         Left            =   420
         TabIndex        =   15
         Top             =   375
         UseMnemonic     =   0   'False
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Effects "
      Height          =   855
      Left            =   75
      TabIndex        =   12
      Top             =   90
      Width           =   6975
      Begin VB.ComboBox cmbBrightness 
         Height          =   330
         ItemData        =   "frmSector.frx":000C
         Left            =   5580
         List            =   "frmSector.frx":000E
         TabIndex        =   2
         Text            =   "0"
         Top             =   315
         Width           =   885
      End
      Begin VB.CommandButton cmdSelectType 
         Caption         =   "Select Effect..."
         Height          =   345
         Left            =   2355
         TabIndex        =   1
         Top             =   300
         Width           =   1335
      End
      Begin DoomBuilder.ctlValueBox txtType 
         Height          =   360
         Left            =   1455
         TabIndex        =   0
         Top             =   285
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   635
         MaxLength       =   5
         Min             =   -32768
         EmptyAllowed    =   -1  'True
         Unsigned        =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Brightness:"
         Height          =   210
         Left            =   4605
         TabIndex        =   16
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   825
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Sector Effect:"
         Height          =   210
         Left            =   255
         TabIndex        =   13
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   1005
      End
   End
   Begin VB.Label lblMakeUndo 
      Height          =   210
      Left            =   360
      TabIndex        =   22
      Top             =   4095
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmSector"
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

Private Function CheckSectorBrightness() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first sector's brightness
     CheckSectorBrightness = sectors(Indices(LBound(Indices))).Brightness
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the brightness is different
          If (sectors(Indices(i)).Brightness <> CheckSectorBrightness) Then
               CheckSectorBrightness = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorHCeiling() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first
     CheckSectorHCeiling = sectors(Indices(LBound(Indices))).hceiling
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (sectors(Indices(i)).hceiling <> CheckSectorHCeiling) Then
               CheckSectorHCeiling = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorHFloor() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first
     CheckSectorHFloor = sectors(Indices(LBound(Indices))).hfloor
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (sectors(Indices(i)).hfloor <> CheckSectorHFloor) Then
               CheckSectorHFloor = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorTag() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first sector's tag
     CheckSectorTag = sectors(Indices(LBound(Indices))).tag
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the tag is different
          If (sectors(Indices(i)).tag <> CheckSectorTag) Then
               CheckSectorTag = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorTCeiling() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first
     CheckSectorTCeiling = UCase$(sectors(Indices(LBound(Indices))).tceiling)
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (UCase$(sectors(Indices(i)).tceiling) <> CheckSectorTCeiling) Then
               CheckSectorTCeiling = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorTFloor() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first
     CheckSectorTFloor = UCase$(sectors(Indices(LBound(Indices))).tfloor)
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if different
          If (UCase$(sectors(Indices(i)).tfloor) <> CheckSectorTFloor) Then
               CheckSectorTFloor = ""
               Exit Function
          End If
     Next i
End Function

Private Function CheckSectorType() As String
     Dim i As Long
     Dim Indices As Variant
     
     'Get selection indices
     Indices = selected.Items
     
     'Set result to first sector's type
     CheckSectorType = sectors(Indices(LBound(Indices))).special
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Check if the type is different
          If (sectors(Indices(i)).special <> CheckSectorType) Then
               CheckSectorType = ""
               Exit Function
          End If
     Next i
End Function

Private Sub cmbBrightness_GotFocus()
     SelectAllText cmbBrightness
End Sub

Private Sub cmbBrightness_KeyPress(KeyAscii As Integer)
     If (KeyAscii <> 8) And _
        ((KeyAscii < 48) Or (KeyAscii > 57)) And _
        (KeyAscii <> 43) And (KeyAscii <> 45) Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
     Unload Me
     Set frmSector = Nothing
End Sub

Private Sub cmdNextTag_Click()
     txtTag.Text = NextUnusedTag
End Sub

Private Sub cmdOK_Click()
     Dim Indices As Variant
     Dim i As Long
     Dim s As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Make undo
     If (lblMakeUndo.Caption = "") Then CreateUndo "sector edit"
     
     'Get the selection indices
     Indices = selected.Items
     
     'Go for all selected sectors
     For i = LBound(Indices) To UBound(Indices)
          
          'Get the sector index
          s = Indices(i)
          
          'Apply type if a type is specified
          If (txtType.Text <> "") Then sectors(s).special = Val(txtType.Text)
          
          'Apply tag if a tag is specified
          If (txtTag.Text <> "") Then sectors(s).tag = txtTag.RelativeValue(sectors(s).tag)
          
          'Apply brightness if a brightness is specified
          If (cmbBrightness.Text <> "") Then sectors(s).Brightness = RelativeBrightness(sectors(s).Brightness)
          
          'Apply heights if specified
          If (txtHCeiling.Text <> "") Then sectors(s).hceiling = txtHCeiling.RelativeValue(sectors(s).hceiling)
          If (txtHFloor.Text <> "") Then sectors(s).hfloor = txtHFloor.RelativeValue(sectors(s).hfloor)
          
          'Apply textures if speicified
          If (txtTCeiling.Text <> "") Then sectors(s).tceiling = txtTCeiling.Text
          If (txtTFloor.Text <> "") Then sectors(s).tfloor = txtTFloor.Text
     Next i
     
     'Map is modified
     mapnodeschanged = True
     mapchanged = True
     
     'Make build defaults if requested
     If (chkSetDefault.Value = vbChecked) Then
          
          'Set the build defaults
          Config("defaultsector")("brightness") = Val(cmbBrightness.Text)
          Config("defaultsector")("hceiling") = Val(txtHCeiling.Value)
          Config("defaultsector")("hfloor") = Val(txtHFloor.Value)
          Config("defaultsector")("tceiling") = txtTCeiling.Text
          Config("defaultsector")("tfloor") = txtTFloor.Text
     End If
     
     'Reset mousepointer
     Screen.MousePointer = vbNormal
     
     'Leave here
     Unload Me
     Set frmSector = Nothing
End Sub

Private Sub cmdSelectType_Click()
     txtType.Text = SelectSectorEffect(txtType.Text, Me)
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
     Dim i As Long
     Dim Levels As Variant
     
     'Check if only one selected
     If (numselected = 1) Then Caption = Caption & " - Sector " & selected.Items(0)
     
     'Fill combo with default brightness levels
     Levels = mapconfig("sectorbrightness").Keys
     For i = LBound(Levels) To UBound(Levels)
          cmbBrightness.AddItem Levels(i)
     Next i
     
     'Sector type
     txtType.Text = CheckSectorType
     
     'Sector tag
     txtTag.Text = CheckSectorTag
     
     'Sector Brightness
     cmbBrightness.Text = CheckSectorBrightness
     
     'Floor and Ceiling
     txtHCeiling.Text = CheckSectorHCeiling
     txtHFloor.Text = CheckSectorHFloor
     txtTCeiling.Text = CheckSectorTCeiling
     txtTFloor.Text = CheckSectorTFloor
     
     'Set the relativescroll property
     'NOTE: This doesnt work nicely
     'txtHCeiling.RelativeScroll = (numselected > 1)
     'txtHFloor.RelativeScroll = (numselected > 1)
End Sub

Public Function RelativeBrightness(ByVal OriginalValue As Long) As Long
     On Local Error Resume Next
     
     'Check if theres anything given
     If (Replace$(Replace$(cmbBrightness.Text, "-", ""), "+", "") <> "") Then
          
          'Check if the value is relative
          If (left$(cmbBrightness.Text, 2) = "--") Or (left$(cmbBrightness.Text, 2) = "++") Then
               
               'Add/Subtract to original
               RelativeBrightness = OriginalValue + Val(Mid$(cmbBrightness.Text, 2))
          Else
               
               'Apply normally
               RelativeBrightness = Val(cmbBrightness.Text)
          End If
     Else
          
          'Keep original value
          RelativeBrightness = OriginalValue
     End If
End Function


Private Sub imgTCeiling_Click()
     txtTCeiling.Text = SelectFlat(txtTCeiling.Text, Me)
End Sub

Private Sub imgTFloor_Click()
     txtTFloor.Text = SelectFlat(txtTFloor.Text, Me)
End Sub

Private Sub txtHCeiling_Change()
     
     'Display the height
     If (Trim$(txtHCeiling.Text) <> "") And (Trim$(txtHFloor.Text) <> "") Then
          lblHeight.Caption = Val(txtHCeiling.Value) - Val(txtHFloor.Value)
     Else
          lblHeight.Caption = "-"
     End If
End Sub

Private Sub txtHCeiling_GotFocus()
     SelectAllText txtHCeiling
End Sub


Private Sub txtHFloor_Change()
     
     'Display the height
     If (Trim$(txtHCeiling.Text) <> "") And (Trim$(txtHFloor.Text) <> "") Then
          lblHeight.Caption = Val(txtHCeiling.Value) - Val(txtHFloor.Value)
     Else
          lblHeight.Caption = "-"
     End If
End Sub

Private Sub txtHFloor_GotFocus()
     SelectAllText txtHFloor
End Sub


Private Sub txtTag_GotFocus()
     SelectAllText txtTag
End Sub


Private Sub txtTCeiling_Change()
     
     'Set the flat in the preview box
     GetScaledFlatPicture txtTCeiling.Text, imgTCeiling
End Sub

Private Sub txtTCeiling_GotFocus()
     SelectAllText txtTCeiling
End Sub


Private Sub txtTCeiling_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteFlatName KeyCode, Shift, txtTCeiling
End Sub

Private Sub txtTCeiling_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtTCeiling.Text = GetNearestFlatName(txtTCeiling.Text)
End Sub


Private Sub txtTFloor_Change()
     
     'Set the flat in the preview box
     GetScaledFlatPicture txtTFloor.Text, imgTFloor
End Sub

Private Sub txtTFloor_GotFocus()
     SelectAllText txtTFloor
End Sub


Private Sub txtTFloor_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Complete texture name
     If Val(Config("autocompletetypetex")) Then CompleteFlatName KeyCode, Shift, txtTFloor
End Sub

Private Sub txtTFloor_Validate(Cancel As Boolean)
     
     'Find closest match if preferred
     If Val(Config("autocompletetex")) Then txtTFloor.Text = GetNearestFlatName(txtTFloor.Text)
End Sub


Private Sub txtType_GotFocus()
     SelectAllText txtType
End Sub


