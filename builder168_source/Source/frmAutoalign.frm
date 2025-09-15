VERSION 5.00
Begin VB.Form frmAutoalign 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autoalign Textures"
   ClientHeight    =   4785
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
   Icon            =   "frmAutoalign.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   368
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkY 
      Caption         =   "Align Y offsets"
      Height          =   255
      Left            =   3525
      TabIndex        =   8
      Top             =   3390
      Width           =   1755
   End
   Begin VB.CheckBox chkX 
      Caption         =   "Align X offsets"
      Height          =   255
      Left            =   3525
      TabIndex        =   7
      Top             =   3075
      Width           =   1755
   End
   Begin VB.CheckBox chkBack 
      Caption         =   "Start on Back sides"
      Height          =   255
      Left            =   3525
      TabIndex        =   2
      Top             =   2760
      Width           =   1755
   End
   Begin VB.CheckBox chkFront 
      Caption         =   "Start on Front sides"
      Height          =   255
      Left            =   3525
      TabIndex        =   1
      Top             =   2445
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3510
      TabIndex        =   4
      Top             =   4230
      Width           =   1755
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   3
      Top             =   3885
      Width           =   1755
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H8000000C&
      Height          =   3345
      Left            =   210
      ScaleHeight     =   219
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   201
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1245
      Width           =   3075
      Begin VB.Image imgTexture 
         Height          =   960
         Left            =   900
         Stretch         =   -1  'True
         Top             =   690
         Width           =   960
      End
   End
   Begin VB.ListBox lstTextures 
      Height          =   1110
      Left            =   3525
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1245
      Width           =   1755
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmAutoalign.frx":000C
      Height          =   855
      Left            =   180
      TabIndex        =   5
      Top             =   165
      Width           =   5265
   End
End
Attribute VB_Name = "frmAutoalign"
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

Private Sub cmdCancel_Click()
     tag = ""
     Hide
End Sub

Private Sub cmdOK_Click()
     tag = "OK"
     Hide
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
     Dim AddedTextures As New Dictionary
     Dim Indices As Variant
     Dim tName As String
     Dim i As Long
     
     'Disable checkboxes
     chkFront.Enabled = False
     chkBack.Enabled = False
     
     'Go for all selected lines
     Indices = selected.Items
     For i = LBound(Indices) To UBound(Indices)
          
          'Check for front side
          If (linedefs(Indices(i)).s1 > -1) Then
               
               'Enable front
               chkFront.Enabled = True
               chkFront.Value = vbChecked
               
               'Add lower textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s1).Lower
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
               
               'Add middle textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s1).Middle
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
               
               'Add upper textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s1).Upper
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
          End If
          
          'Check for back side
          If (linedefs(Indices(i)).s2 > -1) Then
               
               'Enable front
               chkBack.Enabled = True
               chkBack.Value = vbChecked
               
               'Add lower textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s2).Lower
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
               
               'Add middle textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s2).Middle
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
               
               'Add upper textures from sidedef
               tName = sidedefs(linedefs(Indices(i)).s2).Upper
               If (AddedTextures.Exists(tName) = False) And IsTextureName(tName) Then
                    AddedTextures.Add tName, 0
                    lstTextures.AddItem tName
               End If
          End If
     Next i
     
     'Check for textures
     If (lstTextures.ListCount > 0) Then
          
          'Select first
          lstTextures.ListIndex = 0
     Else
          
          'Disable ok
          cmdOK.Enabled = False
     End If
End Sub

Private Sub lstTextures_Click()
     
     'Show texture preview
     GetScaledTexturePictureEx lstTextures.List(lstTextures.ListIndex), imgTexture, picPreview.Width - 20, picPreview.Height - 20
End Sub


Private Sub lstTextures_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Same as clicking
     lstTextures_Click
End Sub


