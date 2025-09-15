VERSION 5.00
Begin VB.Form frmMakeSector 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Sector"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
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
   Icon            =   "frmMakeSector.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   173
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkSnap 
      Caption         =   "Snap vertices to the grid"
      Height          =   255
      Left            =   2955
      TabIndex        =   4
      Top             =   930
      Width           =   2235
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2805
      TabIndex        =   6
      Top             =   2115
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Default         =   -1  'True
      Height          =   345
      Left            =   1110
      TabIndex        =   5
      Top             =   2115
      Width           =   1575
   End
   Begin VB.VScrollBar scrDiameter 
      Height          =   360
      LargeChange     =   8
      Left            =   2415
      Max             =   2
      Min             =   9999
      SmallChange     =   8
      TabIndex        =   3
      Top             =   1305
      Value           =   32
      Width           =   240
   End
   Begin VB.TextBox txtDiameter 
      Height          =   315
      Left            =   1425
      MaxLength       =   4
      TabIndex        =   2
      Text            =   "32"
      Top             =   1335
      Width           =   975
   End
   Begin VB.VScrollBar scrVertices 
      Height          =   360
      Left            =   2415
      Max             =   3
      Min             =   9999
      TabIndex        =   1
      Top             =   855
      Value           =   4
      Width           =   240
   End
   Begin VB.TextBox txtVertices 
      Height          =   315
      Left            =   1425
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "4"
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMakeSector.frx":000C
      Height          =   435
      Left            =   135
      TabIndex        =   9
      Top             =   135
      Width           =   5265
   End
   Begin VB.Label lblDiameter 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Radius:"
      Height          =   210
      Left            =   705
      TabIndex        =   8
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label lblVertices 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vertices:"
      Height          =   210
      Left            =   585
      TabIndex        =   7
      Top             =   930
      UseMnemonic     =   0   'False
      Width           =   660
   End
End
Attribute VB_Name = "frmMakeSector"
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

Private Sub cmdCreate_Click()
     
     'Validate both fields
     txtVertices_Validate False
     txtDiameter_Validate False
     
     tag = 1
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


Private Sub scrDiameter_Change()
     txtDiameter.Text = scrDiameter.Value
End Sub

Private Sub scrVertices_Change()
     txtVertices.Text = scrVertices.Value
End Sub

Private Sub txtDiameter_GotFocus()
     SelectAllText txtDiameter
End Sub

Private Sub txtDiameter_KeyPress(KeyAscii As Integer)
     If ((KeyAscii <> 8) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
End Sub

Private Sub txtDiameter_Validate(Cancel As Boolean)
     On Local Error Resume Next
     scrDiameter.Value = Val(txtDiameter.Text)
     txtDiameter.Text = scrDiameter.Value
End Sub

Private Sub txtVertices_GotFocus()
     SelectAllText txtVertices
End Sub

Private Sub txtVertices_KeyPress(KeyAscii As Integer)
     If ((KeyAscii <> 8) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
End Sub

Private Sub txtVertices_Validate(Cancel As Boolean)
     On Local Error Resume Next
     scrVertices.Value = Val(txtVertices.Text)
     txtVertices.Text = scrVertices.Value
End Sub
