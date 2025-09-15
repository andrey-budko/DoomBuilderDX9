VERSION 5.00
Begin VB.Form frmGrid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Grid"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
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
   Icon            =   "frmGrid.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   167
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   352
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbGridsizeY 
      Height          =   330
      ItemData        =   "frmGrid.frx":000C
      Left            =   1410
      List            =   "frmGrid.frx":0025
      TabIndex        =   1
      Text            =   "100"
      Top             =   1200
      Width           =   1035
   End
   Begin DoomBuilder.ctlValueBox txtGridX 
      Height          =   360
      Left            =   3930
      TabIndex        =   2
      Top             =   750
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   635
      Max             =   9999
      Min             =   -9999
   End
   Begin VB.CheckBox chkShowGrid 
      Caption         =   "Show grid lines"
      Height          =   240
      Left            =   435
      TabIndex        =   6
      Top             =   255
      Width           =   1875
   End
   Begin VB.ComboBox cmbGridsizeX 
      Height          =   330
      ItemData        =   "frmGrid.frx":0045
      Left            =   1410
      List            =   "frmGrid.frx":005E
      TabIndex        =   0
      Text            =   "100"
      Top             =   780
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   1032
      TabIndex        =   4
      Top             =   2010
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2704
      TabIndex        =   5
      Top             =   2010
      Width           =   1545
   End
   Begin DoomBuilder.ctlValueBox txtGridY 
      Height          =   360
      Left            =   3930
      TabIndex        =   3
      Top             =   1170
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   635
      Max             =   9999
      Min             =   -9999
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Grid size in Y:"
      Height          =   210
      Left            =   285
      TabIndex        =   10
      Top             =   1245
      UseMnemonic     =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grid offset Y:"
      Height          =   210
      Left            =   2820
      TabIndex        =   9
      Top             =   1245
      Width           =   990
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Grid offset X:"
      Height          =   210
      Left            =   2835
      TabIndex        =   8
      Top             =   825
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Grid size in X:"
      Height          =   210
      Left            =   285
      TabIndex        =   7
      Top             =   825
      UseMnemonic     =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "frmGrid"
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

Private Sub cmbGridsizeX_GotFocus()
     SelectAllText cmbGridsizeX
End Sub

Private Sub cmbGridsizeX_KeyPress(KeyAscii As Integer)
     If ((KeyAscii <> 8) And (KeyAscii <> 46) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
End Sub

Private Sub cmbGridsizeY_GotFocus()
     SelectAllText cmbGridsizeY
End Sub

Private Sub cmbGridsizeY_KeyPress(KeyAscii As Integer)
     If ((KeyAscii <> 8) And (KeyAscii <> 46) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdOK_Click()
     
     'Change grid size
     gridsizex = Val(cmbGridsizeX.Text)
     gridsizey = Val(cmbGridsizeY.Text)
     gridx = Val(txtGridX.Value)
     gridy = Val(txtGridY.Value)
     Config("gridshow") = chkShowGrid.Value
     
     'Limits
     If (gridsizex < 2) Then gridsizex = 2
     If (gridsizex > 1024) Then gridsizex = 1024
     
     'Update status
     UpdateStatusBar
     
     'Redraw map
     RedrawMap
     
     'Leave here
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
     chkShowGrid.Value = Config("gridshow")
     cmbGridsizeX.Text = gridsizex
     cmbGridsizeY.Text = gridsizey
     txtGridX.Value = gridx
     txtGridY.Value = gridy
End Sub

Private Sub txtGridX_GotFocus()
     SelectAllText txtGridX
End Sub


Private Sub txtGridY_GotFocus()
     SelectAllText txtGridY
End Sub


