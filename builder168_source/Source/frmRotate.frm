VERSION 5.00
Begin VB.Form frmRotate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rotate Selection"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
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
   Icon            =   "frmRotate.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3045
      TabIndex        =   3
      Top             =   945
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3195
      TabIndex        =   1
      Top             =   210
      Width           =   1155
   End
   Begin DoomBuilder.ctlValueBox txtAngle 
      Height          =   375
      Left            =   1515
      TabIndex        =   0
      Top             =   195
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   360
      MaxLength       =   4
      Min             =   -360
      SmallChange     =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rotation angle:"
      Height          =   210
      Left            =   345
      TabIndex        =   2
      Top             =   270
      Width           =   1065
   End
End
Attribute VB_Name = "frmRotate"
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
     
     'Perform undo
     PerformUndo True
     
     'Remove redo (as if nothing happend)
     WithdrawRedo
     
     'Close
     Unload Me
End Sub

Private Sub cmdOK_Click()
     
     'When used during pasting, remove the undo
     If (submode = ESM_PASTING) Then WithdrawUndo
     
     'Round vertices
     RoundVertices vertexes(0), numvertexes
     
     'Map changed
     mapchanged = True
     If (mode <> EM_THINGS) Then mapnodeschanged = True
     
     'Close
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
     
     'Move to left top of parent
     left = frmMain.left + 50 * Screen.TwipsPerPixelX
     top = frmMain.top + 100 * Screen.TwipsPerPixelY
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Check if cancelling
     If (UnloadMode = 0) Then cmdCancel_Click
End Sub

Private Sub txtAngle_Change()
     
     'Check if we have a value to rotate by
     If (Trim$(txtAngle.Text) <> "") Then
          
          'Perform undo
          PerformUndo True
          
          'Remove redo (as if nothing happend)
          WithdrawRedo
          
          'Make undo
          CreateUndo "rotate"
          
          'Rotate with the change difference
          If (mode = EM_THINGS) Then
               RotateThings -Val(txtAngle.Value) / PiDiv
          Else
               RotateVertices -Val(txtAngle.Value) / PiDiv
          End If
          
          'Redraw map
          RedrawMap
     End If
End Sub

Private Sub txtAngle_GotFocus()
     SelectAllText txtAngle
End Sub


