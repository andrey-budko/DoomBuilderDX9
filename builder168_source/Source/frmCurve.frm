VERSION 5.00
Begin VB.Form frmCurve 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Curve Selection"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
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
   Icon            =   "frmCurve.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   128
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkCircleSegment 
      Caption         =   "Force circular segment"
      Height          =   255
      Left            =   870
      TabIndex        =   3
      Top             =   1500
      Width           =   1995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3240
      TabIndex        =   7
      Top             =   2835
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3390
      TabIndex        =   4
      Top             =   180
      Width           =   1155
   End
   Begin DoomBuilder.ctlValueBox txtVertices 
      Height          =   375
      Left            =   1710
      TabIndex        =   0
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   200
      MaxLength       =   4
      Min             =   1
   End
   Begin DoomBuilder.ctlValueBox txtDistance 
      Height          =   375
      Left            =   1710
      TabIndex        =   1
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   10000
      MaxLength       =   6
      Min             =   -10000
      SmallChange     =   8
   End
   Begin DoomBuilder.ctlValueBox txtAngle 
      Height          =   375
      Left            =   1710
      TabIndex        =   2
      Top             =   1020
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Max             =   180
      MaxLength       =   6
      Min             =   1
      SmallChange     =   15
      Value           =   "180"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Delta angle:"
      Height          =   210
      Left            =   765
      TabIndex        =   8
      Top             =   1095
      Width           =   840
   End
   Begin VB.Label lblDistance 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Curve distance:"
      Height          =   210
      Left            =   465
      TabIndex        =   6
      Top             =   675
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Vertices per line:"
      Height          =   210
      Left            =   360
      TabIndex        =   5
      Top             =   255
      Width           =   1230
   End
End
Attribute VB_Name = "frmCurve"
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

Private Sub chkCircleSegment_Click()
     
     'Enabled/Disable available options
     If (chkCircleSegment.Value = vbChecked) Then
          txtDistance.Enabled = False
          lblDistance.ForeColor = vbGrayText
     Else
          txtDistance.Enabled = True
          lblDistance.ForeColor = vbButtonText
     End If
     
     'Remake the curves
     CreateCurve
End Sub

Private Sub cmdCancel_Click()
     
     'Perform undo
     PerformUndo True
     
     'Remove redo (as if nothing happend)
     WithdrawRedo
     
     'Remake selection
     ReselectLinedefs
     
     'Close
     Unload Me
End Sub

Private Sub cmdOK_Click()
     
     'Round vertices
     RoundVertices vertexes(0), numvertexes
     
     'Remake selection
     ReselectLinedefs
     
     'Map changed
     mapchanged = True
     mapnodeschanged = True
     
     'Close
     Unload Me
End Sub

Private Sub CreateCurve()
     Dim Theta As Single
     
     'Check if we have values
     If (Trim$(txtVertices.Text) <> "") And (Trim$(txtDistance.Text) <> "") And (Trim$(txtAngle.Text) <> "") Then
          
          'Perform undo
          PerformUndo True
          
          'Remove redo (as if nothing happend)
          WithdrawRedo
          
          'Make undo
          CreateUndo "curve"
          
          'Limit Theta
          If (Val(txtAngle.Value) < 1) Then
               Theta = 1
          ElseIf (Val(txtAngle.Value) > 180) Then
               Theta = 180
          Else
               Theta = Val(txtAngle.Value)
          End If
          
          'Make the curves
          CurveLines Val(txtVertices.Value), -Val(txtDistance.Value), Theta * pi / 180, (chkCircleSegment.Value = vbChecked)
          
          'Redraw map
          RedrawMap
     End If
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
     Dim lx As Single, ly As Single
     
     'Move to left top of parent
     left = frmMain.left + 50 * Screen.TwipsPerPixelX
     top = frmMain.top + 100 * Screen.TwipsPerPixelY
     
     'Default distance to first linedef's length
     lx = vertexes(linedefs(CLng(selected.Keys(0))).v2).x - vertexes(linedefs(CLng(selected.Keys(0))).v1).x
     ly = vertexes(linedefs(CLng(selected.Keys(0))).v2).y - vertexes(linedefs(CLng(selected.Keys(0))).v1).y
     txtDistance.Value = CLng(Sqr(lx * lx + ly * ly) * 0.5)
     
     'Default vertices
     txtVertices.Value = 8
     
     txtAngle.Value = 180
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Check if cancelling
     If (UnloadMode = 0) Then cmdCancel_Click
End Sub


Private Sub txtAngle_Change()
     CreateCurve
End Sub

Private Sub txtAngle_GotFocus()
     SelectAllText txtAngle
End Sub


Private Sub txtDistance_Change()
     CreateCurve
End Sub

Private Sub txtDistance_GotFocus()
     SelectAllText txtDistance
End Sub


Private Sub txtVertices_Change()
     CreateCurve
End Sub

Private Sub txtVertices_GotFocus()
     SelectAllText txtVertices
End Sub


