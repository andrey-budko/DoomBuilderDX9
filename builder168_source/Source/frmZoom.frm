VERSION 5.00
Begin VB.Form frmZoom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Zoom"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4350
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
   Icon            =   "frmZoom.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   2213
      TabIndex        =   2
      Top             =   990
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   593
      TabIndex        =   1
      Top             =   990
      Width           =   1545
   End
   Begin VB.ComboBox cmbZoom 
      Height          =   330
      ItemData        =   "frmZoom.frx":000C
      Left            =   2745
      List            =   "frmZoom.frx":001F
      TabIndex        =   0
      Text            =   "100"
      Top             =   330
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "%"
      Height          =   210
      Left            =   3825
      TabIndex        =   4
      Top             =   360
      UseMnemonic     =   0   'False
      Width           =   150
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter or select a Zoom level:"
      Height          =   210
      Left            =   525
      TabIndex        =   3
      Top             =   375
      UseMnemonic     =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "frmZoom"
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

Private Sub cmbZoom_GotFocus()
     SelectAllText cmbZoom
End Sub

Private Sub cmbZoom_KeyPress(KeyAscii As Integer)
     If ((KeyAscii <> 8) And (KeyAscii <> 46) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdOK_Click()
     Dim Xdiff As Long, Ydiff As Long
     Dim NewZ As Single
     
     'Calculate new zoom
     NewZ = CSng(cmbZoom) * 0.01
     If NewZ > 100 Then NewZ = 100
     If NewZ < 0.05 Then NewZ = 0.05
     Xdiff = ((ScreenWidth / NewZ) - (ScreenWidth / ViewZoom)) / 2
     Ydiff = ((ScreenHeight / NewZ) - (ScreenHeight / ViewZoom)) / 2
     
     'Check view
     ChangeView ViewLeft - Xdiff, ViewTop - Ydiff, NewZ
     
     'Redraw map
     RedrawMap
     
     'Update stauts bar
     UpdateStatusBar
     
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
     cmbZoom = Format$(ViewZoom * 100, "0.##")
End Sub
