VERSION 5.00
Begin VB.Form frm3D 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Doom Builder 3D Mode"
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2100
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   HasDC           =   0   'False
   Icon            =   "frm3D.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   116
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   140
End
Attribute VB_Name = "frm3D"
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

Private LastMouseButton As Integer
Private LastMouseShift As Integer
Private LastMouseX As Single, LastMouseY As Single

Private Sub Form_DblClick()
     
     'Redo last mousebutton
     Form_MouseDown LastMouseButton, LastMouseShift, LastMouseX, LastMouseY
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim ShortcutCode As Long
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Ignore shift keys alone
     If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 18) Then Exit Sub
     
     'Make the shortcut code from keycode and shift
     ShortcutCode = KeyCode Or (Shift * (2 ^ 16))
     
     'Check how we should process data
     If TextureSelecting Then
          
          'Perform the action associated with the key
          KeydownTextureSelect ShortcutCode
          
     Else
          
          'Perform the action associated with the key
          Keydown3D ShortcutCode
     End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     
     'Check how we should process data
     If TextureSelecting Then
          
          'Perform the action associated with the key
          KeypressTextureSelect KeyAscii
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     Dim ShortcutCode As Long
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Ignore shift keys alone
     If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 18) Then Exit Sub
     
     'Make the shortcut code from keycode and shift
     ShortcutCode = KeyCode Or (Shift * (2 ^ 16))
     
     'Check how we should process data
     If Not TextureSelecting Then
          
          'Perform the action associated with the key
          Keyrelease3D ShortcutCode
     End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ShortcutCode As Long
     
     'Keep last button and coords
     LastMouseButton = Button
     LastMouseShift = Shift
     LastMouseX = x
     LastMouseY = y
     
     'Make the shortcut code from keycode and shift
     Select Case Button
          Case vbLeftButton: ShortcutCode = MOUSE_BUTTON_0 Or (Shift * (2 ^ 16))
          Case vbMiddleButton: ShortcutCode = MOUSE_BUTTON_2 Or (Shift * (2 ^ 16))
          Case vbRightButton: ShortcutCode = MOUSE_BUTTON_1 Or (Shift * (2 ^ 16))
     End Select
     
     'Check how we should process data
     If TextureSelecting Then
          
          'Perform the action associated with the key
          KeydownTextureSelect ShortcutCode
     Else
          
          'Perform the action associated with the key
          Keydown3D ShortcutCode
     End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     'Keep last coords
     LastMouseX = x
     LastMouseY = y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ShortcutCode As Long
     
     'Keep last coords
     LastMouseX = x
     LastMouseY = y
     
     'Make the shortcut code from keycode and shift
     Select Case Button
          Case vbLeftButton: ShortcutCode = MOUSE_BUTTON_0 Or (Shift * (2 ^ 16))
          Case vbMiddleButton: ShortcutCode = MOUSE_BUTTON_2 Or (Shift * (2 ^ 16))
          Case vbRightButton: ShortcutCode = MOUSE_BUTTON_1 Or (Shift * (2 ^ 16))
     End Select
     
     'Check how we should process data
     If Not TextureSelecting Then
          
          'Perform the action associated with the key
          Keyrelease3D ShortcutCode
     End If
End Sub
