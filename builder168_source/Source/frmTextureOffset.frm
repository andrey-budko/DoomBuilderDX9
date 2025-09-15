VERSION 5.00
Begin VB.Form frmTextureOffset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Texture Offset"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
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
   Icon            =   "frmTextureOffset.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   421
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   408
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picLower 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      HasDC           =   0   'False
      Height          =   1980
      Left            =   90
      MousePointer    =   15  'Size All
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4245
      Width           =   3900
   End
   Begin VB.PictureBox picMiddle 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      HasDC           =   0   'False
      Height          =   1980
      Left            =   90
      MousePointer    =   15  'Size All
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2175
      Width           =   3900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4170
      TabIndex        =   2
      Top             =   525
      Width           =   1845
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4170
      TabIndex        =   1
      Top             =   105
      Width           =   1845
   End
   Begin VB.PictureBox picUpper 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      HasDC           =   0   'False
      Height          =   1980
      Left            =   90
      MousePointer    =   15  'Size All
      ScaleHeight     =   128
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   3900
   End
End
Attribute VB_Name = "frmTextureOffset"
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


Public offsetx As Long
Public offsety As Long

Private UpperTexture As clsImage
Private MiddleTexture As clsImage
Private LowerTexture As clsImage
Private UpperBitmap As StdPicture
Private MiddleBitmap As StdPicture
Private LowerBitmap As StdPicture

Private GrabOffsetX As Long
Private GrabOffsetY As Long

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
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     Set UpperBitmap = Nothing
     Set UpperTexture = Nothing
     Set MiddleBitmap = Nothing
     Set MiddleTexture = Nothing
     Set LowerBitmap = Nothing
     Set LowerTexture = Nothing
End Sub

Public Sub Init(ByVal Upper As String, ByVal Middle As String, ByVal Lower As String, ByVal StartX As Long, ByVal StartY As Long)
     
     'Check the texture can be found
     If (textures.Exists(Upper) = True) Then
          
          'Load texture bitmap
          Set UpperTexture = textures(Upper)
          Set UpperBitmap = UpperTexture.Picture(True)
     End If
     
     'Check the texture can be found
     If (textures.Exists(Middle) = True) Then
          
          'Load texture bitmap
          Set MiddleTexture = textures(Middle)
          Set MiddleBitmap = MiddleTexture.Picture(True)
     End If
     
     'Check the texture can be found
     If (textures.Exists(Lower) = True) Then
          
          'Load texture bitmap
          Set LowerTexture = textures(Lower)
          Set LowerBitmap = LowerTexture.Picture(True)
     End If
     
     'Set start X and Y
     offsetx = StartX
     offsety = StartY
End Sub

Private Sub picLower_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     picUpper_MouseDown Button, Shift, x, y
End Sub

Private Sub picLower_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     picUpper_MouseMove Button, Shift, x, y
End Sub

Private Sub picLower_Paint()
     RenderLower
End Sub

Private Sub picMiddle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     picUpper_MouseDown Button, Shift, x, y
End Sub

Private Sub picMiddle_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     picUpper_MouseMove Button, Shift, x, y
End Sub

Private Sub picMiddle_Paint()
     RenderMiddle
End Sub

Private Sub picUpper_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     'Only drag when left button is hold
     If (Button = vbLeftButton) Then
          
          'Set the offsets where the texture was grabbed
          GrabOffsetX = x
          GrabOffsetY = y
     End If
End Sub

Private Sub picUpper_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     
     'Only drag when left button is hold
     If (Button = vbLeftButton) Then
          
          'Change offsets
          offsetx = offsetx + (x - GrabOffsetX)
          offsety = offsety + (y - GrabOffsetY)
          
          'Set the offsets where the texture was grabbed
          GrabOffsetX = x
          GrabOffsetY = y
          
          'Render panels
          RenderPanels
     End If
End Sub

Private Sub picUpper_Paint()
     RenderUpper
End Sub

Private Sub RenderLower()
     Dim x As Long
     Dim y As Long
     
     'Render lower panel
     If Not (LowerTexture Is Nothing) And Not (LowerBitmap Is Nothing) Then
          For x = (0 - LowerTexture.width) To (picLower.ScaleWidth + LowerTexture.width) Step LowerTexture.width
               For y = (0 - LowerTexture.height) To (picLower.ScaleHeight + LowerTexture.height) Step LowerTexture.height
                    
                    'Render texture
                    picLower.PaintPicture LowerBitmap, x + (offsetx Mod LowerTexture.width), y + (offsety Mod LowerTexture.height)
               Next y
          Next x
     End If
End Sub

Private Sub RenderMiddle()
     Dim x As Long
     Dim y As Long
     
     'Render middle panel
     If Not (MiddleTexture Is Nothing) And Not (MiddleBitmap Is Nothing) Then
          For x = (0 - MiddleTexture.width) To (picMiddle.ScaleWidth + MiddleTexture.width) Step MiddleTexture.width
               For y = (0 - MiddleTexture.height) To (picMiddle.ScaleHeight + MiddleTexture.height) Step MiddleTexture.height
                    
                    'Render texture
                    picMiddle.PaintPicture MiddleBitmap, x + (offsetx Mod MiddleTexture.width), y + (offsety Mod MiddleTexture.height)
               Next y
          Next x
     End If
End Sub

Private Sub RenderPanels()
     
     'Render upper panel
     RenderUpper
     
     'Render middle panel
     RenderMiddle
     
     'Render lower panel
     RenderLower
End Sub

Private Sub RenderUpper()
     Dim x As Long
     Dim y As Long
     
     'Render upper panel
     If Not (UpperTexture Is Nothing) And Not (UpperBitmap Is Nothing) Then
          For x = (0 - UpperTexture.width) To (picUpper.ScaleWidth + UpperTexture.width) Step UpperTexture.width
               For y = (0 - UpperTexture.height) To (picUpper.ScaleHeight + UpperTexture.height) Step UpperTexture.height
                    
                    'Render texture
                    picUpper.PaintPicture UpperBitmap, x + (offsetx Mod UpperTexture.width), y + (offsety Mod UpperTexture.height)
               Next y
          Next x
     End If
End Sub
