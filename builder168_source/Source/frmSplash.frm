VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
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
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   200
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWebsite 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "www.codeimp.com"
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   1440
      MouseIcon       =   "frmSplash.frx":7840
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2460
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblAbout1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Pascal vd Heiden"
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   30
      TabIndex        =   2
      Top             =   2250
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CodeImp Doom Builder version 0.00 build 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   30
      TabIndex        =   1
      Top             =   2040
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   4755
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Initializing..."
      ForeColor       =   &H00C0C0C0&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   2550
      UseMnemonic     =   0   'False
      Width           =   4680
   End
End
Attribute VB_Name = "frmSplash"
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

Private Sub Form_Deactivate()
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     
     'Draw the splash
     'AutoRedraw = True
     'PaintPicture Picture, 0, 0, ScaleWidth, ScaleHeight
     'AutoRedraw = False
     
     'Position the status label
     lblStatus.Move (ScaleWidth - lblStatus.width) \ 2, (ScaleHeight - lblStatus.height) / 1.1
     
     'Splash screen is displayed
     SplashDisplayed = True
End Sub

Private Sub Form_LostFocus()
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Splash screen is gone
     SplashDisplayed = False
End Sub

Private Sub Form_Resize()
     
     'Draw the splash
     AutoRedraw = True
     PaintPicture Picture, 0, 0, ScaleWidth, ScaleHeight
     AutoRedraw = False
     
     'Position the status label
     lblStatus.Move (ScaleWidth - lblStatus.width) \ 2, (ScaleHeight - lblStatus.height) / 1.1
End Sub

Private Sub lblAbout1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub lblVersion_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     'When not in loading process, unload this dialog
     If Not Loading Then Unload Me: Set frmSplash = Nothing
End Sub

Private Sub lblWebsite_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     'When not in loading process, show website and unload this dialog
     If Not Loading Then
          
          'Change mousepointer
          Screen.MousePointer = vbHourglass
          
          'Hide dialog
          Hide
          DoEvents
          
          'Go to xode multimedia website
          Execute "http://www.codeimp.com", "", SW_SHOW, False
          
          'Change mousepointer
          Screen.MousePointer = vbNormal
          
          'Unload dialog
          Unload Me: Set frmSplash = Nothing
     End If
End Sub
