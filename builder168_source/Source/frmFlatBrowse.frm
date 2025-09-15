VERSION 5.00
Begin VB.Form frmFlatBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Flat"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
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
   Icon            =   "frmFlatBrowse.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   469
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   658
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      HasDC           =   0   'False
      Height          =   585
      Left            =   0
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   658
      TabIndex        =   2
      Top             =   6450
      Width           =   9870
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
         Height          =   345
         Left            =   8205
         TabIndex        =   3
         Top             =   150
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   -3000
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   1575
   End
   Begin VB.PictureBox picList 
      BackColor       =   &H00000000&
      Height          =   4905
      Left            =   45
      ScaleHeight     =   323
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   540
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   8160
      Begin VB.PictureBox picItem 
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         HasDC           =   0   'False
         Height          =   1260
         Index           =   0
         Left            =   2520
         ScaleHeight     =   84
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   70
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   1050
         Begin VB.Label lblTexture 
            Alignment       =   2  'Center
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   6
            Top             =   1005
            UseMnemonic     =   0   'False
            Width           =   960
         End
         Begin VB.Image imgTexture 
            Height          =   960
            Index           =   0
            Left            =   45
            Stretch         =   -1  'True
            Top             =   45
            Width           =   960
         End
      End
      Begin VB.VScrollBar scrScroll 
         CausesValidation=   0   'False
         Height          =   3405
         LargeChange     =   10
         Left            =   6825
         Max             =   100
         TabIndex        =   4
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmFlatBrowse"
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


'Selected texture (0 = the dash!)
Public SelectedIndex As Long
Public SelectedName As String

Private Const BORDERSPACING As Long = 3
Private FlatNames() As String
Private NumFlats As Long

Private Rows As Long
Private Cols As Long

Public OriginalMessageHandler As Long

Private Sub ArrangeBoxes()
     Dim i As Long
     Dim x As Long
     Dim y As Long
     Dim bx As Long
     Dim by As Long
     Dim bwidth As Long
     Dim bheight As Long
     
     'Calculate total width and height of blocks
     bwidth = picItem(0).Width
     bheight = picItem(0).Height
     
     'Determine number of boxes in width and height
     bx = (picList.ScaleWidth - scrScroll.Width) \ bwidth
     by = picList.ScaleHeight \ bheight
     Rows = by
     Cols = bx
     
     'Go for all boxes
     i = 0
     For y = 0 To (by - 1)
          For x = 0 To (bx - 1)
               
               'Load controls
               On Local Error Resume Next
               Load picItem(i)
               Load imgTexture(i)
               Load lblTexture(i)
               On Local Error GoTo 0
               
               'Position controls
               picItem(i).Move bwidth * x, bheight * y
               Set imgTexture(i).Container = picItem(i)
               Set lblTexture(i).Container = picItem(i)
               lblTexture(i).Move 3, 67
               imgTexture(i).Move BORDERSPACING, BORDERSPACING
               lblTexture(i).Visible = True
               imgTexture(i).Visible = True
               
               'Next control
               i = i + 1
          Next x
     Next y
End Sub

Private Sub cmdCancel_Click()
     tag = 0
     Hide
End Sub

Private Sub cmdSelect_Click()
     tag = 1
     Hide
End Sub

Public Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     Dim ScrollV As Long
     Dim OldSelected As Long
     Dim ci As Long
     Dim i As Long
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Keep old selection
     OldSelected = SelectedIndex
     
     'Check what key is pressed
     Select Case KeyCode
          
          Case 107, 187    '+
               
               'Scroll up
               If (scrScroll.Value - 2 >= scrScroll.Min) Then
                    scrScroll.Value = scrScroll.Value - 2
               Else
                    scrScroll.Value = scrScroll.Min
               End If
               Exit Sub
               
          Case 109, 189  '-
               
               'Scroll down
               If (scrScroll.Value + 2 <= scrScroll.Max) Then
                    scrScroll.Value = scrScroll.Value + 2
               Else
                    scrScroll.Value = scrScroll.Max
               End If
               Exit Sub
               
          Case vbKeyPageUp
               
               SelectedIndex = SelectedIndex - Cols * Rows
               If (SelectedIndex < 0) Then SelectedIndex = 0
               
          Case vbKeyPageDown
               
               SelectedIndex = SelectedIndex + Cols * Rows
               If (SelectedIndex >= NumFlats) Then SelectedIndex = NumFlats - 1
               
          Case vbKeyHome
               
               SelectedIndex = 0
               
          Case vbKeyEnd
               
               SelectedIndex = NumFlats - 1
               
          Case vbKeyUp
               
               SelectedIndex = SelectedIndex - Cols
               If (SelectedIndex < 0) Then SelectedIndex = 0
               
          Case vbKeyDown
               
               SelectedIndex = SelectedIndex + Cols
               If (SelectedIndex >= NumFlats) Then SelectedIndex = NumFlats - 1
               
          Case vbKeyLeft
               
               SelectedIndex = SelectedIndex - 1
               If (SelectedIndex < 0) Then SelectedIndex = 0
               
          Case vbKeyRight
               
               SelectedIndex = SelectedIndex + 1
               If (SelectedIndex >= NumFlats) Then SelectedIndex = NumFlats - 1
               
          Case Else
               
               'Check if we can jump to a flat
               'Go for all flat names
               For i = 0 To (flats.Count - 1)
                    
                    'Check if the name starts with this char
                    If (StrComp(left$(flats.Keys(i), 1), Chr$(KeyCode), vbTextCompare) = 0) Then
                         
                         'Select this texture
                         SelectedIndex = i
                         Exit For
                    End If
               Next i
               
     End Select
     
     'Select texture name
     SelectedName = FlatNames(SelectedIndex)
     
     'Check if old selection is within view
     ci = OldSelected - scrScroll.Value * Cols
     If ((ci >= picItem.LBound) And (ci <= picItem.UBound)) Then
          
          'Deseect old
          picItem(ci).BackColor = vbBlack    'vbWindowBackground
          lblTexture(ci).BackColor = vbBlack 'vbWindowBackground
          lblTexture(ci).ForeColor = vbWhite 'vbWindowText
     End If
     
     'Check if the selection is above view
     ci = SelectedIndex - scrScroll.Value * Cols
     If (ci < picItem.LBound) Then
          
          'Scroll to selection
          ScrollV = (SelectedIndex \ Cols)
          If (ScrollV > scrScroll.Max) Then ScrollV = scrScroll.Max
          If (ScrollV < scrScroll.Min) Then ScrollV = scrScroll.Min
          scrScroll.Value = ScrollV
          
     'Check if the selection is below view
     ElseIf (ci > picItem.UBound) Then
          
          'Scroll to selection
          ScrollV = (SelectedIndex \ Cols) - (Rows - 1)
          If (ScrollV > scrScroll.Max) Then ScrollV = scrScroll.Max
          If (ScrollV < scrScroll.Min) Then ScrollV = scrScroll.Min
          scrScroll.Value = ScrollV
          
     'Otherwise the selection is inside view
     Else
          
          'Select new
          picItem(ci).BackColor = vbHighlight
          lblTexture(ci).BackColor = vbHighlight
          lblTexture(ci).ForeColor = vbHighlightText
     End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     Dim i As Long
     Dim Keys As Variant
     
     'None selected
     SelectedIndex = -1
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Keep original messages handler
          OriginalMessageHandler = GetWindowLong(Me.hWnd, GWL_WNDPROC)
          
          'Set our own messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FlatMessageHandler
     End If
     
     'Get the texture names
     Keys = flats.Keys
     
     'Allocate memory for string names
     NumFlats = UBound(Keys) - LBound(Keys)
     ReDim FlatNames(0 To NumFlats)
     
     'Make string array from texture names
     For i = LBound(Keys) To UBound(Keys)
          FlatNames(i) = Keys(i)
     Next i
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Restore original messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, OriginalMessageHandler
     End If
End Sub

Private Sub Form_Resize()
     Dim ScrollMax As Long
     Dim ScrollV As Long
     
     'Resize list
     picList.Width = ScaleWidth - picList.left * 2
     picList.Height = ScaleHeight - picBottom.Height - picList.top
     
     'Reposition scrollbar
     scrScroll.left = picList.ScaleWidth - scrScroll.Width
     scrScroll.Height = picList.ScaleHeight
     
     'Rearrange controls
     ArrangeBoxes
     
     'Set the scrollbar max
     ScrollMax = (NumFlats \ Cols) + 1 - Rows
     If (ScrollMax < 0) Then ScrollMax = 0
     scrScroll.Max = ScrollMax
     scrScroll.LargeChange = Rows
     
     'Determine scroll position to show selection
     ScrollV = (SelectedIndex \ Cols) - 2
     If (ScrollV < 0) Then ScrollV = 0
     If (ScrollV > ScrollMax) Then ScrollV = ScrollMax
     scrScroll.Value = ScrollV
     
     'Fill the controls with flats
     ShowFlats
End Sub

Private Sub imgTexture_DblClick(Index As Integer)
     cmdSelect_Click
End Sub

Private Sub imgTexture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     picItem_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub lblTexture_DblClick(Index As Integer)
     cmdSelect_Click
End Sub

Private Sub lblTexture_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     picItem_MouseDown Index, Button, Shift, x, y
End Sub

Private Sub picItem_DblClick(Index As Integer)
     cmdSelect_Click
End Sub

Private Sub picItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim ci As Long
     
     'Check if old selection is within view
     ci = SelectedIndex - scrScroll.Value * Cols
     If ((ci >= picItem.LBound) And (ci <= picItem.UBound)) Then
          
          'Deseect old
          picItem(ci).BackColor = vbBlack    'vbWindowBackground
          lblTexture(ci).BackColor = vbBlack 'vbWindowBackground
          lblTexture(ci).ForeColor = vbWhite 'vbWindowText
     End If
     
     'Select this texture
     SelectedIndex = Index + scrScroll.Value * Cols
     SelectedName = FlatNames(SelectedIndex)
     
     'Selection color
     picItem(Index).BackColor = vbHighlight
     lblTexture(Index).BackColor = vbHighlight
     lblTexture(Index).ForeColor = vbHighlightText
     
     'Focus to picturebox
     On Error Resume Next
     picList.SetFocus
End Sub

Private Sub scrScroll_Change()
     
     'Refill controls
     ShowFlats
     DoEvents
End Sub

Private Sub scrScroll_Scroll()
     
     'Refill controls
     ShowFlats
     DoEvents
End Sub

Public Sub SetTextureSelection(ByVal FlatName As String)
     Dim i As Long
     
     'Go for all Flat names
     For i = LBound(FlatNames) To UBound(FlatNames)
          
          'Select if name matches
          If (StrComp(FlatNames(i), FlatName, vbTextCompare) = 0) Then
               
               'This Flat is now selected
               SelectedIndex = i
               SelectedName = FlatNames(i)
               Exit For
          End If
     Next i
End Sub

Private Sub ShowFlats()
     Dim Shown As Long
     Dim offset As Long
     Dim Flat As clsImage
     Dim w As Long, h As Long
     Dim x As Long, y As Long
     Dim i As Long
     Dim ci As Long
     
     'Hide list, this solves flickering problems
     'picList.Visible = False
     LockWindowUpdate Me.hWnd
     
     'Calculate number of flats we can show
     Shown = Cols * Rows
     
     'Calculate index offset
     offset = scrScroll.Value * Cols
     
     'Go for all flats to be shown
     For i = offset To (offset + Shown - 1)
          
          'Get control index
          ci = i - offset
          
          'Determine x an y
          y = ci \ Cols
          x = ci - y * Cols
          
          'Check if within bounds
          If (i <= NumFlats) Then
               
               'Clear picture
               Set imgTexture(ci).Picture = Nothing
               
               'Set flat name
               lblTexture(ci).Caption = FlatNames(i)
               
               'Do not crash here
               On Error Resume Next
               
               'Get flat object
               Set Flat = flats(FlatNames(i))
               
               'Set picture
               Set imgTexture(ci).Picture = Flat.Picture
               
               'Position
               Flat.GetScale 64, 64, w, h, False
               imgTexture(ci).Move BORDERSPACING + (64 - w) \ 2, BORDERSPACING + (64 - h) \ 2, w, h
               
               'Continue error handling
               On Error GoTo 0
               
               'Check if selected
               If (SelectedIndex = i) Then
                    
                    'Selection color
                    picItem(ci).BackColor = vbHighlight
                    lblTexture(ci).BackColor = vbHighlight
                    lblTexture(ci).ForeColor = vbHighlightText
               Else
                    
                    'Normal color
                    picItem(ci).BackColor = vbBlack    'vbWindowBackground
                    lblTexture(ci).BackColor = vbBlack 'vbWindowBackground
                    lblTexture(ci).ForeColor = vbWhite 'vbWindowText
               End If
               
               'Show flat
               picItem(ci).Visible = True
          Else
               
               'Clear flat
               Set imgTexture(ci).Picture = Nothing
               picItem(ci).Visible = False
          End If
     Next i
     
     'Show list
     LockWindowUpdate 0
     'picList.Visible = True
     'picList.Refresh
End Sub
