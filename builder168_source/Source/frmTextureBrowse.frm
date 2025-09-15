VERSION 5.00
Begin VB.Form frmTextureBrowse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Texture"
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
   Icon            =   "frmTextureBrowse.frx":0000
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
      TabStop         =   0   'False
      Top             =   6450
      Width           =   9870
      Begin VB.CommandButton cmdSelect 
         Caption         =   "Select"
         Default         =   -1  'True
         Height          =   345
         Left            =   8205
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   150
         Width           =   1575
      End
      Begin VB.Label lblViewSort 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Viewing used textures only. Press TAB to view all textures."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3135
         TabIndex        =   9
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   4905
      End
      Begin VB.Label lblTextureSize 
         Height          =   210
         Left            =   1380
         TabIndex        =   5
         ToolTipText     =   "Width and Height of the selected texture"
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   1200
      End
      Begin VB.Label lblTextureName 
         Height          =   210
         Left            =   300
         TabIndex        =   4
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   990
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
         TabIndex        =   7
         TabStop         =   0   'False
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
            TabIndex        =   8
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
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmTextureBrowse"
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


'Dictionary being browsed
Public collection As Dictionary

'Selected texture (0 = the dash!)
Public selectedindex As Long
Public SelectedName As String
Public ShowAll As Boolean

Private Const BORDERSPACING As Long = 3
Private itemnames() As String
Private useditemnames() As String
Private numitems As Long
Private numuseditems As Long
Private curitemnames() As String
Private curnumitems As Long

Private Rows As Long
Private cols As Long

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
     bwidth = picItem(0).width
     bheight = picItem(0).height
     
     'Determine number of boxes in width and height
     bx = (picList.ScaleWidth - scrScroll.width) \ bwidth
     by = picList.ScaleHeight \ bheight
     Rows = by
     cols = bx
     
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
               lblTexture(i).visible = True
               imgTexture(i).visible = True
               
               'Next control
               i = i + 1
          Next x
     Next y
End Sub

Public Sub Initialize(ByVal browseflats As Boolean)
     Dim useditems As New Dictionary
     Dim Keys As Variant
     Dim starti As Long
     Dim i As Long
     Dim ScrollMax As Long
     
     'None selected
     selectedindex = -1
     
     'Check if using flats or textures
     If (browseflats) Then
          
          'Set information for Flats
          Set collection = flats
          numitems = collection.Count
          Caption = "Select Flat"
          lblViewSort.Caption = "Viewing used flats only. Press TAB to view all flats."
     Else
          
          'Set information for Textures
          Set collection = textures
          numitems = collection.Count + 1    '1 extra for the -
          Caption = "Select Texture"
          lblViewSort.Caption = "Viewing used textures only. Press TAB to view all textures."
     End If
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Keep original messages handler
          OriginalMessageHandler = GetWindowLong(Me.hWnd, GWL_WNDPROC)
          
          'Set our own messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf TextureMessageHandler
     End If
     
     'Get the key names
     Keys = collection.Keys
     
     'Allocate memory for string names
     ReDim itemnames(0 To numitems - 1)
     
     'First texture is nothing
     If (browseflats = False) Then
          itemnames(0) = "-"
          starti = 1
     Else
          starti = 0
     End If
     
     'Make string array from names
     For i = starti To numitems - 1
          itemnames(i) = Keys(i - starti)
     Next i
     
     'Check if we should select used names from sidedefs (textures)
     If (browseflats = False) Or (Val(mapconfig("mixtexturesflats")) = vbChecked) Then
          
          'Go for all sidedefs
          For i = 0 To numsidedefs - 1
               If (useditems.Exists(sidedefs(i).Upper) = False) Then If (collection.Exists(sidedefs(i).Upper)) Then useditems.Add sidedefs(i).Upper, 1
               If (useditems.Exists(sidedefs(i).Middle) = False) Then If (collection.Exists(sidedefs(i).Middle)) Then useditems.Add sidedefs(i).Middle, 1
               If (useditems.Exists(sidedefs(i).Lower) = False) Then If (collection.Exists(sidedefs(i).Lower)) Then useditems.Add sidedefs(i).Lower, 1
          Next i
     End If
     
     'Check if we should select used names from sectors (flats)
     If (browseflats = True) Or (Val(mapconfig("mixtexturesflats")) = vbChecked) Then
          
          'Go for all sector
          For i = 0 To numsectors - 1
               If (useditems.Exists(sectors(i).tfloor) = False) Then If (collection.Exists(sectors(i).tfloor)) Then useditems.Add sectors(i).tfloor, 1
               If (useditems.Exists(sectors(i).tceiling) = False) Then If (collection.Exists(sectors(i).tceiling)) Then useditems.Add sectors(i).tceiling, 1
          Next i
     End If
     
     'Are there used items?
     If (useditems.Count > 0) And (Val(Config("alwaysalltextures")) = vbUnchecked) Then
          
          'Sort used items
          Set useditems = SortDictionary(useditems)
          
          'When using textures, add 1 for the -
          If (browseflats) Then numuseditems = useditems.Count Else numuseditems = useditems.Count + 1
          
          'Allocate memory for string names
          ReDim useditemnames(0 To numuseditems - 1)
          Keys = useditems.Keys
          
          'First texture is nothing
          If (browseflats = False) Then
               useditemnames(0) = "-"
               starti = 1
          Else
               starti = 0
          End If
          
          'Make string array from texture names
          For i = starti To numuseditems - 1
               
               'Add to array
               useditemnames(i) = Keys(i - starti)
          Next i
          
          'Set the current collection
          curitemnames() = useditemnames()
          curnumitems = numuseditems
     Else
          
          'Show all textures
          ShowAll = True
          lblViewSort.visible = False
          curitemnames() = itemnames()
          curnumitems = numitems
     End If
     
     'Resize list
     picList.width = ScaleWidth - picList.left * 2
     picList.height = ScaleHeight - picBottom.height - picList.top
     
     'Reposition scrollbar
     scrScroll.left = picList.ScaleWidth - scrScroll.width
     scrScroll.height = picList.ScaleHeight
     
     'Rearrange controls
     ArrangeBoxes
     
     'Set the scrollbar max
     ScrollMax = (curnumitems \ cols) + 1 - Rows
     If (ScrollMax < 0) Then ScrollMax = 0
     scrScroll.Max = ScrollMax
     scrScroll.LargeChange = Rows
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
     Dim ScrollMax As Long
     Dim OldSelected As Long
     Dim thisimage As clsImage
     Dim ci As Long
     Dim i As Long
     
     'Adjust shift mask
     CurrentShiftMask = Shift
     
     'Keep old selection
     OldSelected = selectedindex
     
     'Check what key is pressed
     Select Case KeyCode
          
          Case vbKeyTab
               
               'Check if hiding unused items
               If (ShowAll = False) Then
                    
                    'Switch to all items
                    lblViewSort.visible = False
                    ShowAll = True
                    
                    'Switch collections
                    curitemnames() = itemnames()
                    curnumitems = numitems
                    
                    'Set the scrollbar max
                    ScrollMax = (curnumitems \ cols) + 1 - Rows
                    If (ScrollMax < 0) Then ScrollMax = 0
                    scrScroll.Max = ScrollMax
                    scrScroll.LargeChange = Rows
                    
                    'Select same texture
                    SetSelection SelectedName
                    
                    'Show items
                    ShowItems
               End If
          
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
               
               selectedindex = selectedindex - cols * Rows
               If (selectedindex < 0) Then selectedindex = 0
               
          Case vbKeyPageDown
               
               selectedindex = selectedindex + cols * Rows
               If (selectedindex >= curnumitems) Then selectedindex = curnumitems - 1
               
          Case vbKeyHome
               
               selectedindex = 0
               
          Case vbKeyEnd
               
               selectedindex = curnumitems - 1
               
          Case vbKeyUp
               
               selectedindex = selectedindex - cols
               If (selectedindex < 0) Then selectedindex = 0
               
          Case vbKeyDown
               
               selectedindex = selectedindex + cols
               If (selectedindex >= curnumitems) Then selectedindex = curnumitems - 1
               
          Case vbKeyLeft
               
               selectedindex = selectedindex - 1
               If (selectedindex < 0) Then selectedindex = 0
               
          Case vbKeyRight
               
               selectedindex = selectedindex + 1
               If (selectedindex >= curnumitems) Then selectedindex = curnumitems - 1
               
          Case Else
               
               'Check if we can jump to a texture
               'Go for all texture names
               For i = 0 To (curnumitems - 1)
                    
                    'Check if the name starts with this char
                    If (StrComp(left$(curitemnames(i), 1), Chr$(KeyCode), vbTextCompare) = 0) Then
                         
                         'Select this texture
                         selectedindex = i
                         Exit For
                    End If
               Next i
               
     End Select
     
     'Select texture name
     If (selectedindex > -1) Then SelectedName = curitemnames(selectedindex)
     
     'Check if not the dash name
     If (SelectedName <> "-") And (SelectedName <> "") Then
          
          'Known texture?
          If (collection.Exists(SelectedName)) Then
               
               'Get texture object
               Set thisimage = collection(SelectedName)
               
               'Show details
               lblTextureName = SelectedName
               lblTextureSize = thisimage.width & " x " & thisimage.height
               
               'Clean up
               Set thisimage = Nothing
          Else
               
               'No details
               lblTextureName = ""
               lblTextureSize = ""
          End If
     Else
          
          'No details
          lblTextureName = ""
          lblTextureSize = ""
     End If
     
     'Check if old selection is within view
     ci = OldSelected - scrScroll.Value * cols
     If ((ci >= picItem.LBound) And (ci <= picItem.UBound)) Then
          
          'Deseect old
          picItem(ci).BackColor = vbBlack    'vbWindowBackground
          lblTexture(ci).BackColor = vbBlack 'vbWindowBackground
          lblTexture(ci).ForeColor = vbWhite 'vbWindowText
     End If
     
     'Check if the selection is above view
     ci = selectedindex - scrScroll.Value * cols
     If (ci < picItem.LBound) Then
          
          'Scroll to selection
          ScrollV = (selectedindex \ cols)
          If (ScrollV > scrScroll.Max) Then ScrollV = scrScroll.Max
          If (ScrollV < scrScroll.Min) Then ScrollV = scrScroll.Min
          scrScroll.Value = ScrollV
          
     'Check if the selection is below view
     ElseIf (ci > picItem.UBound) Then
          
          'Scroll to selection
          ScrollV = (selectedindex \ cols) - (Rows - 1)
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


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Check if we are allowed to do subclassing
     If (CommandSwitch("-nosubclass") = False) Then
          
          'Restore original messages handler
          SetWindowLong Me.hWnd, GWL_WNDPROC, OriginalMessageHandler
     End If
End Sub

Private Sub Form_Resize()
     
     'Fill the controls with textures
     ShowItems
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
     Dim thisimage As clsImage
     
     'Check if old selection is within view
     ci = selectedindex - scrScroll.Value * cols
     If ((ci >= picItem.LBound) And (ci <= picItem.UBound)) Then
          
          'Deseect old
          picItem(ci).BackColor = vbBlack    'vbWindowBackground
          lblTexture(ci).BackColor = vbBlack 'vbWindowBackground
          lblTexture(ci).ForeColor = vbWhite 'vbWindowText
     End If
     
     'Select this texture
     selectedindex = Index + scrScroll.Value * cols
     SelectedName = curitemnames(selectedindex)
     
     'Selection color
     picItem(Index).BackColor = vbHighlight
     lblTexture(Index).BackColor = vbHighlight
     lblTexture(Index).ForeColor = vbHighlightText
     
     'Check if not the dash name
     If (SelectedName <> "-") And (SelectedName <> "") Then
          
          'Known texture?
          If (collection.Exists(SelectedName)) Then
               
               'Get texture object
               Set thisimage = collection(SelectedName)
               
               'Show details
               lblTextureName = SelectedName
               lblTextureSize = thisimage.width & " x " & thisimage.height
               
               'Clean up
               Set thisimage = Nothing
          Else
               
               'No details
               lblTextureName = ""
               lblTextureSize = ""
          End If
     Else
          
          'No details
          lblTextureName = ""
          lblTextureSize = ""
     End If
     
     'Focus away
     On Error Resume Next
     picList.SetFocus
End Sub

Private Sub scrScroll_Change()
     
     'Refill controls
     ShowItems
     DoEvents
End Sub

Private Sub scrScroll_Scroll()
     
     'Refill controls
     ShowItems
     DoEvents
End Sub

Public Sub SetSelection(ByVal itemname As String)
     Dim ScrollV As Long
     Dim thisimage As clsImage
     Dim i As Long
     
     'Just the name
     SelectedName = itemname
     
     'Go for all texture names
     For i = 0 To curnumitems - 1
          
          'Select if name matches
          If (StrComp(curitemnames(i), itemname, vbTextCompare) = 0) Then
               
               'This item is now selected
               selectedindex = i
               SelectedName = curitemnames(i)
               
               'Check if not the dash name
               If (SelectedName <> "-") And (SelectedName <> "") Then
                    
                    'Get texture object
                    Set thisimage = collection(SelectedName)
                    
                    'Show details
                    lblTextureName = SelectedName
                    lblTextureSize = thisimage.width & " x " & thisimage.height
                    
                    'Clean up
                    Set thisimage = Nothing
               Else
                    
                    'No details
                    lblTextureName = ""
                    lblTextureSize = ""
               End If
               
               'Determine scroll position to show selection
               ScrollV = (selectedindex \ cols) - 2
               If (ScrollV < 0) Then ScrollV = 0
               If (ScrollV > scrScroll.Max) Then ScrollV = scrScroll.Max
               scrScroll.Value = ScrollV
               
               'Leave the search
               Exit For
          End If
     Next i
End Sub

Private Sub ShowItems()
     Dim Shown As Long
     Dim offset As Long
     Dim thisimage As clsImage
     Dim w As Long, h As Long
     Dim x As Long, y As Long
     Dim i As Long
     Dim ci As Long
     
     'Hide list, this solves flickering problems
     'picList.Visible = False
     LockWindowUpdate Me.hWnd
     
     'Calculate number of textures we can show
     Shown = cols * Rows
     
     'Calculate index offset
     offset = scrScroll.Value * cols
     
     'Go for all textures to be shown
     For i = offset To (offset + Shown - 1)
          
          'Get control index
          ci = i - offset
          
          'Determine x an y
          y = ci \ cols
          x = ci - y * cols
          
          'Check if within bounds
          If (i < curnumitems) Then
               
               'Clear picture
               Set imgTexture(ci).Picture = Nothing
               
               'Set texture name
               lblTexture(ci).Caption = curitemnames(i)
               
               'Check if this texture is the dash
               If (curitemnames(i) = "-") Then
                    
                    'Position
                    imgTexture(ci).Move BORDERSPACING, BORDERSPACING, 64, 64
               Else
                    
                    'Do not crash here
                    On Error Resume Next
                    
                    'Get texture object
                    Set thisimage = collection(curitemnames(i))
                    
                    'Set picture
                    Set imgTexture(ci).Picture = thisimage.Picture
                    
                    'Position
                    thisimage.GetScale 64, 64, w, h, False
                    imgTexture(ci).Move BORDERSPACING + (64 - w) \ 2, BORDERSPACING + (64 - h) \ 2, w, h
                    
                    'Continue error handling
                    On Error GoTo 0
               End If
               
               'Check if selected
               If (selectedindex = i) Then
                    
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
               
               'Show texture
               picItem(ci).visible = True
          Else
               
               'Clear texture
               Set imgTexture(ci).Picture = Nothing
               picItem(ci).visible = False
          End If
     Next i
     
     'Show list
     LockWindowUpdate 0
     'picList.Visible = True
     'picList.Refresh
End Sub
