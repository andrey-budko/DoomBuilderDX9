VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmThingType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Thing"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
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
   Icon            =   "frmThingType.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picThing 
      BackColor       =   &H8000000C&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      HasDC           =   0   'False
      Height          =   1380
      Left            =   4425
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   88
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Sector Floor Texture"
      Top             =   150
      Width           =   1380
      Begin VB.Image imgThing 
         Height          =   1260
         Left            =   60
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1200
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2505
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1665
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4245
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4980
      Width           =   1665
   End
   Begin VB.Frame fraThing 
      Height          =   4695
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   4095
      Begin MSComctlLib.TreeView trvThings 
         Height          =   4215
         Left            =   180
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7435
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   529
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "imglstThings"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lstThings 
         Height          =   4215
         Left            =   180
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   7435
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imglstThings"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Category"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Num"
            Object.Width           =   1111
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imglstThings 
      Left            =   0
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":0B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":10DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":1674
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":1C0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":21A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":2742
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":2CDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":3276
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":3810
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":3DAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":4344
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":48DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":4E78
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmThingType.frx":5412
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblThingType 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   210
      Left            =   5130
      TabIndex        =   15
      Top             =   1635
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   210
      Left            =   4425
      TabIndex        =   14
      Top             =   1635
      UseMnemonic     =   0   'False
      Width           =   405
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Hangs:"
      Height          =   210
      Left            =   4425
      TabIndex        =   13
      Top             =   2715
      Width           =   510
   End
   Begin VB.Label lblThingHangs 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   210
      Left            =   5130
      TabIndex        =   12
      Top             =   2715
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Blocking:"
      Height          =   210
      Left            =   4425
      TabIndex        =   11
      Top             =   2445
      Width           =   645
   End
   Begin VB.Label lblThingBlocks 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   210
      Left            =   5130
      TabIndex        =   10
      Top             =   2445
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Height:"
      Height          =   210
      Left            =   4425
      TabIndex        =   9
      Top             =   2175
      Width           =   495
   End
   Begin VB.Label lblThingHeight 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   210
      Left            =   5130
      TabIndex        =   8
      Top             =   2175
      Width           =   90
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Width:"
      Height          =   210
      Left            =   4425
      TabIndex        =   7
      Top             =   1905
      Width           =   450
   End
   Begin VB.Label lblThingWidth 
      AutoSize        =   -1  'True
      Caption         =   "?"
      Height          =   210
      Left            =   5130
      TabIndex        =   6
      Top             =   1905
      Width           =   90
   End
End
Attribute VB_Name = "frmThingType"
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

Public Sub HighlightThing(ByVal ThingIndex As Long)
     
     'Do not give an error when the item cant be found
     On Local Error Resume Next
     trvThings.SelectedItem.selected = False
     trvThings.nodes("T" & ThingIndex).selected = True
     trvThings.nodes("T" & ThingIndex).EnsureVisible
     trvThings_NodeClick trvThings.SelectedItem
     lstThings.SelectedItem.selected = False
     lstThings.ListItems("T" & ThingIndex).selected = True
     lstThings.ListItems("T" & ThingIndex).EnsureVisible
     lstThings_ItemClick lstThings.SelectedItem
     On Local Error GoTo 0
End Sub

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


Private Sub Form_Load()
     
     'Check if showing tree or list
     If (Val(Config("thingstree")) = vbChecked) Then
          
          'Fill things tree
          FillThingsTree trvThings
          trvThings.visible = True
     Else
          
          'Fill things list
          FillThingsList lstThings
          lstThings.visible = True
     End If
End Sub


Private Sub lstThings_DblClick()
     
     'OK
     If cmdOK.Enabled Then cmdOK_Click
End Sub

Private Sub lstThings_ItemClick(ByVal Item As MSComctlLib.ListItem)
     lstThings.tag = Trim$(Item.tag)
End Sub


Private Sub trvThings_DblClick()
     
     'Check if node is a leaf
     If (trvThings.SelectedItem.Children = 0) Then
          
          'OK
          If cmdOK.Enabled Then cmdOK_Click
     End If
End Sub

Private Sub trvThings_NodeClick(ByVal Node As MSComctlLib.Node)
     
     'Check if node is a leaf
     If (Node.Children = 0) Then
          
          'Apply selection
          lstThings.tag = Trim$(Node.tag)
          
          'Erase thing preview
          Set imgThing.Picture = Nothing
          
          'Show thing preview if possible
          GetScaledSpritePicture Val(lstThings.tag), imgThing, picThing.ScaleWidth, picThing.ScaleHeight, False
          
          'Show thing properties
          lblThingType.Caption = CStr(Val(lstThings.tag))
          lblThingWidth.Caption = GetThingWidth(Val(lstThings.tag))
          lblThingHeight.Caption = GetThingHeight(Val(lstThings.tag))
          lblThingHangs.Caption = YesNo(GetThingHangs(Val(lstThings.tag)))
          lblThingBlocks.Caption = GetThingBlockingDesc(GetThingBlocking(Val(lstThings.tag)))
     End If
End Sub


