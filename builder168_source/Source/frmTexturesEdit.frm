VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTexturesEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Texture Management"
   ClientHeight    =   7545
   ClientLeft      =   735
   ClientTop       =   600
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTexturesEdit.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   503
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraPanel 
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   2
      Left            =   165
      TabIndex        =   31
      Top             =   510
      Visible         =   0   'False
      Width           =   10095
      Begin VB.ListBox lstFlats 
         Height          =   4980
         IntegralHeight  =   0   'False
         ItemData        =   "frmTexturesEdit.frx":000C
         Left            =   120
         List            =   "frmTexturesEdit.frx":000E
         Sorted          =   -1  'True
         TabIndex        =   37
         Top             =   90
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Caption         =   " Selected Flat "
         Height          =   6255
         Left            =   1980
         TabIndex        =   35
         Top             =   -15
         Width           =   8070
         Begin VB.PictureBox picFlat 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   5790
            Left            =   180
            ScaleHeight     =   382
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   512
            TabIndex        =   36
            Top             =   300
            Width           =   7740
         End
      End
      Begin VB.CommandButton cmdFlatImport 
         Caption         =   "Import Flat"
         Height          =   345
         Left            =   120
         TabIndex        =   34
         Top             =   5175
         Width           =   1695
      End
      Begin VB.CommandButton cmdFlatDelete 
         Caption         =   "Delete Flat"
         Height          =   345
         Left            =   120
         TabIndex        =   33
         Top             =   5865
         Width           =   1695
      End
      Begin VB.CommandButton cmdFlatRename 
         Caption         =   "Rename Flat"
         Height          =   345
         Left            =   120
         TabIndex        =   32
         Top             =   5520
         Width           =   1695
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   345
      Left            =   8580
      TabIndex        =   23
      Top             =   7095
      Width           =   1785
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   300
      Left            =   120
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   530
      TabIndex        =   21
      Top             =   7110
      Width           =   7980
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   15
         Picture         =   "frmTexturesEdit.frx":0010
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: Editing textures, patches or flats will apply to all maps in the wad file you are currently editing!"
         ForeColor       =   &H80000017&
         Height          =   210
         Left            =   375
         TabIndex        =   22
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   7455
      End
   End
   Begin VB.Frame fraPanel 
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   165
      TabIndex        =   20
      Top             =   510
      Visible         =   0   'False
      Width           =   10095
      Begin VB.CommandButton cmdPatchRename 
         Caption         =   "Rename Patch"
         Height          =   345
         Left            =   120
         TabIndex        =   29
         Top             =   5520
         Width           =   1695
      End
      Begin VB.CommandButton cmdPatchDelete 
         Caption         =   "Delete Patch"
         Height          =   345
         Left            =   120
         TabIndex        =   28
         Top             =   5865
         Width           =   1695
      End
      Begin VB.CommandButton cmdPatchImport 
         Caption         =   "Import Patch"
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   5175
         Width           =   1695
      End
      Begin VB.Frame Frame1 
         Caption         =   " Selected Patch "
         Height          =   6255
         Left            =   1980
         TabIndex        =   25
         Top             =   -15
         Width           =   8070
         Begin VB.PictureBox picPatch 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   5790
            Left            =   180
            ScaleHeight     =   382
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   512
            TabIndex        =   26
            Top             =   300
            Width           =   7740
         End
      End
      Begin VB.ListBox lstPatches 
         Height          =   4980
         IntegralHeight  =   0   'False
         ItemData        =   "frmTexturesEdit.frx":059A
         Left            =   120
         List            =   "frmTexturesEdit.frx":059C
         Sorted          =   -1  'True
         TabIndex        =   24
         Top             =   90
         Width           =   1695
      End
   End
   Begin VB.Frame fraPanel 
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   0
      Left            =   165
      TabIndex        =   1
      Tag             =   "Textures"
      Top             =   510
      Width           =   10095
      Begin VB.CommandButton cmdTextureDelete 
         Caption         =   "Delete Texture"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   5865
         Width           =   1695
      End
      Begin VB.CommandButton cmdTextureCopy 
         Caption         =   "Copy Texture"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   5175
         Width           =   1695
      End
      Begin VB.CommandButton cmdTextureNew 
         Caption         =   "New Texture"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   4830
         Width           =   1695
      End
      Begin VB.Frame fraSelectedTexture 
         Caption         =   " Selected Texture "
         Height          =   6255
         Left            =   1980
         TabIndex        =   3
         Top             =   -15
         Width           =   8070
         Begin VB.CommandButton cmdTexturePatchUp 
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   9.75
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2475
            TabIndex        =   19
            ToolTipText     =   "Move selection Up"
            Top             =   5820
            Width           =   408
         End
         Begin VB.CommandButton cmdTexturePatchDown 
            Caption         =   "6"
            BeginProperty Font 
               Name            =   "Marlett"
               Size            =   9.75
               Charset         =   2
               Weight          =   500
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2880
            TabIndex        =   18
            ToolTipText     =   "Move selection Down"
            Top             =   5820
            Width           =   408
         End
         Begin VB.CommandButton cmdTexturePatchRemove 
            Caption         =   "Remove"
            Height          =   315
            Left            =   1710
            TabIndex        =   11
            Top             =   5820
            Width           =   765
         End
         Begin VB.CommandButton cmdTexturePatchCopy 
            Caption         =   "Copy"
            Height          =   315
            Left            =   945
            TabIndex        =   10
            Top             =   5820
            Width           =   765
         End
         Begin VB.CommandButton cmdTexturePatchAdd 
            Caption         =   "Add"
            Height          =   315
            Left            =   180
            TabIndex        =   9
            Top             =   5820
            Width           =   765
         End
         Begin MSComctlLib.ListView lstTexturePatches 
            Height          =   1455
            Left            =   180
            TabIndex        =   8
            Top             =   4305
            Width           =   3105
            _ExtentX        =   5477
            _ExtentY        =   2566
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Patch Name"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Coordinates"
               Object.Width           =   1764
            EndProperty
         End
         Begin VB.PictureBox picTexture 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   3900
            Left            =   180
            ScaleHeight     =   256
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   512
            TabIndex        =   4
            Top             =   300
            Width           =   7740
         End
         Begin VB.Frame fraTexturePatch 
            Caption         =   " Texture Patch "
            Height          =   1845
            Left            =   3420
            TabIndex        =   12
            Top             =   4260
            Width           =   4485
            Begin DoomBuilder.ctlValueBox ctlValueBox1 
               Height          =   375
               Left            =   1440
               TabIndex        =   16
               Top             =   975
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   661
               Max             =   9999
            End
            Begin VB.ComboBox cmbTexturePatch 
               Height          =   330
               Left            =   1440
               TabIndex        =   14
               Text            =   "Combo"
               Top             =   540
               Width           =   1890
            End
            Begin DoomBuilder.ctlValueBox ctlValueBox2 
               Height          =   375
               Left            =   2430
               TabIndex        =   17
               Top             =   975
               Width           =   885
               _ExtentX        =   1561
               _ExtentY        =   661
               Max             =   9999
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Coordinates:"
               Height          =   210
               Left            =   420
               TabIndex        =   15
               Top             =   1035
               Width           =   915
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Patch name:"
               Height          =   210
               Left            =   450
               TabIndex        =   13
               Top             =   585
               Width           =   885
            End
         End
      End
      Begin VB.ListBox lstTextures 
         Height          =   4650
         IntegralHeight  =   0   'False
         ItemData        =   "frmTexturesEdit.frx":059E
         Left            =   120
         List            =   "frmTexturesEdit.frx":05A0
         Sorted          =   -1  'True
         TabIndex        =   2
         Top             =   90
         Width           =   1695
      End
      Begin VB.CommandButton cmdTextureRename 
         Caption         =   "Rename Texture"
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   5520
         Width           =   1695
      End
   End
   Begin MSComctlLib.TabStrip tbsPanel 
      Height          =   6825
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   12039
      TabWidthStyle   =   2
      ShowTips        =   0   'False
      TabFixedWidth   =   3175
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Textures"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Patches"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Flats"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTexturesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'    XODE Multimedia Doom Builder
'    by Pascal 'gherkin' vd Heiden
'    gherkin@xodemultimedia.com
'    www.xodemultimedia.com
'
'    Copyright (c) 2003 XODE Multimedia
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


Private Sub tbsPanel_Click()
     Dim i As Long
     
     'Show the frame
     For i = fraPanel.LBound To fraPanel.UBound
          
          'Check if this tab is selected
          If (i = tbsPanel.SelectedItem.Index - 1) Then
               
               'Show the frame
               fraPanel(i).Visible = True
               
               'Leave here
               Exit For
          End If
     Next i
     
     'Hide all other frames
     For i = fraPanel.LBound To fraPanel.UBound
          
          'Hide frame if not selected
          If (i <> tbsPanel.SelectedItem.Index - 1) Then fraPanel(i).Visible = False
     Next i
End Sub


