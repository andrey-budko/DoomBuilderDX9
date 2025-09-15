VERSION 5.00
Begin VB.Form frmMapOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map Options"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
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
   Icon            =   "frmMapOptions.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   336
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtFlatDir 
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   3525
      Width           =   3420
   End
   Begin VB.CommandButton cmdBrowseFlatDir 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   3720
      TabIndex        =   15
      Top             =   3510
      Width           =   1110
   End
   Begin VB.TextBox txtTexDir 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   2805
      Width           =   3420
   End
   Begin VB.CommandButton cmdBrowseTexDir 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   3720
      TabIndex        =   12
      Top             =   2790
      Width           =   1110
   End
   Begin VB.CommandButton cmdBrowseWAD 
      Caption         =   "Browse..."
      Height          =   345
      Left            =   3720
      TabIndex        =   3
      Top             =   2070
      Width           =   1110
   End
   Begin VB.TextBox txtWAD 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   2085
      Width           =   3420
   End
   Begin VB.TextBox txtMapLumpName 
      Height          =   315
      Left            =   1380
      MaxLength       =   8
      TabIndex        =   1
      Top             =   1320
      Width           =   1425
   End
   Begin VB.ComboBox cmbGameConfig 
      Height          =   330
      IntegralHeight  =   0   'False
      ItemData        =   "frmMapOptions.frx":000C
      Left            =   1380
      List            =   "frmMapOptions.frx":000E
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   810
      Width           =   3450
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   510
      Left            =   60
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   327
      TabIndex        =   6
      Top             =   60
      Width           =   4935
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   45
         Picture         =   "frmMapOptions.frx":0010
         Top             =   90
         Width           =   240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning: These settings control the way your map is saved. Be sure to configure these correctly."
         ForeColor       =   &H80000017&
         Height          =   450
         Left            =   375
         TabIndex        =   7
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   4470
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   2595
      TabIndex        =   5
      Top             =   4185
      Width           =   1515
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   345
      Left            =   930
      TabIndex        =   4
      Top             =   4185
      Width           =   1515
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Additional Flats from directory:"
      Height          =   210
      Left            =   240
      TabIndex        =   17
      Top             =   3285
      Width           =   2205
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Additional Textures from directory:"
      Height          =   210
      Left            =   240
      TabIndex        =   14
      Top             =   2565
      Width           =   2505
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "(example: MAP01)"
      Height          =   210
      Left            =   3000
      TabIndex        =   11
      Top             =   1365
      UseMnemonic     =   0   'False
      Width           =   1320
   End
   Begin VB.Label lblLumpName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Level name:"
      Height          =   210
      Left            =   375
      TabIndex        =   10
      Top             =   1365
      UseMnemonic     =   0   'False
      Width           =   870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Additional Textures and Flats from WAD file:"
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   1845
      Width           =   3195
   End
   Begin VB.Label lblGameConfig 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Configuration:"
      Height          =   210
      Left            =   240
      TabIndex        =   8
      Top             =   870
      UseMnemonic     =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "frmMapOptions"
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


Public Loading As Boolean

Private Sub cmbGameConfig_Change()
     Dim GameCFGFile As String
     Dim GameCFG As New clsConfiguration
     
     'Set OK button enabled/disabled
     cmdOK.Enabled = (Trim$(cmbGameConfig.Text) <> "") And (Trim$(txtMapLumpName) <> "")
     
     'Check if not loading
     If (Not Loading) And (Val(tag) = 1) Then
          
          'Change the extra wad file to the default for this config
          GameCFGFile = GetGameConfigFile(cmbGameConfig.Text)
          If (Trim$(GameCFGFile) <> "") Then
               
               'Load config
               GameCFG.LoadConfiguration GameCFGFile
               
               'Get the file
               txtWAD.Text = GameCFG.ReadSetting("texturesfile", "")
               
               'Change the default map lump name
               txtMapLumpName.Text = GameCFG.ReadSetting("defaultlumpname", "")
          End If
     End If
End Sub

Private Sub cmbGameConfig_Click()
     cmbGameConfig_Change
End Sub

Private Sub cmbGameConfig_KeyUp(KeyCode As Integer, Shift As Integer)
     cmbGameConfig_Change
End Sub

Private Sub cmdBrowseFlatDir_Click()
     Dim NewFolder As String
     
     'Browse for new file
     NewFolder = SelectFolder(Me.hWnd, "Select additional flats directory")
     
     'Check if not cancelled
     If (Trim$(NewFolder) <> "") Then
          
          'Set the new file in textbox
          txtFlatDir.Text = NewFolder
          txtFlatDir.SelStart = Len(txtFlatDir.Text)
          txtFlatDir.SetFocus
     End If
End Sub

Private Sub cmdBrowseTexDir_Click()
     Dim NewFolder As String
     
     'Browse for new file
     NewFolder = SelectFolder(Me.hWnd, "Select additional textures directory")
     
     'Check if not cancelled
     If (Trim$(NewFolder) <> "") Then
          
          'Set the new file in textbox
          txtTexDir.Text = NewFolder
          txtTexDir.SelStart = Len(txtTexDir.Text)
          txtTexDir.SetFocus
     End If
End Sub


Private Sub cmdBrowseWAD_Click()
     Dim NewFile As String
     
     'Browse for new file
     NewFile = OpenFile(Me.hWnd, "Select Extra WAD File", "Doom/Heretic/Hexen WAD Files   *.wad|*.wad|All Files|*.*", "", cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check if not cancelled
     If (Trim$(NewFile) <> "") Then
          
          'Set the new file in textbox
          txtWAD.Text = NewFile
          txtWAD.SelStart = Len(txtWAD.Text)
          txtWAD.SetFocus
     End If
End Sub

Private Sub cmdCancel_Click()
     tag = 0
     Hide
End Sub

Private Sub cmdOK_Click()
     
     'Check if the lump name changed
     If (StrComp(Trim$(maplumpname), Trim$(txtMapLumpName.Text), vbTextCompare) <> 0) Then
          
          'Check if a map is open
          If (mapfile <> "") Then
               
               'Check if the map is in a file
               If mapsaved Then
                    
                    'Check if the given lump name exists in file
                    If (FindLumpIndex(MapWAD, 1, Trim$(txtMapLumpName.Text)) > 0) Then
                         
                         'Lump already exists, ask confirmation
                         If (MsgBox("The map lump name you entered already exists in the current WAD file." & vbLf & "Saving your map will replace that lump or map with the current map." & vbLf & "Do you want to continue?", vbExclamation Or vbYesNo) = vbNo) Then Exit Sub
                    End If
               End If
          End If
     End If
     
     tag = 1
     Hide
End Sub

Private Sub Command1_Click()

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
     If (UnloadMode = 0) Then
          Cancel = True
          cmdCancel_Click
     End If
End Sub

Private Sub txtFlatDir_GotFocus()
     SelectAllText txtFlatDir
End Sub


Private Sub txtMapLumpName_Change()
     cmdOK.Enabled = (Trim$(cmbGameConfig.Text) <> "") And (Trim$(txtMapLumpName) <> "")
End Sub

Private Sub txtMapLumpName_GotFocus()
     SelectAllText txtMapLumpName
End Sub


Private Sub txtMapLumpName_KeyPress(KeyAscii As Integer)
     KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
End Sub

Private Sub txtTexDir_GotFocus()
     SelectAllText txtTexDir
End Sub


Private Sub txtWAD_GotFocus()
     SelectAllText txtWAD
End Sub


