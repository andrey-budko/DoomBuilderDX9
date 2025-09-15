VERSION 5.00
Begin VB.Form frmExportPicture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Picture"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
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
   Icon            =   "frmExportPicture.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   366
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5430
      TabIndex        =   15
      Top             =   4980
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3720
      TabIndex        =   14
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Frame fraGrid 
      Caption         =   "   "
      Height          =   1470
      Left            =   3900
      TabIndex        =   21
      Top             =   3240
      Width           =   3105
      Begin VB.CommandButton cmdGridSettings 
         Caption         =   "Grid Settings..."
         Enabled         =   0   'False
         Height          =   345
         Left            =   570
         TabIndex        =   6
         Top             =   885
         Width           =   2025
      End
      Begin VB.CheckBox chkGrid64 
         Caption         =   "Show 64 mappixels grid"
         Enabled         =   0   'False
         Height          =   255
         Left            =   570
         TabIndex        =   5
         Top             =   450
         Width           =   2145
      End
      Begin VB.CheckBox chkShowGrid 
         Caption         =   "Show Grid"
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   0
         Width           =   1125
      End
   End
   Begin VB.Frame fraThings 
      Caption         =   "   "
      Height          =   1470
      Left            =   3900
      TabIndex        =   20
      Top             =   1590
      Width           =   3105
      Begin VB.CommandButton cmdFilterSettings 
         Caption         =   "Filter Settings..."
         Height          =   345
         Left            =   570
         TabIndex        =   29
         Top             =   885
         Width           =   2025
      End
      Begin VB.CheckBox chkThingDimmed 
         Caption         =   "Dimmed and in background"
         Height          =   255
         Left            =   570
         TabIndex        =   3
         Top             =   450
         Width           =   2295
      End
      Begin VB.CheckBox chkShowThings 
         Caption         =   "Show Things"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   0
         Value           =   1  'Checked
         Width           =   1260
      End
   End
   Begin VB.Frame fraVertices 
      Caption         =   "   "
      Height          =   1350
      Left            =   3900
      TabIndex        =   16
      Top             =   135
      Width           =   3105
      Begin DoomBuilder.ctlValueBox txtVertexSize 
         Height          =   375
         Left            =   1050
         TabIndex        =   1
         Top             =   435
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         Max             =   999
         Min             =   1
         SmallChange     =   2
      End
      Begin VB.CheckBox chkShowVertices 
         Caption         =   "Show Vertices"
         Height          =   255
         Left            =   180
         TabIndex        =   0
         Top             =   0
         Width           =   1410
      End
      Begin VB.Label lblPixelsVertices 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   210
         Left            =   2040
         TabIndex        =   18
         Top             =   495
         Width           =   420
      End
      Begin VB.Label lblVertexSize 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   210
         Left            =   570
         TabIndex        =   17
         Top             =   495
         Width           =   360
      End
   End
   Begin VB.Frame fraLines 
      Caption         =   "   "
      Height          =   1350
      Left            =   195
      TabIndex        =   19
      Top             =   135
      Width           =   3510
      Begin VB.CheckBox chkShowLineNormals 
         Caption         =   "Show front indicators"
         Height          =   255
         Left            =   570
         TabIndex        =   30
         Top             =   825
         Width           =   2595
      End
      Begin VB.CheckBox chkShowLengths 
         Caption         =   "Show line lengths"
         Height          =   255
         Left            =   570
         TabIndex        =   8
         Top             =   480
         Width           =   2595
      End
      Begin VB.CheckBox chkShowLines 
         Caption         =   "Show Lines and Sectors "
         Height          =   255
         Left            =   180
         TabIndex        =   7
         Top             =   0
         Value           =   1  'Checked
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Size and Resolution "
      Height          =   3120
      Left            =   195
      TabIndex        =   22
      Top             =   1590
      Width           =   3510
      Begin VB.OptionButton optScaled 
         Caption         =   "Scaled manually"
         Height          =   255
         Left            =   510
         TabIndex        =   12
         Top             =   1950
         Width           =   2235
      End
      Begin VB.OptionButton optResolution 
         Caption         =   "Scaled to fit resolution"
         Height          =   255
         Left            =   510
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   2235
      End
      Begin DoomBuilder.ctlValueBox txtWidth 
         Height          =   375
         Left            =   1380
         TabIndex        =   10
         Top             =   870
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         Min             =   1
      End
      Begin DoomBuilder.ctlValueBox txtHeight 
         Height          =   375
         Left            =   1380
         TabIndex        =   11
         Top             =   1290
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         Min             =   1
      End
      Begin DoomBuilder.ctlValueBox txtScale 
         Height          =   375
         Left            =   1395
         TabIndex        =   13
         Top             =   2325
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         Enabled         =   0   'False
         Max             =   10000
         Min             =   1
      End
      Begin VB.Label lblPercent 
         AutoSize        =   -1  'True
         Caption         =   "percent"
         ForeColor       =   &H80000011&
         Height          =   210
         Left            =   2385
         TabIndex        =   28
         Top             =   2385
         Width           =   555
      End
      Begin VB.Label lblScale 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Scale:"
         ForeColor       =   &H80000011&
         Height          =   210
         Left            =   855
         TabIndex        =   27
         Top             =   2385
         Width           =   450
      End
      Begin VB.Label lblPixelsHeight 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   210
         Left            =   2370
         TabIndex        =   26
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label lblPixelsWidth 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   210
         Left            =   2370
         TabIndex        =   25
         Top             =   930
         Width           =   420
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Height:"
         Height          =   210
         Left            =   795
         TabIndex        =   24
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Width:"
         Height          =   210
         Left            =   840
         TabIndex        =   23
         Top             =   930
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmExportPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkShowGrid_Click()
     chkGrid64.Enabled = (chkShowGrid.Value = vbChecked)
     cmdGridSettings.Enabled = (chkShowGrid.Value = vbChecked)
End Sub

Private Sub chkShowLines_Click()
     chkShowLengths.Enabled = (chkShowLines.Value = vbChecked)
End Sub

Private Sub chkShowThings_Click()
     chkThingDimmed.Enabled = (chkShowThings.Value = vbChecked)
     cmdFilterSettings.Enabled = (chkShowThings.Value = vbChecked)
End Sub


Private Sub chkShowVertices_Click()
     txtVertexSize.Enabled = (chkShowVertices.Value = vbChecked)
     
     If (chkShowVertices.Value = vbChecked) Then
          lblVertexSize.ForeColor = vbButtonText
          lblPixelsVertices.ForeColor = vbButtonText
     Else
          lblVertexSize.ForeColor = vbGrayText
          lblPixelsVertices.ForeColor = vbGrayText
     End If
End Sub


Private Sub cmdCancel_Click()
     Hide
End Sub

Private Sub cmdFilterSettings_Click()
     
     'Show filter dialog
     Load frmThingFilter
     frmThingFilter.Show 1, Me
End Sub

Private Sub cmdGridSettings_Click()
     
     'Show Grid Settings
     Load frmGrid
     frmGrid.Show 1, Me
End Sub


Private Sub cmdOK_Click()
     frmExportPicture.tag = "OK"
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
     
     'Defaults
     txtVertexSize.Value = 3
     txtWidth.Value = 640
     txtHeight.Value = 480
     txtScale.Value = 50
End Sub

Private Sub optResolution_Click()
     
     'Enable/disable controls
     lblWidth.ForeColor = vbButtonText
     lblHeight.ForeColor = vbButtonText
     lblPixelsWidth.ForeColor = vbButtonText
     lblPixelsHeight.ForeColor = vbButtonText
     txtWidth.Enabled = True
     txtHeight.Enabled = True
     lblScale.ForeColor = vbGrayText
     lblPercent.ForeColor = vbGrayText
     txtScale.Enabled = False
End Sub


Private Sub optScaled_Click()
     
     'Enable/disable controls
     lblWidth.ForeColor = vbGrayText
     lblHeight.ForeColor = vbGrayText
     lblPixelsWidth.ForeColor = vbGrayText
     lblPixelsHeight.ForeColor = vbGrayText
     txtWidth.Enabled = False
     txtHeight.Enabled = False
     lblScale.ForeColor = vbButtonText
     lblPercent.ForeColor = vbButtonText
     txtScale.Enabled = True
End Sub


Private Sub txtHeight_GotFocus()
     SelectAllText txtHeight
End Sub


Private Sub txtScale_GotFocus()
     SelectAllText txtScale
End Sub


Private Sub txtVertexSize_GotFocus()
     SelectAllText txtVertexSize
End Sub


Private Sub txtWidth_GotFocus()
     SelectAllText txtWidth
End Sub


