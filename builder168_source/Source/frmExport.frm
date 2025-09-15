VERSION 5.00
Begin VB.Form frmExport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export Map"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
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
   Icon            =   "frmExport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3413
      TabIndex        =   6
      Top             =   2595
      Width           =   1545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   1643
      TabIndex        =   5
      Top             =   2595
      Width           =   1545
   End
   Begin VB.CheckBox chkCompressSidedefs 
      Caption         =   "Compress Sidedefs"
      Height          =   255
      Left            =   2288
      TabIndex        =   4
      Top             =   1860
      Width           =   2025
   End
   Begin VB.TextBox txtTargetFile 
      BackColor       =   &H80000000&
      Height          =   315
      Left            =   893
      TabIndex        =   3
      Text            =   "lalalala la la"
      Top             =   1380
      Width           =   4815
   End
   Begin VB.PictureBox picWarning 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      HasDC           =   0   'False
      Height          =   750
      Left            =   60
      ScaleHeight     =   48
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   430
      TabIndex        =   0
      Top             =   60
      Width           =   6480
      Begin VB.Image imgWarning 
         Height          =   240
         Left            =   15
         Picture         =   "frmExport.frx":000C
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmExport.frx":0596
         ForeColor       =   &H80000017&
         Height          =   600
         Left            =   375
         TabIndex        =   1
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   6000
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exporting map and resources to the following file:"
      Height          =   210
      Left            =   893
      TabIndex        =   2
      Top             =   1140
      Width           =   3585
   End
End
Attribute VB_Name = "frmExport"
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

Private Sub cmdCancel_Click()
     tag = ""
     Hide
End Sub


Private Sub cmdOK_Click()
     tag = "OK"
     Hide
End Sub


