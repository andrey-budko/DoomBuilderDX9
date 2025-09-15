VERSION 5.00
Begin VB.Form frmErrorLog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Errors and Warnings"
   ClientHeight    =   3285
   ClientLeft      =   360
   ClientTop       =   1500
   ClientWidth     =   9240
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
   Icon            =   "frmErrorLog.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   219
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   616
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   7500
      TabIndex        =   2
      Top             =   2835
      Width           =   1575
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Height          =   345
      Left            =   945
      TabIndex        =   1
      Top             =   2835
      Width           =   1575
   End
   Begin VB.TextBox txtErrors 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   945
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   8130
   End
   Begin VB.Image imgCritical 
      Height          =   480
      Left            =   225
      Picture         =   "frmErrorLog.frx":000C
      Top             =   915
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   225
      Picture         =   "frmErrorLog.frx":08D6
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblDesc 
      AutoSize        =   -1  'True
      Caption         =   "The following errors and warnings occurred while loading the map or configuration:"
      Height          =   210
      Left            =   945
      TabIndex        =   3
      Top             =   180
      Width           =   6060
   End
End
Attribute VB_Name = "frmErrorLog"
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

Private Sub cmdClose_Click()
     Unload Me
     Set frmErrorLog = Nothing
End Sub

Private Sub cmdSave_Click()
     Dim Result As String
     Dim FileBuffer As Integer
     Dim GameFile As String
     
     'Show save dialog
     Result = SaveFile(Me.hWnd, "Save Errors As", "Text Files   *.txt|*.txt", "", cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist)
     
     'Check if not cancelled
     If Result <> "" Then
          
          'Save the log here
          FileBuffer = FreeFile
          Open Result For Append As #FileBuffer
          
          'Get engine and game config file
          GameFile = Dir(GetGameConfigFile(mapgame))
          
          'Output header
          Print #FileBuffer, "Errors and Warnings logged at: " & CDate(Now) & "   Current map is: " & UCase$(mapfilename) & " (" & UCase$(maplumpname) & ")"
          Print #FileBuffer, "Game configuration is: " & GameFile
          Print #FileBuffer, "========================================================================================="
          
          'Output errors
          Print #FileBuffer, txtErrors.Text
          
          'Empty line
          Print #FileBuffer, ""
          
          'Close
          Close #FileBuffer
     End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


