Attribute VB_Name = "modTextureEdit"
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


'API Declarations
Public Declare Sub DrawPalettedExample Lib "builder.dll" (ByRef Source As Long, ByRef target As BITMAPRGB, ByVal pixels As Long, ByVal width As Long, ByRef palette As BITMAPRGB, ByVal numcolors As Long, ByVal TransColor As Long, ByVal transdiff As Long, ByVal backcolor As Long, ByVal Method As Long)
Public Declare Function GetNearestColor Lib "builder.dll" (ByVal color As Long, ByRef palette As BITMAPRGB, ByVal numcolors As Long) As Long
Public Declare Function GetDarkerColor Lib "builder.dll" (ByVal color As Long, ByRef palette As BITMAPRGB, ByVal numcolors As Long) As Long
Public Declare Function GetLighterColor Lib "builder.dll" (ByVal color As Long, ByRef palette As BITMAPRGB, ByVal numcolors As Long) As Long


'Temporary storage for picture colors
Public TexturePreviewColors() As Long
Public TexturePreviewWidth As Long
Public TexturePreviewHeight As Long

'Previous picture file opened
Public LastPictureFile As String

Public Function SelectImportPicture(ByVal Owner As Form, ByVal AllowTransparency As Boolean, ByRef Method As Long, ByRef TransColor As Long, ByRef TransRange As Long, ByRef LumpName As String) As Boolean
     Dim Result As String
     
     'Browse for picture
     Result = OpenFile(Owner.hWnd, "Select Picture", "All supported picture formats|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf|All Files|*.*", LastPictureFile, cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     
     'Check result
     If (Result <> "") Then
          
          'Load the picture dialog
          Load frmPictureImport
          
          'Disable transparency
          If (AllowTransparency = False) Then
               With frmPictureImport
                    .lblTrans.ForeColor = vbGrayText
                    .lblTransRange.ForeColor = vbGrayText
                    .txtTransRange.Enabled = False
                    .cmdTransColor.Enabled = False
                    .cmdTransColor.backcolor = vbGrayText
               End With
          End If
          
          'Load the picture
          Set frmPictureImport.picOriginal.Picture = LoadPicture(Result)
          
          'Show the dialog
          frmPictureImport.Show 1, Owner
          
          'Check result
          If (Val(frmPictureImport.tag) = 1) Then
               
               'Set the values
               With frmPictureImport
                    
                    'Return method
                    If (.optNearest.Value) Then
                         Method = 0
                    ElseIf (.optDarker.Value) Then
                         Method = 1
                    ElseIf (.optLighter.Value) Then
                         Method = 2
                    End If
                    
                    'Return transparency color
                    TransColor = .cmdTransColor.backcolor
                    
                    'Return transparency range
                    TransRange = Val(.txtTransRange.Text)
                    
                    'Return lump name
                    LumpName = .txtLumpName.Text
               End With
               
               'Return success
               SelectImportPicture = True
          Else
               
               'Cancelled
               SelectImportPicture = False
          End If
     Else
          
          'Cancelled
          SelectImportPicture = False
     End If
End Function


