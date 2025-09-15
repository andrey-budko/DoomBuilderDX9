Attribute VB_Name = "modCommonDialog"
Attribute VB_Description = "Common Dialog API Module"
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


'File Open/Save structure
Private Type cdlOpenFileName
     lStructSize As Long
     Owner As Long
     hInstance As Long
     Filter As String
     CustomFilter As String
     MaxCustFilter As Long
     FilterIndex As Long
     File As String
     MaxFile As Long
     FileTitle As String
     MaxFileTitle As Long
     InitialDir As String
     Title As String
     flags As Long
     FileOffset As Integer
     FileExtension As Integer
     DefExt As String
     CustData As Long
     Hook As Long
     TemplateName As String
End Type

'Color structure
Private Type cdlColor
     lStructSize As Long
     hwndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As Long
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type

'Folder structure
Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type


'Declarations
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As cdlOpenFileName) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pSavefilename As cdlOpenFileName) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As cdlColor) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)


'Constants
Public Enum cdlFileOpenConstants
     cdlOFNAllowMultiselect = 512
     cdlOFNCreatePrompt = 8192
     cdlOFNExplorer = 524288
     cdlOFNExtensionDifferent = 1024
     cdlOFNFileMustExist = 4096
     cdlOFNHelpButton = 16
     cdlOFNHideReadOnly = 4
     cdlOFNLongNames = 2097152
     cdlOFNNoChangeDir = 8
     cdlOFNNoDereferenceLinks = 1048576
     cdlOFNNoLongNames = 262144
     cdlOFNNoReadOnlyReturn = 32768
     cdlOFNNoValidate = 256
     cdlOFNOverwritePrompt = 2
     cdlOFNPathMustExist = 2048
     cdlOFNReadOnly = 1
     cdlOFNShareAware = 16384
End Enum

Public Enum cdlColorConstants
     cdlCCFullOpen = 2
     cdlCCHelpButton = 8
     cdlCCPreventFullOpen = 4
     cdlCCRGBInit = 1
End Enum

Public Function OpenFile(ByVal hWnd As Long, ByRef DialogTitle As String, ByRef Filter As String, ByRef InitFilename As String, ByVal flags As cdlFileOpenConstants) As String
     On Local Error Resume Next
     Dim FileStruct As cdlOpenFileName
     
     With FileStruct
          
          'Set all dialog parameters
          .lStructSize = Len(FileStruct)
          .hInstance = App.hInstance
          .Owner = hWnd
          .Title = DialogTitle
          .Filter = Replace$(Filter, "|", vbNullChar)
          .flags = flags
          .File = Space$(257) & vbNullChar
          .MaxFile = Len(.File)
          If InitFilename <> "" And Dir(InitFilename) <> "" Then .InitialDir = left$(InitFilename, Len(InitFilename) - Len(Dir(InitFilename)))
          
          'Show the dialog and return the result
          If GetOpenFileName(FileStruct) Then OpenFile = StripNullChar(.File)
     End With
End Function

Public Function SaveFile(ByVal hWnd As Long, ByRef DialogTitle As String, ByRef Filter As String, ByRef InitFilename As String, ByVal flags As cdlFileOpenConstants, Optional ByRef FilterIndex As Long) As String
     On Local Error Resume Next
     Dim FileStruct As cdlOpenFileName
     
     With FileStruct
          
          'Set all dialog parameters
          .lStructSize = Len(FileStruct)
          .hInstance = App.hInstance
          .Owner = hWnd
          .Title = DialogTitle
          .Filter = Replace$(Filter, "|", vbNullChar)
          .flags = flags
          .File = Space$(257) & vbNullChar
          .MaxFile = Len(.File)
          If InitFilename <> "" And Dir(InitFilename) <> "" Then .InitialDir = left$(InitFilename, Len(InitFilename) - Len(Dir(InitFilename)))
          
          'Show the dialog and return the result
          If GetSaveFileName(FileStruct) Then SaveFile = StripNullChar(.File)
          
          'Return FilterIndex
          FilterIndex = .FilterIndex
     End With
End Function

Public Function SelectColor(ByVal hWnd As Long, ByRef InitColor As Long, ByVal flags As cdlColorConstants, ByRef CustomColors() As Long) As Long
     Dim ColorStruct As cdlColor
     
     With ColorStruct
          
          'Set all dialog parameters
          .lStructSize = Len(ColorStruct)
          .hInstance = App.hInstance
          .hwndOwner = hWnd
          .rgbResult = InitColor
          .lpCustColors = VarPtr(CustomColors(0))
          .flags = flags
          
          'Show the dialog and return the result
          If ChooseColor(ColorStruct) Then SelectColor = .rgbResult Else SelectColor = -1
     End With
End Function

Public Function SelectFolder(hWnd As Long, Prompt As String) As String
     On Local Error Resume Next
     Dim lngIDList As Long
     Dim lngResult As Long
     Dim strPath As String
     Dim udtBI As BrowseInfo
     
     'Set browse information
     With udtBI
          .hwndOwner = hWnd
          .lpszTitle = lstrcat(Prompt, "")
          .ulFlags = 1
     End With
     
     'Select folder
     lngIDList = SHBrowseForFolder(udtBI)
     
     'Check if not cancelled
     If lngIDList Then
          
          'Get the selected path
          strPath = String$(260, 0)
          lngResult = SHGetPathFromIDList(lngIDList, strPath)
          
          'Free used memory
          Call CoTaskMemFree(lngIDList)
          
          'Remove null terminator
          strPath = StripNullChar(strPath)
     End If
     
     'Return the result
     SelectFolder = strPath
End Function

Private Function StripNullChar(ByRef FixedString As String) As String
     Dim NullPos As Long
     
     'Remove null terminator
     NullPos = InStr(FixedString, vbNullChar)
     If NullPos Then StripNullChar = left$(FixedString, NullPos - 1) Else StripNullChar = FixedString
End Function
