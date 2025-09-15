Attribute VB_Name = "modGeneral"
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


'ShellExecute constants
'http://msdn.microsoft.com/library/en-us/shellcc/platform/Shell/reference/functions/shellexecute.asp
Public Enum ENUM_SHELLWINDOWSTYLE
     SW_HIDE = 0
     SW_SHOW = 5
     SW_DEFAULT = 10
     SW_MAXIMIZED = 3
     SW_MINIMIZED = 2
     SW_MINIMIZED_NOACTIVE = 7
     SW_NA = 8
     SW_NOACTIVE = 4
     SW_NORMAL = 1
End Enum

'ShellExecuteEx mask constants
Public Enum ENUM_SHELLEXECUTEMASK
     SEE_MASK_CLASSKEY = &H3
     SEE_MASK_CLASSNAME = &H1
     SEE_MASK_CONNECTNETDRV = &H80
     SEE_MASK_DOENVSUBST = &H200
     SEE_MASK_FLAG_DDEWAIT = &H100
     SEE_MASK_FLAG_NO_UI = &H400
     SEE_MASK_HOTKEY = &H20
     SEE_MASK_ICON = &H10
     SEE_MASK_IDLIST = &H4
     SEE_MASK_INVOKEIDLIST = &HC
     SEE_MASK_NOCLOSEPROCESS = &H40
End Enum

'Statusbar Icons
Public Enum ENUM_STATUSICON
     STI_READY = 1
     STI_WAITING = 2
     STI_BUSY = 3
End Enum

'Map lump types
Public Enum ENUM_MAPLUMPTYPES
     ML_UNKNOWN = 0
     ML_REQUIRED = 1
     ML_RESPECTED = 2
     ML_NODEBUILD = 4
     ML_EMPTYALLOWED = 8
     ML_CUSTOMTEXT = 4096
     ML_CUSTOMACS = 8192
     ML_CUSTOMDEHACKED = 12288
     ML_CUSTOMFS = 16384
     ML_CUSTOMDED = 20480
     ML_CUSTOM = 61440             'All bits for the custom types
End Enum

'Saving modes
Public Enum ENUM_SAVEMODES
     SM_SAVE
     SM_SAVEAS
     SM_SAVEINTO
     SM_EXPORT
     SM_TEST
End Enum


'ShellExecuteEx structure
'http://msdn.microsoft.com/library/en-us/shellcc/platform/shell/reference/structures/shellexecuteinfo.asp
Private Type SHELLEXECUTEINFO
     cbSize As Long
     fMask As Long
     hWnd As Long
     lpVerb As String
     lpFile As String
     lpParameters As String
     lpDirectory As String
     nShow As Long
     hInstApp As Long
     lpIDList As Long
     lpClass As String
     hkeyClass As Long
     dwHotKey As Long
     hIcon As Long
     hProcess As Long
End Type

'GetVersionEx structure
'http://msdn.microsoft.com/library/en-us/sysinfo/base/getversionex.asp
Private Type OSVERSIONINFO
     dwOSVersionInfoSize As Long
     dwMajorVersion As Long
     dwMinorVersion As Long
     dwBuildNumber As Long
     dwPlatformId As Long
     szCSDVersion As String * 128
End Type

'POINT
Public Type POINT
     x As Long
     y As Long
End Type

'BOX
Public Type BOX
     left As Long
     top As Long
     right As Long
     bottom As Long
     front As Long
     back As Long
End Type

'Configuration types
Public Const BUILDER_CONFIG_TYPE As String = "Doom Builder Configuration"
Public Const SHORTCUTS_CONFIG_TYPE As String = "Doom Builder Shortcuts Configuration"
Public Const GAME_CONFIG_TYPE As String = "Doom Builder Game Configuration"
Public Const SCRIPT_CONFIG_TYPE As String = "Doom Builder Script Configuration"
Public Const SETTINGS_CONFIG_TYPE As String = "Doom Builder Map Settings Configuration"

'Declarations
Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExA" (ByRef ExecInfo As SHELLEXECUTEINFO) As Boolean
Public Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32.dll" (ByVal hWnd As Long, ByRef lpPoint As POINT) As Long
Public Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT) As Long
Public Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Public Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Public Declare Function SetSysColors Lib "user32.dll" (ByVal nChanges As Long, lpSysColor As Long, lpColorValues As Long) As Long
Public Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString1 As Long) As Long
Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Public Declare Function SetWindowPos Lib "user32.dll" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Sub ClipCursor Lib "user32.dll" (ByRef lpRect As Any)
Public Declare Sub GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef lpRect As RECT)
Public Declare Sub OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long)
Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)


'Topmost window constants
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

'Windows color indices
Public Const WCOLOR_APPWORKSPACE As Long = 12
Public Const WCOLOR_HIGHLIGHT As Long = 13

'Max files to keep in history
Public Const MAX_RECENT_FILES As Long = 8

'Windows messages
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_ACTIVATEAPP As Long = &H1C

'Messages handler number
Public Const GWL_WNDPROC As Long = -4

'Pi
Public Const pi As Single = 3.14159265358979
Public Const PiDiv As Single = 57.2957795130823
Public Const PiDivMul As Single = 0.017453292519943

'Directories
Public TempPath As String

'Program Status
Public Loading As Boolean
Public SplashDisplayed As Boolean
Public StatusDisplayed As Boolean
Public Scrolling As Boolean
Public PreviousWindowstate As Integer
Public RunningWindows2000 As Boolean
Public ScriptEditor As Boolean

'Configuration
Private Configfile As clsConfiguration
Public Config As Dictionary

'Game configurations by their title
Public AllGameConfigs As Dictionary

'Last dialog settings
Public LastFindType As Long

'Misc
Public CustomColors(0 To 15) As Long
Public OptionsCancelled As Boolean
Public CurrentShiftMask As Long


Public Sub CopyBytes(ByRef source() As Byte, ByRef target() As Byte, ByVal sourceoffset As Long, ByVal targetoffset As Long, ByVal Count As Long)
     Dim i As Long
     For i = 0 To Count - 1
          target(targetoffset + i) = source(sourceoffset + i)
     Next i
End Sub


Public Sub FillBytes(ByRef target() As Byte, ByVal offset As Long, ByVal Count As Long, ByVal Value As Byte)
     Dim i As Long
     For i = 0 To Count - 1
          target(offset + i) = Value
     Next i
End Sub



Public Sub CorrectDefaultTextures()
     
     'Check if textures are available
     If Not (alltextures Is Nothing) Then
          
          'This sets the first known texture on the default textures where none specified
          If (IsTextureName(Config("defaulttexture")("upper")) = False) Then Config("defaulttexture")("upper") = alltextures.Keys(0)
          If (IsTextureName(Config("defaulttexture")("middle")) = False) Then Config("defaulttexture")("middle") = alltextures.Keys(0)
          If (IsTextureName(Config("defaulttexture")("lower")) = False) Then Config("defaulttexture")("lower") = alltextures.Keys(0)
     End If
End Sub


Public Sub InitializeStartupDefaults()
     Dim i As Long
     
     'Initialize databases
     Set selected = New Dictionary
     Set dragselected = New Dictionary
     
     'Always have one entry here
     ReDim changedlines(0)
     numchangedlines = 0
     
     'Go for all custom colors
     For i = 0 To 15
          
          'Set color default
          CustomColors(i) = RGB(255 - i * 17, 255 - i * 17, 255 - i * 17)
     Next i
     
     'Default error checking settings
     IgnoreWarningsOption = vbUnchecked
     InvalidTexturesOption = vbChecked
     LineErrorsOption = vbChecked
     MissingTexturesOption = vbUnchecked
     PlayerStartsOption = vbChecked
     UnclosedSectorsOption = vbChecked
     VertexErrorsOption = vbChecked
     ZeroLengthLinesOption = vbChecked
     ThingErrorsOption = vbChecked
End Sub

Public Function NextPowerOf2(ByVal Value As Long) As Long
     Dim p As Long
     Dim v As Long
     
     'Start with power 0
     p = 0
     
     Do
          v = 2 ^ p
          p = p + 1
          
     'Continue until power is same or higher than value
     Loop Until (v >= Value)
     
     'Return result
     NextPowerOf2 = v
End Function

Public Sub SetTopMostWindow(hWnd As Long, Topmost As Boolean)
     
     'Make at top?
     If (Topmost = True) Then
          SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     Else
          SetWindowPos hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
     End If
End Sub
Public Sub CenterViewAt(ByRef target As RECT, ByVal Zoom As Boolean, Optional ByVal ZoomBorder As Long = 30, Optional ByVal DefaultZoom As Single = 0.75, Optional MaxZoom As Single = 2)
     Dim WidthZoom As Single, HeightZoom As Single
     Dim ZScreenWidth As Long, ZScreenHeight As Long
     Dim ZZoom As Single
     Dim TargetRect As RECT
     
     'NOTE: Target must be in mappixel coordinates
     
     'Copy the rect
     TargetRect = target
     
     'Zoom to target?
     If (Zoom) Then
          
          'Check if the rect has a size
          If (Abs(TargetRect.right - TargetRect.left) > 0) Or (Abs(TargetRect.bottom - TargetRect.top) > 0) Then
               
               'Calculate Zoom
               If (Abs(TargetRect.right - TargetRect.left) > 0) Then WidthZoom = (ScreenWidth - ZoomBorder * 2) / Abs(TargetRect.right - TargetRect.left) Else WidthZoom = MaxZoom
               If (Abs(TargetRect.bottom - TargetRect.top) > 0) Then HeightZoom = (ScreenHeight - ZoomBorder * 2) / Abs(TargetRect.bottom - TargetRect.top) Else HeightZoom = MaxZoom
               
               'Determine Zoom to use
               If (WidthZoom < HeightZoom) And (HeightZoom > 0) Then
                    ZZoom = WidthZoom
               ElseIf (HeightZoom > 0) Then
                    ZZoom = HeightZoom
               Else
                    ZZoom = DefaultZoom
               End If
          Else
               
               'Use default Zoom
               ZZoom = DefaultZoom
          End If
     Else
          
          'Use current zoom
          ZZoom = ViewZoom
     End If
     
     'Limit to max zoom
     If (ZZoom > MaxZoom) Then ZZoom = MaxZoom
     
     'Calculate ZScreenWidth and ZScreenHeight
     ZScreenWidth = ScreenWidth / ZZoom
     ZScreenHeight = ScreenHeight / ZZoom
     
     'Center rect in viewport of calculated zoom
     TargetRect.left = TargetRect.left - (ZScreenWidth - Abs(TargetRect.right - TargetRect.left)) * 0.5
     TargetRect.top = TargetRect.top - (ZScreenHeight - Abs(TargetRect.bottom - TargetRect.top)) * 0.5
     
     'Set the viewport
     ChangeView TargetRect.left, TargetRect.top, ZZoom
End Sub

Public Sub SelectAllText(ByRef Txt As Control)
     
     'Select text
     Txt.SelStart = 0
     Txt.SelLength = Len(Txt.Text)
End Sub


Public Function DetectWindowsVersion() As Boolean
     Dim Info As OSVERSIONINFO
     
     'Set the structure length
     Info.dwOSVersionInfoSize = Len(Info)
     
     'Get the Windows version
     GetVersionEx Info
     
     'Store True when running on a Windows 2000 based version or newer
     RunningWindows2000 = (Info.dwMajorVersion >= 5)
End Function

Public Sub ErrorLog_DisplayAndFlush()
     
     'Check if errors occurred
     If (Trim$(frmErrorLog.txtErrors.Text) <> "") Then
          
          'Add empty lines
          frmErrorLog.txtErrors = " " & vbCrLf & frmErrorLog.txtErrors & " "
          
          'Make the default sound
          Beep
          
          'Show the errors and warnings dialog
          frmErrorLog.txtErrors.SelStart = Len(frmErrorLog.txtErrors.Text)
          frmErrorLog.Show 1, frmMain
     End If
     
     'Unload the dialog
     Unload frmErrorLog
     Set frmErrorLog = Nothing
End Sub

Public Sub ErrorLog_Flush()
     
     'Unload the dialog
     Unload frmErrorLog
     Set frmErrorLog = Nothing
End Sub


Public Sub ErrorLog_Add(ByRef Message As String, ByVal critical As Boolean)
     
     'Report error
     frmErrorLog.txtErrors = frmErrorLog.txtErrors & " " & Message & vbCrLf
     
     'Check if critical icon must be shown
     If (critical) Then
          
          'Set icon when not already set
          If (frmErrorLog.imgIcon.Picture <> frmErrorLog.imgCritical.Picture) Then Set frmErrorLog.imgIcon.Picture = frmErrorLog.imgCritical.Picture
     End If
End Sub

Public Sub ErrorLog_Load()
     
     'Load error dialog
     Load frmErrorLog
End Sub


Public Sub FillThingsList(ByRef lstThings As ListView)
     Dim ThingCollection As Dictionary
     Dim Cats() As Variant
     Dim CurCat As String
     Dim ThingKeys As Variant
     Dim CurKey As String
     Dim c As Long
     Dim t As Long
     Dim nt As ListItem
     
     'Get the categories
     Cats = mapconfig("thingtypes").Keys
     
     'Go through all categories
     For c = LBound(Cats) To UBound(Cats)
          
          'Get category
          CurCat = Cats(c)
          
          'Get the collection
          Set ThingCollection = mapconfig("thingtypes")(CurCat)
          
          'Get things
          ThingKeys = ThingCollection.Keys
          
          'Go for all things
          For t = LBound(ThingKeys) To UBound(ThingKeys)
               
               'Get the key
               CurKey = ThingKeys(t)
               
               'Check if not one of the category properties
               If IsNumeric(CurKey) Then
                    
                    'Add thing to list
                    Set nt = lstThings.ListItems.Add(, "T" & CurKey, ThingCollection(CurKey)("title"), , ThingCollection("color") + 1)
                    
                    'Add stuff
                    With nt
                         .tag = CStr(CurKey)
                         .ToolTipText = .Text
                         .ListSubItems.Add , , ThingCollection("title")
                         .ListSubItems.Add , , Space$(5 - Len(CStr(CurKey))) & CStr(CurKey)
                    End With
               End If
          Next t
     Next c
     
     'Sort the list
     On Error Resume Next
     lstThings.SortKey = Abs(Val(Config("thingssort"))) - 1
     lstThings.SortOrder = Abs(Val(Config("thingssort")) < 0)
     On Error GoTo 0
End Sub


Public Sub FillThingsTree(ByRef trvThings As TreeView)
     Dim ThingCollection As Dictionary
     Dim Cats() As Variant
     Dim CurCat As String
     Dim ThingKeys As Variant
     Dim CurKey As String
     Dim c As Long
     Dim t As Long
     
     'Get the categories
     Cats = mapconfig("thingtypes").Keys
     
     'Go through all categories
     For c = LBound(Cats) To UBound(Cats)
          
          'Get category
          CurCat = Cats(c)
          
          'Get the collection
          Set ThingCollection = mapconfig("thingtypes")(CurCat)
          
          'Add category to list
          trvThings.nodes.Add , , CurCat, ThingCollection("title"), ThingCollection("color") + 1
          If (Val(ThingCollection("sort")) = 1) Then trvThings.nodes(CurCat).Sorted = True Else trvThings.nodes(CurCat).Sorted = False
          
          'Get things
          ThingKeys = ThingCollection.Keys
          
          'Go for all things
          For t = LBound(ThingKeys) To UBound(ThingKeys)
               
               'Get the key
               CurKey = ThingKeys(t)
               
               'Check if not one of the category properties
               If IsNumeric(CurKey) Then
                    
                    'Add thing to list
                    trvThings.nodes.Add CurCat, tvwChild, "T" & CurKey, ThingCollection(CurKey)("title"), ThingCollection("color") + 1
                    trvThings.nodes("T" & CurKey).tag = CurKey
               End If
          Next t
     Next c
End Sub


Public Function GetEpisodeNum() As Long
     
     'Check if map is in E#M# format
     If (maplumpname Like "E#M#") Then
          
          'Get the Episode number
          GetEpisodeNum = Val(Mid$(maplumpname, 2, 1))
     Else
          
          'Return 1
          GetEpisodeNum = 1
     End If
End Function

Public Function GetMapNum() As Long
     
     'Check if map is in E#M# format
     If (maplumpname Like "E#M#") Then
          
          'Get the Map number
          GetMapNum = Mid$(maplumpname, 4, 1)
          
     'Check if map is in MAP## format
     ElseIf (maplumpname Like "MAP##") Then
          
          'Get the Map number
          GetMapNum = Val(Mid$(maplumpname, 4, 2))
     Else
          
          'Return 1
          GetMapNum = 1
     End If
End Function


Public Function GetNearestTextureName(ByVal PartName As String) As String
     Dim i As Long
     Dim Names As Variant
     
     'Dont find one when no part given
     If (LenB(PartName) > 0) Then
          
          'Find the first name that partly matches
          Names = textures.Keys
          For i = 0 To (textures.Count - 1)
               
               'Check if it matches
               If (StrComp(PartName, left$(Names(i), Len(PartName)), vbTextCompare) = 0) Then
                    
                    'This matches, return the name
                    GetNearestTextureName = Names(i)
                    Exit Function
               End If
          Next i
          
          'Nothing found, return given name
          GetNearestTextureName = PartName
     End If
End Function

Public Sub CompleteTextureName(ByVal KeyCode As Integer, ByVal Shift As Integer, ByRef Txt As TextBox)
     Dim nTexture As String
     Dim SelStart As Long
     
     'No CTRL or ALT
     If ((Shift And vbCtrlMask) = 0) And ((Shift And vbAltMask) = 0) Then
          
          'Any 'usual' key?
          If ((KeyCode >= vbKeyA) And (KeyCode <= vbKeyZ)) Or _
             ((KeyCode >= vbKey0) And (KeyCode <= vbKey9)) Or _
             (KeyCode = 189) Or (KeyCode = 187) Or (KeyCode = 219) Or _
             (KeyCode = 221) Or (KeyCode = 191) Or (KeyCode = 220) Or _
             (KeyCode = 111) Or (KeyCode = 106) Or (KeyCode = 109) Or _
             (KeyCode = 107) Then
               
               'Anything typed?
               If (Txt.Text <> "") Then
               
                    'Find the name of the first matching texture
                    nTexture = GetNearestTextureName(Txt.Text)
                    
                    'Anything close?
                    If (nTexture <> "") Then
                         
                         'Apply texture name and select it
                         With Txt
                              SelStart = Len(.Text)
                              .Text = nTexture
                              .SelStart = SelStart
                              .SelLength = Len(.Text) - SelStart
                         End With
                    End If
               End If
          End If
     End If
End Sub


Public Sub CompleteFlatName(ByVal KeyCode As Integer, ByVal Shift As Integer, ByRef Txt As TextBox)
     Dim nFlat As String
     Dim SelStart As Long
     
     'No CTRL or ALT
     If ((Shift And vbCtrlMask) = 0) And ((Shift And vbAltMask) = 0) Then
          
          'Any 'usual' key?
          If ((KeyCode >= vbKeyA) And (KeyCode <= vbKeyZ)) Or _
             ((KeyCode >= vbKey0) And (KeyCode <= vbKey9)) Or _
             (KeyCode = 189) Or (KeyCode = 187) Or (KeyCode = 219) Or _
             (KeyCode = 221) Or (KeyCode = 191) Or (KeyCode = 220) Or _
             (KeyCode = 111) Or (KeyCode = 106) Or (KeyCode = 109) Or _
             (KeyCode = 107) Then
               
               'Anything typed?
               If (Txt.Text <> "") Then
               
                    'Find the name of the first matching flat
                    nFlat = GetNearestFlatName(Txt.Text)
                    
                    'Anything close?
                    If (nFlat <> "") Then
                         
                         'Apply texture name and select it
                         With Txt
                              SelStart = Len(.Text)
                              .Text = nFlat
                              .SelStart = SelStart
                              .SelLength = Len(.Text) - SelStart
                         End With
                    End If
               End If
          End If
     End If
End Sub



Public Function GetNearestFlatName(ByVal PartName As String) As String
     Dim i As Long
     Dim Names As Variant
     
     'Dont find one when no part given
     If (LenB(PartName) > 0) Then
          
          'Find the first name that partly matches
          Names = flats.Keys
          For i = 0 To (flats.Count - 1)
               
               'Check if it matches
               If (StrComp(PartName, left$(Names(i), Len(PartName)), vbTextCompare) = 0) Then
                    
                    'This matches, return the name
                    GetNearestFlatName = Names(i)
                    Exit Function
               End If
          Next i
          
          'Nothing found, return given name
          GetNearestFlatName = PartName
     End If
End Function


Public Function SelectAction(ByVal LinedefType As String, ByRef Parent As Form) As String
     Dim GenEffect As Long
     Dim c As Long, i As Long
     Dim Bit As Long
     
     'Load types dialog
     Load frmLinedefType
     
     'Select current linedef type
     If (LenB(LinedefType) <> 0) Then
          
          'Check if its generalized
          If IsGenLinedefEffect(LinedefType) Then
               
               'Show the panel by default
               frmLinedefType.tbsPanel.Tabs("generalized").selected = True
               
               'Set the category
               frmLinedefType.cmbCategory.ListIndex = GetGenLinedefCategoryIndex(LinedefType)
               
               'Get ungeneralized effect flags
               GenEffect = UngenLinedefEffect(LinedefType)
               
               'Go for all combos
               For c = 0 To 7
                    
                    'Check if available
                    If frmLinedefType.cmbOption(c).Enabled Then
                         
                         'Go for all items
                         For i = 0 To (frmLinedefType.cmbOption(c).ListCount - 1)
                              
                              'Get the bit value
                              Bit = frmLinedefType.cmbOption(c).ItemData(i)
                              
                              'Check if this bit is set in the flags
                              If (GenEffect And Bit) = Bit Then
                                   
                                   'Select this item in the combo
                                   frmLinedefType.cmbOption(c).ListIndex = i
                              End If
                         Next i
                    End If
               Next c
          Else
               
               'Do not give an error when the item cant be found
               On Local Error Resume Next
               frmLinedefType.lstTypes.ListItems("L" & LinedefType).selected = True
               frmLinedefType.lstTypes.ListItems("L" & LinedefType).EnsureVisible
               frmLinedefType.trvTypes.nodes("L" & LinedefType).selected = True
               frmLinedefType.trvTypes.nodes("L" & LinedefType).EnsureVisible
               On Local Error GoTo 0
          End If
     End If
     
     'Show linedef types dialog
     frmLinedefType.Show 1, Parent
     
     'Check if not cancelled
     If (frmLinedefType.tag = "1") Then
          
          'Check if chosen for generlized linedef
          If (frmLinedefType.tbsPanel.SelectedItem.Key = "generalized") Then
               
               'Go for all combos
               GenEffect = 0
               For c = 0 To 7
                    
                    'Check if available
                    If frmLinedefType.cmbOption(c).Enabled Then
                         
                         'Add the bits
                         GenEffect = GenEffect Or frmLinedefType.cmbOption(c).ItemData(frmLinedefType.cmbOption(c).ListIndex)
                    End If
               Next c
               
               'Make the generalized value
               LinedefType = MakeGenLinedefEffect(GenEffect, frmLinedefType.cmbCategory.ListIndex)
          Else
               
               'Apply selection
               On Local Error Resume Next
               If (Val(Config("linestree")) = vbUnchecked) Then
                    LinedefType = Trim$(frmLinedefType.lstTypes.SelectedItem.Text)
               Else
                    LinedefType = Trim$(frmLinedefType.trvTypes.SelectedItem.tag)
               End If
               On Local Error GoTo 0
          End If
     End If
     
     'Unload dialog
     Unload frmLinedefType: Set frmLinedefType = Nothing
     
     'Return new type
     SelectAction = LinedefType
End Function

Public Function SelectSectorEffect(ByVal SectorEffect As String, ByRef Parent As Form) As String
     Dim c As Long, i As Long
     Dim Bit As Long, GenEffect As Long
     
     'Load types dialog
     Load frmSectorType
     
     'Select current sector type
     If (LenB(SectorEffect) <> 0) Then
          
          'Check if using generalized sector effects
          If (Val(mapconfig("generalizedsectors")) <> 0) Then
               
               'Go for all combos
               For c = 0 To 7
                    
                    'Check if available
                    If frmSectorType.cmbOption(c).Enabled Then
                         
                         'Go for all items
                         For i = 0 To (frmSectorType.cmbOption(c).ListCount - 1)
                              
                              'Get the bit value
                              Bit = frmSectorType.cmbOption(c).ItemData(i)
                              
                              'Check if this bit is set in the flags
                              If (Val(SectorEffect) And Bit) = Bit Then
                                   
                                   'Select this item in the combo
                                   frmSectorType.cmbOption(c).ListIndex = i
                              End If
                         Next i
                    End If
               Next c
          End If
          
          'Do not give an error when the item cant be found
          On Local Error Resume Next
          frmSectorType.lstTypes.ListItems("L" & SectorEffect).selected = True
          
          'Check if an error occurred (item doesnt exist)
          If (Err.number <> 0) Then
               
               'Show second panel if using generalized sector effects
               If (Val(mapconfig("generalizedsectors")) <> 0) Then frmSectorType.tbsPanels.Tabs("generalized").selected = True
          Else
               
               'Show the selected item
               frmSectorType.lstTypes.ListItems("L" & SectorEffect).EnsureVisible
          End If
          On Local Error GoTo 0
     End If
     
     'Show sector types dialog
     frmSectorType.Show 1, Parent
     
     'Check if not cancelled
     If (frmSectorType.tag = "1") Then
          
          'Check if chosen for generlized linedef
          If (frmSectorType.tbsPanels.SelectedItem.Key = "generalized") Then
               
               'Go for all combos
               GenEffect = 0
               For c = 0 To 7
                    
                    'Check if available
                    If frmSectorType.cmbOption(c).Enabled Then
                         
                         'Add the bits
                         GenEffect = GenEffect Or frmSectorType.cmbOption(c).ItemData(frmSectorType.cmbOption(c).ListIndex)
                    End If
               Next c
               
               'Make the generalized value
               SectorEffect = GenEffect
          Else
               
               'Apply selection
               SectorEffect = Trim$(frmSectorType.lstTypes.SelectedItem.Text)
          End If
     End If
     
     'Unload dialog
     Unload frmSectorType
     Set frmSectorType = Nothing
     
     'Return the new value
     SelectSectorEffect = SectorEffect
End Function

Public Function SelectTexture(ByVal TextureName As String, ByRef Parent As Form) As String
     
     'Load dialog
     Load frmTextureBrowse
     frmTextureBrowse.Initialize False
     
     'Select this texture
     frmTextureBrowse.SetSelection TextureName
     
     'Show dialog
     frmTextureBrowse.Show 1, Parent
     
     'Set new texture if not cancelled
     If (frmTextureBrowse.tag = "1") Then SelectTexture = frmTextureBrowse.SelectedName Else SelectTexture = TextureName
     
     'Unload dialog
     Unload frmTextureBrowse
     Set frmTextureBrowse = Nothing
End Function


Public Function SelectFlat(ByVal FlatName As String, ByRef Parent As Form) As String
     
     'Load dialog
     Load frmTextureBrowse
     frmTextureBrowse.Initialize True
     
     'Select this texture
     frmTextureBrowse.SetSelection FlatName
     
     'Show dialog
     frmTextureBrowse.Show 1, Parent
     
     'Set new texture if not cancelled
     If (frmTextureBrowse.tag = "1") Then SelectFlat = frmTextureBrowse.SelectedName Else SelectFlat = FlatName
     
     'Unload dialog
     Unload frmTextureBrowse
     Set frmTextureBrowse = Nothing
End Function


Public Function SelectThing(ByVal thingtype As String, ByRef Parent As Form) As String
     
     'Load dialog
     Load frmThingType
     
     'Select this thing
     frmThingType.HighlightThing Val(thingtype)
     
     'Show dialog
     frmThingType.Show 1, Parent
     
     'Set new thing if not cancelled
     If (frmThingType.tag = "1") Then
          SelectThing = Val(frmThingType.lstThings.tag)
     Else
          SelectThing = thingtype
     End If
     
     'Unload dialog
     Unload frmThingType
     Set frmThingType = Nothing
End Function

Public Function StringOf(ByVal ptrString As Long, Optional ByVal Length As Long) As String
     
     'Get the string length from pointer
     If (Length = 0) Then Length = lstrlen(ptrString)
     
     'Allocate string for VB
     StringOf = Space$(Length)
     
     'Set the string
     CopyMemory ByVal StringOf, ByVal ptrString, Length
End Function


Public Sub AddRecentFile(ByVal filepathname As String)
     Dim i As Long
     
     'Go backwards for all recent files
     'and move them down (1 will be lost)
     For i = (MAX_RECENT_FILES - 1) To 1 Step -1
          
          'Check if set
          If Config("recent").Exists(CStr(i)) Then
               
               'Move it down by 1
               Config("recent")(CStr(i + 1)) = Config("recent")(CStr(i))
          End If
     Next i
     
     'Add the new file
     Config("recent")(CStr(1)) = filepathname
End Sub

Public Function ATan2(x As Single, y As Single) As Single
     
     Select Case x
          Case Is > 0
               If y > 0 Then
                    ATan2 = Atn(y / x)
               Else
                    ATan2 = Atn(y / x) + pi + pi
               End If
          Case 0
               If y > 0 Then
                    ATan2 = pi * 0.5
               Else
                    ATan2 = pi * 1.5
               End If
          Case Is < 0
               ATan2 = Atn(y / x) + pi
     End Select
End Function

Public Sub ChangeView(ByVal offsetx As Long, ByVal offsety As Long, ByVal Zoom As Single)
     
     'Not allowed during 3D mode
     If (Running3D) Then Exit Sub
     
     'Check if the zoom changes
     If (Zoom <> ViewZoom) Then
          
          'Terminate last thing pointer
          DestroyBitmapPointer ThingBitmapData
          
          'Determine thing size
          If (Zoom > 0.6) Then
               thingsize = 3
          ElseIf (Zoom > 0.3) Then
               thingsize = 2
          ElseIf (Zoom > 0.2) Then
               thingsize = 1
          Else
               thingsize = 0
          End If
          
          'Get a pointer to the new thing bitmap
          CreateBitmapPointer frmMain.picThings(thingsize), ThingBitmapData, ThingDescriptor
     End If
     
     'Keep the view
     ViewLeft = offsetx
     ViewTop = offsety
     ViewZoom = Zoom
     
     'Set the scalemode
     With ScreenTarget
          .ScaleMode = vbUser
          .ScaleLeft = offsetx
          .ScaleTop = offsety
          .ScaleWidth = ScreenWidth / Zoom
          .ScaleHeight = ScreenHeight / Zoom
     End With
     
     'Set the renderer viewport transformation
     Render_Scale offsetx, offsety, Zoom
     
     'Change the vertex block size depending on the zoom
     vertexsize = (Config("vertexsize") - 1) + 1.8 * Zoom
     If vertexsize > 4 + (Config("vertexsize") - 1) Then vertexsize = 4 + (Config("vertexsize") - 1)
     
     'Change the length of the linedef indicators
     If (Config("indicatorscaled")) Then
          indicatorsize = Config("indicatorsize")
     Else
          indicatorsize = (ScreenTarget.ScaleWidth / ScreenWidth + 0.1) * Config("indicatorsize")
     End If
End Sub

Public Sub CleanUpTemporaries()
     On Local Error Resume Next
     
     'Kill if there are any temporary files
     If (LenB(Dir(TempPath & "wad*.tmp")) <> 0) Then Kill TempPath & "wad*.tmp"
End Sub

Public Function Combined(ByRef Original As Dictionary, ByRef Patch As Dictionary, Optional ByRef target As Dictionary) As Dictionary
     Dim OriginalKeys As Variant
     Dim PatchKeys As Variant
     Dim i As Long
     
     'Check if a target is given
     If (target Is Nothing) Then
          
          'Create new target dictionary
          Set Combined = New Dictionary
     Else
          
          'Use target dictionary
          Set Combined = target
     End If
     
     'Get the keys of both dictionaries
     OriginalKeys = Original.Keys
     PatchKeys = Patch.Keys
     
     'Go for all original keys
     For i = LBound(OriginalKeys) To UBound(OriginalKeys)
          
          'Remove if already exists in target
          If (Combined.Exists(CStr(OriginalKeys(i)))) Then Combined.Remove CStr(OriginalKeys(i))
          
          'Check if exists in Patch
          If (Patch.Exists(CStr(OriginalKeys(i)))) Then
               
               'Check if this is another dictionary
               If (IsObject(Original(CStr(OriginalKeys(i))))) Then
                    
                    'Add combination from Original and Patch
                    Combined.Add CStr(OriginalKeys(i)), Combined(Original(CStr(OriginalKeys(i))), Patch(CStr(OriginalKeys(i))))
               Else
                    
                    'Copy value from Patch
                    Combined.Add CStr(OriginalKeys(i)), Patch(CStr(OriginalKeys(i)))
               End If
          Else
               
               'Check if this is another dictionary
               If (IsObject(Original(CStr(OriginalKeys(i))))) Then
                    
                    'Add deepcopy from Original
                    Combined.Add CStr(OriginalKeys(i)), Combined(Original(CStr(OriginalKeys(i))), New Dictionary)
               Else
                    
                    'Copy value from Original
                    Combined.Add CStr(OriginalKeys(i)), Original(CStr(OriginalKeys(i)))
               End If
          End If
     Next i
     
     'Go for all patch keys
     For i = LBound(PatchKeys) To UBound(PatchKeys)
          
          'Check if not already exists in target
          If (Combined.Exists(PatchKeys(i)) = False) Then
               
               'Check if this is another dictionary
               If (IsObject(Patch(CStr(PatchKeys(i))))) Then
                    
                    'Add deepcopy from Patch
                    Combined.Add CStr(PatchKeys(i)), Combined(Patch(CStr(PatchKeys(i))), New Dictionary)
               Else
                    
                    'Copy value from Patch
                    Combined.Add CStr(PatchKeys(i)), Patch(CStr(PatchKeys(i)))
               End If
          End If
     Next i
End Function

Public Function CommandSwitch(ByVal Switch As String) As Boolean
     Dim qs As Long, qe As Long
     Dim nCmd As String
     
     'Get quote positions
     qs = InStr(Command, """")
     qe = InStr(qs + 1, Command, """")
     
     'Cut the filename from string
     If (qs > 0) Then
          nCmd = left$(Command, qs - 1) & Mid$(Command, qe + 1)
     Else
          nCmd = Command
     End If
     
     'Check if the switch exists
     CommandSwitch = (InStr(1, nCmd, Switch, vbTextCompare) <> 0)
End Function

Public Function DetectNewGameConfigs() As Boolean
     On Local Error Resume Next
     Dim Filename As String
     Dim TempCfg As New clsConfiguration
     
     'Ensure the "iwads" item is a dictionary
     If (Config.Exists("iwads") = False) Then
          
          'Add the item
          Config.Add "iwads", New Dictionary
          
     'Ensure the "iwads" item is a dictionary
     ElseIf (VarType(Config("iwads")) <> vbObject) Then
     
          'Re-add the item
          Config.Remove "iwads"
          Config.Add "iwads", New Dictionary
     End If
     
     
     'Find first file
     Filename = Dir(App.Path & "\*.cfg")
     
     'Continue until no more files found
     Do Until (LenB(Filename) = 0)
          
          'Clear errors
          Err.Clear
          
          'Load this configuration file
          TempCfg.NewConfiguration
          TempCfg.LoadConfiguration App.Path & "\" & Filename
          
          'Check for errors
          If (Err.number = 0) Then
               
               'Check if this file is a game configuration
               If (TempCfg.ReadSetting("type", "") = GAME_CONFIG_TYPE) Then
                    
                    'Check for errors
                    If (Err.number = 0) Then
                         
                         'Check if this is the game has an IWAD entry in the config
                         If (Config("iwads").Exists(LCase$(Filename)) = False) Then
                              
                              'New configs found!
                              DetectNewGameConfigs = True
                              
                              'Leave search
                              Exit Do
                         End If
                    Else
                         
                         'Could not load this configuration
                         MsgBox "The configuration file " & Filename & " has errors and cannot be parsed." & "Syntax error on line " & TempCfg.CurrentScanLine & ".", vbExclamation
                    End If
               End If
          End If
          
          'Find next file
          Filename = Dir()
     Loop
End Function

Public Sub DisableMapEditing()
     Dim i As Long
     
     'Menu items
     With frmMain
          .itmFile(4).Enabled = False
          .itmFile(5).Enabled = False
          .itmFile(6).Enabled = False
          .itmFile(2).Enabled = False
          .itmFile(11).Enabled = False
          .itmFile(12).Enabled = False
          .itmFile(8).Enabled = False
          .itmFile(9).Enabled = False
          
          .mnuEdit.visible = False
          .itmEditUndo.Enabled = False
          .itmEditUndo.Caption = "&Undo"
          .itmEditRedo.Enabled = False
          .itmEditRedo.Caption = "&Redo"
          .itmEditMode(0).Enabled = False
          .itmEditMode(1).Enabled = False
          .itmEditMode(2).Enabled = False
          .itmEditMode(3).Enabled = False
          .itmEditMode(4).Enabled = False
          .itmEditMode(5).Enabled = False
          .itmEditMapOptions.Enabled = False
          .itmEditCut.Enabled = False
          .itmEditCopy.Enabled = False
          .itmEditPaste.Enabled = False
          .itmEditDelete.Enabled = False
          .itmEditFind.Enabled = False
          .itmEditReplace.Enabled = False
          .itmEditResize.Enabled = False
          .itmEditCenterView.Enabled = False
          
          .itmToolsFindErrors.Enabled = False
          .itmToolsClearTextures.Enabled = False
          .itmToolsFixTextures.Enabled = False
          .itmToolsFixZeroLinedefs.Enabled = False
          .itmToolsReloadResources.Enabled = False
          
          .mnuVertices.visible = False
          .mnuLines.visible = False
          .mnuSectors.visible = False
          .mnuThings.visible = False
          .mnuPrefabs.visible = False
          .mnuScripts.visible = False
     End With
     
     'Toolbar buttons
     With frmMain.tlbToolbar
          .Buttons("FileSaveMap").Enabled = False
          
          .Buttons("ModeMove").Enabled = False
          .Buttons("ModeVertices").Enabled = False
          .Buttons("ModeLines").Enabled = False
          .Buttons("ModeSectors").Enabled = False
          .Buttons("ModeThings").Enabled = False
          .Buttons("Mode3D").Enabled = False
          
          .Buttons("FileBuild").Enabled = False
          .Buttons("FileTest").Enabled = False
          
          .Buttons("EditUndo").Enabled = False
          .Buttons("EditRedo").Enabled = False
          .Buttons("EditGrid").Enabled = False
          .Buttons("EditSnap").Enabled = False
          .Buttons("EditStitch").Enabled = False
          .Buttons("EditFlipH").Enabled = False
          .Buttons("EditFlipV").Enabled = False
          .Buttons("EditRotate").Enabled = False
          .Buttons("EditResize").Enabled = False
          .Buttons("EditCenterView").Enabled = False
          
          .Buttons("LinesFlip").visible = False
          .Buttons("LinesCurve").visible = False
          .Buttons("SectorsJoin").visible = False
          .Buttons("SectorsMerge").visible = False
          .Buttons("SectorsGradientBrightness").visible = False
          .Buttons("SectorsGradientFloors").visible = False
          .Buttons("SectorsGradientCeilings").visible = False
          .Buttons("ThingsFilter").visible = False
          
          .Buttons("PrefabsInsert").Enabled = False
          .Buttons("PrefabsInsertPrevious").Enabled = False
          
     End With
     
     'Statusbar panels
     For i = 1 To frmMain.stbStatus.Panels.Count
          frmMain.stbStatus.Panels(i).visible = False
     Next i
     
     'Info panels
     With frmMain
          .fraBackSidedef.visible = False
          .fraFrontSidedef.visible = False
          .fraLinedef.visible = False
          .fraSector.visible = False
          .fraSectorCeiling.visible = False
          .fraSectorFloor.visible = False
          .fraThing.visible = False
          .fraVertex.visible = False
          
          .fraSVertex.visible = False
          .fraSLinedef.visible = False
          .fraSFrontSidedef.visible = False
          .fraSBackSidedef.visible = False
          .fraSSector.visible = False
          .fraSSectorCeiling.visible = False
          .fraSSectorFloor.visible = False
          .fraSThing.visible = False
          
          'Tooltip
          .picMap.ToolTipText = ""
          .lblMode.visible = False
          .lblBarText.Caption = ""
          .cmdToggleBar.visible = False
          .cmdToggleSBar.visible = False
     End With
End Sub

Public Sub DisplayStatus(ByRef StatusText As String)
     
     'Check if splash screen shown
     If SplashDisplayed Then
          
          'Check if the text changes at all
          If (frmSplash.lblStatus <> StatusText) Then
               
               'Update status on the splash screen
               frmSplash.lblStatus = StatusText
               DoEvents
          End If
     End If
     
     'Check if status scren shown
     If StatusDisplayed Then
          
          'Check if the text changes at all
          If (frmStatus.lblStatus <> StatusText) Then
               
               'Update status on the status screen
               frmStatus.lblStatus = StatusText
               frmStatus.lblStatus.Refresh
          End If
     End If
End Sub

Public Sub EnableMapEditing()
     Dim i As Long
     
     'Menu items
     With frmMain
          .itmFile(4).Enabled = True
          .itmFile(5).Enabled = True
          .itmFile(6).Enabled = True
          .itmFile(2).Enabled = True
          .itmFile(11).Enabled = True
          .itmFile(12).Enabled = True
          .itmFile(8).Enabled = True
          .itmFile(9).Enabled = True
          
          .mnuEdit.visible = True
          .itmEditMode(0).Enabled = True
          .itmEditMode(1).Enabled = True
          .itmEditMode(2).Enabled = True
          .itmEditMode(3).Enabled = True
          .itmEditMode(4).Enabled = True
          .itmEditMode(5).Enabled = True
          .itmEditMapOptions.Enabled = True
          .itmEditCut.Enabled = True
          .itmEditCopy.Enabled = True
          .itmEditPaste.Enabled = True
          .itmEditDelete.Enabled = True
          .itmEditFind.Enabled = True
          .itmEditReplace.Enabled = True
          .itmEditResize.Enabled = True
          .itmEditCenterView.Enabled = True
          
          .itmToolsFindErrors.Enabled = True
          .itmToolsClearTextures.Enabled = True
          .itmToolsFixTextures.Enabled = True
          .itmToolsFixZeroLinedefs.Enabled = True
          .itmToolsReloadResources.Enabled = True
          
          .mnuTools.visible = True
          .mnuPrefabs.visible = True
          
          'Infobar
          .cmdToggleBar.visible = True
          .cmdToggleSBar.visible = True
     End With
     
     'Toolbar buttons
     With frmMain.tlbToolbar
          .Buttons("FileSaveMap").Enabled = True
          
          .Buttons("ModeMove").Enabled = True
          .Buttons("ModeVertices").Enabled = True
          .Buttons("ModeLines").Enabled = True
          .Buttons("ModeSectors").Enabled = True
          .Buttons("ModeThings").Enabled = True
          .Buttons("Mode3D").Enabled = True
          
          .Buttons("FileBuild").Enabled = True
          .Buttons("FileTest").Enabled = True
          
          .Buttons("EditGrid").Enabled = True
          .Buttons("EditSnap").Enabled = True
          .Buttons("EditStitch").Enabled = True
          .Buttons("EditFlipH").Enabled = True
          .Buttons("EditFlipV").Enabled = True
          .Buttons("EditRotate").Enabled = True
          .Buttons("EditResize").Enabled = True
          .Buttons("EditCenterView").Enabled = True
          
          .Buttons("PrefabsInsert").Enabled = True
          .Buttons("PrefabsInsertPrevious").Enabled = True
     End With
     
     'Statusbar panels
     For i = 1 To frmMain.stbStatus.Panels.Count
          frmMain.stbStatus.Panels(i).visible = True
     Next i
     
     'Enable stuff for current mode
     frmMain.itmEditMode_Click CInt(mode)
     frmMain.lblMode.visible = True
End Sub

Public Function Execute(ByRef Filename As String, ByRef Parameters As String, ByVal WindowStyle As ENUM_SHELLWINDOWSTYLE, ByVal WaitForProcess As Boolean) As Boolean
     On Local Error Resume Next
     Dim ExecInfo As SHELLEXECUTEINFO
     
     'Check if we should add local path
     If (InStr(Filename, "\") = 0) And _
        (InStr(Filename, "/") = 0) And _
        (InStr(Filename, ":") = 0) Then
          
          'Add local path to filename
          Filename = App.Path & "\" & Filename
     End If
     
     'Make short path/file name
     If (Trim$(GetShortFileName(Filename)) <> "") Then Filename = GetShortFileName(Filename)
     
     'Remove .pif file if any (because it would override the way we run the program)
     If (Dir(left$(Filename, Len(Filename) - 4) & ".pif") <> "") Then Kill left$(Filename, Len(Filename) - 4) & ".pif"
     
     'Fill structure
     With ExecInfo
          .cbSize = Len(ExecInfo)
          .fMask = SEE_MASK_NOCLOSEPROCESS
          .lpFile = Filename
          .lpParameters = Parameters
          .lpDirectory = PathOf(Filename)
          .nShow = WindowStyle
     End With
     
     'Execute the file
     Execute = ShellExecuteEx(ExecInfo)
     
     'Check if we should wait
     If (WaitForProcess = True) And (ExecInfo.hProcess <> 0) Then
          
          'Wait for the process to end
          While WaitForSingleObject(ExecInfo.hProcess, 20): DoEvents: Wend
     End If
End Function

Private Sub FindGameConfigurations()
     On Local Error Resume Next
     Dim Filename As String
     Dim TempCfg As New clsConfiguration
     
     'Display status
     DisplayStatus "Loading game configurations..."
     
     'Create dictionary
     Set AllGameConfigs = New Dictionary
     
     'Find first file
     Filename = Dir(App.Path & "\*.cfg")
     
     'Continue until no more files found
     Do Until (LenB(Filename) = 0)
          
          'No errors
          Err.Clear
          
          'Load this configuration file
          TempCfg.NewConfiguration
          TempCfg.LoadConfiguration App.Path & "\" & Filename
          
          'Check if no errors during reading
          If (Err.number = 0) Then
               
               'Check if this file is a game configuration
               If (TempCfg.ReadSetting("type", "") = GAME_CONFIG_TYPE) Then
                    
                    'Add to database
                    AllGameConfigs.Add TempCfg.ReadSetting("game"), App.Path & "\" & Filename
               End If
          Else
               
               'Show warning
               MsgBox "The configuration file " & Filename & " has errors and cannot be parsed." & vbLf & "Syntax error on line " & TempCfg.CurrentScanLine & ": " & Err.Description, vbExclamation
          End If
          
          'Find next file
          Filename = Dir()
     Loop
End Sub

Public Function FindLumpIndex(ByRef WadFile As clsWAD, ByVal StartIndex As Long, ByVal LumpName As String, Optional ByVal Range As Long) As Long
     Dim i As Long
     Dim EndIndex As Long
     'Dim ll As Long
     
     'Leave when file is closed
     If (LenB(WadFile.Filename) = 0) Then Exit Function
     
     'Check if a range is given
     If (Range > 0) Then
          
          'Set the end to the range
          EndIndex = StartIndex + Range - 1
          
          'If the end hits end of file, set it to the end
          If (EndIndex > WadFile.LumpCount) Then EndIndex = WadFile.LumpCount
     Else
          
          'End at the end of table
          EndIndex = WadFile.LumpCount
     End If
     
     'Get the next lump
     LumpName = Padded$(UCase$(LumpName), 8)
     'll = Len(LumpName)
     For i = StartIndex To EndIndex
          
          'Check if this is the lump being searched
          If StrComp(WadFile.LumpnamePadded(i), LumpName, vbBinaryCompare) = 0 Then
               
               'Return this index
               FindLumpIndex = i
               
               'Leave the search
               Exit For
          End If
     Next i
End Function

Public Function FindValueInArray(ByRef LongsArray As Variant, ByRef Value As Variant) As Long
     On Local Error GoTo OutOfRange
     Dim i As Long
     
     'Go for all array items
     For i = LBound(LongsArray) To UBound(LongsArray)
          
          'Check if this is the value being searched
          If (LongsArray(i) = Value) Then
               
               'Return the index
               FindValueInArray = i
               
               'Leave here
               Exit Function
          End If
     Next i
     
OutOfRange:
     'Nothing found, return -1
     FindValueInArray = -1
End Function

Public Function GetCurrentIWADFile(Optional ByVal Gameconfig As String) As String
     Dim GameConfigFile As String
     
     'Get game config file
     If (LenB(Gameconfig) = 0) Then GameConfigFile = GetGameConfigFile(mapgame) Else GameConfigFile = GetGameConfigFile(Gameconfig)
     
     'Get the IWAD
     If (LenB(GameConfigFile) <> 0) Then GetCurrentIWADFile = Config("iwads")(LCase$(Dir(GameConfigFile)))
End Function

Public Function GetFileName(ByRef filepathname As String) As String
     Dim SeperatorNewPos As Long
     Dim SeperatorLastPos As Long
     
     'Get the last seperator position
     SeperatorLastPos = InStrRev(filepathname, "\")
     SeperatorNewPos = InStrRev(filepathname, "/")
     If SeperatorNewPos > SeperatorLastPos Then SeperatorLastPos = SeperatorNewPos
     SeperatorNewPos = InStrRev(filepathname, ":")
     If SeperatorNewPos > SeperatorLastPos Then SeperatorLastPos = SeperatorNewPos
     
     'Return the filename only
     GetFileName = Mid$(filepathname, SeperatorLastPos + 1)
End Function

Public Function GetGameConfigFile(ByVal Gameconfig As String) As String
     On Local Error Resume Next
     
     'Return the filename for this config
     If AllGameConfigs.Exists(Gameconfig) Then GetGameConfigFile = AllGameConfigs(Gameconfig)
End Function

Public Function GetGenLinedefCategory(ByVal effect As Long) As Dictionary
     Dim c As Long
     Dim Cats As Variant
     Dim Cat As Dictionary
     
     'Check if generalized linedefs are in use
     If (Val(mapconfig("generalizedlinedefs")) <> 0) Then
          
          'Go for all generalized type categories
          Cats = mapconfig("gen_linedeftypes").Items
          For c = LBound(Cats) To UBound(Cats)
               
               'Get category
               Set Cat = Cats(c)
               
               'Check if the effect lies in this range
               If (effect >= Cat("offset")) And (effect < Cat("offset") + Cat("length")) Then
                    
                    'Effect is in this category, return the object
                    Set GetGenLinedefCategory = Cat
                    
                    'Leave the search
                    Exit For
               End If
          Next c
     End If
End Function

Public Function GetGenLinedefCategoryIndex(ByVal effect As Long) As Long
     Dim c As Long
     Dim Cats As Variant
     Dim Cat As Dictionary
     
     'Check if generalized linedefs are in use
     If (Val(mapconfig("generalizedlinedefs")) <> 0) Then
          
          'Go for all generalized type categories
          Cats = mapconfig("gen_linedeftypes").Items
          For c = LBound(Cats) To UBound(Cats)
               
               'Get category
               Set Cat = Cats(c)
               
               'Check if the effect lies in this range
               If (effect >= Cat("offset")) And (effect < Cat("offset") + Cat("length")) Then
                    
                    'Effect is in this category, return the index
                    GetGenLinedefCategoryIndex = c
                    
                    'Leave the search
                    Exit For
               End If
          Next c
     End If
End Function

Public Sub GetLineSideSpot(ByVal ld As Long, ByVal distance As Single, ByVal front As Boolean, ByRef sx As Single, ByRef sy As Single)
     Dim lx As Single, ly As Single
     Dim bx As Single, by As Single
     Dim sl As Single
     
     'Calculate distance to middle of line
     lx = (vertexes(linedefs(ld).v2).x - vertexes(linedefs(ld).v1).x)
     ly = (-vertexes(linedefs(ld).v2).y + vertexes(linedefs(ld).v1).y)
          
     'Calculate middle of line
     bx = vertexes(linedefs(ld).v1).x + lx * 0.5
     by = -vertexes(linedefs(ld).v1).y + ly * 0.5
     
     'Get slope length for normalization
     sl = Sqr(lx * lx + ly * ly)
     
     'Calculate sector check spot
     If (sl <> 0) Then
          If front Then
               sx = bx - (ly / sl) * distance
               sy = by + (lx / sl) * distance
          Else
               sx = bx + (ly / sl) * distance
               sy = by - (lx / sl) * distance
          End If
     Else
          sx = bx
          sy = by
     End If
End Sub

Public Function GetMapLumpType(ByVal LumpName As String, Optional ByVal DetectMapHeader As Boolean = True) As ENUM_MAPLUMPTYPES
     Dim Lumpnames As Variant
     Dim lname As String
     Dim i As Long
     
     'Go for all defined lump names
     Lumpnames = mapconfig("maplumpnames").Keys
     For i = LBound(Lumpnames) To UBound(Lumpnames)
          
          'Get string
          lname = CStr(Lumpnames(i))
          If (lname = "~") And (DetectMapHeader) Then lname = maplumpname
          
          'Check if matches
          If StrComp(left$(LumpName, Len(lname)), lname, vbBinaryCompare) = 0 Then
               
               'Return definition
               GetMapLumpType = CInt(mapconfig("maplumpnames")(Lumpnames(i)))
               
               'Leave now
               Exit Function
          End If
     Next i
End Function

Public Function GetRecentFileIndex(ByVal filepathname As String) As Long
     Dim i As Long
     
     'Go for all recent files
     For i = 1 To MAX_RECENT_FILES
          
          'Check if set
          If Config("recent").Exists(CStr(i)) Then
               
               'Check if matches
               If (StrComp(Config("recent")(CStr(i)), filepathname, vbTextCompare) = 0) Then
                    
                    'Return this index now
                    GetRecentFileIndex = i
                    
                    'Leave the search
                    Exit For
               End If
          End If
     Next i
End Function

Function GetShortFileName(ByVal Filename As String) As String
     Dim buffer As String, Length As Long
     
     'Make buffer
     buffer = Space$(300)
     Length = GetShortPathName(Filename, buffer, Len(buffer))
     
     'Return the result
     GetShortFileName = left$(buffer, Length)
End Function

Public Function GetThingAngleDesc(ByVal ThingAngle As Long) As String
     Dim da As Long
     
     'Divide the angle to simpler integer
     da = CLng(ThingAngle / 45)
     
     'Return the correct description
     Select Case da
          Case 0: GetThingAngleDesc = "East"
          Case 1: GetThingAngleDesc = "North East"
          Case 2: GetThingAngleDesc = "North"
          Case 3: GetThingAngleDesc = "North West"
          Case 4: GetThingAngleDesc = "West"
          Case 5: GetThingAngleDesc = "South West"
          Case 6: GetThingAngleDesc = "South"
          Case 7: GetThingAngleDesc = "South East"
     End Select
End Function

Public Function GetThingTypeCategory(ByVal thingtype As Long) As String
     
     'Return nothing by default
     GetThingTypeCategory = ""
     
     'Check if this thing number is in this category
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this category
          GetThingTypeCategory = CStr(mapconfig("__things")(CStr(thingtype))("category"))
     End If
End Function

Public Function GetThingTypeCategoryIndex(ByVal thingtype As Long) As Long
     Dim i As Long
     Dim ThingCats As Variant
     
     'Return nothing by default
     GetThingTypeCategoryIndex = -1
     
     'Get all thing categories (keys)
     ThingCats = mapconfig("thingtypes").Keys
     
     'Go for all thing categories
     For i = LBound(ThingCats) To UBound(ThingCats)
          
          'Check if this thing number is in this category
          If (mapconfig("thingtypes")(ThingCats(i)).Exists(CStr(thingtype))) Then
               
               'Return this category
               GetThingTypeCategoryIndex = i
               
               'Leave the category search
               Exit For
          End If
     Next i
End Function


Public Function GetThingTypeDesc(ByVal thingtype As Long, Optional ByVal DefaultDesc As String = "Unknown") As String
     
     'Default return
     GetThingTypeDesc = DefaultDesc
     
     'Check if this thing number is in this category
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this description
          GetThingTypeDesc = mapconfig("__things")(CStr(thingtype))("title")
     End If
End Function

Public Function GetThingTypeSpriteName(ByVal thingtype As Long) As String
     Dim ThingCfg As Dictionary
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Get the thing config
          Set ThingCfg = mapconfig("__things")(CStr(thingtype))
          
          'Has a sprite been set?
          If (ThingCfg.Exists("sprite") = True) Then
               
               'Return this sprite name
               GetThingTypeSpriteName = ThingCfg("sprite")
          End If
     End If
End Function


Public Function GetThingWidth(ByVal thingtype As Long) As Long
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this width
          GetThingWidth = mapconfig("__things")(CStr(thingtype))("width")
     End If
End Function

Public Function GetThingHeight(ByVal thingtype As Long) As Long
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this height
          GetThingHeight = mapconfig("__things")(CStr(thingtype))("height")
     End If
End Function


Public Function GetThingError(ByVal thingtype As Long) As Long
     
     'Default return
     GetThingError = 1
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this errorlevel
          GetThingError = mapconfig("__things")(CStr(thingtype))("error")
     End If
End Function


Public Function GetThingHangs(ByVal thingtype As Long) As Long
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this hangs
          GetThingHangs = mapconfig("__things")(CStr(thingtype))("hangs")
     End If
End Function



Public Function GetThingBlocking(ByVal thingtype As Long) As Long
     
     'Check if this thing number is defined
     If (mapconfig("__things").Exists(CStr(thingtype))) Then
          
          'Return this blocking
          GetThingBlocking = mapconfig("__things")(CStr(thingtype))("blocking")
     End If
End Function



Public Function GetThingBlockingDesc(ByVal blocking As Long) As String
     
     'Return description
     Select Case blocking
          Case 0: GetThingBlockingDesc = "No"
          Case 1: GetThingBlockingDesc = "Completely"
          Case 2: GetThingBlockingDesc = "True-Height"
          Case Else: GetThingBlockingDesc = CStr(blocking) & "?"
     End Select
End Function

Public Sub GetWindowsTempPath()
     
     'Get windows temp path
     TempPath = Space$(256)
     GetTempPath 255, TempPath
     TempPath = Trim$(Replace(TempPath, vbNullChar, ""))
     If (right$(TempPath, 1) <> "\") Then TempPath = TempPath & "\"
End Sub

Public Function Hexadecimal(ByVal Value As Long, ByVal paddinglength As Long) As String
     On Local Error Resume Next
     Dim Hexvalue As String
     
     'Make hexadecimal value
     Hexvalue = Hex(Value)
     
     'Pad with zero's
     Hexadecimal = String$(paddinglength - Len(Hexvalue), "0") & Hexvalue
End Function

Public Function IsGenLinedefEffect(ByVal effect As Long) As Boolean
     Dim c As Long
     Dim Cats As Variant
     Dim Cat As Dictionary
     
     'Check if generalized linedefs are in use
     If (Val(mapconfig("generalizedlinedefs")) <> 0) Then
          
          'Go for all generalized type categories
          Cats = mapconfig("gen_linedeftypes").Items
          For c = LBound(Cats) To UBound(Cats)
               
               'Get category
               Set Cat = Cats(c)
               
               'Check if the effect lies in this range
               If (effect >= Cat("offset")) And (effect < Cat("offset") + Cat("length")) Then
                    
                    'Effect is in this category, return true
                    IsGenLinedefEffect = True
                    
                    'Leave the search
                    Exit For
               End If
          Next c
     End If
End Function

Public Function IsTextureName(ByVal TextureName As String) As Long
     
     'Check if this can be concidered a valid texture name
     'NOTE: The input must be unpadded!
     
     If (LenB(TextureName) <> 0) Then If (AscW(TextureName) <> 45) Then IsTextureName = True
End Function

Private Function LoadConfiguration() As Boolean
     On Local Error GoTo ConfigError
     Dim ShortcutKeys As Variant
     Dim ShortcutsConfig As New clsConfiguration
     Dim i As Long
     
     'Show status
     DisplayStatus "Loading configuration..."
     
     'Create object
     Set Configfile = New clsConfiguration
     
     'Load configuration from file
     Configfile.LoadConfiguration App.Path & "\Builder.cfg"
     
     'Reference the Config to the object orientated file structure
     Set Config = Configfile.Root(True)
     
     'Check if configuration type is specified
     If Not Config.Exists("type") Then
          
          'Add configuration type
          Config.Add "type", BUILDER_CONFIG_TYPE
     Else
          
          'Valid configuration
          If Config("type") <> BUILDER_CONFIG_TYPE Then MsgBox "Warning: The configuration file 'builder.cfg' is not a valid configuration." & vbLf & "If you experience any problems, you may need to reinstall Doom Builder.", vbCritical
     End If
     
     'Load the shortcuts configuration for verification
     ShortcutsConfig.LoadConfiguration App.Path & "\Shortcuts.cfg"
     
     'Check if configuration type is specified
     If ShortcutsConfig.ReadSetting("type", "") <> SHORTCUTS_CONFIG_TYPE Then
          
          'Invalid configuration
          MsgBox "Warning: The configuration file 'shortcuts.cfg' is not a valid configuration." & vbLf & "If you experience any problems, you may need to reinstall Doom Builder.", vbCritical
     End If
     
     
     'Go for all shortcuts
     ShortcutKeys = Config("shortcuts").Keys
     For i = LBound(ShortcutKeys) To UBound(ShortcutKeys)
          
          'Make sure this value is of integer type
          Config("shortcuts")(ShortcutKeys(i)) = Val(Config("shortcuts")(ShortcutKeys(i)))
     Next i
     
     'Correct values that changed from previous versions and
     'add missing values from previous versions
     If (Val(Config("videoviewdistance")) < 500) Then Config("videoviewdistance") = 3000
     If (Config.Exists("keywordslistwidth") = False) Then Config.Add "keywordslistwidth", 200
     If (Config.Exists("insertfulldefinition") = False) Then Config.Add "insertfulldefinition", vbChecked
     If (Config("palette").Exists("CLR_SCRIPTBACKGROUND") = False) Then Config("palette").Add "CLR_SCRIPTBACKGROUND", 16777215
     If (Config("palette").Exists("CLR_SCRIPTTEXT") = False) Then Config("palette").Add "CLR_SCRIPTTEXT", 0
     If (Config("palette").Exists("CLR_SCRIPTCOMMENT") = False) Then Config("palette").Add "CLR_SCRIPTCOMMENT", 8947848
     If (Config("palette").Exists("CLR_SCRIPTKEYWORD") = False) Then Config("palette").Add "CLR_SCRIPTKEYWORD", 255
     If (Config("palette").Exists("CLR_SCRIPTSTRING") = False) Then Config("palette").Add "CLR_SCRIPTSTRING", 32768
     If (Config("palette").Exists("CLR_SCRIPTLINENUMBERS") = False) Then Config("palette").Add "CLR_SCRIPTLINENUMBERS", 12632256
     If (Config("palette").Exists("CLR_SCRIPTCONSTANT") = False) Then Config("palette").Add "CLR_SCRIPTCONSTANT", 8388608
     If (Config("palette").Exists("CLR_THINGTAG") = False) Then Config("palette").Add "CLR_THINGTAG", Config("palette")("CLR_LINETAG")
     If (Config.Exists("scriptwindow") = False) Then Config.Add "scriptwindow", New Dictionary
     If (Config("shortcuts").Exists("select1sided") = False) Then Config("shortcuts").Add "select1sided", 131121
     If (Config("shortcuts").Exists("select2sided") = False) Then Config("shortcuts").Add "select2sided", 131122
     If (Config("shortcuts").Exists("linesautoalign") = False) Then Config("shortcuts").Add "linesautoalign", 65
     If (Config("buildnodes") = 1) Then Config("buildnodes") = 0
     If (Config("buildnodes") = 3) Then Config("buildnodes") = 2
     If (Config.Exists("copytagdraw") = False) Then Config.Add "copytagdraw", vbUnchecked
     If (Config.Exists("copytagpaste") = False) Then Config.Add "copytagpaste", vbUnchecked
     If (Config.Exists("buildexportcompression") = False) Then Config.Add "buildexportcompression", vbChecked
     If (Config("shortcuts").Exists("helpfaq") = False) Then Config("shortcuts").Add "helpfaq", 112
     If (Config("shortcuts").Exists("editmove") = False) Then Config("shortcuts").Add "editmove", 77
     If (Config("shortcuts").Exists("mode3dcopyoffsets") = False) Then Config("shortcuts").Add "mode3dcopyoffsets", 65603
     If (Config("shortcuts").Exists("mode3dpasteoffsets") = False) Then Config("shortcuts").Add "mode3dpasteoffsets", 65622
     If (Config("shortcuts").Exists("editquickmove") = False) Then Config("shortcuts").Add "editquickmove", 32
     If (Config("shortcuts").Exists("editcenterview") = False) Then Config("shortcuts").Add "editcenterview", 131104
     If (Config("shortcuts").Exists("togglebar") = False) Then Config("shortcuts").Add "togglebar", 192
     If (Config.Exists("storeeditinginfo") = False) Then Config.Add "storeeditinginfo", vbChecked
     If (Config.Exists("pasteadjustsheights") = False) Then Config.Add "pasteadjustsheights", vbChecked
     If (Config("shortcuts").Exists("mode3dautoaligny") = False) Then Config("shortcuts").Add "mode3dautoaligny", 65601
     If (Config("shortcuts").Exists("errorcheck") = False) Then Config("shortcuts").Add "errorcheck", 115
     If (Config.Exists("windowedvideo") = False) Then Config.Add "windowedvideo", vbChecked
     If (Config.Exists("modekeys3d") = False) Then Config.Add "modekeys3d", vbUnchecked
     If (Config.Exists("standardtexturebrowse") = False) Then Config.Add "standardtexturebrowse", vbUnchecked
     If (Config.Exists("linessectorsinfo") = False) Then Config.Add "linessectorsinfo", vbChecked
     If (Config("palette").Exists("CLR_LINEBLOCKSOUND") = False) Then Config("palette").Add "CLR_LINEBLOCKSOUND", 6907800
     If (Config("shortcuts").Exists("reversedrawing") = False) Then Config("shortcuts").Add "reversedrawing", 8
     If (Config("shortcuts").Exists("mode3dtexalignresetx") = False) Then Config("shortcuts").Add "mode3dtexalignresetx", 65618
     If (Config("shortcuts").Exists("mode3dtexalignresety") = False) Then Config("shortcuts").Add "mode3dtexalignresety", 131154
     If (Config("shortcuts").Exists("mode3dthingstoggle") = False) Then Config("shortcuts").Add "mode3dthingstoggle", 84
     If (Config("shortcuts").Exists("thingrotatecw") = False) Then Config("shortcuts").Add "thingrotatecw", 190
     If (Config("shortcuts").Exists("thingrotateccw") = False) Then Config("shortcuts").Add "thingrotateccw", 188
     If (Config("shortcuts").Exists("mode3dthingrotatecw") = False) Then Config("shortcuts").Add "mode3dthingrotatecw", 190
     If (Config("shortcuts").Exists("mode3dthingrotateccw") = False) Then Config("shortcuts").Add "mode3dthingrotateccw", 188
     If (Config("palette").Exists("CLR_MAPBOUNDARY") = False) Then Config("palette").Add "CLR_MAPBOUNDARY", 16711680
     If (Config("shortcuts").Exists("mode3dinsert") = False) Then Config("shortcuts").Add "mode3dinsert", 45
     If (Config.Exists("alwaysalltextures") = False) Then Config.Add "alwaysalltextures", vbUnchecked
     
     
     'Success
     LoadConfiguration = True
     Exit Function
     
     
ConfigError:
     
     MsgBox "Error " & Err.number & " in LoadConfiguration: " & Err.Description, vbCritical
End Function

Public Function SortDictionary(ByRef SourceDictionary As Dictionary) As Dictionary
     Dim ItemCount As Long
     Dim Keys As Variant
     Dim Item1 As Long, Item2 As Long
     Dim TempItem As String
     
     'Create new dictionary
     Set SortDictionary = New Dictionary
     
     ItemCount = SourceDictionary.Count
     Keys = SourceDictionary.Keys
     
     'Loop through the collection
     For Item1 = 0 To (ItemCount - 2)
          
          'Loop from the current item to the end
          For Item2 = Item1 To (ItemCount - 1)
               
               'Swap if item from Item1 is more then item from Item2
               If (Keys(Item1) > Keys(Item2)) Then
                    
                    'Swap Item1 with Item2
                    TempItem = Keys(Item2)
                    Keys(Item2) = Keys(Item1)
                    Keys(Item1) = TempItem
               End If
          Next Item2
     Next Item1
     
     'Build new dictionary
     For Item1 = 0 To (ItemCount - 1)
          SortDictionary.Add Keys(Item1), SourceDictionary(Keys(Item1))
     Next Item1
End Function


Public Function SortDictionaryByValue(ByRef SourceDictionary As Dictionary) As Dictionary
     Dim ItemCount As Long
     Dim Keys As Variant
     Dim Values As Variant
     Dim Item1 As Long, Item2 As Long
     Dim TempItem As String
     Dim TempValue As String
     
     'Create new dictionary
     Set SortDictionaryByValue = New Dictionary
     
     ItemCount = SourceDictionary.Count
     Keys = SourceDictionary.Keys
     Values = SourceDictionary.Items
     
     'Loop through the collection
     For Item1 = 0 To (ItemCount - 2)
          
          'Loop from the current item to the end
          For Item2 = Item1 To (ItemCount - 1)
               
               'Swap if item from Item1 is more then item from Item2
               If (Values(Item1) > Values(Item2)) Then
                    
                    'Swap Item1 with Item2
                    TempItem = Keys(Item2)
                    TempValue = Values(Item2)
                    Keys(Item2) = Keys(Item1)
                    Values(Item2) = Values(Item1)
                    Keys(Item1) = TempItem
                    Values(Item1) = TempValue
               End If
          Next Item2
     Next Item1
     
     'Build new dictionary
     For Item1 = 0 To (ItemCount - 1)
          SortDictionaryByValue.Add Keys(Item1), Values(Item1)
     Next Item1
End Function



Public Function CalculateMapRect() As RECT
     Dim i As Long
     Dim MapRect As RECT
     
     'Go for all vertices
     For i = 0 To (numvertexes - 1)
          
          'Check if this is the first vertex
          If (i = 0) Then
               
               'Start with map size measurement from first vertex
               MapRect.left = vertexes(i).x
               MapRect.top = -vertexes(i).y
               MapRect.right = vertexes(i).x
               MapRect.bottom = -vertexes(i).y
          Else
               
               'Measure map size
               If vertexes(i).x < MapRect.left Then MapRect.left = vertexes(i).x
               If -vertexes(i).y < MapRect.top Then MapRect.top = -vertexes(i).y
               If vertexes(i).x > MapRect.right Then MapRect.right = vertexes(i).x
               If -vertexes(i).y > MapRect.bottom Then MapRect.bottom = -vertexes(i).y
          End If
     Next i
     
     'Return the rect
     CalculateMapRect = MapRect
End Function

Public Function CalculateSectorRect(ByVal sector As Long) As RECT
     Dim sd As Long
     Dim SectorRect As RECT
     Dim FirstVertex As Long
     Dim linedef As MAPLINEDEF
     
     'Go for all sidedefs
     For sd = 0 To (numsidedefs - 1)
          
          'Sidedef refers to this sector?
          If (sidedefs(sd).sector = sector) Then
               
               'Get linedef
               linedef = linedefs(sidedefs(sd).linedef)
               
               'Check if this is the first vertex
               If (FirstVertex = False) Then
                    
                    'First vertex
                    With SectorRect
                         .left = vertexes(linedef.v1).x
                         .right = vertexes(linedef.v1).x
                         .top = vertexes(linedef.v1).y
                         .bottom = vertexes(linedef.v1).y
                    End With
                    
                    'First done
                    FirstVertex = True
               Else
                    
                    'Apply first vertex
                    With SectorRect
                         If (vertexes(linedef.v1).x < .left) Then .left = vertexes(linedef.v1).x
                         If (vertexes(linedef.v1).x > .right) Then .right = vertexes(linedef.v1).x
                         If (vertexes(linedef.v1).y > .top) Then .top = vertexes(linedef.v1).y
                         If (vertexes(linedef.v1).y < .bottom) Then .bottom = vertexes(linedef.v1).y
                    End With
               End If
               
               'Apply second vertex
               With SectorRect
                    If (vertexes(linedef.v2).x < .left) Then .left = vertexes(linedef.v2).x
                    If (vertexes(linedef.v2).x > .right) Then .right = vertexes(linedef.v2).x
                    If (vertexes(linedef.v2).y > .top) Then .top = vertexes(linedef.v2).y
                    If (vertexes(linedef.v2).y < .bottom) Then .bottom = vertexes(linedef.v2).y
               End With
          End If
     Next sd
     
     'Flip Y
     SectorRect.top = -SectorRect.top
     SectorRect.bottom = -SectorRect.bottom
     
     'Return the rect
     CalculateSectorRect = SectorRect
End Function


Public Function CalculateLinedefRect(ByVal linedef As Long) As RECT
     Dim LinedefRect As RECT
     
     'First vertex
     With LinedefRect
          .left = vertexes(linedefs(linedef).v1).x
          .right = vertexes(linedefs(linedef).v1).x
          .top = vertexes(linedefs(linedef).v1).y
          .bottom = vertexes(linedefs(linedef).v1).y
     End With
     
     'Apply second vertex
     With LinedefRect
          If (vertexes(linedefs(linedef).v2).x < .left) Then .left = vertexes(linedefs(linedef).v2).x
          If (vertexes(linedefs(linedef).v2).x > .right) Then .right = vertexes(linedefs(linedef).v2).x
          If (vertexes(linedefs(linedef).v2).y > .top) Then .top = vertexes(linedefs(linedef).v2).y
          If (vertexes(linedefs(linedef).v2).y < .bottom) Then .bottom = vertexes(linedefs(linedef).v2).y
     End With
     
     'Flip Y
     LinedefRect.top = -LinedefRect.top
     LinedefRect.bottom = -LinedefRect.bottom
     
     'Return the rect
     CalculateLinedefRect = LinedefRect
End Function


Private Sub Main()
     Dim SplashStart As Long
     Dim qs As Long, qe As Long
     Dim Filename As String
     Dim NewConfigs As Boolean
     
     'Indicate we're in loading sequence
     Loading = True
     
     'Load the splash dialog
     Load frmSplash
     
     'Show splash dialog
     frmSplash.Show
     frmSplash.Refresh
     
     'Change to local path
     ChDrive left$(App.Path, 1)
     ChDir App.Path
     
     'Get the windows version
     DetectWindowsVersion
     
     'Get the windows temp path
     GetWindowsTempPath
     
     'Clean up temporary files
     CleanUpTemporaries
     
     'Keep the splash screen start time
     SplashStart = GetTickCount
     
     'Load configuration, quit when no success
     If Not LoadConfiguration Then End
     
     'Find game configurations
     FindGameConfigurations
     
     'Check for new game configurations
     'DisplayStatus "Validating game configurations..."
     'DoEvents
     'NewConfigs = DetectNewGameConfigs
     
     'Setting up defaults
     DisplayStatus "Setting defaults..."
     
     'Setup default settings
     InitializeStartupDefaults
     
     'Setting up defaults
     DisplayStatus "Creating renderer palette..."
     
     'Create rendering palette
     CreateRendererPalette
     
     
     'Load the 2D interface
     DisplayStatus "Loading interface..."
     Load frmMain
     
     'Make the splash float over main window
     frmSplash.Show 0, frmMain
     
     'Initialize the clipboard
     InitializeClipboard
     ClipboardCleanup
     
     'Show the 2D interface
     DisplayStatus "Initializing interface..."
     frmMain.Show
     frmMain.Refresh
     DoEvents
     
     'We're done loading
     Loading = False
     
     'Check if theres nothing more to do
     If (InStr(Command, """") = 0) And (NewConfigs = False) Then
          
          'Ensure the splash shows long enough
          DisplayStatus ""
          While ((SplashStart + 2000 > GetTickCount) And SplashDisplayed): DoEvents: Sleep 10: Wend
          
     End If
     
     'Unload the splash dialog
     Unload frmSplash: Set frmSplash = Nothing
     
     'Check if we should ask use to configure IWADs
     If NewConfigs Then
          
          'Ask the user now
          MsgBox "Doom Builder has detected new game configuration files." & vbLf & _
                 "Please click OK to configure the IWAD file locations for them now.", vbInformation
          
          'Show configuration
          frmMain.ShowConfiguration 3
     End If
     
     'Check if we should load a WAD now
     If (InStr(Command, """") <> 0) Then
          
          'Get quote positions
          qs = InStr(Command, """")
          qe = InStr(qs + 1, Command, """")
          
          'Get filename
          Filename = Mid$(Command, qs + 1, qe - qs - 1)
          
          'Check if we should add current path
          If (InStr(Filename, "\") = 0) And _
             (InStr(Filename, "/") = 0) And _
             (InStr(Filename, ":") = 0) Then
               
               'Add current path to filename
               If (right$(CurDir, 1) = "\") Then
                    Filename = CurDir & Filename
               Else
                    Filename = CurDir & "\" & Filename
               End If
          End If
          
          'Load the select map dialog
          Load frmMapSelect
          
          'Set the tag and caption
          frmMapSelect.tag = Filename
          frmMapSelect.Caption = "Select Map from " & Dir(Filename)
          
          'Show the select dialog
          frmMapSelect.Show 1, frmMain
     End If
End Sub

Public Function MainMessageHandler(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim k As Long, s As Long
     
     'Check what message to ahndle
     Select Case wMsg
          
          Case WM_MOUSEWHEEL       'Mousewheel scroll up or down
               
               'Check if the mousewheel went up or down
               If (wParam > 0) Then
                    
                    'Split keycode and shift
                    'k = (Config("shortcuts")("zoomin") And &HFFF)
                    's = (Config("shortcuts")("zoomin") And &HFF0000) \ 2 ^ 16
                    k = MOUSE_SCROLL_UP
                    s = CurrentShiftMask
               Else
                    
                    'Split keycode and shift
                    'k = (Config("shortcuts")("zoomout") And &HFFF)
                    's = (Config("shortcuts")("zoomout") And &HFF0000) \ 2 ^ 16
                    k = MOUSE_SCROLL_DOWN
                    s = CurrentShiftMask
               End If
               
               'Mousehweel up, zoom in
               frmMain.Form_KeyDown CInt(k), CInt(s)
               
          Case WM_ACTIVATEAPP      'Application is activated/deactivated
               
               'Check if activated
               If (wParam <> 0) Then
                    
                    'When in windowed 3D Mode
                    If (Running3D) And (Val(Config("windowedvideo")) <> 0) Then
                         
                         'Capture mouse now
                         CaptureMouse
                         
                         'Normal processing
                         DelayVideoFrames = False
                    End If
                    
               'Otherwise it is deactivated
               Else
                    
                    'When in windowed 3D Mode
                    If (Running3D) And (Val(Config("windowedvideo")) <> 0) Then
                         
                         'Free the mouse
                         FreeMouse
                         
                         'Delayed processing to save CPU time
                         DelayVideoFrames = True
                    End If
               End If
               
     End Select
     
     'Pass the message on to the original handler
     MainMessageHandler = CallWindowProc(frmMain.OriginalMessageHandler, hWnd, wMsg, wParam, lParam)
End Function


Public Function MakeGenLinedefEffect(ByVal effect As Long, ByVal CategoryIndex As Long) As Long
     Dim Cats As Variant
     Dim Cat As Dictionary
     
     'Get category
     Cats = mapconfig("gen_linedeftypes").Items
     Set Cat = Cats(CategoryIndex)
     
     'Return the value changed with the offset
     MakeGenLinedefEffect = effect + Val(Cat("offset"))
End Function

Public Function MakeTempFile(Optional CreateFile As Boolean = True) As String
     
     'Make sure the path is set
     If (LenB(TempPath) = 0) Then GetWindowsTempPath
     
     'Make a temp file
     MakeTempFile = Space$(255)
     GetTempFileName TempPath, "wad", 0, MakeTempFile
     MakeTempFile = GetShortFileName(Trim$(Replace$(MakeTempFile, vbNullChar, "")))
     
     'Remove the file if we dont need it
     If (CreateFile = False) Then Kill MakeTempFile
End Function

Public Function MenuNameForShortcut(ByVal CurrentCaption As String, ByVal ShortcutItem As String) As String
     Dim Tabpos As Long
     Dim OriginalCaption As String
     Dim k As Long, s As Long
     
     'Find tab position
     Tabpos = InStr(CurrentCaption, vbTab)
     
     'Check if a tab was found
     If (Tabpos > 0) Then
          
          'Use without shortcut or tab
          OriginalCaption = left$(CurrentCaption, Tabpos - 1)
     Else
          
          'No shortcut, use entire current caption
          OriginalCaption = CurrentCaption
     End If
     
     'Check if shortcut is unbound
     If (Config("shortcuts")(ShortcutItem) = 0) Then
          
          'Only show caption
          MenuNameForShortcut = OriginalCaption
     Else
          
          'Split keycode and shift
          k = (Val(Config("shortcuts")(ShortcutItem)) And &HFFF)
          s = (Val(Config("shortcuts")(ShortcutItem)) And &HFF0000) \ 2 ^ 16
          
          'Show caption with tab and shortcut
          MenuNameForShortcut = OriginalCaption & vbTab & NameForKeycode(k, s)
     End If
End Function

Public Function NameForKeycode(ByVal KeyCode As Integer, ByVal Shift As Integer) As String
     Dim Prefix As String
     
     'Make Shift prefix
     If (Shift And vbAltMask) Then Prefix = Prefix & "Alt+"
     If (Shift And vbCtrlMask) Then Prefix = Prefix & "Ctrl+"
     If (Shift And vbShiftMask) Then Prefix = Prefix & "Shift+"
     
     'Return the name for the key
     Select Case KeyCode
          Case 0: NameForKeycode = ""
          Case 27: NameForKeycode = Prefix & "Esc"
          Case 112: NameForKeycode = Prefix & "F1"
          Case 113: NameForKeycode = Prefix & "F2"
          Case 114: NameForKeycode = Prefix & "F3"
          Case 115: NameForKeycode = Prefix & "F4"
          Case 116: NameForKeycode = Prefix & "F5"
          Case 117: NameForKeycode = Prefix & "F6"
          Case 118: NameForKeycode = Prefix & "F7"
          Case 119: NameForKeycode = Prefix & "F8"
          Case 120: NameForKeycode = Prefix & "F9"
          Case 121: NameForKeycode = Prefix & "F10"
          Case 122: NameForKeycode = Prefix & "F11"
          Case 123: NameForKeycode = Prefix & "F12"
          Case 192: NameForKeycode = Prefix & "~"
          Case 48: NameForKeycode = Prefix & "0"
          Case 49: NameForKeycode = Prefix & "1"
          Case 50: NameForKeycode = Prefix & "2"
          Case 51: NameForKeycode = Prefix & "3"
          Case 52: NameForKeycode = Prefix & "4"
          Case 53: NameForKeycode = Prefix & "5"
          Case 54: NameForKeycode = Prefix & "6"
          Case 55: NameForKeycode = Prefix & "7"
          Case 56: NameForKeycode = Prefix & "8"
          Case 57: NameForKeycode = Prefix & "9"
          Case 189: NameForKeycode = Prefix & "-"
          Case 187: NameForKeycode = Prefix & "="
          Case 8: NameForKeycode = Prefix & "Backspace"
          Case 9: NameForKeycode = Prefix & "Tab"
          Case 81: NameForKeycode = Prefix & "Q"
          Case 87: NameForKeycode = Prefix & "W"
          Case 69: NameForKeycode = Prefix & "E"
          Case 82: NameForKeycode = Prefix & "R"
          Case 84: NameForKeycode = Prefix & "T"
          Case 89: NameForKeycode = Prefix & "Y"
          Case 85: NameForKeycode = Prefix & "U"
          Case 73: NameForKeycode = Prefix & "I"
          Case 79: NameForKeycode = Prefix & "O"
          Case 80: NameForKeycode = Prefix & "P"
          Case 219: NameForKeycode = Prefix & "["
          Case 221: NameForKeycode = Prefix & "]"
          Case 13: NameForKeycode = Prefix & "Enter"
          Case 65: NameForKeycode = Prefix & "A"
          Case 83: NameForKeycode = Prefix & "S"
          Case 68: NameForKeycode = Prefix & "D"
          Case 70: NameForKeycode = Prefix & "F"
          Case 71: NameForKeycode = Prefix & "G"
          Case 72: NameForKeycode = Prefix & "H"
          Case 74: NameForKeycode = Prefix & "J"
          Case 75: NameForKeycode = Prefix & "K"
          Case 76: NameForKeycode = Prefix & "L"
          Case 186: NameForKeycode = Prefix & ";"
          Case 222: NameForKeycode = Prefix & "'"
          Case 90: NameForKeycode = Prefix & "Z"
          Case 88: NameForKeycode = Prefix & "X"
          Case 67: NameForKeycode = Prefix & "C"
          Case 86: NameForKeycode = Prefix & "V"
          Case 66: NameForKeycode = Prefix & "B"
          Case 78: NameForKeycode = Prefix & "N"
          Case 77: NameForKeycode = Prefix & "M"
          Case 188: NameForKeycode = Prefix & ","
          Case 190: NameForKeycode = Prefix & "."
          Case 191: NameForKeycode = Prefix & "/"
          Case 220: NameForKeycode = Prefix & "\"
          Case 32: NameForKeycode = Prefix & "Space"
          Case 45: NameForKeycode = Prefix & "Ins"
          Case 46: NameForKeycode = Prefix & "Del"
          Case 36: NameForKeycode = Prefix & "Home"
          Case 35: NameForKeycode = Prefix & "End"
          Case 33: NameForKeycode = Prefix & "PgUp"
          Case 34: NameForKeycode = Prefix & "PgDwn"
          Case 37: NameForKeycode = Prefix & "Left"
          Case 38: NameForKeycode = Prefix & "Up"
          Case 39: NameForKeycode = Prefix & "Right"
          Case 40: NameForKeycode = Prefix & "Down"
          Case 111: NameForKeycode = Prefix & "Num /"
          Case 109: NameForKeycode = Prefix & "Num -"
          Case 107: NameForKeycode = Prefix & "Num +"
          Case 96: NameForKeycode = Prefix & "Num 0"
          Case 97: NameForKeycode = Prefix & "Num 1"
          Case 98: NameForKeycode = Prefix & "Num 2"
          Case 99: NameForKeycode = Prefix & "Num 3"
          Case 100: NameForKeycode = Prefix & "Num 4"
          Case 101: NameForKeycode = Prefix & "Num 5"
          Case 102: NameForKeycode = Prefix & "Num 6"
          Case 103: NameForKeycode = Prefix & "Num 7"
          Case 104: NameForKeycode = Prefix & "Num 8"
          Case 105: NameForKeycode = Prefix & "Num 9"
          Case 110: NameForKeycode = Prefix & "Num ."
          Case MOUSE_BUTTON_0: NameForKeycode = Prefix & "Mouse1"
          Case MOUSE_BUTTON_1: NameForKeycode = Prefix & "Mouse2"
          Case MOUSE_BUTTON_2: NameForKeycode = Prefix & "Mouse3"
          Case MOUSE_BUTTON_3: NameForKeycode = Prefix & "Mouse4"
          Case MOUSE_BUTTON_4: NameForKeycode = Prefix & "Mouse5"
          Case MOUSE_BUTTON_5: NameForKeycode = Prefix & "Mouse6"
          Case MOUSE_BUTTON_6: NameForKeycode = Prefix & "Mouse7"
          Case MOUSE_BUTTON_7: NameForKeycode = Prefix & "Mouse8"
          Case MOUSE_SCROLL_DOWN: NameForKeycode = Prefix & "ScrollDown"
          Case MOUSE_SCROLL_UP: NameForKeycode = Prefix & "ScrollUp"
          Case Else: NameForKeycode = Prefix & "Key " & KeyCode
     End Select
End Function

Public Function NextThingTag() As Long
     Dim t As Long
     Dim Used As New Dictionary
     
     'Go for all things
     For t = 0 To (numthings - 1)
          
          'Check if this has a tag
          If (things(t).tag <> 0) Then
               
               'Check if not already added
               If (Used.Exists(CStr(things(t).tag)) = False) Then
                    
                    'Add to used list
                    Used.Add CStr(things(t).tag), things(t).tag
               End If
          End If
     Next t
     
     'Count up until a free value is found
     NextThingTag = 1
     Do While Used.Exists(CStr(NextThingTag)): NextThingTag = NextThingTag + 1: Loop
End Function

Public Function NextUnusedTag() As Long
     Dim ld As Long
     Dim s As Long
     Dim Used As New Dictionary
     
     'Go for all linedefs
     For ld = 0 To (numlinedefs - 1)
          
          'Check if this has a tag
          If (linedefs(ld).tag <> 0) Then
               
               'Check if not already added
               If (Used.Exists(CStr(linedefs(ld).tag)) = False) Then
                    
                    'Add to used list
                    Used.Add CStr(linedefs(ld).tag), linedefs(ld).tag
               End If
          End If
     Next ld
     
     'Go for all sectors
     For s = 0 To (numsectors - 1)
          
          'Check if this has a tag
          If (sectors(s).tag <> 0) Then
               
               'Check if not already added
               If (Used.Exists(CStr(sectors(s).tag)) = False) Then
                    
                    'Add to used list
                    Used.Add CStr(sectors(s).tag), sectors(s).tag
               End If
          End If
     Next s
     
     'Count up until a free value is found
     NextUnusedTag = 1
     Do While Used.Exists(CStr(NextUnusedTag)): NextUnusedTag = NextUnusedTag + 1: Loop
End Function

Public Function OnOff(ByVal of As Long) As String
     If (of) Then OnOff = "On" Else OnOff = "Off"
End Function

Public Sub OpenADDWADFile()
     
     'Errors will be outputted to log
     On Error Resume Next
     
     'Check if ADDWAD is given
     If (Trim$(addwadfile) <> "") Then
          
          'Check if file exists
          If (Dir(addwadfile) <> "") Then
               
               'Open associated ADDWAD
               Err.Clear
               AddWAD.OpenFile addwadfile, True
               
               'Check for errors
               If (Err.number <> 0) Then
                    
                    'Add warning message
                    ErrorLog_Add "WARNING: Could not open the additional WAD file """ & GetFileName(addwadfile) & """", False
                    
                    'Make a temporary file to use
                    AddWAD.CloseFile
               End If
          Else
               
               'Add warning message
               ErrorLog_Add "WARNING: Could not find the additional WAD file """ & GetFileName(addwadfile) & """. Check your configuration!", False
               
               'Make a temporary file to use
               AddWAD.CloseFile
          End If
     Else
          
          'No additional file
          AddWAD.CloseFile
     End If
End Sub

Public Sub OpenIWADFile()
     Dim CurIWADFile As String
     
     'Errors will be outputted to log
     On Error Resume Next
     
     'Get configured IWAD file
     CurIWADFile = GetCurrentIWADFile
     
     'Check if the IWAD was configured
     If (Trim$(CurIWADFile) <> "") Then
          
          'Check if IWAD exists
          If (Dir(CurIWADFile) <> "") Then
               
               'Open associated IWAD
               Err.Clear
               IWAD.OpenFile CurIWADFile, True
               
               'Check for errors
               If (Err.number <> 0) Then
                    
                    'Add warning message
                    ErrorLog_Add "WARNING: Could not open the IWAD file """ & GetFileName(GetCurrentIWADFile) & """", False
                    
                    'Make a temporary file to use
                    IWAD.NewFile MakeTempFile(False), True
               End If
          Else
               
               'Add warning message
               ErrorLog_Add "WARNING: Could not find the IWAD file """ & GetFileName(GetCurrentIWADFile) & """. Check your configuration!", False
               
               'Make a temporary file to use
               IWAD.NewFile MakeTempFile(False), True
          End If
     Else
          
          'Add warning message
          ErrorLog_Add "WARNING: You have no IWAD file set for this configuration!", False
          
          'Make a temporary file to use
          IWAD.NewFile MakeTempFile(False), True
     End If
End Sub

Public Function Padded(ByRef Src As String, ByVal Length As Long) As String
     
     'Check the length of Src
     If Len(Src) < Length Then
          
          'Make an 8 byte Null padded string
          Padded = Src & String$(Length - Len(Src), vbNullChar)
     ElseIf Len(Src) = Length Then
          
          'Src just fits in an 8 bytes string
          Padded = Src
     Else
          
          'Chop Src to 8 bytes
          Padded = left$(Src, Length)
     End If
End Function

Public Function PathOf(ByVal Filename As String) As String
     On Local Error GoTo NoPath
     Dim FileTitle As String
     
     'Get file title
     FileTitle = Dir(Filename)
     
     'Return path only
     PathOf = left$(Filename, Len(Filename) - Len(FileTitle))
     
NoPath:
     
     'Leave now
     Exit Function
End Function

Public Function point_in_polygon(ByRef lines As Variant, ByVal numlines As Long, ByVal x As Single, ByVal y As Single) As Boolean
     Dim ld As MAPLINEDEF
     Dim v1x As Single, v1y As Single
     Dim v2x As Single, v2y As Single
     Dim MinY As Single, MaxY As Single, MaxX As Single
     Dim ldi As Long
     Dim xinters As Single
     Dim Count As Long
     
     'Go for selected linedefs
     For ldi = 0 To (numlines - 1)
          
          'Get linedef
          ld = linedefs(lines(ldi))
          
          'Get vertex coordinates
          v1x = vertexes(ld.v1).x
          v1y = vertexes(ld.v1).y
          v2x = vertexes(ld.v2).x
          v2y = vertexes(ld.v2).y
          
          'Determine smallest/largest values
          If (v1y < v2y) Then MinY = v1y Else MinY = v2y
          If (v1y > v2y) Then MaxY = v1y Else MaxY = v2y
          If (v1x > v2x) Then MaxX = v1x Else MaxX = v2x
          
          'Check for intersection
          If (y > MinY) And (y <= MaxY) Then
               If (x <= MaxX) Then
                    If (v1y <> v2y) Then
                         xinters = (y - v1y) * (v2x - v1x) / (v2y - v1y) + v1x
                         If (v1x = v2x) Or (x <= xinters) Then Count = Count + 1
                    End If
               End If
          End If
     Next ldi
     
     'Return result
     point_in_polygon = (Count Mod 2)
End Function

Public Function PointInRect(ByVal x As Long, ByVal y As Long, ByRef r As RECT) As Long
     
     'Return true when xy is inside r, otherwise return false
     PointInRect = (x >= r.left) And (x <= r.right) And (y >= r.top) And (y <= r.bottom)
     
End Function

Public Sub RemoveRecentFile(ByVal Index As Long)
     Dim i As Long
     
     'Go backwards for all recent files from
     'index and move them up (file at index will be lost)
     For i = (Index + 1) To MAX_RECENT_FILES
          
          'Check if set
          If Config("recent").Exists(CStr(i)) Then
               
               'Move up by 1
               Config("recent")(CStr(i - 1)) = Config("recent")(CStr(i))
          Else
               
               'Remove previous
               If Config("recent").Exists(CStr(i - 1)) Then Config("recent").Remove CStr(i - 1)
          End If
     Next i
     
     'Remove last item
     If Config("recent").Exists(CStr(MAX_RECENT_FILES)) Then Config("recent").Remove CStr(MAX_RECENT_FILES)
End Sub

Public Function RequiresS1Lower(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Check for other sidedef
     If (linedefs(ld).s1 > -1) And (linedefs(ld).s2 > -1) Then
          
          'Check if floor on the other side is not F_SKY1
          If (Trim$(sectors(sidedefs(linedefs(ld).s2).sector).tfloor) <> "F_SKY1") Then
               
               'Return result depending on sector floor heights
               RequiresS1Lower = (sectors(sidedefs(linedefs(ld).s2).sector).hfloor > sectors(sidedefs(linedefs(ld).s1).sector).hfloor)
          End If
     End If
End Function

Public Function RequiresS1Middle(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Return result depending on backside
     RequiresS1Middle = (linedefs(ld).s2 = -1)
End Function

Public Function RequiresS1Upper(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Check for other sidedef
     If (linedefs(ld).s1 > -1) And (linedefs(ld).s2 > -1) Then
          
          'Check if ceiling on the other side is not F_SKY1
          If (Trim$(sectors(sidedefs(linedefs(ld).s2).sector).tceiling) <> "F_SKY1") Then
               
               'Return result depending on sector heights
               RequiresS1Upper = (sectors(sidedefs(linedefs(ld).s2).sector).hceiling < sectors(sidedefs(linedefs(ld).s1).sector).hceiling)
          End If
     End If
End Function

Public Function RequiresS2Lower(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Check for other sidedef
     If (linedefs(ld).s1 > -1) And (linedefs(ld).s2 > -1) Then
          
          'Check if floor on the other side is not F_SKY1
          If (Trim$(sectors(sidedefs(linedefs(ld).s1).sector).tfloor) <> "F_SKY1") Then
               
               'Return result depending on sector floor heights
               RequiresS2Lower = (sectors(sidedefs(linedefs(ld).s1).sector).hfloor > sectors(sidedefs(linedefs(ld).s2).sector).hfloor)
          End If
     End If
End Function

Public Function RequiresS2Middle(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Return result depending on backside
     RequiresS2Middle = (linedefs(ld).s1 = -1)
End Function

Public Function RequiresS2Upper(ByVal ld As Long) As Boolean
     On Local Error Resume Next
     
     'Check for other sidedef
     If (linedefs(ld).s1 > -1) And (linedefs(ld).s2 > -1) Then
          
          'Check if ceiling on the other side is not F_SKY1
          If (Trim$(sectors(sidedefs(linedefs(ld).s1).sector).tceiling) <> "F_SKY1") Then
               
               'Return result depending on sector heights
               RequiresS2Upper = (sectors(sidedefs(linedefs(ld).s1).sector).hceiling < sectors(sidedefs(linedefs(ld).s2).sector).hceiling)
          End If
     End If
End Function

Public Function ShortedPathText(ByRef Text As String, ByVal MaxPixels As Long, Optional Bold As Boolean) As String
     Dim i As Long
     
     'Set the bold
     frmMain.FontBold = Bold
     
     'Check if the text doesnt fit
     If (frmMain.TextWidth(Text) > MaxPixels) Then
          
          'Go for all characters
          For i = 4 To Len(Text)
               
               'Check if this is too long
               If (frmMain.TextWidth(left$(Text, 3) & "..." & right$(Text, i)) > MaxPixels) Then
                    
                    'Return shortened text after this char
                    ShortedPathText = left$(Text, 3) & "..." & right$(Text, i - 1)
                    Exit Function
               End If
          Next i
     End If
     
     'Return all text
     ShortedPathText = Text
End Function

Public Function ShortedText(ByRef Text As String, ByVal MaxPixels As Long, Optional Bold As Boolean) As String
     Dim i As Long
     
     'Set the bold
     frmMain.FontBold = Bold
     
     'Check if the text doesnt fit
     If (frmMain.TextWidth(Text) > MaxPixels) Then
          
          'Go for all characters
          For i = 1 To Len(Text)
               
               'Check if this is too long
               If (frmMain.TextWidth(left$(Text, i) & "...") > MaxPixels) Then
                    
                    'Return shortened text before this char
                    ShortedText = left$(Text, i - 1) & "..."
                    Exit Function
               End If
          Next i
     End If
     
     'Return all text
     ShortedText = Text
End Function

Public Function side_of_line(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, ByVal vx As Single, ByVal vy As Single) As Single
     
     'return < 0 for front (right) side, > 0 for back (left) side and 0 for on the line
     side_of_line = (vy - y1) * (x2 - x1) - (vx - x1) * (y2 - y1)
     
End Function

Public Function SnappedToGridX(ByVal x As Single) As Single
     Dim offset As Single
     
     'Calculate offset
     offset = (gridx Mod gridsizex)
     
     'Integer Divide and Multiply by grid size to align with it
     SnappedToGridX = CLng((x - offset) / gridsizex) * gridsizex + offset
End Function

Public Function SnappedToGridY(ByVal y As Single) As Single
     Dim offset As Single
     
     'Calculate offset
     offset = (gridy Mod gridsizey)
     
     'Integer Divide and Multiply by grid size to align with it
     SnappedToGridY = CLng((y + offset) / gridsizey) * gridsizey - offset
End Function

Public Function StringFromBytes(ByRef ByteArray() As Byte) As String
     Dim c As Long
     Dim NewString As String
     
     'This function creates a variable-length string from byte array
     
     NewString = Space$(UBound(ByteArray) - LBound(ByteArray))
     For c = LBound(ByteArray) To UBound(ByteArray)
          
          'Check if end of string
          If ByteArray(c) Then
               
               'Set character
               Mid$(NewString, c + 1, 1) = ChrW$(ByteArray(c))
          Else
               
               'Leave the loop
               Exit For
          End If
     Next c
     
     StringFromBytes = left$(NewString, c)
End Function

Public Sub Terminate()
     Dim Tmr As Single
     
     'Cleanup undo/redo memory
     TerminateUndoRedo
     
     'Terminate last thing pointer
     DestroyBitmapPointer ThingBitmapData
     
     'Stop the map screen renderer
     TerminateMapRenderer
     
     'Clean up temporary files
     CleanUpTemporaries
     
     'Save configuration
     Configfile.SaveConfiguration App.Path & "\Builder.cfg"
     
     'No more editing
     DisableMapEditing
     
     'Discard events
     DoEvents
     DoEvents
     
     'No more errors from here
     On Error Resume Next
     
     'Free the mouse
     FreeMouse
     
     'Get rid of status window
     Unload frmStatus
     
     'Get rid of 3D window
     Unload frm3D
     
     'Unload everything
     Tmr = Timer
     Do While (Forms.Count > 1)
          Unload Forms(0)
          Set Forms(0) = Nothing
          
          'End if impossible to unload
          If (Tmr + 0.2 < Timer) Then End
     Loop
End Sub

Public Function IsLoaded(ByRef AnyForm As Form) As Boolean
     Dim TempForm As Form
     
     'Go for all forms that are loaded
     For Each TempForm In Forms
          
          'Check if the given form is among them
          If TempForm Is AnyForm Then
               
               'Return True
               IsLoaded = True
               Exit Function
          End If
     Next
End Function


Public Function TextureMessageHandler(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim k As Long, s As Long
     
     'Check what message to ahndle
     Select Case wMsg
          
          Case WM_MOUSEWHEEL       'Mousewheel scroll up or down
               
               'Check if the mousewheel went up or down
               If (wParam > 0) Then
                    
                    'Split keycode and shift
                    'k = (Config("shortcuts")("zoomin") And &HFFF)
                    's = (Config("shortcuts")("zoomin") And &HFF0000) \ 2 ^ 16
                    k = 107
               Else
                    
                    'Split keycode and shift
                    'k = (Config("shortcuts")("zoomout") And &HFFF)
                    's = (Config("shortcuts")("zoomout") And &HFF0000) \ 2 ^ 16
                    k = 109
               End If
               
               'Mousehweel up, zoom in
               frmTextureBrowse.Form_KeyDown CInt(k), CInt(s)
               
     End Select
     
     'Pass the message on to the original handler
     TextureMessageHandler = CallWindowProc(frmTextureBrowse.OriginalMessageHandler, hWnd, wMsg, wParam, lParam)
End Function

Public Function ThingFiltered(ByVal Index As Long) As Long
     
     'Check if filter is on
     If (filterthings = True) Then
          
          'Test thing category
          If ((filtersettings.category = -1) Or (things(Index).category = filtersettings.category)) Then
               
               'Check filter mode
               Select Case filtersettings.filtermode
                    Case 0: ThingFiltered = ((things(Index).Flags And filtersettings.Flags) <> 0) Or (things(Index).Flags = 0)
                    Case 1: ThingFiltered = ((things(Index).Flags And filtersettings.Flags) = filtersettings.Flags)
                    Case 2: ThingFiltered = (things(Index).Flags = filtersettings.Flags)
               End Select
          Else
               
               'Wrong category
               ThingFiltered = False
          End If
     Else
          
          'Always display thing
          ThingFiltered = True
     End If
End Function

Public Function UngenLinedefEffect(ByVal effect As Long) As Long
     Dim Cat As Dictionary
     
     'Get category
     Set Cat = GetGenLinedefCategory(effect)
     
     'Return the value changed with the offset
     UngenLinedefEffect = effect - Val(Cat("offset"))
End Function

Public Function UnPadded(ByRef Src As String) As String
     Dim nt As Long
     
     'Find null terminator
     nt = InStr(Src, vbNullChar)
     
     'Check if a null terminator is found
     If (nt > 0) Then
          
          'Return string up to the null terminator
          UnPadded = Trim$(left$(Src, nt - 1))
     Else
          
          'Return original
          UnPadded = Trim$(Src)
     End If
End Function

Public Sub UpdateRecentFilesMenu()
     Dim i As Long, mi As Long
     Dim DisplayName As String
     
     'Hide first item
     frmMain.itmFileRecent(0).visible = False
     
     'Go for all items
     For i = 1 To frmMain.itmFileRecent.UBound
          
          'Unload this item
          Unload frmMain.itmFileRecent(i)
     Next i
     
     'Go for all recent files
     For i = 1 To MAX_RECENT_FILES
          
          'Check if set
          If Config("recent").Exists(CStr(i)) Then
               
               'Make an item
               If (mi > 0) Then Load frmMain.itmFileRecent(mi)
               
               'Show item
               With frmMain.itmFileRecent(mi)
                    
                    'Make the name to display
                    DisplayName = ShortedPathText(CStr(Config("recent")(CStr(i))), 140, False)
                    
                    'Set the filename on tag
                    .tag = Config("recent")(CStr(i))
                    
                    'Set the display name
                    .Caption = "&" & (mi + 1) & "  " & UCase$(left$(DisplayName, 1)) & Mid$(DisplayName, 2)
                    
                    'Show item
                    .visible = True
               End With
               
               'Next item
               mi = mi + 1
          End If
     Next i
End Sub

Public Sub UpdateScriptLumpsMenu()
     Dim i As Long
     Dim MapLumpNames As Variant
     Dim PresentName As String
     Dim mi As Long
     
     'Go for all items
     For i = 1 To frmMain.itmScriptEdit.UBound
          
          'Unload this item
          Unload frmMain.itmScriptEdit(i)
     Next i
     
     'Go for all maplumpnames
     MapLumpNames = mapconfig("maplumpnames").Keys
     For i = LBound(MapLumpNames) To UBound(MapLumpNames)
          
          'Make presenting name
          PresentName = MapLumpNames(i)
          If (PresentName = "~") Then PresentName = maplumpname
          
          'Check if this is meant for scripting
          If (GetMapLumpType(PresentName) And ML_CUSTOM) Then
               
               'Make a menu item for this
               If (mi > 0) Then Load frmMain.itmScriptEdit(mi)
               
               'Set the properties
               With frmMain.itmScriptEdit(mi)
                    .Caption = "Edit " & PresentName & " lump..."
                    .tag = MapLumpNames(i)
                    .visible = True
               End With
               
               'Next menu item
               mi = mi + 1
          End If
     Next i
     
     'Hide menu if no lumps to edit
     frmMain.mnuScripts.visible = (mi > 0)
End Sub

Public Sub UpdateStatusBar()
     
     'Update the panels
     With frmMain.stbStatus
          .Panels("numvertexes").Text = numvertexes & " vertices"
          .Panels("numlinedefs").Text = numlinedefs & " linedefs"
          .Panels("numsidedefs").Text = numsidedefs & " sidedefs"
          .Panels("numsectors").Text = numsectors & " sectors"
          .Panels("numthings").Text = numthings & " things"
          
          'Check what to show for grid
          If gridsizex = gridsizey Then
               .Panels("gridsize").Text = "Grid: " & gridsizex
          Else
               .Panels("gridsize").Text = "Grid: " & gridsizex & ", " & gridsizey
          End If
          
          .Panels("snapmode").Text = "AutoSnap: " & UCase$(OnOff(snapmode))
          .Panels("stitchmode").Text = "AutoStitch: " & UCase$(OnOff(stitchmode))
          .Panels("viewzoom").Text = "Zoom: " & CLng(ViewZoom * 100) & "%"
     End With
End Sub

Public Sub UpdateThingImageColor(ByVal ThingIndex As Long)
     Dim c As Long
     Dim ThingCats As Variant
     Dim a As Long
     
     'Get category keys
     ThingCats = mapconfig("thingtypes").Keys
     
     'Default to unknown thing
     things(ThingIndex).image = TI_UNKNOWN
     
     'Default to thing unknown color
     things(ThingIndex).Color = CLR_THINGUNKNOWN
     
     'Go for all thing categories
     For c = LBound(ThingCats) To UBound(ThingCats)
          
          'Check if this thing number is in this category
          If (mapconfig("thingtypes")(ThingCats(c)).Exists(CStr(things(ThingIndex).thing))) Then
               
               'Check if the thing is supposed to have an arrow
               If (mapconfig("thingtypes")(ThingCats(c))(CStr(things(ThingIndex).thing))("arrow")) Then
                    
                    'Get the angle
                    a = things(ThingIndex).angle
                    
                    'Make the angle 0 - 360
                    While (a < 0): a = a + 360: Wend
                    While (a > 360): a = a - 360: Wend
                    
                    'Set the image up on direction
                    things(ThingIndex).image = CLng(a / 45)
                    If (things(ThingIndex).image = 8) Then things(ThingIndex).image = 0
               Else
                    
                    'Set the image to a dot
                    things(ThingIndex).image = TI_DOT
               End If
               
               'Set the thing color as specified in the category
               things(ThingIndex).Color = PALETTE_16COLORS_OFFSET + mapconfig("thingtypes")(ThingCats(c))("color")
               
               'Leave the category search
               Exit For
          End If
     Next c
End Sub

Public Sub UpdateThingSize(ByVal ThingIndex As Long)
     
     'Apply thing sizes
     things(ThingIndex).size = GetThingWidth(things(ThingIndex).thing)
     things(ThingIndex).height = GetThingHeight(things(ThingIndex).thing)
     things(ThingIndex).hangs = GetThingHangs(things(ThingIndex).thing)
End Sub


Public Sub UpdateThingCategory(ByVal ThingIndex As Long)
     
     'Apply category
     things(ThingIndex).category = GetThingTypeCategoryIndex(things(ThingIndex).thing)
End Sub


Public Function WaitForSingleFile(ByVal Filename As String, ByVal Timeout As Long, ByVal AccessTimeout As Long) As Boolean
     On Error Resume Next
     Dim FileBuffer As Long
     Dim BeginTime As Long
     Dim ErrorNumber As Long
     
     'Return True when the file has become available
     
     'Get the begin time
     BeginTime = GetTickCount
     
     'Wait for the file to exist
     Do: Loop Until (Dir(Filename) <> "") Or ((BeginTime + Timeout < GetTickCount) And (Timeout > 0))
     
     'Get the begin time
     BeginTime = GetTickCount
     
     'Now access the file
     Do
          'Close when file is opened
          If FileBuffer Then Close #FileBuffer
          
          'Clear errors
          Err.Clear
          
          'Check if file exists
          If (Dir(Filename) = "") Then
               
               'File cant be found
               Err.Raise 53
          Else
               
               'Try opening the file for exclusive access
               FileBuffer = FreeFile
               Open Filename For Binary Access Read Write Lock Read Write As #FileBuffer
          End If
          
          'Get any errors
          ErrorNumber = Err.number
          
     'Continue until no errors or timeout
     Loop Until (ErrorNumber = 0) Or (ErrorNumber = 53) Or ((BeginTime + AccessTimeout < GetTickCount) And (AccessTimeout > 0))
     
     'Close the file
     Close #FileBuffer
     
     'Return result
     WaitForSingleFile = (ErrorNumber = 0)
End Function

Public Sub WriteLogLine(ByRef line As String)
     Dim filebuf As Integer
     
     'Open log file
     filebuf = FreeFile
     Open App.Path & "\Builder.log" For Append As #filebuf
     
     'Write the line
     Print #filebuf, line
     
     'Close file
     Close #filebuf
End Sub

Public Function YesNo(ByVal yn As Long) As String
     If (yn) Then YesNo = "Yes" Else YesNo = "No"
End Function
