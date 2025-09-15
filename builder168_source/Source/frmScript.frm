VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Begin VB.Form frmScript 
   Caption         =   "Map Script"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   9630
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HasDC           =   0   'False
   Icon            =   "frmScript.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   507
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   642
   Begin MSComctlLib.ImageList imglstScriptIcons 
      Left            =   135
      Top             =   3675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":1658
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CodeSenseCtl.CodeSense csnScript 
      Height          =   2400
      Left            =   5325
      OleObjectBlob   =   "frmScript.frx":1BF2
      TabIndex        =   9
      Top             =   570
      Width           =   3030
   End
   Begin VB.Frame fraControls 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   8880
      Begin VB.CommandButton cmdScriptCompile 
         Caption         =   "Compile"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5340
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   105
         Width           =   1515
      End
      Begin VB.CommandButton cmdScriptExport 
         Caption         =   "Export Script..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3765
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   105
         Width           =   1515
      End
      Begin VB.CommandButton cmdScriptImport 
         Caption         =   "Import Script..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2190
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   105
         Width           =   1515
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   105
         Width           =   1515
      End
   End
   Begin VB.Frame fraNoScript 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   75
      TabIndex        =   4
      Top             =   495
      Visible         =   0   'False
      Width           =   5025
      Begin VB.CommandButton cmdScriptMake 
         Caption         =   "Make Script"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1725
      End
      Begin VB.Label lblNoScript 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "This script does not yet exist in the currently loaded map."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   5
         Top             =   1110
         Width           =   4980
      End
   End
   Begin VB.Label lblLumpname 
      Caption         =   "MAP01"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   150
      TabIndex        =   7
      Top             =   3195
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "mnuEdit"
      Visible         =   0   'False
      Begin VB.Menu itmUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu itmRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu itmLine1 
         Caption         =   "-"
      End
      Begin VB.Menu itmCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu itmCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu itmPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu itmDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu itmLine2 
         Caption         =   "-"
      End
      Begin VB.Menu itmFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu itmReplace 
         Caption         =   "Replace..."
         Shortcut        =   ^H
      End
      Begin VB.Menu itmLine3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu itmToggleBookmark 
         Caption         =   "Toggle &Bookmark"
         Shortcut        =   ^{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu itmNextBookmark 
         Caption         =   "Go to Next Bookmark"
         Shortcut        =   {F2}
         Visible         =   0   'False
      End
      Begin VB.Menu itmPrevBookmark 
         Caption         =   "Go to Previous Bookmark"
         Shortcut        =   +{F2}
         Visible         =   0   'False
      End
      Begin VB.Menu itmClearBookmarks 
         Caption         =   "Clear Bookmarks"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmScript"
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


'Virtual Keys
Private Const VK_F2 = &H71

'Scripting
Private ScriptConfig As Dictionary
Private CodeList As CodeSenseCtl.ICodeList
Private FuncPosInfo As FUNC_POSITION_INFO
Private CodeTip As Object

'Previous keydown keycode
Private LastKeyCode As Long

'Misc
Private AllowUpdate As Boolean
Private scriptchanged As Boolean

'Position info
Private Type FUNC_POSITION_INFO
     FunctionName As String
     ArgumentIndex As Long
End Type

'Keywords list for CodeList
Private KeywordsList() As String
Private KeywordsIcons() As Byte


Private Function AcceptCodeListSelection() As Boolean
     On Error GoTo DestroyList
     
     'CodeList exists?
     If Not (CodeList Is Nothing) Then
          
          'CodeList displayed?
          If (CodeList.hWnd <> 0) Then
               
               'Anything selected?
               If (CodeList.SelectedItem > -1) Then
                    
                    'Replace word with keyword
                    ReplaceCurrentWord CodeList.GetItemText(CodeList.SelectedItem)
                    
                    'Remove the list
                    CodeList.Destroy
                    Set CodeList = Nothing
                    
                    'Code accepted and inserted
                    AcceptCodeListSelection = True
               End If
          Else
               
               'Destroy list
               Set CodeList = Nothing
          End If
     End If
     
     'Leave now
     Exit Function
     
     
DestroyList:
     
     'CodeList exists?
     If Not (CodeList Is Nothing) Then
          
          'CodeList displayed?
          If (CodeList.hWnd <> 0) Then CodeList.Destroy
          
          'Destroy it
          Set CodeList = Nothing
     End If
End Function

Public Sub Compile()
     Dim ScriptType As ENUM_MAPLUMPTYPES
     Dim LumpName As String
     Dim lumpindex As Long
     
     'Get the script type
     ScriptType = GetMapLumpType(lblLumpName.Caption)
     If (Trim$(lblLumpName.Caption) = "~") Then
          ScriptType = ScriptType Or ML_REQUIRED
          LumpName = maplumpname
     Else
          LumpName = lblLumpName.Caption
     End If
     
     'Check if this is an ACS script
     If (ScriptType And ML_CUSTOMACS) = ML_CUSTOMACS Then
          
          'Find and remove the original lump
          lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
          If (lumpindex > 0) Then TempWAD.DeleteLump lumpindex
          
          'Compile and save the script to original lump
          TempWAD.AddLump CompiledACS(csnScript.Text), LumpName
     End If
     
     'Save changes to temp file
     TempWAD.WriteChanges
     
     'Map changed
     scriptchanged = True
     mapchanged = True
End Sub

Private Function CurrentWordPosition() As Position
     Dim WordPos As New Position
     Dim SearchPos As New Position
     Dim CurrentWord As String
     Dim WordLength As Long
     Dim TextLine As String
     Dim TextChar As String
     Dim i As Long
     
     'Initialize positions
     CurrentWord = csnScript.CurrentWord
     WordLength = csnScript.CurrentWordLength
     WordPos.LineNo = csnScript.GetSel(True).EndLineNo
     WordPos.ColNo = csnScript.GetSel(True).EndColNo - 1
     SearchPos.LineNo = csnScript.GetSel(True).EndLineNo
     SearchPos.ColNo = csnScript.GetSel(True).EndColNo
     
     'Go back along the word until its not current word anymore
     For i = 0 To WordLength
          
          'Check if not at beginning of line
          If (SearchPos.ColNo > 0) Then
               
               'Go back one character
               SearchPos.ColNo = SearchPos.ColNo - 1
               
               'Get the character at this position
               TextLine = csnScript.GetLine(SearchPos.LineNo)
               If (SearchPos.ColNo > 0) And (SearchPos.ColNo <= Len(TextLine)) Then TextChar = Mid$(TextLine, SearchPos.ColNo, 1) Else TextChar = " "
               
               'Test the word here
               If (csnScript.GetWord(SearchPos) = CurrentWord) And (InStr(1, ScriptConfig("delimiters"), TextChar, vbBinaryCompare) <= 0) Then
                    
                    'Still the same word, store position
                    WordPos.ColNo = SearchPos.ColNo - 1
                    WordPos.LineNo = SearchPos.LineNo
               Else
                    
                    'End of word! Leave now
                    Exit For
               End If
          End If
     Next i
     
     'Return position for current word
     Set CurrentWordPosition = WordPos
End Function
Private Function FunctionPositionInfo() As FUNC_POSITION_INFO
     Dim BracketLevel As Long
     Dim ArgumentIndex As Long
     Dim OriginalPos As Range
     Dim line As Long
     Dim Pos As Long
     Dim char As Long
     Dim TopIndex As Long
     Dim Range As Long
     Dim BracketOpen As Long
     Dim BracketClose As Long
     Dim ArgumentDelimiter As Long
     Dim Terminator As Long
     
     'Do not check further than 200 chars
     'otherwise this may take up too much time when not inside a function
     Const MaxRange As Long = 200
     
     'For this to work we must have the following defined
     If (ScriptConfig("functionopen") <> "") And (ScriptConfig("functionclose") <> "") And _
        (ScriptConfig("argumentdelimiter") <> "") And (ScriptConfig("terminator") <> "") Then
          
          'Check if no selection has been made
          If (csnScript.SelLength = 0) Then
               
               'Get character codes
               BracketOpen = AscW(ScriptConfig("functionopen"))
               BracketClose = AscW(ScriptConfig("functionclose"))
               ArgumentDelimiter = AscW(ScriptConfig("argumentdelimiter"))
               Terminator = AscW(ScriptConfig("terminator"))
               
               'Lock the control
               LockWindowUpdate csnScript.hWnd
               
               'Get current position
               Set OriginalPos = csnScript.GetSel(False)
               line = OriginalPos.EndLineNo
               Pos = OriginalPos.EndColNo
               If (Pos = 0) Then Pos = 1
               TopIndex = csnScript.TopIndex
               
               'Go from current position back to beginning
               Do
                    Do
                         'When meeting ) then increase bracket level
                         'When meeting ( then decrease bracket level
                         'When bracket level goes -1, then the next word should be the function name
                         'Only when at bracket level 0, count the comma's for argument index
                         
                         'Move caret
                         csnScript.SetCaretPos line, Pos
                         
                         'Check if this is a scope change
                         If (csnScript.CurrentToken = cmTokenTypeScopeBegin) Or _
                            (csnScript.CurrentToken = cmTokenTypeScopeEnd) Then
                              
                              'Function broken? Leave now
                              Range = MaxRange
                              
                         'Check if not in a string or comment
                         ElseIf (csnScript.CurrentToken <> cmTokenTypeMultiLineComment) And _
                                (csnScript.CurrentToken <> cmTokenTypeSingleLineComment) And _
                                (csnScript.CurrentToken <> cmTokenTypeString) Then
                              
                              'Get character at pos
                              char = AscW(Mid$(csnScript.GetLine(line), Pos, 1) & " ")
                              
                              'Is this a ) ?
                              If (char = BracketClose) Then
                                   
                                   'Increase bracket level
                                   BracketLevel = BracketLevel + 1
                                   
                              'Is this a ( ?
                              ElseIf (char = BracketOpen) Then
                                   
                                   'Decrease bracket level
                                   BracketLevel = BracketLevel - 1
                                   
                                   'Out of our current brackets?
                                   If (BracketLevel < 0) Then
                                        
                                        'Check if can have a word before this
                                        If (Pos > 0) Then
                                             
                                             'Go to the word before the bracket
                                             csnScript.SetCaretPos line, Pos - 1
                                             
                                             'Keyword before the bracket?
                                             If (csnScript.CurrentToken = cmTokenTypeKeyword) Then
                                                  
                                                  'Set the function name and argument index
                                                  With FunctionPositionInfo
                                                       .FunctionName = csnScript.CurrentWord
                                                       .ArgumentIndex = ArgumentIndex
                                                  End With
                                                  
                                                  'No further scanning needed
                                                  Range = MaxRange
                                                  
                                             'Only a delimiter before the bracket?
                                             ElseIf (InStr(1, ScriptConfig("delimiters"), left$(csnScript.CurrentWord, 1), vbBinaryCompare) > 0) Then
                                                  
                                                  'Change the bracketlevel to one more
                                                  BracketLevel = BracketLevel + 1
                                                  ArgumentIndex = 0
                                                  
                                             'Otherwise this is an unknown keyword
                                             Else
                                                  
                                                  'No further scanning needed
                                                  Range = MaxRange
                                             End If
                                        Else
                                             
                                             'No further scanning needed
                                             Range = MaxRange
                                        End If
                                   End If
                                   
                              'Is this a , ?
                              ElseIf (char = ArgumentDelimiter) Then
                                   
                                   'Only count at 0 bracket level
                                   If (BracketLevel = 0) Then
                                        
                                        'Increase argument index
                                        ArgumentIndex = ArgumentIndex + 1
                                   End If
                                   
                              'Is this a terminator?
                              ElseIf (char = Terminator) Then
                                   
                                   'Function broken? Stop scanning now
                                   Range = MaxRange
                              End If
                         End If
                         
                         'One char back
                         Pos = Pos - 1
                         Range = Range + 1
                         
                    'Continue until at beginning of line
                    Loop While (Pos > 0) And (Range < MaxRange)
                    
                    'Go one line up
                    line = line - 1
                    Pos = csnScript.GetLineLength(line)
                    If (Pos = 0) Then Pos = 1
                    
               'Continue until at the beginning
               Loop While (line > -1) And (Range < MaxRange)
               
               'Return caret to original position
               csnScript.SetCaretPos OriginalPos.EndLineNo, OriginalPos.EndColNo
               csnScript.TopIndex = TopIndex
               
               'Unlock the control
               LockWindowUpdate 0
          End If
     End If
End Function

Private Sub ReplaceCurrentWord(ByVal NewWord As String)
     Dim CaretPos As Position
     Dim WordRange As New Range
     
     'Make current word range
     Set CaretPos = CurrentWordPosition
     WordRange.StartColNo = CaretPos.ColNo
     WordRange.StartLineNo = CaretPos.LineNo
     WordRange.EndLineNo = CaretPos.LineNo
     WordRange.EndColNo = CaretPos.ColNo + csnScript.CurrentWordLength
     
     'Check if no word at the cursor
     If (csnScript.CurrentWordLength = 0) Then
          
          'CurrentWordPosition gives the word position, not the cursor position
          'We want the actual cursor position for insertion
          Set CaretPos = New Position
          CaretPos.ColNo = csnScript.GetSel(True).StartColNo
          CaretPos.LineNo = csnScript.GetSel(True).StartLineNo
          
          'Insert without replacing anything
          WordRange.StartColNo = WordRange.EndColNo
          csnScript.InsertText NewWord, CaretPos
          
          'Move caret to end of inserted word
          csnScript.SetCaretPos CaretPos.LineNo, CaretPos.ColNo + Len(NewWord)
          
     'Check if the current word is only a delimiter
     ElseIf (InStr(1, ScriptConfig("delimiters"), csnScript.CurrentWord, vbBinaryCompare) > 0) Then
          
          'Insert keyword after the delimiter
          WordRange.StartColNo = WordRange.EndColNo
          csnScript.ReplaceText NewWord, WordRange
     Else
          
          'Replace word with keyword
          csnScript.ReplaceText NewWord, WordRange
     End If
End Sub

Public Sub Save()
     
     'Check if any script given
     If (csnScript.visible = True) And (csnScript.Text <> "") Then
          
          'Save the script
          SaveLumpScript
     Else
          
          'Remove any unneeded lumps
          DeleteLumpScript
     End If
End Sub

Private Sub SetupScriptControl()
     Dim ScriptGlobals As New CodeSenseCtl.Globals
     Dim ScriptLang As New CodeSenseCtl.Language
     Dim List As Variant
     Dim Keywrd As String
     Dim HKey As New HotKey
     Dim i As Long
     
     'Unregister built-in languages
     ScriptGlobals.UnregisterAllLanguages
     
     'Check if a script configuration is loaded
     If Not (ScriptConfig Is Nothing) Then
          
          'Reserve memory for keywords
          ReDim KeywordsList(0 To ScriptConfig("keywords").Count + ScriptConfig("constants").Count - 1)
          ReDim KeywordsIcons(0 To ScriptConfig("keywords").Count + ScriptConfig("constants").Count - 1)
          
          'Create language
          With ScriptLang
               
               'Create language settings
               .Style = cmLangStyleProcedural
               .CaseSensitive = CBool(ScriptConfig("casesensitive"))
               .EscapeChar = ScriptConfig("escape")
               .MultiLineComments1 = ScriptConfig("commentopen")
               .MultiLineComments2 = ScriptConfig("commentclose")
               .ScopeKeywords1 = ScriptConfig("scopeopen")
               .ScopeKeywords2 = ScriptConfig("scopeclose")
               .SingleLineComments = ScriptConfig("linecomment")
               .StringDelims = ScriptConfig("string")
               .TerminatorChar = ScriptConfig("terminator")
               
               'Create keywords
               List = ScriptConfig("keywords").Keys
               For i = LBound(List) To UBound(List)
                    
                    'Replace periods by spaces
                    Keywrd = Replace(List(i), ".", " ")
                    
                    'Add keyword to syntax highlighter
                    .Keywords = .Keywords & Keywrd & vbLf
                    
                    'Check how to add for listing
                    Select Case Val(ScriptConfig("insertcase"))
                         Case 0: KeywordsList(i) = Keywrd
                         Case 1: KeywordsList(i) = LCase$(Keywrd)
                         Case 2: KeywordsList(i) = UCase$(Keywrd)
                    End Select
                    
                    'Function icon
                    KeywordsIcons(i) = 2
                    
                    'If not case sensitive, then make keywords lowercase
                    If (CBool(ScriptConfig("casesensitive")) = False) Then
                         
                         'Make lowercase key
                         ScriptConfig("keywords").Key(List(i)) = LCase$(List(i))
                         List(i) = LCase$(Keywrd)
                    End If
               Next i
               
               'Create operators (i use this for constants)
               List = ScriptConfig("constants").Keys
               For i = LBound(List) To UBound(List)
                    
                    'Replace periods by spaces
                    Keywrd = Replace(List(i), ".", " ")
                    
                    'Add constant to syntax highlighter
                    .Operators = .Operators & Keywrd & vbLf
                    
                    'Check how to add for listing
                    Select Case Val(ScriptConfig("insertcase"))
                         Case 0: KeywordsList(ScriptConfig("keywords").Count + i) = Keywrd
                         Case 1: KeywordsList(ScriptConfig("keywords").Count + i) = LCase$(Keywrd)
                         Case 2: KeywordsList(ScriptConfig("keywords").Count + i) = UCase$(Keywrd)
                    End Select
                    
                    'Constant icon
                    KeywordsIcons(ScriptConfig("keywords").Count + i) = 3
                    
                    'If not case sensitive, then make keywords lowercase
                    If (CBool(ScriptConfig("casesensitive")) = False) Then
                         
                         'Make lowercase key
                         ScriptConfig("constants").Key(List(i)) = LCase$(List(i))
                         List(i) = LCase$(Keywrd)
                    End If
               Next i
          End With
          
          'Apply the language definition
          ScriptGlobals.RegisterLanguage "Current", ScriptLang
          
          'Use the language in the script control
          csnScript.Language = "Current"
     End If
     
     'Set colors/font
     With csnScript
          
          'Icons
          Set csnScript.ImageList = imglstScriptIcons
          
          'Highlighting?
          If Not (ScriptConfig Is Nothing) Then .ColorSyntax = CBool(Config("syntaxhighlighting")) Else .ColorSyntax = False
          
          'Font
          .Font.Name = "Lucida Console"
          .Font.Bold = True
          .Font.Italic = False
          .Font.size = 9
          .Font.Strikethrough = False
          .Font.Underline = False
          
          'Styles
          .SetFontStyle cmStyComment, cmFontBold
          .SetFontStyle cmStyKeyword, cmFontBold
          .SetFontStyle cmStyLineNumber, cmFontBold
          .SetFontStyle cmStyNumber, cmFontBold
          .SetFontStyle cmStyOperator, cmFontBold
          .SetFontStyle cmStyScopeKeyword, cmFontBold
          .SetFontStyle cmStyString, cmFontBold
          .SetFontStyle cmStyText, cmFontBold
          
          'Foreground colors
          .SetColor cmClrBookmark, LongToBGRLong(Config("palette")("CLR_SCRIPTTEXT"))
          .SetColor cmClrComment, LongToBGRLong(Config("palette")("CLR_SCRIPTCOMMENT"))
          .SetColor cmClrHighlightedLine, LongToBGRLong(Config("palette")("CLR_SCRIPTTEXT"))
          .SetColor cmClrKeyword, LongToBGRLong(Config("palette")("CLR_SCRIPTKEYWORD"))
          .SetColor cmClrLeftMargin, LongToBGRLong(Config("palette")("CLR_SCRIPTLINENUMBERS"))
          .SetColor cmClrLineNumber, LongToBGRLong(Config("palette")("CLR_SCRIPTLINENUMBERS"))
          .SetColor cmClrNumber, LongToBGRLong(Config("palette")("CLR_SCRIPTSTRING"))
          .SetColor cmClrOperator, LongToBGRLong(Config("palette")("CLR_SCRIPTCONSTANT"))
          .SetColor cmClrScopeKeyword, LongToBGRLong(Config("palette")("CLR_SCRIPTTEXT"))
          .SetColor cmClrString, LongToBGRLong(Config("palette")("CLR_SCRIPTSTRING"))
          .SetColor cmClrText, LongToBGRLong(Config("palette")("CLR_SCRIPTTEXT"))
          
          'Background colors
          .SetColor cmClrCommentBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrBookmarkBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrKeywordBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrLineNumberBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrNumberBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrOperatorBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrScopeKeywordBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrStringBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrTextBk, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
          .SetColor cmClrWindow, LongToBGRLong(Config("palette")("CLR_SCRIPTBACKGROUND"))
     End With
     
     'Do not error if key is already (un)registered
     On Error Resume Next
     
     'Unregister Alt+ENTER for Properties
     HKey.Modifiers1 = 4: HKey.VirtKey1 = vbCr
     ScriptGlobals.UnregisterHotKey HKey
     
     'Unregister Ctrl+F2 for Bookmark
     HKey.Modifiers1 = 2: HKey.VirtKey1 = Chr$(VK_F2)
     ScriptGlobals.UnregisterHotKey HKey
     
     'Unregister Ctrl+A for Redo
     HKey.Modifiers1 = 2: HKey.VirtKey1 = "A"
     ScriptGlobals.UnregisterHotKey HKey
     
     'Register Ctrl+Y for Redo
     HKey.Modifiers1 = 2: HKey.VirtKey1 = "Y"
     ScriptGlobals.RegisterHotKey HKey, cmCmdRedo
     
     'Register Ctrl+H for Find/Replace
     HKey.Modifiers1 = 2: HKey.VirtKey1 = "H"
     ScriptGlobals.RegisterHotKey HKey, cmCmdFindReplace
End Sub

Private Sub cmdClose_Click()
     
     'Same as closing
     Unload Me
End Sub

Private Sub cmdScriptCompile_Click()
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Compile
     Compile
     
     'Focus to script
     On Error Resume Next
     csnScript.SetFocus
     
     'Change mousepointer
     Screen.MousePointer = vbDefault
End Sub

Private Sub cmdScriptExport_Click()
     On Local Error Resume Next
     Dim result As String
     Dim FileBuffer As Integer
     Dim AllData As String
     
     'Show save dialog
     result = SaveFile(Me.hWnd, "Export Script", "All Files|*.*", "", cdlOFNExplorer Or cdlOFNHideReadOnly Or cdlOFNLongNames Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt)
     frmMain.Refresh
     frmScript.Refresh
     
     'Check if not cancelled
     If result <> "" Then
          
          'Kill file if exists
          If (Dir(result) <> "") Then Kill result
          
          'Get the script text
          AllData = csnScript.Text
          
          'Open the script file
          FileBuffer = FreeFile
          Open result For Output As #FileBuffer
          
          'Output the script
          Print #FileBuffer, AllData
          
          'Close the script file
          Close #FileBuffer
     End If
     
     'Focus to script
     On Error Resume Next
     csnScript.SetFocus
End Sub

Private Sub cmdScriptImport_Click()
     On Local Error Resume Next
     Dim result As String
     Dim FileBuffer As String
     Dim LineData As String
     Dim AllData As String
     Dim LineReport As Long
     
     'Open dialog
     result = OpenFile(Me.hWnd, "Import Script", "All Files|*.*", "", cdlOFNExplorer Or cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNLongNames)
     frmMain.Refresh
     frmScript.Refresh
     
     'Check if not cancelled
     If result <> "" Then
          
          'Check if any script is already been made
          If csnScript.visible Then
               
               'Leave if not confirmed
               If (MsgBox("This will replace this current script with the contents of the file." & vbLf & "Do you want to continue?", vbQuestion Or vbYesNo) = vbNo) Then Exit Sub
          End If
          
          'Clear all error icons
          For LineReport = 0 To csnScript.LineCount - 1
               csnScript.SetMarginImages LineReport, csnScript.GetMarginImages(LineReport) And Not 2
          Next LineReport
          
          'Open the file
          FileBuffer = FreeFile
          Open result For Input As #FileBuffer
          
          'Continue reading until end of file
          Do Until EOF(FileBuffer)
               
               'Read a line
               Line Input #FileBuffer, LineData
               
               'Add to result
               AllData = AllData & LineData & vbCrLf
          Loop
          
          'Close file
          Close #FileBuffer
          
          'Erase undo/redo
          csnScript.ClearUndoBuffer
          
          'Enable script controls
          csnScript.Text = AllData
          csnScript.visible = True
          
          'Hide panel
          fraNoScript.visible = False
          cmdScriptExport.Enabled = True
     End If
     
     'Focus to script
     On Error Resume Next
     csnScript.SetFocus
End Sub

Private Sub cmdScriptMake_Click()
     Dim ScriptType As ENUM_MAPLUMPTYPES
     
     'Get the script type
     ScriptType = GetMapLumpType(lblLumpName.Caption)
     If (Trim$(lblLumpName.Caption) = "~") Then ScriptType = ScriptType Or ML_REQUIRED
     
     'Erase undo/redo
     csnScript.ClearUndoBuffer
     
     'New script
     csnScript.Text = ""
     csnScript.visible = True
     fraNoScript.visible = False
     cmdScriptExport.Enabled = True
     scriptchanged = True
     
     'ACS must be compiled, enable button
     cmdScriptCompile.Enabled = ((ScriptType And ML_CUSTOMACS) = ML_CUSTOMACS)
End Sub

Private Function CompiledACS(ByVal ACS As String) As String
     On Local Error GoTo CompilingError
     Dim FileBuffer As Integer
     Dim ScriptFile As String
     Dim ObjFile As String
     Dim ErrorFile As String
     Dim AllData As String
     Dim OutputLine As String
     Dim acsFilename As String
     Dim PreviousMouse As Integer
     Dim LineReport As Long
     
     'Change mousepointer
     PreviousMouse = Screen.MousePointer
     Screen.MousePointer = vbHourglass
     
     'Clear all error icons
     For LineReport = 0 To csnScript.LineCount - 1
          csnScript.SetMarginImages LineReport, csnScript.GetMarginImages(LineReport) And Not 2
     Next LineReport
     
     'Load error dialog
     ErrorLog_Load
     frmErrorLog.lblDesc = "The compiler returned the following output about your code:"
     
     'Create a temporary wad files to build in
     ObjFile = App.Path & "\script.o"
     ScriptFile = App.Path & "\script.acs"
     ErrorFile = App.Path & "\acs.err"
     
     
     'Open the script file
     FileBuffer = FreeFile
     Open ScriptFile For Output As #FileBuffer
     
     'Output the script
     Print #FileBuffer, ACS
     
     'Close the script file
     Close #FileBuffer
     
     
     'Run batch to compile the script
     If (Execute(App.Path & "\acc.exe", GetFileName(ScriptFile) & " " & GetFileName(ObjFile), SW_HIDE, True) = False) Then Err.Raise vbObjectError + 2, , "Could not run script compiler!"
     
     
     'Check if the object file exists
     If WaitForSingleFile(ObjFile, 500, 3000) Then
          
          'Open the object file
          FileBuffer = FreeFile
          Open ObjFile For Binary Access Read Lock Read Write As #FileBuffer
          
          'Make string to size of file
          AllData = Space$(LOF(FileBuffer))
          
          'Read the data from file
          Get #FileBuffer, 1, AllData
          CompiledACS = AllData
          
          'Close object file
          Close #FileBuffer
     End If
     
     
     'Check for ACS compiler errors
     If WaitForSingleFile(ErrorFile, 200, 3000) Then
          
          'Open the error file
          FileBuffer = FreeFile
          Open ErrorFile For Input Access Read Lock Read Write As #FileBuffer
          
          'Read all lines until end of file
          Do Until EOF(FileBuffer)
               
               'Read line
               Line Input #FileBuffer, OutputLine
               
               'Check if line begins with "script.acs:"
               If (StrComp(left$(OutputLine, 11), "script.acs:", vbTextCompare) = 0) Then
                    
                    'Get the line number on which the error exists
                    LineReport = Val(Mid$(OutputLine, 12, 10))
                    
                    'Add error icon here
                    csnScript.SetMarginImages (LineReport - 1), csnScript.GetMarginImages(LineReport - 1) Or 2
               End If
               
               'Add to output dialog
               ErrorLog_Add OutputLine, False
          Loop
          
          'Close output file
          Close #FileBuffer
     End If
     
     
     'Check if not compiled
     If (CompiledACS = "") Then
          
          'Add error to output dialog
          ErrorLog_Add vbCrLf & " The ACS compiler did not compile your script.", True
          If (frmErrorLog.imgIcon.Picture <> frmErrorLog.imgCritical.Picture) Then Set frmErrorLog.imgIcon.Picture = frmErrorLog.imgCritical.Picture
     End If
     
     'Kill temporary files
     On Local Error Resume Next
     Kill ObjFile
     Kill ScriptFile
     Kill ErrorFile
     'Kill OutputFile
     'Kill BatchFile
     On Local Error GoTo CompilingError
     
     
CompilingError:
     
     'Check for error
     If (Err.number <> 0) Then
          
          'Add error report
          ErrorLog_Add "ERROR: Fatal error " & Err.number & " while compiling ACS: " & Err.Description, True
          If (frmErrorLog.imgIcon.Picture <> frmErrorLog.imgCritical.Picture) Then Set frmErrorLog.imgIcon.Picture = frmErrorLog.imgCritical.Picture
          frmErrorLog.Show 1, Me
     End If
     
     'Default mousepointer
     Screen.MousePointer = vbNormal
     
     'Show the errors and warnings dialog
     ErrorLog_DisplayAndFlush
     
     'Reset mousepointer
     Screen.MousePointer = PreviousMouse
End Function

Private Function DecompiledACS(ByVal BehaviorLumpIndex As Long) As String
     On Local Error GoTo errorhandler
     Dim FileBuffer As Integer
     Dim ScriptFile As String
     Dim BatchFile As String
     Dim ObjFile As String
     Dim LineData As String
     Dim LineIndex As Long
     
     
     'Create a temporary wad files to build in
     ObjFile = App.Path & "\script.o"
     ScriptFile = App.Path & "\script.acs"
     
     
     'Write the lump data to object file
     TempWAD.ExportLump BehaviorLumpIndex, ObjFile
     
     'Make sure the script file does not exist
     If (Dir(ScriptFile) <> "") Then Kill ScriptFile
     
     
     'Decompile the script
     If (Execute(App.Path & "\deacc.exe", GetFileName(ObjFile) & " " & GetFileName(ScriptFile), SW_HIDE, True) = False) Then MsgBox "Warning: Could not run script decompiler!", vbCritical
     
     
     'Add a description to the script to explain the uglyness of the code
     '(maybe i should do so with my code too ;)
     DecompiledACS = "// Doom Builder could not detect a SCRIPTS lump in the map," & vbCrLf & _
                     "// so it has decompiled this code from the " & TempWAD.LumpName(BehaviorLumpIndex) & " lump." & vbCrLf & _
                     "// Please verify that this code is correct, because it may" & vbCrLf & _
                     "// not always work correctly with newer engines." & vbCrLf
     
     'Clear errors
     Err.Clear
     
     'Wait for the file
     If WaitForSingleFile(ScriptFile, 2000, 3000) Then
          
          'Open the result file
          FileBuffer = FreeFile
          Open ScriptFile For Input Access Read Lock Read Write As #FileBuffer
          
          'Continue reading until end of file or error
          Do Until EOF(FileBuffer)
               
               'Count lines
               LineIndex = LineIndex + 1
               
               'Read a line
               Line Input #FileBuffer, LineData
               
               'Ignore the first 3 lines
               If (LineIndex > 3) Then
                    
                    'Add to result
                    DecompiledACS = DecompiledACS & LineData & vbCrLf
               End If
          Loop
          
          'Close file
          Close #FileBuffer
     Else
          
          'Cannot open result
          DecompiledACS = ""
     End If
     
     
     'Kill temporary files
     On Local Error Resume Next
     Kill ObjFile
     Kill ScriptFile
     On Local Error GoTo 0
     
     'Leave now
     Exit Function
     
     
'Error handler
errorhandler:
     
     'Show and log error message (terminates application)
     MsgBox "Error " & Err.number & " in DecompileACS(): " & Err.Description, vbCritical
End Function

Private Sub DeleteLumpScript()
     Dim ScriptType As ENUM_MAPLUMPTYPES
     Dim lumpindex As Long
     Dim LumpName As String
     
     'Get the script type
     ScriptType = GetMapLumpType(lblLumpName.Caption)
     If (Trim$(lblLumpName.Caption) = "~") Then
          ScriptType = ScriptType Or ML_REQUIRED
          LumpName = maplumpname
     Else
          LumpName = lblLumpName.Caption
     End If
     
     'Check if this is an ACS script
     If (ScriptType And ML_CUSTOMACS) = ML_CUSTOMACS Then
          
          'Find and remove the SCRIPTS lump
          lumpindex = FindLumpIndex(TempWAD, 1, "SCRIPTS")
          If (lumpindex > 0) Then TempWAD.DeleteLump lumpindex: mapchanged = True
          
          'Find and remove the original lump
          lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
          If (lumpindex > 0) Then
               
               'Remove lump
               TempWAD.DeleteLump lumpindex
               mapchanged = True
               
               'Remake empty if original lump is required
               If ((ScriptType And ML_REQUIRED) = ML_REQUIRED) Then TempWAD.AddLump "", LumpName, lumpindex
          End If
     Else
          
          'Find and remove the exisitng lump
          lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
          If (lumpindex > 0) Then
               
               'Remove lump
               TempWAD.DeleteLump lumpindex
               mapchanged = True
               
               'Remake empty if original lump is required
               If ((ScriptType And ML_REQUIRED) = ML_REQUIRED) Then TempWAD.AddLump "", LumpName, lumpindex
          End If
     End If
     
     'Save changes to temp file
     TempWAD.WriteChanges
     mapchanged = scriptchanged
End Sub

Private Sub csnScript_Change(ByVal Control As CodeSenseCtl.ICodeSense)
     scriptchanged = True
End Sub

Private Function csnScript_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
     Dim i As Long
     Dim PossibleKeywords As Long
     Dim FoundPossibility As String
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Get a reference
          Set CodeList = ListCtrl
          
          'Set properties
          Set CodeList.ImageList = imglstScriptIcons
          
          'Go for all keywords
          For i = LBound(KeywordsList) To UBound(KeywordsList)
               
               'Add to list
               CodeList.AddItem KeywordsList(i), KeywordsIcons(i)
               
               'Check if this keyword is a possibility for autocomplete
               If (StrComp(left(KeywordsList(i), Len(csnScript.CurrentWord)), csnScript.CurrentWord, vbTextCompare) = 0) Then
                    
                    'Count the posibility
                    PossibleKeywords = PossibleKeywords + 1
                    
                    'Keep the function name
                    FoundPossibility = KeywordsList(i)
               End If
          Next i
          
          'Check if only 1 possibility
          If (PossibleKeywords = 1) Then
               
               'Replace word with keyword
               ReplaceCurrentWord FoundPossibility
               
               'No list
               csnScript_CodeList = False
               Set CodeList = Nothing
          Else
               
               'Any word typed
               If (csnScript.CurrentWord <> "") Then
                    
                    'Select the best match
                    CodeList.SelectedItem = CodeList.FindString(csnScript.CurrentWord, 1)
               End If
               
               'Show the list
               csnScript_CodeList = True
          End If
     Else
          
          'No list
          csnScript_CodeList = False
     End If
End Function

Private Function csnScript_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
     
     'Destroy list
     'CodeList.Destroy
     Set CodeList = Nothing
End Function


Private Function csnScript_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Accept the selection
          AcceptCodeListSelection
     End If
End Function

Private Function csnScript_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Get a reference
          Set CodeList = ListCtrl
          
          'Check if not holding a shift key
          If (CurrentShiftMask = 0) Then
          'If (LastKeyCode <> vbKeyShift) And (LastKeyCode <> vbKeyControl) Then
               
               'Check if the pressed key is not one of the delimiters
               If (InStr(1, ScriptConfig("delimiters"), Chr$(LastKeyCode), vbBinaryCompare) <= 0) Then
                    
                    'Any word typed
                    If (csnScript.CurrentWord <> "") Then
                         
                         'Select the best match
                         CodeList.SelectedItem = CodeList.FindString(csnScript.CurrentWord, 1)
                    End If
               End If
          End If
     End If
     
     'No.
     csnScript_CodeListSelWord = False
End Function


Private Function csnScript_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
     Dim FunctionDefinition As String
     
     'Default to no CodeTip
     csnScript_CodeTip = cmToolTipTypeNone
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Can we show a tooltip?
          If (FuncPosInfo.FunctionName <> "") Then
               
               'No CodeList shown now?
               If (CodeList Is Nothing) Then
                    
                    'Check if case sensitive
                    If CBool(ScriptConfig("casesensitive")) Then
                         
                         'Get function definition
                         If ScriptConfig("keywords").Exists(FuncPosInfo.FunctionName) Then FunctionDefinition = ScriptConfig("keywords")(FuncPosInfo.FunctionName)
                    Else
                         
                         'Get function definition
                         If ScriptConfig("keywords").Exists(LCase$(FuncPosInfo.FunctionName)) Then FunctionDefinition = ScriptConfig("keywords")(LCase$(FuncPosInfo.FunctionName))
                    End If
                    
                    'If there are brackets in the function, then use highlighting
                    If (InStr(1, FunctionDefinition, ScriptConfig("functionopen"), vbBinaryCompare) > 0) Then
                         
                         'Use highlighting in tooltip
                         csnScript_CodeTip = cmToolTipTypeFuncHighlight
                    Else
                         
                         'No highlighting
                         csnScript_CodeTip = cmToolTipTypeNormal
                    End If
               End If
          End If
     End If
End Function

Private Function csnScript_CodeTipCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip) As Boolean
     
     'Destroy the tooltip
     Set CodeTip = Nothing
End Function

Private Sub csnScript_CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
     Dim FunctionDefinition As String
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Set current tooltip
          Set CodeTip = ToolTipCtrl
          
          'Set Font
          CodeTip.Font.Name = "Verdana"
          CodeTip.Font.size = 9
          
          'Can we show a tooltip?
          If (FuncPosInfo.FunctionName <> "") Then
               
               'Check if case sensitive
               If CBool(ScriptConfig("casesensitive")) Then
                    
                    'Get function definition
                    If ScriptConfig("keywords").Exists(FuncPosInfo.FunctionName) Then FunctionDefinition = ScriptConfig("keywords")(FuncPosInfo.FunctionName)
               Else
                    
                    'Get function definition
                    If ScriptConfig("keywords").Exists(LCase$(FuncPosInfo.FunctionName)) Then FunctionDefinition = ScriptConfig("keywords")(LCase$(FuncPosInfo.FunctionName))
               End If
               
               'Set tooltip text
               CodeTip.TipText = FunctionDefinition
               
               'If there are brackets in the function, then use highlighting
               If (InStr(1, FunctionDefinition, ScriptConfig("functionopen"), vbBinaryCompare) > 0) Then
                    
                    'Use highlighting in tooltip
                    CodeTip.Argument = FuncPosInfo.ArgumentIndex
               End If
          End If
     End If
End Sub


Private Function csnScript_KeyDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
     Dim HelpURL As String
     
     'Update current shift mask
     CurrentShiftMask = Shift
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'ENTER and TAB are not intercepted by KeyPress, so catch it here if it is one of the delimiters
          If ((KeyCode = vbKeyReturn) Or (KeyCode = vbKeyTab)) And _
             (InStr(1, ScriptConfig("delimiters"), Chr$(KeyCode), vbBinaryCompare) > 0) Then
               
               'Cancel the enter key if code was inserted
               csnScript_KeyDown = AcceptCodeListSelection
          End If
          
          'F1?
          If (KeyCode = vbKeyF1) And (LastKeyCode <> KeyCode) Then
               
               'Is the cursor on a known keyword?
               If (csnScript.CurrentToken = cmTokenTypeKeyword) Or (FuncPosInfo.FunctionName <> "") Then
                    
                    'Is the help url configured?
                    If (ScriptConfig("keywordhelp") <> "") Then
                         
                         'Change mousepointer
                         Screen.MousePointer = vbHourglass
                         
                         'Fill in the keyword
                         If (csnScript.CurrentToken = cmTokenTypeKeyword) Then
                              
                              'From current word
                              HelpURL = Replace$(ScriptConfig("keywordhelp"), "%K", csnScript.CurrentWord)
                         Else
                              
                              'From current function
                              HelpURL = Replace$(ScriptConfig("keywordhelp"), "%K", FuncPosInfo.FunctionName)
                         End If
                         
                         'Go to website
                         Execute HelpURL, "", SW_SHOW, False
                         
                         'Change mousepointer
                         Screen.MousePointer = vbNormal
                    End If
               End If
          End If
     End If
     
     'Save last key
     LastKeyCode = KeyCode
End Function

Private Function csnScript_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Check if the pressed key is one of the delimiters
          If (InStr(1, ScriptConfig("delimiters"), Chr$(KeyAscii), vbBinaryCompare) > 0) Then
               
               'Accept the code in CodeList when CodeList is open
               AcceptCodeListSelection
          End If
     End If
End Function

Private Function csnScript_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
     Dim FunctionDefinition As String
     Dim ShowNewCodeTip As Boolean
     
     'Update current shift mask
     CurrentShiftMask = Shift
     
     'Check if we may show a new CodeTip with this key
     If (KeyCode <> vbKeyLeft) And (KeyCode <> vbKeyUp) And _
        (KeyCode <> vbKeyRight) And (KeyCode <> vbKeyDown) And _
        (KeyCode <> vbKeyShift) And (KeyCode <> vbKeyControl) Then ShowNewCodeTip = True
     
     'Check if a script configuration is given
     If Not (ScriptConfig Is Nothing) Then
          
          'Get function info
          FuncPosInfo = FunctionPositionInfo()
          
          'Can we show a tooltip?
          If (FuncPosInfo.FunctionName <> "") Then
               
               'CodeTip exists?
               If Not (CodeTip Is Nothing) Then
                    
                    'CodeTip displayed?
                    If (CodeTip.hWnd <> 0) Then
                         
                         'Check if case sensitive
                         If CBool(ScriptConfig("casesensitive")) Then
                              
                              'Get function definition
                              If ScriptConfig("keywords").Exists(FuncPosInfo.FunctionName) Then FunctionDefinition = ScriptConfig("keywords")(FuncPosInfo.FunctionName)
                         Else
                              
                              'Get function definition
                              If ScriptConfig("keywords").Exists(LCase$(FuncPosInfo.FunctionName)) Then FunctionDefinition = ScriptConfig("keywords")(LCase$(FuncPosInfo.FunctionName))
                         End If
                         
                         'Update tooltip
                         CodeTip.TipText = FunctionDefinition
                         
                         'If there are brackets in the function, then use highlighting
                         If (InStr(1, FunctionDefinition, ScriptConfig("functionopen"), vbBinaryCompare) > 0) Then
                              
                              'Use highlighting in tooltip
                              CodeTip.Argument = FuncPosInfo.ArgumentIndex
                         End If
                    Else
                         
                         'Invoke tooltip command
                         If ShowNewCodeTip Then csnScript.ExecuteCmd cmCmdCodeTip, ""
                    End If
               Else
                    
                    'Invoke tooltip command
                    If ShowNewCodeTip Then csnScript.ExecuteCmd cmCmdCodeTip, ""
               End If
          Else
               
               'CodeTip exists?
               If Not (CodeTip Is Nothing) Then
                    
                    'CodeTip displayed?
                    If (CodeTip.hWnd <> 0) Then CodeTip.Destroy
                    
                    'Destroy tooltip
                    Set CodeTip = Nothing
               End If
          End If
     End If
     
     'No key down
     LastKeyCode = 0
End Function


Private Sub csnScript_KillFocus(ByVal Control As CodeSenseCtl.ICodeSense)
     
     'No key down
     LastKeyCode = 0
End Sub


Private Function csnScript_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
     
     'Show custom popup menu
     PopupMenu mnuEdit
     
     'Do not show the default popup menu
     csnScript_RClick = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
     
     'Adjust shift mask
     CurrentShiftMask = Shift
End Sub


Private Sub Form_Load()
     
     'Script editor is now open
     ScriptEditor = True
     
     'Apply window sizes
     With frmScript
          If (Config("scriptwindow").Exists("left")) Then .left = Config("scriptwindow")("left")
          If (Config("scriptwindow").Exists("top")) Then .top = Config("scriptwindow")("top")
          If (Config("scriptwindow")("width") > 1500) Then .width = Config("scriptwindow")("width")
          If (Config("scriptwindow")("height") > 1500) Then .height = Config("scriptwindow")("height")
          If (Config("scriptwindow").Exists("windowstate")) Then .WindowState = Config("scriptwindow")("windowstate")
     End With
     
     'Allow textbox updating
     AllowUpdate = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
     
     'Do a resize to keep window sizes
     Form_Resize
     
     'Clean up
     If Not (CodeList Is Nothing) Then If (CodeList.hWnd <> 0) Then CodeList.Destroy
     If Not (CodeTip Is Nothing) Then If (CodeTip.hWnd <> 0) Then CodeTip.Destroy
     Set ScriptConfig = Nothing
     Set CodeList = Nothing
     Set CodeTip = Nothing
     Erase KeywordsList()
     
     'Save the script
     Save
     
     'Script editor is now closed
     ScriptEditor = False
     
     'Focus to form if possible
     On Error Resume Next
     frmMain.Show
     frmMain.SetFocus
End Sub


Private Sub Form_Resize()
     On Local Error Resume Next
     
     'Resize script
     csnScript.left = 2
     csnScript.width = ScaleWidth - csnScript.left - 2
     csnScript.height = ScaleHeight - csnScript.top - 2
     
     'Resize frame
     fraNoScript.width = ScaleWidth - fraNoScript.left * 2
     fraNoScript.height = ScaleHeight - fraNoScript.top - fraNoScript.left
     lblNoScript.width = fraNoScript.width * Screen.TwipsPerPixelX
     cmdScriptMake.left = (lblNoScript.width - cmdScriptMake.width) \ 2
     
     'Only do this when visible
     If (frmScript.visible = True) Then
          
          'Save windowstate
          Config("scriptwindow")("windowstate") = frmScript.WindowState
          
          'Check if it has a valid size now
          If (frmScript.WindowState = vbNormal) Then
               
               'Save window size
               Config("scriptwindow")("left") = frmScript.left
               Config("scriptwindow")("top") = frmScript.top
               Config("scriptwindow")("width") = frmScript.width
               Config("scriptwindow")("height") = frmScript.height
          End If
     End If
End Sub

Public Sub LoadLumpScript()
     Dim ScriptType As ENUM_MAPLUMPTYPES
     Dim ScriptConfigFile As New clsConfiguration
     Dim LumpName As String
     Dim lumpindex As Long
     Dim AllData As String
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Get the script type
     If (Trim$(lblLumpName.Caption) = "~") Then
          
          'Script lump is map lump name
          ScriptType = GetMapLumpType(maplumpname)
          
          'Definitely required
          ScriptType = ScriptType Or ML_REQUIRED
          LumpName = "MAP01"
     Else
          
          'Script lump is as specified
          ScriptType = GetMapLumpType(lblLumpName.Caption)
          LumpName = lblLumpName.Caption
     End If
     
          
     'Check if script is DED
     If (ScriptType And ML_CUSTOM) = ML_CUSTOMDED Then
          
          'Load DED definitions
          ScriptConfigFile.LoadConfiguration App.Path & "\DED.cfg"
          
          'Get the object
          Set ScriptConfig = ScriptConfigFile.Root(True)
          
     'Check if script is FS
     ElseIf (ScriptType And ML_CUSTOM) = ML_CUSTOMFS Then
          
          'Load FS definitions
          ScriptConfigFile.LoadConfiguration App.Path & "\FS.cfg"
          
          'Get the object
          Set ScriptConfig = ScriptConfigFile.Root(True)
          
     'Check if script is DEHACKED
     ElseIf (ScriptType And ML_CUSTOM) = ML_CUSTOMDEHACKED Then
          
          'Load DEHACKED definitions
          ScriptConfigFile.LoadConfiguration App.Path & "\Dehacked.cfg"
          
          'Get the object
          Set ScriptConfig = ScriptConfigFile.Root(True)
          
     'Check if script is ACS
     ElseIf (ScriptType And ML_CUSTOM) = ML_CUSTOMACS Then
          
          'Load ACS definitions
          ScriptConfigFile.LoadConfiguration App.Path & "\ACS.cfg"
          
          'Get the object
          Set ScriptConfig = ScriptConfigFile.Root(True)
          
     Else
          
          'No script definitions
          Set ScriptConfig = Nothing
     End If
     
     'Initialize settings and keywords
     SetupScriptControl
     
     'Check if this is an ACS script
     If (ScriptType And ML_CUSTOM) = ML_CUSTOMACS Then
          
          'Find the SCRIPTS lump
          lumpindex = FindLumpIndex(TempWAD, 1, "SCRIPTS")
          
          'Check if found
          If (lumpindex > 0) Then
               
               'Load the SCRIPTS lump if anything
               AllData = TempWAD.GetLump(lumpindex)
          Else
               
               'Find the original lump
               lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
               
               'Check if found and has contents
               If (lumpindex > 0) And (TempWAD.LumpSize(lumpindex) > 0) Then
                    
                    'Load the decompiled script
                    AllData = DecompiledACS(lumpindex)
               End If
          End If
     Else
          
          'Find the lump
          lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
          
          'Check if found
          If (lumpindex > 0) Then
               
               'Load script
               AllData = TempWAD.GetLump(lumpindex)
          End If
     End If
     
     'Check if any script data
     If (Len(AllData) > 0) Then
          
          'Show script
          csnScript.Text = AllData
          csnScript.visible = True
          fraNoScript.visible = False
          cmdScriptExport.Enabled = True
          scriptchanged = False
          
          'ACS must be compiled, enable button
          cmdScriptCompile.Enabled = ((ScriptType And ML_CUSTOMACS) = ML_CUSTOMACS)
     Else
          
          'No script
          csnScript.visible = False
          fraNoScript.visible = True
          lblNoScript.Caption = "This script does not yet exist in the currently loaded map." & vbLf & _
                                "Click the button below to add this script now."
          cmdScriptExport.Enabled = False
          cmdScriptCompile.Enabled = False
          scriptchanged = False
     End If
     
     'Change mousepointer
     Screen.MousePointer = vbDefault
End Sub

Private Sub SaveLumpScript()
     Dim ScriptType As ENUM_MAPLUMPTYPES
     Dim LumpName As String
     Dim lumpindex As Long
     
     'Change mousepointer
     Screen.MousePointer = vbHourglass
     
     'Get the script type
     ScriptType = GetMapLumpType(lblLumpName.Caption)
     If (Trim$(lblLumpName.Caption) = "~") Then
          ScriptType = ScriptType Or ML_REQUIRED
          LumpName = "MAP01"
     Else
          LumpName = lblLumpName.Caption
     End If
     
     'Check if this is an ACS script
     If (ScriptType And ML_CUSTOMACS) = ML_CUSTOMACS Then
          
          'Find and remove the SCRIPTS lump
          lumpindex = FindLumpIndex(TempWAD, 1, "SCRIPTS")
          If (lumpindex > 0) Then TempWAD.DeleteLump lumpindex
          
          'Find and remove the original lump
          'LumpIndex = FindLumpIndex(TempWAD, 1, LumpName)
          'If (LumpIndex > 0) Then TempWAD.DeleteLump LumpIndex
          
          'Save the script to SCRIPTS lump
          TempWAD.AddLump csnScript.Text, "SCRIPTS", lumpindex
          
          'Compile and save the script to original lump
          'TempWAD.AddLump CompiledACS(csnScript.Text), LumpName
     Else
          
          'Find and remove the exisitng lump
          lumpindex = FindLumpIndex(TempWAD, 1, LumpName)
          If (lumpindex > 0) Then TempWAD.DeleteLump lumpindex
          
          'Save the script to lump
          TempWAD.AddLump csnScript.Text, LumpName, lumpindex
     End If
     
     'Save changes to temp file
     TempWAD.WriteChanges
     
     'Map changed
     mapchanged = scriptchanged
     
     'Change mousepointer
     Screen.MousePointer = vbDefault
End Sub

Private Sub itmClearBookmarks_Click()
     csnScript.ExecuteCmd cmCmdBookmarkClearAll
End Sub

Private Sub itmCopy_Click()
     If (csnScript.SelLength > 0) Then csnScript.Copy
End Sub

Private Sub itmCut_Click()
     If (csnScript.SelLength > 0) Then csnScript.Cut
End Sub

Private Sub itmDelete_Click()
     If (csnScript.SelLength > 0) Then csnScript.DeleteSel
End Sub

Private Sub itmFind_Click()
     csnScript.ExecuteCmd cmCmdFind
End Sub


Private Sub itmNextBookmark_Click()
     csnScript.ExecuteCmd cmCmdBookmarkNext
End Sub

Private Sub itmPaste_Click()
     csnScript.Paste
End Sub

Private Sub itmPrevBookmark_Click()
     csnScript.ExecuteCmd cmCmdBookmarkPrev
End Sub

Private Sub itmRedo_Click()
     csnScript.Redo
End Sub

Private Sub itmReplace_Click()
     csnScript.ExecuteCmd cmCmdFindReplace
End Sub

Private Sub itmToggleBookmark_Click()
     csnScript.ExecuteCmd cmCmdBookmarkToggle
End Sub

Private Sub itmUndo_Click()
     csnScript.Undo
End Sub

Private Sub mnuEdit_Click()
     
     'Disable impossible items
     itmCopy.Enabled = (csnScript.SelLength > 0)
     itmCut.Enabled = (csnScript.SelLength > 0)
     itmDelete.Enabled = (csnScript.SelLength > 0)
     itmPaste.Enabled = Clipboard.GetFormat(vbCFText) Or Clipboard.GetFormat(vbCFRTF)
     
End Sub


