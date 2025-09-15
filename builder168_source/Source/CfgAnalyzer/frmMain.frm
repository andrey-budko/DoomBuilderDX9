VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuration analyzer"
   ClientHeight    =   9585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   639
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   537
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopyIPs 
      Caption         =   "Copy IPs"
      Height          =   375
      Left            =   2235
      TabIndex        =   2
      Top             =   9090
      Width           =   1965
   End
   Begin VB.CommandButton cmdCopySetting 
      Caption         =   "Copy Setting"
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   9090
      Width           =   1965
   End
   Begin MSComctlLib.ListView lstResult 
      Height          =   8490
      Left            =   135
      TabIndex        =   0
      Top             =   495
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   14975
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "setting"
         Text            =   "Setting"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "value"
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "usage"
         Text            =   "Usage"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IPs"
         Object.Width           =   4763
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   210
      Left            =   150
      TabIndex        =   3
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AnalyzeConfig(ByRef str As String, ByRef ip As String)
     
     Dim cfg As New clsConfiguration
     
     'Parse configuration
     cfg.InputConfiguration str
     
     'Recursively analyze the config
     AnalyzeStructure cfg.Root(True), "", ip
     
End Sub

Private Sub AnalyzeStructure(ByRef obj As Dictionary, ByRef structname As String, ByRef ip As String)
     
     Dim i As Long, k As Long
     Dim Keys As Variant
     Dim Value As String
     Dim name As String
     Dim item As ListItem
     Dim insertbelow As Long
     Dim counted As Boolean
     
     'Loop through all items
     Keys = obj.Keys
     For i = 0 To obj.Count - 1
          
          'Determine name of next struct
          If (structname = "") Then name = Keys(i) Else name = structname & "\" & Keys(i)
          
          'Check if this is another structure
          If VarType(obj(Keys(i))) = vbObject Then
               
               'Check the name (exclude some structures)
               If (StrComp(name, "recent") <> 0) And _
                  (StrComp(name, "iwads") <> 0) And _
                  (StrComp(name, "mainwindow") <> 0) And _
                  (StrComp(name, "shortcuts") <> 0) And _
                  (StrComp(name, "palette") <> 0) And _
                  (StrComp(name, "defaulttexture") <> 0) And _
                  (StrComp(name, "defaultsector") <> 0) And _
                  (StrComp(name, "scriptwindow") <> 0) Then
                    
                    'Recursively analyze this object
                    AnalyzeStructure obj(Keys(i)), name, ip
               End If
               
          'Check if not a string type (string types do not concern me)
          ElseIf (VarType(obj(Keys(i))) <> vbString) Then
               
               'Make fixed value
               Value = CStr(obj(Keys(i)))
               
               'Begin list search
               counted = False
               insertbelow = lstResult.ListItems.Count
               
               'Go for all items in the results list
               For k = 1 To lstResult.ListItems.Count
                    
                    'Get this list item
                    Set item = lstResult.ListItems(k)
                    
                    'Check if the name is the same
                    If StrComp(item.Text, name) = 0 Then
                         
                         'If no equal can be found
                         'Insert the new setting below this one
                         insertbelow = k
                         
                         'Check if the value is the same
                         If StrComp(item.SubItems(1), Value) = 0 Then
                              
                              'Same name and value, increase count and add ip
                              item.SubItems(2) = Val(item.SubItems(2)) + 1
                              item.SubItems(3) = item.SubItems(3) & " || " & ip
                              counted = True
                              
                              'Done here
                              Exit For
                         End If
                    End If
               Next k
               
               'Check if not counted
               If (counted = False) Then
                    
                    'Add new item
                    Set item = lstResult.ListItems.Add(insertbelow + 1, , name)
                    
                    'Add subitems
                    item.ListSubItems.Add 1, "value", Value
                    item.ListSubItems.Add 2, "count", 1
                    item.ListSubItems.Add 3, "ips", ip
               End If
          End If
     Next i
End Sub

Private Function SpacedNumber(ByVal num As Long) As String
     
     'Add spaces
     SpacedNumber = Space$(2 - Len(CStr(num))) & CStr(num) & Space$(2 - Len(CStr(num)))
End Function


Private Sub cmdCopyIPs_Click()
     
     'Anything selected?
     If Not (lstResult.SelectedItem Is Nothing) Then
          
          'Copy IPs to clipboard
          Clipboard.Clear
          Clipboard.SetText lstResult.SelectedItem.SubItems(3)
     End If
End Sub

Private Sub cmdCopySetting_Click()
     
     'Anything selected?
     If Not (lstResult.SelectedItem Is Nothing) Then
          
          'Copy setting to clipboard
          Clipboard.Clear
          Clipboard.SetText lstResult.SelectedItem.Text
     End If
End Sub

Private Sub Form_Load()
     
     Const OPENINGLINE As String = "Doom Builder configuration submission from"
     Const BREAKLINE As String = "===================================================================="
     
     Dim UsedIPs As New Dictionary
     Dim FB As Integer
     Dim line As String
     Dim cfg As String
     Dim ip As String
     Dim ignored As Long
     Dim counted As Long
     Dim item As ListItem
     Dim k As Long
     Dim p As Single
     
     'Show status
     lblStatus.Caption = "Processing configurations..."
     Show
     Refresh
     
     'Open text file
     FB = FreeFile
     Open App.Path & "\configs.txt" For Input As #FB
     
     'Continue until end of file
     Do Until EOF(FB)
          
          'Read a line
          Line Input #FB, line
          
          'Check if this is a beginning line
          If (Left$(line, Len(OPENINGLINE)) = OPENINGLINE) Then
               
               'Count config
               counted = counted + 1
               
               'Get the ip address
               ip = Trim$(Mid$(line, Len(OPENINGLINE) + 1))
               
               'Check if not previously added
               If (UsedIPs.Exists(ip) = False) Then
                    
                    'Add ip to the list
                    UsedIPs.Add ip, ip
                    
                    'Continue reading until break line
                    Do Until EOF(FB) Or line = BREAKLINE
                         
                         'Read a line
                         Line Input #FB, line
                    Loop
                    
                    'New config starts here
                    cfg = ""
                    
                    'Read a line
                    Line Input #FB, line
                    
                    'Continue reading until next break line
                    Do Until EOF(FB) Or line = BREAKLINE
                         
                         'Add to config
                         cfg = cfg & line & vbCrLf
                         
                         'Read next line
                         Line Input #FB, line
                    Loop
                    
                    'Analyze this config now
                    AnalyzeConfig cfg, ip
               Else
                    
                    'Count ignored configs
                    ignored = ignored + 1
               End If
          End If
     Loop
     
     'Close the file
     Close #FB
     
     'Set up label
     lblStatus.Caption = "Processed " & counted & " configurations (" & ignored & " ignored)"
     
     'Go for all items in the results list
     For k = 1 To lstResult.ListItems.Count
          
          'Get this list item
          Set item = lstResult.ListItems(k)
          
          'Make usage with percentage
          p = (CSng(Val(item.SubItems(2))) / CSng(counted - ignored)) * CSng(100)
          item.SubItems(2) = SpacedNumber(Val(item.SubItems(2))) & "  (" & Format(p, "###0.0") & "%)"
     Next k
End Sub


Private Sub lstResult_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     
     'Check if already sorted by this column
     If lstResult.SortKey = (ColumnHeader.Index - 1) Then
          
          'Reverse sort
          If lstResult.SortOrder = lvwAscending Then
               lstResult.SortOrder = lvwDescending
          Else
               lstResult.SortOrder = lvwAscending
          End If
     Else
          
          'Change sort key
          lstResult.SortKey = ColumnHeader.Index - 1
          lstResult.SortOrder = lvwAscending
          lstResult.Sorted = True
     End If
End Sub


