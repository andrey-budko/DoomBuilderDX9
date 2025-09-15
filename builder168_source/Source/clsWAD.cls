VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWAD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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


'WAD Types
Private Const WADTYPE_IWAD As String = "IWAD"
Private Const WADTYPE_PWAD As String = "PWAD"

'Buffer size for copying data
Private Const BUFFERSIZE As Long = 500000


'WAD Header
Private hdrType As String * 4           'Should be IWAD or PWAD
Private hdrLumps As Long                'Number of Lumps in file
Private hdrTableOffset As Long          'Address of Lumps Table

'WAD Lump Entry
Private Type lmpEntry
     DataOffset As Long                 'Address of Lump Data
     DataLength As Long                 'Length of Lump Data
     LumpName As String * 8             'Name of the Lump
End Type

'WAD Lump Table
Private lmpTable() As lmpEntry          'Will hold the WAD Lump Table

'Open WAD File
Private wadFilename As String           'Empty string if no WAD opened
Private wadFilebuffer As Integer        'Filebuffer of WAD

Public Sub AddLump(ByRef Data As String, ByRef LumpName As String, Optional ByVal LumpIndex As Long = 0)
     '
     ' Data: Data to put in this Lump
     ' LumpName: Name to assign to this Lump
     ' LumpIndex: Position to put this Lump entry at in the table (0 = add to the end)
     '
     
     Dim i As Long
     
     'Check if a file is open
     If (wadFilename = "") Then Err.Raise 75   'Path/File access error
     
     'Put the Data at the startbyte address of the table
     '(this is the end of the data section, where this data
     'must be added to the package)
     Put #wadFilebuffer, hdrTableOffset + 1, Data
     
     'Add a table entry
     hdrLumps = hdrLumps + 1
     ReDim Preserve lmpTable(1 To hdrLumps) As lmpEntry
     
     'Check if we should move entries down
     If (LumpIndex > 0) Then
          
          'Move entries down (this must be done in backward order)
          For i = hdrLumps To LumpIndex + 1 Step -1
               
               'Move down
               lmpTable(i) = lmpTable(i - 1)
          Next i
     Else
          
          'Set the new LumpIndex
          LumpIndex = hdrLumps
     End If
     
     'Fill the table entry
     With lmpTable(LumpIndex)
          .LumpName = NullPadded(UCase$(LumpName))
          .DataOffset = hdrTableOffset
          .DataLength = 0
          .DataLength = Len(Data)
     End With
     
     'Update header info
     hdrTableOffset = hdrTableOffset + Len(Data)
     
'     'Update header
'     WriteHeader
'
'     'Write table
'     WriteTable
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Sub

Private Sub Class_Terminate()
     
     'Make sure files are closed
     CloseFile
End Sub

Public Sub CloseFile()
     
     'Check if a previous file was open
     If wadFilename <> "" Then
          
          'Close file
          Close wadFilebuffer
          wadFilename = ""
     End If
End Sub

Public Sub DeleteLump(ByVal Index As Long)
     '
     ' Index: Lump index to delete
     '
     '    NOTE: This function does not clean up the space in the file.
     '          To clean the wasted space, rebuild the WAD.
     '
     
     Dim i As Long
     
     'Check if a file is open
     If (wadFilename = "") Then Err.Raise 75   'Path/File access error
     
     'Move all entries 1 up
     For i = Index + 1 To hdrLumps
          
          'Move up
          lmpTable(i - 1) = lmpTable(i)
     Next i
     
     'Resize memory
     hdrLumps = hdrLumps - 1
     If (hdrLumps > 0) Then ReDim Preserve lmpTable(1 To hdrLumps) As lmpEntry
     
'     'Update header
'     WriteHeader
'
'     'Write table
'     WriteTable
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Sub

Public Sub ExportLump(ByVal Index As Long, ByRef Filename As String)
     '
     ' Index: Lump index to export
     ' Filename: File to write
     '
     
     Dim nFilebuffer As Integer
     Dim DataBuffer As String
     Dim DataSteps As Long
     Dim i As Long
     
     'Check if a file is open
     If (wadFilename = "") Then Err.Raise 75   'Path/File access error
     
     'Remove the target if it already exists
     If (Dir(Filename) <> "") Then Kill Filename
     
     'Create filebuffer
     nFilebuffer = FreeFile
     
     'Open target file
     Open Filename For Binary Lock Read Write As #nFilebuffer
     
     'Seek to beginning of file
     Seek #wadFilebuffer, lmpTable(Index).DataOffset + 1
     
     'Export the file with steps of BUFFERSIZE
     DataBuffer = Space$(BUFFERSIZE)
     DataSteps = Int(lmpTable(Index).DataLength / BUFFERSIZE)
     For i = 1 To DataSteps
          
          'Copy the data
          Get #wadFilebuffer, , DataBuffer
          Put #nFilebuffer, , DataBuffer
     Next i
     
     'Check if there are more bytes to copy
     If (lmpTable(Index).DataLength > (DataSteps * BUFFERSIZE)) Then
          
          'Copy the data
          DataBuffer = Space$(lmpTable(Index).DataLength - (DataSteps * BUFFERSIZE))
          Get #wadFilebuffer, , DataBuffer
          Put #nFilebuffer, , DataBuffer
     End If
     
     'Close the file
     Close #nFilebuffer
End Sub

Public Property Get FileBuffer() As Integer
     
     'Return the filebuffer
     FileBuffer = wadFilebuffer
End Property

Public Property Get Filename() As String
     
     'Return the filename
     Filename = wadFilename
End Property

Public Function GetLump(ByVal Index As Long) As String
     
     GetLump = Space$(lmpTable(Index).DataLength)
     Get #wadFilebuffer, lmpTable(Index).DataOffset + 1, GetLump
End Function

Public Sub ImportLump(ByRef Filename As String, ByRef LumpName As String, Optional ByVal LumpIndex As Long = 0)
     '
     ' Filename: File to add to the WAD
     ' LumpName: Name to assign to this Lump
     ' LumpIndex: Position to put this Lump entry at in the table (0 = add to the end)
     '
     
     Dim nFilebuffer As Integer
     Dim DataBuffer As String
     Dim DataSteps As Long
     Dim BytesLeft As Long
     Dim i As Long
     
     'Check if a file is open
     If (wadFilename = "") Then Err.Raise 75   'Path/File access error
     
     'Check if the file can be found
     If (Dir(Filename) = "") Then Err.Raise 53 'File not found
     
     'Create filebuffer
     nFilebuffer = FreeFile
     
     'Open original file
     Open Filename For Binary Lock Write As #nFilebuffer
     
     'Go to the startbyte address of the table
     '(this is the end of the data section, where this file
     'must be added to the package)
     Seek #wadFilebuffer, hdrTableOffset + 1
     
     'Import the file with steps of buffersize
     DataBuffer = Space$(BUFFERSIZE)
     DataSteps = Int(LOF(nFilebuffer) / BUFFERSIZE)
     For i = 1 To DataSteps
          
          'Copy the data
          Get #nFilebuffer, , DataBuffer
          Put #wadFilebuffer, , DataBuffer
     Next i
     
     'Check if there are more bytes to copy
     BytesLeft = LOF(nFilebuffer) - (DataSteps * BUFFERSIZE)
     If (BytesLeft > 0) Then
          
          'Copy the data
          DataBuffer = Space$(BytesLeft)
          Get #nFilebuffer, , DataBuffer
          Put #wadFilebuffer, , DataBuffer
     End If
     
     'Add a table entry
     hdrLumps = hdrLumps + 1
     ReDim Preserve lmpTable(1 To hdrLumps) As lmpEntry
     
     'Check if we should move entries down
     If (LumpIndex > 0) Then
          
          'Move entries down (this must be done in backward order)
          For i = hdrLumps To LumpIndex + 1 Step -1
               
               'Move down
               lmpTable(i) = lmpTable(i - 1)
          Next i
     Else
          
          'Set the new LumpIndex
          LumpIndex = hdrLumps
     End If
     
     'Fill the table entry
     With lmpTable(LumpIndex)
          .LumpName = NullPadded(LumpName)
          .DataOffset = hdrTableOffset
          .DataLength = LOF(nFilebuffer)
     End With
     
     'Update header info
     hdrTableOffset = hdrTableOffset + LOF(nFilebuffer)
     
     'Close the file
     Close #nFilebuffer
     
'     'Update header
'     WriteHeader
'
'     'Write table
'     WriteTable
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Sub

Public Property Get IWAD() As Boolean
     
     'Return True for IWAD type
     If hdrType = WADTYPE_IWAD Then IWAD = True
End Property

Public Property Let IWAD(ByVal NewType As Boolean)
     
     'Set True for IWAD, False for PWAD
     If (NewType = True) Then hdrType = WADTYPE_IWAD Else hdrType = WADTYPE_PWAD
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Property

Public Property Get LumpAddress(ByVal Index As Long) As Long
     
     'Return the Lump address
     LumpAddress = lmpTable(Index).DataOffset
End Property

Public Property Let LumpAddress(ByVal Index As Long, ByVal NewAddress As Long)
     
     'Set the Lump address
     lmpTable(Index).DataOffset = NewAddress
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Property

Public Property Get LumpCount() As Long
     
     'Return the number of Lumps
     LumpCount = hdrLumps
End Property

Public Property Let LumpName(ByVal Index As Long, ByRef NewName As String)
     
     'Set the Lump name as padded string
     lmpTable(Index).LumpName = NullPadded(UCase$(NewName))
     
'     'Write table
'     WriteTable
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Property

Public Property Get LumpName(ByVal Index As Long) As String
     
     'Return the Lump name as normal string
     LumpName = UnPadded(lmpTable(Index).LumpName)
End Property

Public Property Get LumpnamePadded(ByVal Index As Long) As String
     
     'Return the Lump name as normal string
     LumpnamePadded = lmpTable(Index).LumpName
End Property

Public Property Get LumpSize(ByVal Index As Long) As Long
     
     'Return the Lump size
     LumpSize = lmpTable(Index).DataLength
End Property

Public Property Let LumpSize(ByVal Index As Long, ByVal NewSize As Long)
     
     'Set the Lump size
     lmpTable(Index).DataLength = NewSize
     
'     'Write table
'     WriteTable
     
     'NOTE: You must use WriteChanges method to write these changes to file!
End Property

Public Sub NewFile(ByVal Filename As String, Optional ByVal CreateIWAD As Boolean)
     '
     ' Filename: WAD file to create
     ' CreateIWAD: True for IWAD type, False for PWAD type
     '
     
     'Close any open files
     CloseFile
     
     'Remove the file if it already exists
     If (Dir(Filename) <> "") Then Kill Filename
     
     'Create filebuffer
     wadFilebuffer = FreeFile
     wadFilename = Filename
     
     'Clear memory
     hdrLumps = 0
     hdrTableOffset = 13
     If (CreateIWAD = True) Then hdrType = WADTYPE_IWAD Else hdrType = WADTYPE_PWAD
     
     'Open package
     Open wadFilename For Binary Lock Write As #wadFilebuffer
     
     'Write the header
     WriteHeader
     
     'Remove table from memory
     ReDim lmpTable(1 To 1) As lmpEntry
End Sub

Private Function NullPadded(ByRef Src As String) As String
     
     'Check the length of Src
     If Len(Src) < 8 Then
          
          'Make an 8 byte Null padded string
          NullPadded = Src & String$(8 - Len(Src), vbNullChar)
     ElseIf Len(Src) = 8 Then
          
          'Src just fits in an 8 bytes string
          NullPadded = Src
     Else
          
          'Chop Src to 8 bytes
          NullPadded = left$(Src, 8)
     End If
End Function

Public Sub OpenFile(ByRef Filename As String, ByVal ReadOnly As Boolean)
     '
     ' Filename: WAD file to open
     '
     
     Dim i As Long
     
     'Close any open files
     CloseFile
     
     'Create filebuffer
     wadFilename = Filename
     wadFilebuffer = FreeFile
     
     'Open package
     If ReadOnly Then
          Open wadFilename For Binary Access Read Lock Write As #wadFilebuffer
     Else
          Open wadFilename For Binary Access Read Write Lock Read Write As #wadFilebuffer
          'Open wadFilename For Binary Lock Write As #wadFilebuffer
     End If
     
     'Read the header
     Get #wadFilebuffer, 1, hdrType
     Get #wadFilebuffer, , hdrLumps
     Get #wadFilebuffer, , hdrTableOffset
     
     'Verify format
     If (hdrType <> WADTYPE_IWAD) And (hdrType <> WADTYPE_PWAD) Then Err.Raise vbObjectError, , "File is not in IWAD or PWAD format"
     
     'Verify number of lumps
     If (hdrLumps < 0) Then Err.Raise vbObjectError, , "Incorrect number of Lumps in file (" & hdrLumps & ")"
     
     'Verify table offset
     If (hdrTableOffset < 12) Or (hdrTableOffset > LOF(wadFilebuffer) + 1) Then Err.Raise vbObjectError, , "Incorrect Lump Table address (" & hdrTableOffset & ")"
     
     'Check if we have a table
     If (hdrLumps > 0) Then
          
          'Create table in memory
          ReDim lmpTable(1 To hdrLumps) As lmpEntry
          
          'Read the table
          Get #wadFilebuffer, hdrTableOffset + 1, lmpTable()
          
          'Go for each item to make uppercase
          For i = 1 To hdrLumps
               
               'Make uppercase
               lmpTable(i).LumpName = Padded(UCase$(UnPadded(lmpTable(i).LumpName)), 8)
          Next i
     End If
End Sub

Public Sub SetLump(ByVal Index As Long, ByRef Data As String)
     
     'Set data in file
     Put #wadFilebuffer, lmpTable(Index).DataOffset + 1, Data
     
     'WARNING: Be carefull with this procedure! It may overwrite other data if you use it incorrectly!
End Sub

Private Function UnPadded(ByRef Src As String) As String
     
     'Remove Null padding from string
     UnPadded = Replace$(Src, vbNullChar, "")
End Function

Public Sub WriteChanges()
     
     'Write Header
     WriteHeader
     
     'Write Table
     WriteTable
End Sub

Private Sub WriteHeader()
     
     'Write the header
     Put #wadFilebuffer, 1, hdrType
     Put #wadFilebuffer, , hdrLumps
     Put #wadFilebuffer, , hdrTableOffset
End Sub

Private Sub WriteTable()
     
     'Write the Table
     Put #wadFilebuffer, hdrTableOffset + 1, lmpTable()
End Sub
