VERSION 5.00
Begin VB.UserControl ctlValueBox 
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1200
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
   LockControls    =   -1  'True
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   80
   ToolboxBitmap   =   "ctlValueBox.ctx":0000
   Begin VB.VScrollBar scrValue 
      Height          =   360
      Left            =   960
      Max             =   0
      Min             =   9999
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   240
   End
   Begin VB.TextBox txtValue 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Text            =   "0"
      Top             =   30
      Width           =   960
   End
End
Attribute VB_Name = "ctlValueBox"
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


'Events
Event Change()
Event MouseDown()

'Properties
Private pEmptyAllowed As Boolean
Private pRelativeAllowed As Boolean
Private pRelativeScroll As Boolean
Private pUnsigned As Boolean

'Misc
Private ReflectChanges As Boolean

'Behavior
Private HasFocus As Boolean
Private CtrlHold As Boolean

Public Property Get EmptyAllowed() As Boolean
     EmptyAllowed = pEmptyAllowed
End Property

Public Property Let EmptyAllowed(ByVal New_EmptyAllowed As Boolean)
     pEmptyAllowed = New_EmptyAllowed
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
     Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     UserControl.Enabled = New_Enabled
     txtValue.Enabled = New_Enabled
     scrValue.Enabled = New_Enabled
     PropertyChanged "Enabled"
End Property

Private Function GetRealValue() As Long
     
     'Check if value must be converted
     If (pUnsigned) Then
          GetRealValue = CLng(scrValue.Value) + 32768
     Else
          GetRealValue = scrValue.Value
     End If
End Function

Private Function GetFakeValue(ByVal realvalue As Long) As Long
     
     'Check if value must be converted
     If (pUnsigned) Then
          GetFakeValue = realvalue - 32768
     Else
          GetFakeValue = realvalue
     End If
End Function


Public Property Get Max() As Integer
Attribute Max.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
     Max = scrValue.Min
End Property

Public Property Let Max(ByVal New_Max As Integer)
     scrValue.Min = New_Max
     PropertyChanged "Max"
End Property

Public Property Get MaxLength() As Long
Attribute MaxLength.VB_Description = "Returns/sets the maximum number of characters that can be entered in a control."
     MaxLength = txtValue.MaxLength
End Property

Public Property Let MaxLength(ByVal New_MaxLength As Long)
     txtValue.MaxLength = New_MaxLength
     PropertyChanged "MaxLength"
End Property

Public Property Get Min() As Integer
Attribute Min.VB_Description = "Returns/sets a scroll bar position's maximum Value property setting."
     Min = scrValue.Max
End Property

Public Property Let Min(ByVal New_Min As Integer)
     scrValue.Max = New_Min
     PropertyChanged "Min"
End Property

Public Property Let RelativeAllowed(ByVal New_RelativeAllowed As Boolean)
     pRelativeAllowed = New_RelativeAllowed
End Property

Public Property Get RelativeAllowed() As Boolean
     RelativeAllowed = pRelativeAllowed
End Property

Public Property Let RelativeScroll(ByVal New_RelativeScroll As Boolean)
     pRelativeScroll = New_RelativeScroll
End Property

Public Property Get RelativeScroll() As Boolean
     RelativeScroll = pRelativeScroll
End Property

Public Function RelativeValue(ByVal OriginalValue As Long) As Long
     On Local Error Resume Next
     
     'Check if theres anything given
     If (Replace$(Replace$(txtValue.Text, "-", ""), "+", "") <> "") Then
          
          'Check if the value is relative
          If (left$(txtValue.Text, 2) = "--") Or (left$(txtValue.Text, 2) = "++") Then
               
               'Add/Subtract to original
               RelativeValue = OriginalValue + Val(Mid$(txtValue.Text, 2))
          Else
               
               'Apply normally
               RelativeValue = Val(txtValue.Text)
          End If
     Else
          
          'Keep original value
          RelativeValue = OriginalValue
     End If
End Function

Public Property Let SelLength(ByVal NewLength As Long)
     txtValue.SelLength = NewLength
End Property

Public Property Get SelLength() As Long
     SelLength = txtValue.SelLength
End Property

Public Property Let SelStart(ByVal NewStart As Long)
     txtValue.SelStart = NewStart
End Property

Public Property Get SelStart() As Long
     SelStart = txtValue.SelStart
End Property

Public Property Let Unsigned(ByVal unsign As Boolean)
     pUnsigned = unsign
End Property

Public Property Get Unsigned() As Boolean
     Unsigned = pUnsigned
End Property

Private Sub scrValue_Change()
     Dim realvalue As Long
     On Local Error Resume Next
     
     'Check if changes to textbox are allowed
     If ReflectChanges Then
          
          'Check if scrolling with relative values
          If pRelativeScroll Then
               
               'Check if positive
               If (realvalue >= 0) Then
                    txtValue.Text = "++" & CStr(GetRealValue)
               Else
                    txtValue.Text = "-" & CStr(GetRealValue)
               End If
          Else
               
               'Normal value
               txtValue.Text = GetRealValue
          End If
          
          'Focus to textbox
          'If HasFocus Then txtValue.SetFocus
          'txtValue.SetFocus
          txtValue.SelStart = Len(txtValue.Text)
     End If
End Sub

Private Sub scrValue_GotFocus()
     'If HasFocus Then txtValue.SetFocus
     txtValue.SetFocus
End Sub

Private Sub scrValue_KeyUp(KeyCode As Integer, Shift As Integer)
     'If HasFocus Then txtValue.SetFocus
     txtValue.SetFocus
End Sub

Public Property Get SmallChange() As Integer
Attribute SmallChange.VB_Description = "Returns/sets amount of change to Value property in a scroll bar when user clicks a scroll arrow."
     SmallChange = scrValue.SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Integer)
     scrValue.SmallChange = New_SmallChange
     PropertyChanged "SmallChange"
End Property

Public Property Get Text() As String
     Text = txtValue.Text
End Property

Public Property Let Text(ByVal New_Value As String)
     On Local Error Resume Next
     txtValue.Text = New_Value
     txtValue_Validate False
     PropertyChanged "Value"
End Property

Private Sub txtValue_Change()
     Dim Cancel As Boolean
     RaiseEvent Change
     
     'Check if anything given
     If (Replace$(Replace$(txtValue.Text, "-", ""), "+", "") <> "") Then
          
          'Check if within range
          If (GetFakeValue(Val(txtValue.Text)) >= scrValue.Min) And _
             (GetFakeValue(Val(txtValue.Text)) <= scrValue.Max) Then
               
               'Validate
               txtValue_Validate Cancel
          End If
     End If
End Sub

Private Sub txtValue_GotFocus()
     HasFocus = True
End Sub

Private Sub txtValue_KeyDown(KeyCode As Integer, Shift As Integer)
     If (Shift And vbCtrlMask) = vbCtrlMask Then CtrlHold = True
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
     
     'When CTRL is hold, allow the key
     If (CtrlHold) Then Exit Sub
     
     'Check if key is allowed
     If ((KeyAscii <> 8) And (KeyAscii <> 45) And ((KeyAscii <> 43) Or (pRelativeAllowed = False)) And ((KeyAscii < 48) Or (KeyAscii > 57))) Then KeyAscii = 0
     
     'Check if relative is allowed
     If pRelativeAllowed Then
          
          'Check if key is for relative change
          If (KeyAscii = 45) Or (KeyAscii = 43) Then
               
               'Check if already enough of these signs
               If (InStr(txtValue.Text, "++") > 0) Or (InStr(txtValue.Text, "--") > 0) Then KeyAscii = 0
          End If
     Else
          
          'Check if key is for relative change
          If (KeyAscii = 43) Then KeyAscii = 0
          If (KeyAscii = 45) Then
               
               'Check if already enough of these signs
               If (InStr(txtValue.Text, "-") > 0) Then KeyAscii = 0
          End If
     End If
End Sub

Private Sub txtValue_KeyUp(KeyCode As Integer, Shift As Integer)
     If (Shift And vbCtrlMask) = vbCtrlMask Then CtrlHold = False
End Sub

Private Sub txtValue_LostFocus()
     HasFocus = False
End Sub

Private Sub txtValue_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     RaiseEvent MouseDown
End Sub

Private Sub txtValue_Validate(Cancel As Boolean)
     Dim StrippedValue As String
     On Local Error Resume Next
     
     'Check if theres anything given
     If (Replace$(Replace$(txtValue.Text, "-", ""), "+", "") <> "") Then
          
          'Check if the value should be stripped for validating
          If (left$(txtValue.Text, 2) = "--") Or (left$(txtValue.Text, 2) = "++") Then
               StrippedValue = Mid$(txtValue.Text, 2)
               ReflectChanges = False
          Else
               StrippedValue = txtValue.Text
               ReflectChanges = True
          End If
          
          'Validate with the scrollbar
          If (GetFakeValue(Val(StrippedValue)) > scrValue.Min) Then scrValue.Value = scrValue.Min
          If (GetFakeValue(Val(StrippedValue)) < scrValue.Max) Then scrValue.Value = scrValue.Max
          
          'Set the scrollbar
          scrValue.Value = GetFakeValue(Val(StrippedValue))
          
          'Reflect changes in textbox
          If ReflectChanges Then txtValue.Text = GetRealValue
     End If
     
     'Put the old value back
     If Not pEmptyAllowed Then txtValue.Text = GetRealValue
     
     'Allow changes
     ReflectChanges = True
End Sub

Private Sub UserControl_Initialize()
     ReflectChanges = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
     scrValue.Min = PropBag.ReadProperty("Max", 32767)
     txtValue.MaxLength = PropBag.ReadProperty("MaxLength", 0)
     scrValue.Max = PropBag.ReadProperty("Min", 0)
     scrValue.SmallChange = PropBag.ReadProperty("SmallChange", 1)
     txtValue.Text = PropBag.ReadProperty("Value", "0")
     pEmptyAllowed = PropBag.ReadProperty("EmptyAllowed", False)
     pRelativeAllowed = PropBag.ReadProperty("RelativeAllowed", False)
     pRelativeScroll = PropBag.ReadProperty("RelativeScroll", False)
     pUnsigned = PropBag.ReadProperty("Unsigned", False)
End Sub

Private Sub UserControl_Resize()
     On Local Error Resume Next
     txtValue.width = UserControl.ScaleWidth - scrValue.width - 1
     scrValue.left = UserControl.ScaleWidth - scrValue.width
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
     Call PropBag.WriteProperty("Max", scrValue.Min, 32767)
     Call PropBag.WriteProperty("MaxLength", txtValue.MaxLength, 0)
     Call PropBag.WriteProperty("Min", scrValue.Max, 0)
     Call PropBag.WriteProperty("SmallChange", scrValue.SmallChange, 1)
     Call PropBag.WriteProperty("Value", txtValue.Text, "0")
     Call PropBag.WriteProperty("EmptyAllowed", pEmptyAllowed, False)
     Call PropBag.WriteProperty("RelativeAllowed", pRelativeAllowed, False)
     Call PropBag.WriteProperty("RelativeScroll", pRelativeScroll, False)
     Call PropBag.WriteProperty("Unsigned", pUnsigned, False)
End Sub

Public Property Get Value() As String
Attribute Value.VB_Description = "Returns/sets the value of an object."
     Value = txtValue.Text
End Property

Public Property Let Value(ByVal New_Value As String)
     Dim realvalue As Long
     Dim fakevalue As Integer
     On Local Error Resume Next
     
     'Set the new value
     txtValue.Text = New_Value
     
'     'Limit the value
'     If (Val(txtValue.Text) > scrValue.Min) Then scrValue.Value = scrValue.Min
'     If (Val(txtValue.Text) < scrValue.Max) Then scrValue.Value = scrValue.Max
'
'     'Check if the value must be converted
'     If (pUnsigned) Then
'          realvalue = ItoL(scrValue.Value)
'     Else
'          realvalue = scrValue.Value
'     End If
'
'     'Set the scrollbar
'     scrValue.Value = realvalue
     txtValue_Validate False
     
     PropertyChanged "Value"
End Property
