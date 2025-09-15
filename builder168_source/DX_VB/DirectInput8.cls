VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectInput8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub di_Release Lib "dx_vb" (ByVal pDI As Long)
Private Declare Sub di_CreateInputDeviceMouse Lib "dx_vb" (ByVal pDI As Long, ByRef pDID As Long)

Private mePtr As Long

Friend Function CreateDeviceMouse() As DirectInputDevice8
  Dim pDID As Long

  di_CreateInputDeviceMouse Ptr, pDID
  If pDID <> 0 Then
    Set CreateDeviceMouse = New DirectInputDevice8
    CreateDeviceMouse.Ptr = pDID
  End If
End Function

Private Sub Class_Terminate()
  If Ptr <> 0 Then
    di_Release Ptr
  End If
End Sub

Public Property Let Ptr(p As Long)
  If mePtr <> 0 Then di_Release mePtr
  mePtr = p
End Property

Public Property Get Ptr() As Long
  Ptr = mePtr
End Property
