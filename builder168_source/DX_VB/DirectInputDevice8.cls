VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DirectInputDevice8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub did_SetCooperativeLevel Lib "dx_vb" (ByVal pDID As Long, ByVal hWnd As Long, ByVal Flags As CONST_DISCL)
Private Declare Sub did_SetCommonDataFormatMouse Lib "dx_vb" (ByVal pDID As Long)
Private Declare Sub did_GetDeviceData Lib "dx_vb" (ByVal pDID As Long, ByVal cbObjectData As Long, ByRef rgDOD As DIDEVICEOBJECTDATA, ByRef dwInOut As Long, ByVal dwFlags As Long)
Private Declare Sub did_SetProperty Lib "dx_vb" (ByVal pDID As Long, ByRef GuidProp As GUID, ByRef pDiph As DIPROPHEADER)
Private Declare Sub did_SetPropertyBufferSize Lib "dx_vb" (ByVal pDID As Long, ByVal Size As Long)
Private Declare Sub did_Acquire Lib "dx_vb" (ByVal pDID As Long)
Private Declare Sub did_Unacquire Lib "dx_vb" (ByVal pDID As Long)
Private Declare Sub did_Release Lib "dx_vb" (ByVal pDID As Long)

Private mePtr As Long

Public Sub SetCooperativeLevel(ByVal hWnd As Long, ByVal Flags As CONST_DISCL)
  did_SetCooperativeLevel Ptr, hWnd, Flags
End Sub

Friend Sub SetCommonDataFormatMouse()
  did_SetCommonDataFormatMouse Ptr
End Sub

Friend Sub GetDeviceData(ByVal cbObjectData As Long, ByRef rgDOD As DIDEVICEOBJECTDATA, ByRef dwInOut As Long, ByVal dwFlags As Long)
  did_GetDeviceData Ptr, cbObjectData, rgDOD, dwInOut, dwFlags
End Sub

Friend Sub SetProperty(ByRef GuidProp As GUID, ByRef pDiph As DIPROPHEADER)
  did_SetProperty Ptr, GuidProp, pDiph
End Sub

Friend Sub SetPropertyBufferSize(ByVal Size As Long)
  did_SetPropertyBufferSize Ptr, Size
End Sub

Public Sub Acquire()
  did_Acquire (Ptr)
End Sub

Public Sub Unacquire()
  did_Unacquire Ptr
End Sub

Private Sub Class_Terminate()
  If Ptr <> 0 Then
    did_Release Ptr
  End If
End Sub

Public Property Let Ptr(p As Long)
  If mePtr <> 0 Then did_Release mePtr
  mePtr = p
End Property

Public Property Get Ptr() As Long
  Ptr = mePtr
End Property
