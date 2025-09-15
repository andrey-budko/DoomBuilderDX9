VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Direct3D9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub d3d_CreateDevice Lib "dx_vb" (ByVal pD3d As Long, ByVal hWnd As Long, ByVal Flags As CONST_D3DCREATEFLAGS, ByRef pp As D3DPRESENT_PARAMETERS, ByRef pDev As Long)
Private Declare Sub d3d_Release Lib "dx_vb" (ByVal pD3d As Long)
Private Declare Sub d3d_GetDeviceCaps Lib "dx_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByRef Caps As D3DCAPS9)
Private Declare Function d3d_GetAdapterCount Lib "dx_vb" (ByVal pD3d As Long) As Long
Private Declare Sub d3d_GetAdapterIdentifier Lib "dx_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal Flags As Long, ByRef pIdentifier As D3DADAPTER_IDENTIFIER9)
Private Declare Sub d3d_EnumAdapterModes Lib "dx_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal Format As D3DFORMAT, ByVal mode As Long, ByRef pMode As D3DDISPLAYMODE)
Private Declare Function d3d_GetAdapterModeCount Lib "dx_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal Format As D3DFORMAT) As Long
Private Declare Function d3d_CheckDeviceFormat Lib "dx_vb" (ByVal pD3d As Long, ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByVal AdapterFormat As D3DFORMAT, ByVal Usage As Long, ByVal RType As D3DRESOURCETYPE, ByVal CheckFormat As D3DFORMAT) As Long

Private mePtr As Long

Friend Sub GetDeviceCaps(ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByRef Caps As D3DCAPS9)
  d3d_GetDeviceCaps Ptr, Adapter, DeviceType, Caps
End Sub

Friend Function GetAdapterCount()
  GetAdapterCount = d3d_GetAdapterCount(Ptr)
End Function

Friend Function GetAdapterIdentifier(ByVal Adapter As Long, ByVal Flags As Long, ByRef pIdentifier As D3DADAPTER_IDENTIFIER9)
  d3d_GetAdapterIdentifier Ptr, Adapter, Flags, pIdentifier
End Function

Friend Function EnumAdapterModes(ByVal Adapter As Long, ByVal Format As D3DFORMAT, ByVal mode As Long, ByRef pMode As D3DDISPLAYMODE)
  d3d_EnumAdapterModes Ptr, Adapter, Format, mode, pMode
End Function

Friend Function GetAdapterModeCount(ByVal Adapter As Long, ByVal Format As D3DFORMAT)
  GetAdapterModeCount = d3d_GetAdapterModeCount(Ptr, Adapter, Format)
End Function

Friend Function CheckDeviceFormat(ByVal Adapter As Long, ByVal DeviceType As D3DDEVTYPE, ByVal AdapterFormat As D3DFORMAT, ByVal Usage As Long, ByVal RType As D3DRESOURCETYPE, ByVal CheckFormat As D3DFORMAT)
  CheckDeviceFormat = d3d_CheckDeviceFormat(Ptr, Adapter, DeviceType, AdapterFormat, Usage, RType, CheckFormat)
End Function

Friend Function CreateDevice(ByVal hWnd As Long, ByVal Flags As CONST_D3DCREATEFLAGS, ByRef d3dpp As D3DPRESENT_PARAMETERS) As Direct3DDevice9
Dim pDev As Long
  d3d_CreateDevice Ptr, hWnd, Flags, d3dpp, pDev
  If pDev <> 0 Then
    Set CreateDevice = New Direct3DDevice9
    CreateDevice.Ptr = pDev
  End If
End Function

Public Property Let Ptr(p As Long)
  If mePtr <> 0 Then d3d_Release mePtr
  mePtr = p
End Property

Public Property Get Ptr() As Long
  Ptr = mePtr
End Property

Private Sub Class_Terminate()
  If Ptr <> 0 Then
    d3d_Release Ptr
  End If
End Sub
