#include <d3d9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void d3d_Create( LPDIRECT3D9* g_pD3D )
{
	*g_pD3D = Direct3DCreate9( D3D_SDK_VERSION );
}

extern "C" __declspec(dllexport) void d3d_Release( LPDIRECT3D9 g_pD3D )
{
  if( g_pD3D != NULL)
      g_pD3D->Release();
}

extern "C" __declspec(dllexport) void d3d_CreateDevice( LPDIRECT3D9 g_pD3D, HWND hWnd, DWORD Flags, D3DPRESENT_PARAMETERS* d3dpp, LPDIRECT3DDEVICE9* pDev )
{
  if( FAILED( g_pD3D->CreateDevice( D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, Flags, d3dpp, pDev ) ) ) pDev = 0;
}

extern "C" __declspec(dllexport) void d3d_GetDeviceCaps( LPDIRECT3D9 g_pD3D, UINT Adapter, D3DDEVTYPE DeviceType, D3DCAPS9* pCaps)
{
	g_pD3D->GetDeviceCaps (Adapter, DeviceType, pCaps);
}

extern "C" __declspec(dllexport) UINT d3d_GetAdapterCount( LPDIRECT3D9 g_pD3D )
{
  return g_pD3D->GetAdapterCount ();
}

extern "C" __declspec(dllexport) void d3d_GetAdapterIdentifier( LPDIRECT3D9 g_pD3D, UINT Adapter, DWORD Flags, D3DADAPTER_IDENTIFIER9 *pIdentifier )
{
  g_pD3D->GetAdapterIdentifier (Adapter, Flags, pIdentifier);
}

extern "C" __declspec(dllexport) void d3d_EnumAdapterModes( LPDIRECT3D9 g_pD3D, UINT Adapter, D3DFORMAT Format, UINT Mode, D3DDISPLAYMODE *pMode )
{
  g_pD3D->EnumAdapterModes(Adapter, Format, Mode, pMode);
}

extern "C" __declspec(dllexport) UINT d3d_GetAdapterModeCount( LPDIRECT3D9 g_pD3D, UINT Adapter, D3DFORMAT Format )
{
  return g_pD3D->GetAdapterModeCount(Adapter, Format);
}

extern "C" __declspec(dllexport) HRESULT d3d_CheckDeviceFormat( LPDIRECT3D9 g_pD3D, UINT Adapter, D3DDEVTYPE DeviceType, D3DFORMAT AdapterFormat, DWORD Usage, D3DRESOURCETYPE RType, D3DFORMAT CheckFormat )
{
  return g_pD3D->CheckDeviceFormat(Adapter, DeviceType, AdapterFormat, Usage, RType, CheckFormat);
}
