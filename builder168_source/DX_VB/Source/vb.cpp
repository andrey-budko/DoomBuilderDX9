#include <d3d9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) int vb_GetData( LPDIRECT3DVERTEXBUFFER9 g_pVB, UINT Offset, UINT Size, void* pbData, DWORD Flags )
{
  void* pVertices;
	if( FAILED( g_pVB->Lock( Offset, Size, (void**)&pVertices, Flags ) ) ) return 0;
  memcpy( pbData, pVertices, Size );
  g_pVB->Unlock();
  return 1;
}

extern "C" __declspec(dllexport) int vb_SetData( LPDIRECT3DVERTEXBUFFER9 g_pVB, UINT Offset, UINT Size, void* pbData, DWORD Flags )
{
  void* pVertices;
	if( FAILED( g_pVB->Lock( Offset, Size, (void**)&pVertices, Flags ) ) ) return 0;
  memcpy( pVertices, pbData, Size );
  g_pVB->Unlock();
  return 1;
}

extern "C" __declspec(dllexport) void vb_Release( LPDIRECT3DVERTEXBUFFER9 g_pVB )
{
  if( g_pVB != NULL )
      g_pVB->Release();
}
