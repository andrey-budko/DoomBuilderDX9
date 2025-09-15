#include <d3d9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) int ib_GetData( LPDIRECT3DINDEXBUFFER9 g_pIB, UINT Offset, UINT Size, void* pbData, DWORD Flags )
{
  void* pIndices;
	if( FAILED( g_pIB->Lock( Offset, Size, (void**)&pIndices, Flags ) ) ) return 0;
  memcpy( pbData, pIndices, Size );
  g_pIB->Unlock();
  return 1;
}

extern "C" __declspec(dllexport) void ib_Release( LPDIRECT3DINDEXBUFFER9 g_pIB )
{
  if( g_pIB != NULL )
      g_pIB->Release();
}

extern "C" __declspec(dllexport) int ib_SetData( LPDIRECT3DINDEXBUFFER9 g_pIB, UINT Offset, UINT Size, void* pbData, DWORD Flags )
{
  void* pIndices;
	if( FAILED( g_pIB->Lock( Offset, Size, (void**)&pIndices, Flags ) ) ) return 0;
  memcpy( pIndices, pbData, Size );
  g_pIB->Unlock();
  return 1;
}
