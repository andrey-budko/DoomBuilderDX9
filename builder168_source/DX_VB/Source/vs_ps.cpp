#include <d3d9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void decl_Release( LPDIRECT3DVERTEXDECLARATION9 pDecl )
{
  if( pDecl != NULL )
      pDecl->Release();
}

extern "C" __declspec(dllexport) void ps_Release( LPDIRECT3DPIXELSHADER9 pShader )
{
  if( pShader != NULL )
      pShader->Release();
}

extern "C" __declspec(dllexport) void vs_Release( LPDIRECT3DVERTEXSHADER9 pShader )
{
  if( pShader != NULL )
      pShader->Release();
}
