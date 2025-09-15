#include <d3d9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void dev_SetGammaRamp( LPDIRECT3DDEVICE9 pDev, UINT iSwapChain, DWORD Flags, D3DGAMMARAMP* pRamp )
{
  pDev->SetGammaRamp( iSwapChain, Flags, pRamp );
}

extern "C" __declspec(dllexport) HRESULT dev_TestCooperativeLevel( LPDIRECT3DDEVICE9 pDev )
{
  return pDev->TestCooperativeLevel();
}

extern "C" __declspec(dllexport) void dev_Reset( LPDIRECT3DDEVICE9 pDev, D3DPRESENT_PARAMETERS* pPP )
{
  pDev->Reset( pPP );
}

extern "C" __declspec(dllexport) void dev_SetClipPlane( LPDIRECT3DDEVICE9 pDev, DWORD Index, float* pPlane )
{
	pDev->SetClipPlane( Index, pPlane );
}

extern "C" __declspec(dllexport) void dev_CreateVolumeTexture( LPDIRECT3DDEVICE9 pDev, UINT Width, UINT Height, UINT Depth, UINT Levels, 
																 DWORD Usage, D3DFORMAT Format, D3DPOOL Pool, LPDIRECT3DVOLUMETEXTURE9* ppTex )
{
	if( FAILED( pDev->CreateVolumeTexture( Width, Height, Depth, Levels, Usage, Format, Pool, ppTex, NULL ) ) ) *ppTex = 0;
}

extern "C" __declspec(dllexport) void dev_SetPixelShaderConstantB( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, BOOL* pData, UINT BoolCount )
{
	pDev->SetPixelShaderConstantB( StartRegister, pData, BoolCount );
}

extern "C" __declspec(dllexport) void dev_SetPixelShaderConstantF( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, float* pData, UINT Vec4fCount )
{
	pDev->SetPixelShaderConstantF( StartRegister, pData, Vec4fCount );
}

extern "C" __declspec(dllexport) void dev_SetPixelShaderConstantI( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, int* pData, UINT Vec4Count )
{
	pDev->SetPixelShaderConstantI( StartRegister, pData, Vec4Count );
}

extern "C" __declspec(dllexport) void dev_SetVertexShaderConstantB( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, BOOL* pData, UINT BoolCount )
{
	pDev->SetVertexShaderConstantB( StartRegister, pData, BoolCount );
}

extern "C" __declspec(dllexport) void dev_SetVertexShaderConstantF( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, float* pData, UINT Vec4fCount )
{
	pDev->SetVertexShaderConstantF( StartRegister, pData, Vec4fCount );
}

extern "C" __declspec(dllexport) void dev_SetVertexShaderConstantI( LPDIRECT3DDEVICE9 pDev, UINT StartRegister, int* pData, UINT Vec4Count )
{
	pDev->SetVertexShaderConstantI( StartRegister, pData, Vec4Count );
}

extern "C" __declspec(dllexport) void dev_SetVertexDeclaration( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DVERTEXDECLARATION9 pDecl )
{
	pDev->SetVertexDeclaration( pDecl );
}

extern "C" __declspec(dllexport) void dev_CreateVertexDeclaration( LPDIRECT3DDEVICE9 pDev, D3DVERTEXELEMENT9* pVertexElement, LPDIRECT3DVERTEXDECLARATION9* ppDecl )
{
	if( FAILED( pDev->CreateVertexDeclaration( pVertexElement, ppDecl ) ) ) *ppDecl = 0;
}

extern "C" __declspec(dllexport) void dev_SetVertexShader( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DVERTEXSHADER9 pShader )
{
	pDev->SetVertexShader( pShader );
}

extern "C" __declspec(dllexport) void dev_CreateVertexShader( LPDIRECT3DDEVICE9 pDev, DWORD* pFunction, LPDIRECT3DVERTEXSHADER9* ppShader )
{
	if( FAILED( pDev->CreateVertexShader( pFunction, ppShader ) ) ) *ppShader = 0;
}

extern "C" __declspec(dllexport) void dev_SetPixelShader( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DPIXELSHADER9 pShader )
{
	pDev->SetPixelShader( pShader );
}

extern "C" __declspec(dllexport) void dev_CreatePixelShader( LPDIRECT3DDEVICE9 pDev, DWORD* pFunction, LPDIRECT3DPIXELSHADER9* ppShader )
{
	if( FAILED( pDev->CreatePixelShader( pFunction, ppShader ) ) ) *ppShader = 0;
}

extern "C" __declspec(dllexport) void dev_GetRenderTarget( LPDIRECT3DDEVICE9 pDev, DWORD RenderTargetIndex, LPDIRECT3DSURFACE9* ppSurface )
{
	pDev->GetRenderTarget( RenderTargetIndex, ppSurface );
}

extern "C" __declspec(dllexport) void dev_SetRenderTarget( LPDIRECT3DDEVICE9 pDev, DWORD RenderTargetIndex, LPDIRECT3DSURFACE9 pSurface )
{
	pDev->SetRenderTarget( RenderTargetIndex, pSurface );
}

extern "C" __declspec(dllexport) void dev_GetDepthStencilSurface( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DSURFACE9* ppSurface )
{
	pDev->GetDepthStencilSurface( ppSurface );
}

extern "C" __declspec(dllexport) void dev_SetDepthStencilSurface( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DSURFACE9 pSurface )
{
	pDev->SetDepthStencilSurface( pSurface );
}

extern "C" __declspec(dllexport) void dev_CreateDepthStencilSurface( LPDIRECT3DDEVICE9 pDev, UINT Width, UINT Height, 
																 D3DFORMAT Format, D3DMULTISAMPLE_TYPE MultiSample, DWORD MultisampleQuality, BOOL Discard,
																 LPDIRECT3DSURFACE9* ppSurface )
{
	if( FAILED( pDev->CreateDepthStencilSurface( Width, Height, Format, MultiSample, MultisampleQuality, Discard, ppSurface, NULL ) ) ) *ppSurface = 0;
}

extern "C" __declspec(dllexport) void dev_CreateTexture( LPDIRECT3DDEVICE9 pDev, UINT Width, UINT Height, UINT Levels, 
																 DWORD Usage, D3DFORMAT Format, D3DPOOL Pool, LPDIRECT3DTEXTURE9* ppTex )
{
	if( FAILED( pDev->CreateTexture( Width, Height, Levels, Usage, Format, Pool, ppTex, NULL ) ) ) *ppTex = 0;
}

extern "C" __declspec(dllexport) void dev_DrawIndexedPrimitive( LPDIRECT3DDEVICE9 pDev, D3DPRIMITIVETYPE PrimitiveType,
																 INT BaseVertexIndex, UINT MinVertexIndex, UINT NumVertices, UINT StartIndex, UINT PrimitiveCount )
{
	pDev->DrawIndexedPrimitive( PrimitiveType, BaseVertexIndex, MinVertexIndex,NumVertices, StartIndex, PrimitiveCount );
}

extern "C" __declspec(dllexport) void dev_DrawIndexedPrimitiveUp( LPDIRECT3DDEVICE9 pDev, D3DPRIMITIVETYPE PrimitiveType,
																 UINT MinVertexIndex, UINT NumVertices, UINT PrimitiveCount, void* pIndexData, D3DFORMAT IndexDataFormat,
																 void* pVertexStreamZeroData, UINT VertexStreamZeroStride )
{
	pDev->DrawIndexedPrimitiveUP( PrimitiveType, MinVertexIndex, NumVertices, PrimitiveCount, pIndexData, IndexDataFormat, pVertexStreamZeroData, VertexStreamZeroStride );
}

extern "C" __declspec(dllexport) void dev_DrawPrimitiveUp( LPDIRECT3DDEVICE9 pDev, D3DPRIMITIVETYPE PrimitiveType,
																 UINT PrimitiveCount, void* pVertexStreamZeroData, UINT VertexStreamZeroStride )
{
	pDev->DrawPrimitiveUP( PrimitiveType, PrimitiveCount, pVertexStreamZeroData, VertexStreamZeroStride );
}

extern "C" __declspec(dllexport) void dev_SetIndices( LPDIRECT3DDEVICE9 pDev, LPDIRECT3DINDEXBUFFER9 g_pIB )
{
	pDev->SetIndices( g_pIB );
}

extern "C" __declspec(dllexport) void dev_CreateIndexBuffer( LPDIRECT3DDEVICE9 pDev, UINT Length, 
																 DWORD Usage, D3DFORMAT Format, D3DPOOL Pool, LPDIRECT3DINDEXBUFFER9* g_pIB )
{
	if( FAILED( pDev->CreateIndexBuffer( Length, Usage, Format, Pool, g_pIB, NULL ) ) ) *g_pIB = 0;
}

extern "C" __declspec(dllexport) void dev_LightEnable( LPDIRECT3DDEVICE9 pDev, DWORD LightIndex, BOOL bEnable )
{
	pDev->LightEnable( LightIndex, bEnable );
}

extern "C" __declspec(dllexport) void dev_SetLight( LPDIRECT3DDEVICE9 pDev, DWORD Index, D3DLIGHT9* pLight )
{
	pDev->SetLight( Index, pLight );
}

extern "C" __declspec(dllexport) void dev_SetMaterial( LPDIRECT3DDEVICE9 pDev, D3DMATERIAL9* pMtrl )
{
	pDev->SetMaterial( pMtrl );
}

extern "C" __declspec(dllexport) void dev_SetSamplerState( LPDIRECT3DDEVICE9 pDev, DWORD Sampler, D3DSAMPLERSTATETYPE Type, DWORD Value)
{
  pDev->SetSamplerState( Sampler, Type, Value );
}

extern "C" __declspec(dllexport) void dev_GetSamplerState( LPDIRECT3DDEVICE9 pDev, DWORD Sampler, D3DSAMPLERSTATETYPE Type, DWORD* pValue)
{
  pDev->GetSamplerState( Sampler, Type, pValue );
}

extern "C" __declspec(dllexport) void dev_SetTextureStageState( LPDIRECT3DDEVICE9 pDev, DWORD Stage, D3DTEXTURESTAGESTATETYPE Type, DWORD Value)
{
  pDev->SetTextureStageState( Stage, Type, Value );
}

extern "C" __declspec(dllexport) void dev_GetTextureStageState( LPDIRECT3DDEVICE9 pDev, DWORD Stage, D3DTEXTURESTAGESTATETYPE Type, DWORD* pValue)
{
  pDev->GetTextureStageState( Stage, Type, pValue );
}

extern "C" __declspec(dllexport) void dev_MultiplyTransform( LPDIRECT3DDEVICE9 pDev, D3DTRANSFORMSTATETYPE State, D3DMATRIX* pMatrix )
{
	pDev->MultiplyTransform( State, pMatrix );
}

extern "C" __declspec(dllexport) void dev_GetTransform( LPDIRECT3DDEVICE9 pDev, D3DTRANSFORMSTATETYPE State, D3DMATRIX* pMatrix )
{
	pDev->GetTransform( State, pMatrix );
}

extern "C" __declspec(dllexport) void dev_SetTransform( LPDIRECT3DDEVICE9 pDev, D3DTRANSFORMSTATETYPE State, D3DMATRIX* pMatrix )
{
	pDev->SetTransform( State, pMatrix );
}

extern "C" __declspec(dllexport) void dev_CreateVertexBuffer( LPDIRECT3DDEVICE9 pDev, UINT Size,
															   DWORD Usage, DWORD FVF, D3DPOOL Pool, LPDIRECT3DVERTEXBUFFER9* g_pVB )
{
	if( FAILED( pDev->CreateVertexBuffer( Size, Usage, FVF, Pool, g_pVB, NULL ) ) ) *g_pVB = 0;
}

extern "C" __declspec(dllexport) void dev_SetStreamSource( LPDIRECT3DDEVICE9 pDev, UINT StreamNumber,
																			LPDIRECT3DVERTEXBUFFER9 g_pVB, UINT Offset, UINT Stride )
{
	pDev->SetStreamSource( StreamNumber, g_pVB, Offset, Stride );
}

extern "C" __declspec(dllexport) void dev_SetFVF( LPDIRECT3DDEVICE9 pDev, DWORD FVF )
{
  pDev->SetFVF( FVF );
}

extern "C" __declspec(dllexport) void dev_DrawPrimitive( LPDIRECT3DDEVICE9 pDev, D3DPRIMITIVETYPE PrimitiveType,
																			UINT StartVertex, UINT PrimitiveCount )
{
	pDev->DrawPrimitive( PrimitiveType, StartVertex, PrimitiveCount );
}

extern "C" __declspec(dllexport) void dev_SetTexture( LPDIRECT3DDEVICE9 pDev, DWORD Stage, LPDIRECT3DTEXTURE9 pTex )
{
  pDev->SetTexture( Stage, pTex );
}

extern "C" __declspec(dllexport) void dev_SetRenderState( LPDIRECT3DDEVICE9 pDev, D3DRENDERSTATETYPE State, DWORD Value )
{
  pDev->SetRenderState( State, Value );
}

extern "C" __declspec(dllexport) void dev_GetRenderState( LPDIRECT3DDEVICE9 pDev, D3DRENDERSTATETYPE State, DWORD* pValue )
{
  pDev->GetRenderState ( State, pValue );
}

extern "C" __declspec(dllexport) void dev_Clear( LPDIRECT3DDEVICE9 pDev, DWORD Flags,	D3DCOLOR Color, float Z, DWORD Stencil )
{
	pDev->Clear( 0, NULL, Flags, Color, Z, Stencil );
}

extern "C" __declspec(dllexport) int dev_BeginScene( LPDIRECT3DDEVICE9 pDev )
{
	if( SUCCEEDED( pDev->BeginScene() ) ) return 1;
	return 0;
}

extern "C" __declspec(dllexport) void dev_EndScene( LPDIRECT3DDEVICE9 pDev )
{
	pDev->EndScene();
}

extern "C" __declspec(dllexport) void dev_Present( LPDIRECT3DDEVICE9 pDev )
{
  pDev->Present( NULL, NULL, NULL, NULL );
}

extern "C" __declspec(dllexport) void dev_Release( LPDIRECT3DDEVICE9 pDev )
{
  if( pDev != NULL) 
      pDev->Release();
}
