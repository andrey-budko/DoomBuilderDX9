#include <d3dx9.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void D3DX_CreateTextureFromFile( LPDIRECT3DDEVICE9 pDev, LPCWSTR pFile, LPDIRECT3DTEXTURE9* ppTex )
{
  D3DXCreateTextureFromFile( pDev, pFile, ppTex ); 
}

extern "C" __declspec(dllexport) void D3DX_CreateTextureFromFileEx( LPDIRECT3DDEVICE9 pDev, LPCWSTR pFile, UINT Width, UINT Height,
                                                                    UINT MipLevels, DWORD Usage, D3DFORMAT Format, D3DPOOL Pool,
                                                                    DWORD Filter, DWORD MipFilter, D3DCOLOR ColorKey, D3DXIMAGE_INFO *pSrcInfo,
                                                                    PALETTEENTRY *pPalette, LPDIRECT3DTEXTURE9* ppTex )
{
  D3DXCreateTextureFromFileEx( pDev, pFile, Width, Height, MipLevels, Usage, Format, Pool, Filter, MipFilter, ColorKey, pSrcInfo, pPalette, ppTex );
}

extern "C" __declspec(dllexport) void D3DX_GetImageInfoFromFile( LPCWSTR pFile, D3DXIMAGE_INFO* pInfo )
{
  D3DXGetImageInfoFromFile( pFile, pInfo );
}

extern "C" __declspec(dllexport) void D3DX_CreateTextureFromFileInMemory( LPDIRECT3DDEVICE9 pDev, LPCVOID pData, UINT DataSize, LPDIRECT3DTEXTURE9* ppTex )
{
  D3DXCreateTextureFromFileInMemory ( pDev, pData, DataSize, ppTex ); 
}

extern "C" __declspec(dllexport) void D3DX_CreateTextureFromFileInMemoryEx( LPDIRECT3DDEVICE9 pDev, LPCVOID pData, UINT DataSize, UINT Width, UINT Height,
                                                                    UINT MipLevels, DWORD Usage, D3DFORMAT Format, D3DPOOL Pool,
                                                                    DWORD Filter, DWORD MipFilter, D3DCOLOR ColorKey, D3DXIMAGE_INFO *pSrcInfo,
                                                                    PALETTEENTRY *pPalette, LPDIRECT3DTEXTURE9* ppTex )
{
  D3DXCreateTextureFromFileInMemoryEx( pDev, pData, DataSize, Width, Height, MipLevels, Usage, Format, Pool, Filter, MipFilter, ColorKey, pSrcInfo, pPalette, ppTex );
}
