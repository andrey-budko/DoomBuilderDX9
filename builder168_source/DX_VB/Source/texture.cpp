#include <d3d9.h>
#pragma warning( disable : 4996 )

#define byte unsigned char

extern "C" __declspec(dllexport) void surf_GetDesc( LPDIRECT3DSURFACE9 pSurface, D3DSURFACE_DESC* pDesc )
{
  pSurface->GetDesc( pDesc ); 
}

extern "C" __declspec(dllexport) void tex_GetLevelDesc( LPDIRECT3DTEXTURE9 pTex, UINT Level, D3DSURFACE_DESC* pDesc )
{
  pTex->GetLevelDesc( Level, pDesc ); 
}

extern "C" __declspec(dllexport) void tex_GenerateMipSubLevels( LPDIRECT3DTEXTURE9 pTex )
{
	pTex->GenerateMipSubLevels();
}

extern "C" __declspec(dllexport) void tex_GetSurfaceLevel( LPDIRECT3DTEXTURE9 pTex, UINT Level, LPDIRECT3DSURFACE9* ppSurface )
{
  pTex->GetSurfaceLevel( Level, ppSurface );
}

extern "C" __declspec(dllexport) void tex_Release( LPDIRECT3DTEXTURE9 pTex )
{
  if( pTex != NULL )
      pTex->Release();
}

extern "C" __declspec(dllexport) void surf_Release( LPDIRECT3DSURFACE9 pSurface )
{
  if( pSurface != NULL )
      pSurface->Release();
}

extern "C" __declspec(dllexport) int surf_SetData( LPDIRECT3DSURFACE9 pSurface, RECT* pRect, UINT Size, byte* pbData )
{
	D3DLOCKED_RECT LockedRect;
	D3DSURFACE_DESC ud;

	if( FAILED( pSurface->LockRect( &LockedRect, pRect, 0 ) ) ) return 0;
	pSurface->GetDesc(&ud);
	UINT bpp = LockedRect.Pitch / ud.Width;
	if( pRect == 0)
	{
		memcpy( LockedRect.pBits, pbData, Size );
	}
	else
	{
//		UINT L=pRect->left;
//		UINT T=pRect->top;
		UINT W=pRect->right - pRect->left;
		UINT H=pRect->bottom - pRect->top;
		for( UINT y=0; y<H; y++ )
		{
			//memcpy( ((byte*)LockedRect.pBits) + (T + y) * LockedRect.Pitch + L * bpp, pbData + W * y * bpp, W * bpp);
			memcpy( ((byte*)LockedRect.pBits) + y * LockedRect.Pitch, pbData + W * y * bpp, W * bpp );
		}
	}
	pSurface->UnlockRect();
  return 1;
}

extern "C" __declspec(dllexport) int surf_GetData( LPDIRECT3DSURFACE9 pSurface, RECT* pRect, UINT Size, byte* pbData )
{
	D3DLOCKED_RECT LockedRect;
	D3DSURFACE_DESC ud;

	if( FAILED( pSurface->LockRect( &LockedRect, pRect, 0 ) ) ) return 0;
	pSurface->GetDesc(&ud);
	UINT bpp = LockedRect.Pitch / ud.Width;
	if( pRect == 0)
	{
		memcpy( pbData, LockedRect.pBits, Size );
	}
	else
	{
//		UINT L=pRect->left;
//		UINT T=pRect->top;
		UINT W=pRect->right - pRect->left;
		UINT H=pRect->bottom - pRect->top;
		for( UINT y=0; y<H; y++ )
		{
			//memcpy( pbData + W * y * bpp, ((byte*)LockedRect.pBits) + (T + y) * LockedRect.Pitch + L * bpp, W * bpp );
			memcpy( pbData + W * y * bpp, ((byte*)LockedRect.pBits) + y * LockedRect.Pitch, W * bpp );
		}
	}
	pSurface->UnlockRect();
  return 1;
}

extern "C" __declspec(dllexport) int surf_GetPixel( LPDIRECT3DSURFACE9 pSurface, UINT x, UINT y, UINT Size, byte* pbData )
{
	D3DLOCKED_RECT LockedRect;
	D3DSURFACE_DESC ud;
  RECT R;

  R.left = x;
  R.top = y;
  R.right = x + 1;
  R.bottom = y + 1;
	if( FAILED( pSurface->LockRect( &LockedRect, &R, 0 ) ) ) return 0;
	pSurface->GetDesc(&ud);
	memcpy( pbData, LockedRect.pBits, Size );
	pSurface->UnlockRect();
  return 1;
}
