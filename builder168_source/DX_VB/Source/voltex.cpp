#include <d3d9.h>
#pragma warning( disable : 4996 )

#define byte unsigned char

extern "C" __declspec(dllexport) void vtex_GenerateMipSubLevels( LPDIRECT3DVOLUMETEXTURE9 pTex )
{
	pTex->GenerateMipSubLevels ();
}

extern "C" __declspec(dllexport) void vtex_GetVolumeLevel( LPDIRECT3DVOLUMETEXTURE9 pTex, UINT Level, LPDIRECT3DVOLUME9* ppVol )
{
  pTex->GetVolumeLevel( Level, ppVol );
}

extern "C" __declspec(dllexport) void vtex_Release( LPDIRECT3DVOLUMETEXTURE9 pTex )
{
  if( pTex != NULL )
      pTex->Release();
}

extern "C" __declspec(dllexport) void vol_Release( LPDIRECT3DVOLUME9 pVol )
{
  if( pVol != NULL )
      pVol->Release();
}

extern "C" __declspec(dllexport) int vol_SetData( LPDIRECT3DVOLUME9 pVol, D3DBOX* pBox, UINT Size, byte* pbData )
{
	D3DLOCKED_BOX LockedVol;
	D3DVOLUME_DESC ud;

	if( FAILED( pVol->LockBox( &LockedVol, pBox, 0 ) ) ) return 0;
	pVol->GetDesc(&ud);
	UINT bpp = LockedVol.RowPitch / ud.Width;
	if( pBox == 0 )
	{
		memcpy( LockedVol.pBits, pbData, Size );
	}
	else
	{
		//UINT L=pBox->Left;
		//UINT T=pBox->Top;
		//UINT F=pBox->Front;
		UINT W=pBox->Right - pBox->Left;
		UINT H=pBox->Bottom - pBox->Top;
		UINT D=pBox->Back - pBox->Front;
		for( UINT z=0; z<D; z++ )
		{
			for( UINT y=0; y<H; y++ )
			{
				memcpy(
					((byte*)LockedVol.pBits) + z * LockedVol.SlicePitch + y * LockedVol.RowPitch,
					pbData + W * (z * D + y) * bpp,
					W * bpp);
			}
		}
	}
	pVol->UnlockBox();
  return 1;
}

extern "C" __declspec(dllexport) int vol_GetData( LPDIRECT3DVOLUME9 pVol, D3DBOX* pBox, UINT Size, byte* pbData )
{
	D3DLOCKED_BOX LockedVol;
	D3DVOLUME_DESC ud;

	if( FAILED( pVol->LockBox( &LockedVol, pBox, 0 ) ) ) return 0;
	pVol->GetDesc(&ud);
	UINT bpp = LockedVol.RowPitch / ud.Width;
	if( pBox == 0 )
	{
		memcpy( LockedVol.pBits, pbData, Size );
	}
	else
	{
		//UINT L=pBox->Left;
		//UINT T=pBox->Top;
		//UINT F=pBox->Front;
		UINT W=pBox->Right - pBox->Left;
		UINT H=pBox->Bottom - pBox->Top;
		UINT D=pBox->Back - pBox->Front;
		for( UINT z=0; z<D; z++ )
		{
			for( UINT y=0; y<H; y++ )
			{
				memcpy(
					pbData + W * (z * D + y) * bpp,
					((byte*)LockedVol.pBits) + z * LockedVol.SlicePitch + y * LockedVol.RowPitch,
					W * bpp);
			}
		}
	}
	pVol->UnlockBox();
  return 1;
}
