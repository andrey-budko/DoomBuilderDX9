#include <DSound.h>
#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void dsb_SetVolume( LPDIRECTSOUNDBUFFER pDSB, DWORD lVol )
{
	pDSB->SetVolume( (LONG)lVol );
}

extern "C" __declspec(dllexport) void dsb_SetPan( LPDIRECTSOUNDBUFFER pDSB, DWORD lPan )
{
	pDSB->SetPan( (LONG)lPan );
}

extern "C" __declspec(dllexport) void dsb_SetFrequency( LPDIRECTSOUNDBUFFER pDSB, DWORD dwFreq )
{
	pDSB->SetFrequency( dwFreq );
}

extern "C" __declspec(dllexport) void dsb_SetCurrentPosition( LPDIRECTSOUNDBUFFER pDSB, DWORD dwNewPosition )
{
	pDSB->SetCurrentPosition( dwNewPosition );
}

extern "C" __declspec(dllexport) int dsb_SetData( LPDIRECTSOUNDBUFFER pDSB, UINT Offset, UINT Size, void* pbData )
{
	LPVOID lpvWrite;
	DWORD  dwLength;

	if (FAILED (pDSB->Lock( Offset,	Size,	&lpvWrite, &dwLength,	NULL,	NULL,	0 ))) return 0;
	memcpy(lpvWrite, pbData, dwLength);
	pDSB->Unlock(	lpvWrite,	dwLength,	NULL,	0 );
	return 1;
}

extern "C" __declspec(dllexport) int dsb_GetData( LPDIRECTSOUNDBUFFER pDSB, UINT Offset, UINT Size, void* pbData )
{
	LPVOID lpvRead;
	DWORD  dwLength;

	if (FAILED (pDSB->Lock( Offset,	Size,	&lpvRead, &dwLength,	NULL,	NULL,	0 ))) return 0;
	memcpy(lpvRead, pbData, dwLength);
	pDSB->Unlock(	lpvRead,	dwLength,	NULL,	0 );
	return 1;
}

extern "C" __declspec(dllexport) void dsb_Play( LPDIRECTSOUNDBUFFER pDSB, DWORD dwFlags )
{
	pDSB->Play( 0, 0, dwFlags );
}

extern "C" __declspec(dllexport) void dsb_Stop( LPDIRECTSOUNDBUFFER pDSB )
{
	pDSB->Stop();
}

extern "C" __declspec(dllexport) void ds_SetCooperativeLevel( LPDIRECTSOUND8 pDS, HWND hwnd, DWORD dwLevel )
{
	pDS->SetCooperativeLevel( hwnd, dwLevel );
}

extern "C" __declspec(dllexport) void ds_CreateSoundBuffer( LPDIRECTSOUND8 pDS, int nChannels, int BitsPerSample, int Freq, int Size, int Flags, LPDIRECTSOUNDBUFFER* ppDSB )
{
  WAVEFORMATEX WF;
  DSBUFFERDESC BD;
  WF.wFormatTag = WAVE_FORMAT_PCM;
  WF.nChannels = nChannels;
  WF.nSamplesPerSec = Freq;
  WF.wBitsPerSample = BitsPerSample;
  WF.nBlockAlign = WF.wBitsPerSample * WF.nChannels / 8;
  WF.nAvgBytesPerSec = WF.nSamplesPerSec * WF.nBlockAlign;
  WF.cbSize = 0;
  BD.dwSize = sizeof(BD);
  BD.dwFlags = Flags;
  BD.dwBufferBytes = Size;
  BD.dwReserved = 0;
  BD.lpwfxFormat = &WF;

	if( FAILED( pDS->CreateSoundBuffer( &BD, ppDSB, NULL ))) *ppDSB = 0;
}

extern "C" __declspec(dllexport) void ds_DuplicateSoundBuffer( LPDIRECTSOUND8 pDS, LPDIRECTSOUNDBUFFER pDSB, LPDIRECTSOUNDBUFFER* ppDSB )
{
	if( FAILED( pDS->DuplicateSoundBuffer( pDSB, ppDSB ))) *ppDSB = 0;
}

extern "C" __declspec(dllexport) void dsb_Release( LPDIRECTSOUNDBUFFER pDSB )
{
  if( pDSB != NULL)
      pDSB->Release();
}

extern "C" __declspec(dllexport) void ds_Create( LPDIRECTSOUND8* ppDS )
{
	if( FAILED( DirectSoundCreate8( NULL, ppDS, NULL ))) *ppDS = 0;
}

extern "C" __declspec(dllexport) void ds_Release( LPDIRECTSOUND8 pDS )
{
  if( pDS != NULL)
      pDS->Release();
}
