#include <dinput.h>

#pragma warning( disable : 4996 )

extern "C" __declspec(dllexport) void di_Create( LPDIRECTINPUT8* ppDI )
{
	if( FAILED( DirectInput8Create( GetModuleHandle(NULL), DIRECTINPUT_VERSION, IID_IDirectInput8, (VOID**)ppDI, NULL ) ) ) *ppDI = 0;
}

extern "C" __declspec(dllexport) void di_CreateInputDeviceMouse( LPDIRECTINPUT8 pDI, LPDIRECTINPUTDEVICE8* ppDID )
{
  if( FAILED( pDI->CreateDevice( GUID_SysMouse, ppDID, NULL ) ) ) *ppDID = 0;
}

extern "C" __declspec(dllexport) void did_SetCommonDataFormatMouse( LPDIRECTINPUTDEVICE8 pDID )
{
  pDID->SetDataFormat( &c_dfDIMouse ); 
}

extern "C" __declspec(dllexport) void did_SetCooperativeLevel( LPDIRECTINPUTDEVICE8 pDID, HWND hWnd, DWORD Flags )
{
  pDID->SetCooperativeLevel( hWnd, Flags );
}

extern "C" __declspec(dllexport) void did_GetDeviceData( LPDIRECTINPUTDEVICE8 pDID, DWORD cbObjectData, LPDIDEVICEOBJECTDATA rgdod, LPDWORD pdwInOut, DWORD dwFlags )
{
  pDID->GetDeviceData(cbObjectData, rgdod, pdwInOut, dwFlags);
}

extern "C" __declspec(dllexport) void did_SetProperty( LPDIRECTINPUTDEVICE8 pDID, REFGUID rGuidProp, LPCDIPROPHEADER pDiph )
{
  pDID->SetProperty( rGuidProp, pDiph );
}

extern "C" __declspec(dllexport) void did_SetPropertyBufferSize( LPDIRECTINPUTDEVICE8 pDID, DWORD size )
{
  DIPROPDWORD prop;
  prop.diph.dwSize = sizeof(prop);
  prop.diph.dwHeaderSize = sizeof(prop.diph);
  prop.diph.dwObj = 0;
  prop.diph.dwHow = DIPH_DEVICE;
  prop.dwData = size;
  pDID->SetProperty(DIPROP_BUFFERSIZE, &prop.diph);
}

extern "C" __declspec(dllexport) void did_Acquire( LPDIRECTINPUTDEVICE8 pDID )
{
  pDID->Acquire();
}

extern "C" __declspec(dllexport) void did_Unacquire( LPDIRECTINPUTDEVICE8 pDID )
{
  pDID->Unacquire();
}

extern "C" __declspec(dllexport) void did_Release( LPDIRECTINPUTDEVICE8 pDID )
{
  if( pDID != NULL)
  {
    pDID->Release();
  }
}

extern "C" __declspec(dllexport) void di_Release( LPDIRECTINPUT8 pDI )
{
  if( pDI != NULL) 
      pDI->Release();
}
