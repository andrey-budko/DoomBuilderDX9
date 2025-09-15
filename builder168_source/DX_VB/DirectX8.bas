Attribute VB_Name = "DirectX8"
Option Explicit

' CONST ===========================================================

Public Const DS_TRUE As Long = 1
Public Const DS_FALSE As Long = 0

Public Enum CONST_DSBCAPS
  DSBCAPS_PRIMARYBUFFER = &H1
  DSBCAPS_STATIC = &H2
  DSBCAPS_LOCHARDWARE = &H4
  DSBCAPS_LOCSOFTWARE = &H8
  DSBCAPS_CTRL3D = &H10
  DSBCAPS_CTRLFREQUENCY = &H20
  DSBCAPS_CTRLPAN = &H40
  DSBCAPS_CTRLVOLUME = &H80
  DSBCAPS_CTRLPOSITIONNOTIFY = &H100
  DSBCAPS_CTRLFX = &H200
  DSBCAPS_STICKYFOCUS = &H4000
  DSBCAPS_GLOBALFOCUS = &H8000&
  DSBCAPS_GETCURRENTPOSITION2 = &H10000
  DSBCAPS_MUTE3DATMAXDISTANCE = &H20000
  DSBCAPS_LOCDEFER = &H40000
End Enum

Public Enum CONST_DSSCL
  DSSCL_NORMAL = &H1
  DSSCL_PRIORITY = &H2
  DSSCL_EXCLUSIVE = &H3
  DSSCL_WRITEPRIMARY = &H4
End Enum

Public Enum CONST_DISCL
  DISCL_EXCLUSIVE = &H1
  DISCL_NONEXCLUSIVE = &H2
  DISCL_FOREGROUND = &H4
  DISCL_BACKGROUND = &H8
  DISCL_NOWINKEY = &H10
End Enum

Public Type DIDEVICEOBJECTDATA
  lOfs As Long
  lData As Long
  lTimeStamp As Long
  lSequence As Long
  uAppData As Long
End Type

Public Type DIDATAFORMAT
  dwSize As Long
  dwObjSize As Long
  dwFlags As Long
  dwDataSize As Long
  dwNumObjs As Long
End Type

Public Type DIPROPHEADER
  lSize As Long
  lHeaderSize As Long
  lObj As Long
  lHow As Long
End Type

Public Type DIPROPDWORD
  diph As DIPROPHEADER
  dwData As Long
End Type


Public Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(7) As Byte
End Type

Public Type DIMOUSESTATE
    lX As Long
    lY As Long
    lZ As Long
    rgbButtons(4) As Byte
End Type
Public Enum DIMOFS
  DIMOFS_X = 0
  DIMOFS_Y = 4
  DIMOFS_Z = 8
End Enum


' DECLARES ========================================================

Private Declare Sub ds_Create Lib "dx_vb" (ByRef pDS As Long)
Private Declare Sub di_Create Lib "dx_vb" (ByRef pDI As Long)

' FUNCTIONS =======================================================

Public Function DirectInputCreate() As DirectInput8
  Dim pDI As Long

  di_Create pDI
  If pDI <> 0 Then
    Set DirectInputCreate = New DirectInput8
    DirectInputCreate.Ptr = pDI
  End If
End Function

Public Function CreateDirectSound() As DirectSound8
  Dim pDS As Long

  ds_Create pDS
  If pDS <> 0 Then
    Set CreateDirectSound = New DirectSound8
    CreateDirectSound.Ptr = pDS
  End If
End Function
