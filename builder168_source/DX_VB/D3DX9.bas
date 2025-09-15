Attribute VB_Name = "D3DX9"
Option Explicit

Public Enum D3DX_FILTER
  D3DX_FILTER_NONE = 1
  D3DX_FILTER_POINT = 2
  D3DX_FILTER_LINEAR = 3
  D3DX_FILTER_TRIANGLE = 4
  D3DX_FILTER_BOX = 5

  D3DX_FILTER_MIRROR_U = &H10000
  D3DX_FILTER_MIRROR_V = &H20000
  D3DX_FILTER_MIRROR_W = &H40000
  D3DX_FILTER_MIRROR = &H70000

  D3DX_FILTER_DITHER = &H80000
  D3DX_FILTER_DITHER_DIFFUSION = &H100000

  D3DX_FILTER_SRGB_IN = &H200000
  D3DX_FILTER_SRGB_OUT = &H400000
  D3DX_FILTER_SRGB = &H600000
End Enum

Public Enum D3DXIMAGE_FILEFORMAT
  D3DXIFF_BMP = 0
  D3DXIFF_JPG = 1
  D3DXIFF_TGA = 2
  D3DXIFF_PNG = 3
  D3DXIFF_DDS = 4
  D3DXIFF_PPM = 5
  D3DXIFF_DIB = 6
  D3DXIFF_HDR = 7
  D3DXIFF_PFM = 8
End Enum

Public Type D3DXIMAGE_INFO
  Width As Long
  Height As Long
  Depth As Long
  MipLevels As Long
  Format As D3DFORMAT
  ResourceType As D3DRESOURCETYPE
  ImageFileFormat As D3DXIMAGE_FILEFORMAT
End Type

Public Const D3DX_DEFAULT As Long = -1

Private Declare Sub D3DX_CreateTextureFromFile Lib "dx_vb" (ByVal pDev As Long, ByVal pSrcFile As Long, ByRef pTex As Long)
Private Declare Sub D3DX_CreateTextureFromFileEx Lib "dx_vb" (ByVal pDev As Long, ByVal pSrcFile As Long, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As Long, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As Long, ByVal MipFilter As Long, ByVal ColorKey As Long, ByVal pSrcInfo As Long, ByVal pPalette As Long, ByRef pTex As Long)
Private Declare Sub D3DX_CreateTextureFromFileInMemory Lib "dx_vb" (ByVal pDev As Long, ByVal pData As Long, ByVal DataSize As Long, ByRef pTex As Long)
Private Declare Sub D3DX_CreateTextureFromFileInMemoryEx Lib "dx_vb" (ByVal pDev As Long, ByVal pData As Long, ByVal DataSize As Long, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As Long, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As Long, ByVal MipFilter As Long, ByVal ColorKey As Long, ByVal pSrcInfo As Long, ByVal pPalette As Long, ByRef pTex As Long)
Private Declare Sub D3DX_GetImageInfoFromFile Lib "dx_vb" (ByVal pSrcFile As Long, ByRef imgInfo As D3DXIMAGE_INFO)

Public Function CreateTextureFromFileInMemory(d3dDev As Direct3DDevice9, ByVal pData As Long, ByVal DataSize As Long) As Direct3DTexture9
  Dim pTex As Long

  D3DX_CreateTextureFromFileInMemory d3dDev.Ptr, pData, DataSize, pTex
  If pTex <> 0 Then
    Set CreateTextureFromFileInMemory = New Direct3DTexture9
    CreateTextureFromFileInMemory.Ptr = pTex
  End If
End Function

Public Function CreateTextureFromFileInMemoryEx(d3dDev As Direct3DDevice9, ByVal pData As Long, ByVal DataSize As Long, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As D3DUSAGE, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As D3DX_FILTER, ByVal MipFilter As D3DX_FILTER, ByVal ColorKey As Long, Optional ByVal pImageInfo As Long = 0, Optional ByVal pPalette As Long = 0) As Direct3DTexture9
  Dim pTex As Long

  D3DX_CreateTextureFromFileInMemoryEx d3dDev.Ptr, pData, DataSize, Width, Height, MipLevels, Usage, Format, Pool, Filter, MipFilter, ColorKey, pImageInfo, pPalette, pTex
  If pTex <> 0 Then
    Set CreateTextureFromFileInMemoryEx = New Direct3DTexture9
    CreateTextureFromFileInMemoryEx.Ptr = pTex
  End If
End Function

Public Function CreateTextureFromFile(d3dDev As Direct3DDevice9, fName As String) As Direct3DTexture9
  Dim pTex As Long

  D3DX_CreateTextureFromFile d3dDev.Ptr, StrPtr(fName), pTex
  If pTex <> 0 Then
    Set CreateTextureFromFile = New Direct3DTexture9
    CreateTextureFromFile.Ptr = pTex
  End If
End Function

Public Function CreateTextureFromFileEx(d3dDev As Direct3DDevice9, fName As String, ByVal Width As Long, ByVal Height As Long, ByVal MipLevels As Long, ByVal Usage As D3DUSAGE, ByVal Format As D3DFORMAT, ByVal Pool As D3DPOOL, ByVal Filter As D3DX_FILTER, ByVal MipFilter As D3DX_FILTER, ByVal ColorKey As Long, Optional ByVal pImageInfo As Long = 0, Optional ByVal pPalette As Long = 0) As Direct3DTexture9
  Dim pTex As Long

  D3DX_CreateTextureFromFileEx d3dDev.Ptr, StrPtr(fName), Width, Height, MipLevels, Usage, Format, Pool, Filter, MipFilter, ColorKey, pImageInfo, pPalette, pTex
  If pTex <> 0 Then
    Set CreateTextureFromFileEx = New Direct3DTexture9
    CreateTextureFromFileEx.Ptr = pTex
  End If
End Function

Public Function GetImageInfoFromFile(fName As String) As D3DXIMAGE_INFO
  Dim imgInfo As D3DXIMAGE_INFO

  D3DX_GetImageInfoFromFile StrPtr(fName), imgInfo
  GetImageInfoFromFile = imgInfo
End Function
