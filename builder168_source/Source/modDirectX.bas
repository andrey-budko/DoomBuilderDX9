Attribute VB_Name = "modDirectX"
'
'    Doom Builder
'    Copyright (c) 2003 Pascal vd Heiden, www.codeimp.com
'    This program is released under GNU General Public License
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'


'Do not allow any undeclared variables
Option Explicit

'Case sensitive comparisions
Option Compare Binary


'Declarations
Public Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long


'DirectX objects
Public D3D As Direct3D9
Public D3DD As Direct3DDevice9
Public D3D_BB As Direct3DSurface9
Public DI As DirectInput8
Public DIMouse As DirectInputDevice8

'3D Mode
Public Running3D As Boolean

'Texture Format
Public TEXTUREFORMAT As D3DFORMAT

'Vertex Formats
Public Const VERTEXFVF As Long = D3DFVF_XYZ Or D3DFVF_TEX1 'Or D3DFVF_NORMAL
Public Const VERTEXSTRIDE As Long = 5 * 4 '8 * 4
Public Const TLVERTEXFVF As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 ' Or D3DFVF_SPECULAR
Public Const TLVERTEXSTRIDE As Long = 7 * 4 '8 * 4

'Vertex structure
Public Type VERTEX
     x As Single
     y As Single
     Z As Single
     tu As Single
     tv As Single
End Type

'Transformed & Lit Vertex structure
Public Type TLVERTEX
     sx As Single
     sy As Single
     sz As Single
     rhw As Single
     Color As Long
     'Specular As Long
     tu As Single
     tv As Single
End Type

'Video information
Public VideoParams As D3DPRESENT_PARAMETERS

Public Function BitmapData_D3DTexture(ByRef BitmapData() As Byte, ByVal DataWidth As Long, ByVal BitmapHeight As Long, ByRef PaletteData() As BITMAPRGB, ByVal Padding As Boolean) As Direct3DTexture9
     Dim TextureData() As Byte
     Dim TextureHeader As BITMAPFILEHEADER
     Dim TextureInfo As BITMAPINFOHEADER
     Dim TextureMemSize As Long
     Dim TextureWidth As Long
     Dim TextureHeight As Long
     Dim offset As Long
     Dim r As Long
     
     'Check if data padding must be done
     If (Padding) Then
          
          'Make texture size to the power of 2
          TextureWidth = NextPowerOf2(DataWidth)
          TextureHeight = NextPowerOf2(BitmapHeight)
     Else
          
          'Keep original size
          TextureWidth = DataWidth
          TextureHeight = BitmapHeight
     End If
     
     'Create texture header
     With TextureHeader
          .bfType = "BM"
          .bfOffBits = 1078
     End With
     
     'Create texture info
     With TextureInfo
          .biBitCount = 8
          .biClrUsed = 256
          .biPlanes = 1
          .biWidth = TextureWidth
          .biHeight = TextureHeight
          .biSizeImage = TextureWidth * TextureHeight
          .biSize = Len(TextureInfo)
     End With
     
     'Calculate memory size needed
     TextureMemSize = Len(TextureHeader) + Len(TextureInfo) + 256 * 4 + TextureWidth * TextureHeight
     
     'Allocate memory to build texture file in
     ReDim TextureData(0 To TextureMemSize)
     
     'Copy texture header
     CopyMemory TextureData(offset), TextureHeader, Len(TextureHeader)
     offset = offset + Len(TextureHeader)
     
     'Copy texture info
     CopyMemory TextureData(offset), TextureInfo, Len(TextureInfo)
     offset = offset + Len(TextureInfo)
     
     'Copy palette
     CopyMemory TextureData(offset), PaletteData(0), 256 * 4
     offset = offset + 256 * 4
     
     'Check if data padding must be done
     If (Padding) Then
          
          'Fill entire area with transparent bytes
          'FillMemory TextureData(offset), TextureHeight * TextureWidth, TRANSPARENCY_INDEX
          FillBytes TextureData(), offset, TextureHeight * TextureWidth, TRANSPARENCY_INDEX
          
          'Align bitmap to top of texture
          '(the bitmap and texture are upside-down)
          offset = offset + (TextureHeight - BitmapHeight) * TextureWidth
          
          'Go for each row of the bitmap
          For r = 0 To BitmapHeight - 1
               
               'Copy the part of the bitmap data
               'CopyMemory TextureData(offset), BitmapData(r * DataWidth), DataWidth
               CopyBytes BitmapData(), TextureData(), r * DataWidth, offset, DataWidth
               
               'Increase offset
               offset = offset + TextureWidth
          Next r
     Else
          
          'Copy all bitmap data in 1 call
          'CopyMemory TextureData(offset), BitmapData(0), DataWidth * BitmapHeight
          CopyBytes BitmapData(), TextureData(), 0, offset, DataWidth * BitmapHeight
          offset = offset + DataWidth * BitmapHeight
     End If
     
     'Create Direct3D Texture from memory block
     Set BitmapData_D3DTexture = CreateTextureFromFileInMemoryEx( _
                                   D3DD, VarPtr(TextureData(0)), TextureMemSize, D3DX_DEFAULT, _
                                   D3DX_DEFAULT, D3DX_DEFAULT, 0, TEXTUREFORMAT, _
                                   D3DPOOL_MANAGED, D3DX_DEFAULT, D3DX_FILTER_DITHER, _
                                   &HFF000000, ByVal 0, ByVal 0)
     
     'Discard memory block
     Erase TextureData()
End Function

Public Function BitsFromFormat(ByRef Format As D3DFORMAT) As Long
     
     'Return the number of bits each display format has
     Select Case Format
          
          Case D3DFORMAT.D3DFMT_A1R5G5B5
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_A4R4G4B4
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_A8R3G3B2
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_A8R8G8B8
               BitsFromFormat = 32
          
          Case D3DFORMAT.D3DFMT_L6V5U5
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_Q8W8V8U8
               BitsFromFormat = 32
          
          Case D3DFORMAT.D3DFMT_R3G3B2
               BitsFromFormat = 8
          
          Case D3DFORMAT.D3DFMT_R5G6B5
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_R8G8B8
               BitsFromFormat = 24
          
          Case D3DFORMAT.D3DFMT_X1R5G5B5
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_X4R4G4B4
               BitsFromFormat = 16
          
          Case D3DFORMAT.D3DFMT_X8L8V8U8
               BitsFromFormat = 32
          
          Case D3DFORMAT.D3DFMT_X8R8G8B8
               BitsFromFormat = 32
               
     End Select
End Function

Public Sub CreateGammaCorrection(ByVal Gamma As Single, ByVal Brightness As Long)
     Dim i As Long
     Dim r As Long, g As Long, b As Long
     Dim ramp As D3DGAMMARAMP
     
     'Gamma is a value which multiplies the colors
     'Brightness is a value which is added to the colors
     
     'Adjust Brightness for Integer value
     Brightness = Brightness * 257
     
     'Go for all 255 color shades
     For i = 0 To 255
          
          'Create basic colors
          r = 257 * i
          g = 257 * i
          b = 257 * i
          
          'Adjust color with Gamma
          r = r * Gamma
          g = g * Gamma
          b = b * Gamma
          
          'Adjust color with Brightness
          r = r + Brightness
          g = g + Brightness
          b = b + Brightness
          
          'Limit the colors
          If (r > 65535) Then r = 65535
          If (r < 0) Then r = 0
          If (g > 65535) Then g = 65535
          If (g < 0) Then g = 0
          If (b > 65535) Then b = 65535
          If (b < 0) Then b = 0
          
          'Assign the color to the ramp
          ramp.red(i) = CVI(MKL(r))
          ramp.green(i) = CVI(MKL(g))
          ramp.blue(i) = CVI(MKL(b))
     Next i
     
     'Apply the ramp
     D3DD.SetGammaRamp D3DSGR_NO_CALIBRATION, VarPtr(ramp)
End Sub

Public Function InitDirectX() As Boolean
     On Local Error GoTo InitError
     
     'Clear last error
     Err.Clear
     
     'Create Direct3D Device
     Set D3D = DirectX9.CreateDirect3D
     
     'Leave now
     InitDirectX = True
     Exit Function
     
InitError:
     
     'Clean up
     Set D3D = Nothing
     
     'Return false
     InitDirectX = False
End Function

Public Function InitMouse() As Boolean
     On Local Error Resume Next    'If we cant capture it now, we'll do later
     Dim DIProp As DIPROPHEADER
     
     'Initialize Mouse
     Set DIMouse = DI.CreateDeviceMouse

     'DIMouse.SetCommonDataFormat DIFORMAT_MOUSE
     DIMouse.SetCommonDataFormatMouse
     
     'Get cooperative access
     Err.Clear
     If (Val(Config("exclusivemouse"))) Then DIMouse.SetCooperativeLevel frm3D.hWnd, DISCL_EXCLUSIVE Or DISCL_BACKGROUND
     If (Val(Config("exclusivemouse")) = 0) Or (Err.number <> 0) Then DIMouse.SetCooperativeLevel frm3D.hWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND
     Err.Clear
     
     'Set buffer size for mouse
     'DIProp.lHow = DIPH_DEVICE
     'DIProp.lObj = 0
     'DIProp.lData = 20
     'DIMouse.SetProperty "DIPROP_BUFFERSIZE", DIProp
     DIMouse.SetPropertyBufferSize (20)
     
     'Acquire the mouse
     DIMouse.Acquire
     
     If Err.number = 0 Then InitMouse = True
End Function

Public Sub StartDirectX()
     Dim Adapter As Long
     
     'Initialize
     If InitDirectX Then
          
          'Unload status dialog
          Unload frmStatus
          Set frmStatus = Nothing
          
          'Check if going windowed
          If (Val(Config("windowedvideo"))) Then
               
               'Pixels
               frmMain.picMap.ScaleMode = vbPixels
               
               'Create parameters for windowed mode
               With VideoParams
                    .AutoDepthStencilFormat = D3DFMT_D16                   '16-bit depth buffer
                    .BackBufferCount = 1                                   '1 render target
                    .BackBufferFormat = D3DFMT_UNKNOWN
                    .BackBufferHeight = frmMain.picMap.ScaleHeight
                    .BackBufferWidth = frmMain.picMap.ScaleWidth
                    .FullScreen_RefreshRateInHz = 0
                    .EnableAutoDepthStencil = 1                            'automatically draw on the z buffer
                    .SwapEffect = D3DSWAPEFFECT_DISCARD                    'use whatever is fastest technique for flipping render target
                    .Windowed = 1                                          'windowed
                    .hDeviceWindow = frmMain.picMap.hWnd
               End With
               
          Else
               
               'Get the adapter to use
               Adapter = Val(Config("videoadapter"))
               
               'Create Presentation Parameters
               With VideoParams
                    .AutoDepthStencilFormat = D3DFMT_D16                   '16-bit depth buffer
                    .BackBufferCount = 1                                   '1 render target
                    .BackBufferFormat = Config("videoformat")              'video format as configured
                    .BackBufferHeight = Config("videoheight")              'video resolution as configured
                    .BackBufferWidth = Config("videowidth")
                    .FullScreen_RefreshRateInHz = Config("videorate")      'refresh rate as configured
                    .EnableAutoDepthStencil = 1                            'automatically draw on the z buffer
                    .SwapEffect = D3DSWAPEFFECT_DISCARD                    'use whatever is fastest technique for flipping render target
                    .Windowed = 0                                          'no.
                    .hDeviceWindow = frm3D.hWnd
               End With
          End If
          
          'Check if 8-bit paletted texture format is supported
          If (D3D.CheckDeviceFormat(Adapter, D3DDEVTYPE_HAL, Config("videoformat"), 0, D3DRTYPE_TEXTURE, D3DFMT_P8) = D3D_OK) Then
               
               'Use this texture format
               TEXTUREFORMAT = D3DFMT_P8
               
          'Otherwise
          Else
               
               'Use whatever the videocard prefers
               TEXTUREFORMAT = D3DFMT_UNKNOWN
          End If
          
          'Check if not windowed
          If (Val(Config("windowedvideo")) = 0) Then
               
               'Show the render target
               frm3D.Show
               frm3D.SetFocus
               
               'Make it on-top
               SetTopMostWindow frm3D.hWnd, True
          End If
          
          'Create Direct3D Device
          On Local Error Resume Next
          Set D3DD = D3D.CreateDevice(frm3D.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, VideoParams)
          On Local Error GoTo 0
          
          'Check if any errors occurred
          If (D3DD Is Nothing) Then
               
               'Error while starting Direct3D
               Err.Raise vbObjectError + 3, , "Direct3D could not be started." & vbLf & vbLf & "Please ensure that you have the latest DirectX installed and that your videocard and videodrivers support DirectX. Check your configuration for options that may depend on the support of your videocard. Make sure no other application is using your videocard acceleration features and that the acceleration features are enabled."
          Else
               
               'Create a DirectInput object
               On Local Error Resume Next
               Set DI = DirectInputCreate
               On Local Error GoTo 0
               
               'Check if any errors occurred
               If (DI Is Nothing) Then
                    
                    'Error while starting DirectInput
                    Err.Raise vbObjectError + 4, , "DirectInput could not be started." & vbLf & vbLf & "Please ensure that your configuration is correct and that no other application is using direct access to your mouse or keyboard."
               Else
                    
                    'Clear buffers (get rid of noise)
                    D3DD.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Val(Config("palette")("CLR_BACKGROUND")), 1, 0
                    D3DD.Present
                    D3DD.Clear D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Val(Config("palette")("CLR_BACKGROUND")), 1, 0
                    D3DD.Present
                    
                    'Keep a references to the render target
                    'Set D3D_BB = D3DD.GetRenderTarget
                    
                    'Disable the main form when not running windowed
                    'If (Val(Config("windowedvideo")) = 0) Then frmMain.Enabled = False
               End If
          End If
     Else
          
          'Error while initializing DirectX
          Err.Raise vbObjectError + 2, , "DirectX could not be initialized." & vbLf & vbLf & "Please ensure that you have the latest DirectX installed and that your videocard and videodrivers support DirectX."
     End If
End Sub

Public Function StdPicture_D3DTexture(ByRef Bitmap As StdPicture) As Direct3DTexture9
     Dim Tempfile As String
     
     'Check if we must make a temporary file
     Tempfile = MakeTempFile(False)
     
     'Remove the temp file if exists
     If (Dir(Tempfile) <> "") Then Kill Tempfile
     
     'Save picture to temp file
     SavePicture Bitmap, Tempfile
     
     'Create Direct3D Texture from file
     Set StdPicture_D3DTexture = CreateTextureFromFileEx( _
                                   D3DD, Tempfile, D3DX_DEFAULT, D3DX_DEFAULT, _
                                   0, 0, TEXTUREFORMAT, _
                                   D3DPOOL_DEFAULT, D3DX_DEFAULT, _
                                   D3DX_FILTER_DITHER, _
                                   &HFF000000, ByVal 0, ByVal 0)
     
     'Kill the file
     Kill Tempfile
End Function

Public Sub TerminateDirectX()
     On Error Resume Next
     
     'Reenable main form
     If (IsLoaded(frmMain)) Then frmMain.Enabled = True
     
     'Erase Buffer references
     Set D3D_BB = Nothing
     
     'Erase DirectX references
     Set DIMouse = Nothing
     Set DI = Nothing
     Set D3DD = Nothing
     Set D3D = Nothing
     
     'Unload 3D rendering form
     Unload frm3D
     Set frm3D = Nothing
End Sub

Public Function timeExactTime() As Long
     Dim QPFrequency As Currency
     Dim QPCounter As Currency
     
     'Get the CPU's timer frequency
     QueryPerformanceFrequency QPFrequency
     
     If QPFrequency Then
          'Use the CPU's internal clock
          QueryPerformanceCounter QPCounter      'Get the CPU's tick count
          timeExactTime = (QPCounter / QPFrequency) * 1000
     Else
          'The CPU does not have a high performance timer, use the default timing
          timeExactTime = GetTickCount
     End If
End Function

Public Function Vector3D(ByVal x As Single, _
                         ByVal y As Single, _
                         ByVal Z As Single) As D3DVECTOR
    With Vector3D
        .x = x
        .y = y
        .Z = Z
    End With
End Function

