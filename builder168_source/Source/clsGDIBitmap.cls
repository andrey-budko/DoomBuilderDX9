VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDIBitmap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'    Doom Builder
'    Copyright (c) 2003 Pascal vd Heiden, www.codeimp.com
'    This program is released under GNU General Public License
'
'    GDI+ Bitmap Class written by Paul
'    I dont even know who paul is, just found this on the internet.
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


' Class Object will Dispose on setting equal to Nothing (terminate event)
' And you are responsible for Disposing before any new bitmap is created
' Failure to Dispose properly may lead to various problems/crashes.


'API Declarations
Private Declare Sub OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As PICDESC, riid As CLSID, ByVal fOwn As Long, ByRef lplpvObj As Any)

'Event Declarations
Event Error(ByVal lGdiError As Long, ByVal sErrorDesc As String)

'Private properties
Private Bitmap          As Long
Private lToken          As Long
Private lCurErr         As Long
Private lPixelFormat    As Long
Private m_lWidth        As Long
Private m_lHeight       As Long
Private m_lNumCodecs    As Long
Private rc              As RECT
Private tGuids()        As CLSID ' array of GUIDs for codecs
Private colCodecs       As Collection   ' search into codec index on file extension
Private Stream          As IUnknown

Public Function CreateFromScan0(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lStride As Long, ByVal ePixelFormats As PixelFormats, ByRef lDataPtr As Long)
     CreateFromScan0 = Execute(GdipCreateBitmapFromScan0(lWidth, lHeight, lStride, ePixelFormats, ByVal lDataPtr, Bitmap))
     GetPixelFormat
     GetDimension
End Function

Public Function CreateFromHBitmap(ByVal hBitmap As Long) As Long
     CreateFromHBitmap = Execute(GdipCreateBitmapFromHBITMAP(hBitmap, 0, Bitmap))
     GetPixelFormat
     GetDimension
End Function

Public Function CreateFromGraphics(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lGraphicHandle As Long)
     CreateFromGraphics = Execute(GdipCreateBitmapFromGraphics(lWidth, lHeight, lGraphicHandle, Bitmap))
     GetPixelFormat
     GetDimension
End Function

' I purposely did not use GdipCreateBitmapFromResource for now,
' since it doesn't work on anything but bitmap resources.
' This should work on all GDI+ supported image types
' Store images as "CUSTOM" resources with the resource editor addin
Public Function CreateFromResource(ByVal lResID As Long) As Long
     Dim b() As Byte
     b = LoadResData(lResID, "CUSTOM")
     CreateStreamOnHGlobal b(0), True, Stream
     
     If Not (Stream Is Nothing) Then
          CreateFromResource = Execute(GdipLoadImageFromStream(Stream, Bitmap))
          GetPixelFormat
          GetDimension
     End If
End Function

Public Function LoadFromStream(ByRef lStream As IStream) As Long
     LoadFromStream = Execute(GdipLoadImageFromStream(lStream, Bitmap))
     GetPixelFormat
     GetDimension
End Function


Public Function CreateThumbFromImage(ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBitmapHandle As Long) As Long
     CreateThumbFromImage = Execute(GdipGetImageThumbnail(lBitmapHandle, lWidth, lHeight, Bitmap))
     GetPixelFormat
     GetDimension
End Function

Public Property Get CurErr() As Long
     CurErr = lCurErr
End Property

Public Sub Dispose()
     If Bitmap Then GdipDisposeImage Bitmap
     Set Stream = Nothing
End Sub

Private Function EnumEncoders() As Long
     Dim lNumEncoders  As Long
     Dim lEncoderSize  As Long
     Dim lError        As Long
     Dim b()           As Byte
     Dim codecs()      As ImageCodecInfo
     
     lError = GdipGetImageEncodersSize(lNumEncoders, lEncoderSize)
     If lError = 0 Then
       ReDim codecs(lNumEncoders - 1)
       ReDim b(lEncoderSize - 1)
       
       lError = GdipGetImageEncoders(lNumEncoders, lEncoderSize, b(0))
       If lError = 0 Then
         RtlMoveMemory codecs(0), b(0), lNumEncoders * LenB(codecs(0))
         ReDim tGuids(lNumEncoders - 1)
         m_lNumCodecs = lNumEncoders
         Set colCodecs = Nothing
         Set colCodecs = New Collection
         
         Do While lNumEncoders
           lNumEncoders = lNumEncoders - 1
           tGuids(lNumEncoders) = codecs(lNumEncoders).CLSID
           ParseOnChar StringFromPointerW(codecs(lNumEncoders).pwszFilenameExtension), ";", lNumEncoders
         Loop
       Else
         RaiseEvent Error(lError, GdiErrorString(lError))
       End If
     Else
       RaiseEvent Error(lError, GdiErrorString(lError))
     End If
End Function

Friend Property Get EncoderGuid(ByVal lIndex As Long) As CLSID
     EncoderGuid = tGuids(lIndex)
End Property

Private Function Execute(ByVal lReturn As Long) As Long
     If lReturn Then
          lCurErr = lReturn
          RaiseEvent Error(lReturn, GdiErrorString(lReturn))
     Else
          lCurErr = 0
     End If
     Execute = lCurErr
End Function

' do not compare this to a boolean
'  returns -1 for not found, 0-positive GUID index for found
Public Function ExtensionExists(ByRef sKey As String) As Long
     On Error GoTo errorhandler
     ExtensionExists = True ' invalid index
     
     If Not colCodecs Is Nothing Then
       ExtensionExists = colCodecs.Item(sKey)
     End If
     
     Exit Function
errorhandler:
     ' exit silently
End Function

Public Function LoadFromFile(ByVal sFileName As String) As Long
     LoadFromFile = Execute(GdipLoadImageFromFile(sFileName, Bitmap))
     GetPixelFormat
     GetDimension
End Function

' Returns the number of bytes from the stream in the array
' 0 if failure
Friend Function GetBitmapStream(ByRef bBmpBits As String, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
     Dim Stream As IStream
     Dim lBytesRead As Long
     Dim curMax As Currency
     
     ' True allows the global memory to be freed automatically,
     ' after the destruction of the IStream
     If CreateStreamOnHGlobal(ByVal 0, True, Stream) = 0 Then
       If Execute(GdipSaveImageToStream(Bitmap, Stream, tEncoder, ByVal lEncParamPointer)) = 0 Then
         
         ' Here is an alternate method of getting information from the stream
         ' It's not nearly as versitile, and liable to be inaccurate,
         ' as GlobalSize might be rounded.
       
         ' I thought I'd include it, as the IStream interface need not be
         ' defined to use this code, simply; Dim stream as IUnknown
       
         ' Dim p As Long
         ' Dim hGlob As Long
         ' Dim lSizeGlob As Long
       
         ' GetHGlobalFromStream stream, p
         ' RtlMoveMemory hGlob, ByVal p, 4 ' could also GlobalLock here to deref and lock, remember to unlock later
         ' lSizeGlob = GlobalSize(hGlob)
         ' If lSizeGlob Then
         '   ReDim bBmpBits(lSizeGlob - 1)
         '   RtlMoveMemory bBmpBits(0), ByVal hGlob, lSizeGlob
         ' End If
     
         ' Note: you could seek anywhere in the stream and get any portion
         ' of the bitmap stream.
         Stream.Seek 0@, STREAM_SEEK_END, curMax ' get the size of the stream
         curMax = curMax * 10000 ' adjust for int64
     
         Stream.Seek 0@, STREAM_SEEK_SET, 0@ ' return to the beginning
     
         If curMax > 0 Then
           bBmpBits = Space$(curMax)
           Stream.Read ByVal bBmpBits, curMax, lBytesRead ' note that lBytesRead could overflow, please fix this if you're reading more than 2 Gigs :P
           GetBitmapStream = lBytesRead
         End If
       End If
     End If
     
     Set Stream = Nothing
End Function

Private Sub GetDimension()
     Dim sngWidth  As Single
     Dim sngHeight As Single
     
     Execute GdipGetImageDimension(Bitmap, sngWidth, sngHeight)
     m_lWidth = sngWidth
     m_lHeight = sngHeight
     rc.right = m_lWidth
     rc.bottom = m_lHeight
End Sub

Friend Sub GetPalette(ByVal lPalPointer As Long, ByVal lNumColors As Long)
     Dim lSize As Long
     
     lSize = (lNumColors * 4) + 8
     Execute GdipGetImagePalette(Bitmap, ByVal lPalPointer, lSize)
End Sub

Private Sub GetPixelFormat()
     Execute GdipGetImagePixelFormat(Bitmap, lPixelFormat)
End Sub

Public Function GdiErrorString(ByVal lError As Status) As String
     Dim s As String
     
     Select Case lError
       Case GenericError:              s = "Generic Error"
       Case InvalidParameter:          s = "Invalid Parameter"
       Case OutOfMemory:               s = "Out Of Memory"
       Case ObjectBusy:                s = "Object Busy"
       Case InsufficientBuffer:        s = "Insufficient Buffer"
       Case NotImplemented:            s = "Not Implemented"
       Case Win32Error:                s = "Win32 Error"
       Case WrongState:                s = "Wrong State"
       Case Aborted:                   s = "Aborted"
       Case FileNotFound:              s = "File Not Found"
       Case ValueOverflow:             s = "Value Overflow"
       Case AccessDenied:              s = "Access Denied"
       Case UnknownImageFormat:        s = "Unknown Image Format"
       Case FontFamilyNotFound:        s = "FontFamily Not Found"
       Case FontStyleNotFound:         s = "FontStyle Not Found"
       Case NotTrueTypeFont:           s = "Not TrueType Font"
       Case UnsupportedGdiplusVersion: s = "Unsupported Gdiplus Version"
       Case GdiplusNotInitialized:     s = "Gdiplus Not Initialized"
       Case PropertyNotFound:          s = "Property Not Found"
       Case PropertyNotSupported:      s = "Property Not Supported"
       Case Else:                      s = "Unknown GDI+ Error"
     End Select
     
     GdiErrorString = s
End Function

Public Property Get Handle() As Long
     Handle = Bitmap
End Property

Public Property Get Height() As Long
     Height = m_lHeight
End Property

Private Function IPictureFromBitmap(ByVal lBitmap As Long) As IPicture
     Dim lError  As Long
     Dim hBitmap As Long
     Dim picdes  As PICDESC
     
     If Bitmap = 0 Then Exit Function
     
     lError = GdipCreateHBITMAPFromBitmap(lBitmap, hBitmap, 0)
     
     If hBitmap <> 0 Then
       picdes.cbSizeOfStruct = Len(picdes)
       picdes.picType = vbPicTypeBitmap
       picdes.hGdiObj = hBitmap
       OleCreatePictureIndirect picdes, StringToGuid(IPictureCLSID), True, IPictureFromBitmap
     Else
       RaiseEvent Error(lError, GdiErrorString(lError))
     End If
     ' dispose of bitmap in caller
End Function

' included instead of using Split() for the VB5 set :)
Private Sub ParseOnChar(ByRef sIn As String, ByRef sChar As String, ByVal lGuidIndex As Long)
     Dim lStartPosition As Long
     Dim lFoundPosition As Long
     Dim sItem          As String
     
     lFoundPosition = InStr(sIn, sChar)
     lStartPosition = 1
     
     Do While lFoundPosition
       sItem = Mid$(sIn, lStartPosition, lFoundPosition - lStartPosition)
       colCodecs.Add lGuidIndex, sItem
       lStartPosition = lFoundPosition + 1
       lFoundPosition = InStr(lStartPosition, sIn, sChar)
     Loop
     
     sItem = Trim$(Mid$(sIn, lStartPosition))
     If LenB(sItem) Then colCodecs.Add lGuidIndex, sItem
End Sub

Public Property Get Picture() As StdPicture
     Set Picture = IPictureFromBitmap(Bitmap)
End Property

Public Property Get PixelFormat() As Long
     PixelFormat = lPixelFormat
End Property

Friend Property Get Rectangle() As RECT
     Rectangle = rc
End Property

Public Sub SetTransColor(ByVal lIndex As Byte)
     Dim gpal As ColorPalette256
     
     If Bitmap Then
       GetPalette VarPtr(gpal), 256
       gpal.Entries(lIndex).Alpha = 0
       SetPalette VarPtr(gpal)
     End If
End Sub

Friend Function SaveToFile(ByVal sFileName As String, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
     SaveToFile = Execute(GdipSaveImageToFile(Bitmap, sFileName, tEncoder, ByVal lEncParamPointer))
End Function

Friend Function SaveToStream(ByRef sStream As IStream, ByRef tEncoder As CLSID, ByVal lEncParamPointer As Long) As Long
     SaveToStream = Execute(GdipSaveImageToStream(Bitmap, sStream, tEncoder, ByVal lEncParamPointer))
End Function


Public Function SetPalette(ByVal lPalPointer As Long) As Long
     SetPalette = Execute(GdipSetImagePalette(Bitmap, ByVal lPalPointer))
End Function

Private Function StringFromPointerW(ByVal lPointer As Long) As String
     Dim lLength As Long
     
     If lPointer Then
       lLength = lstrlenW(lPointer)
       StringFromPointerW = Space$(lLength)
       RtlMoveMemory ByVal StrPtr(StringFromPointerW), ByVal lPointer, lLength * 2
     End If
End Function

Private Function StringToGuid(ByRef sGuid As String) As CLSID
     CLSIDFromString sGuid, StringToGuid
End Function

Public Property Get Width() As Long
     Width = m_lWidth
End Property

Private Sub Class_Initialize()
  Dim gsi As GdiplusStartupInput
  
  gsi.GdiplusVersion = 1
  
  ' next line used when debugging in VC, a la DebugBreak API
  ' gsi.DebugEventCallback = PassAddress(AddressOf DebugEventProc)
  Execute GdiplusStartup(lToken, gsi, ByVal 0)
  EnumEncoders
End Sub

Private Sub Class_Terminate()
  Dispose
  If lToken Then GdiplusShutdown lToken
End Sub
