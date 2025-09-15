Attribute VB_Name = "modCasting"
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


'Declares
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef Source As Any, ByVal bytes As Long)


Public Function CVA(ByRef arg As String) As Variant
     
     'Convert from String to Any
     CopyMemory CVA, ByVal arg, Len(arg)
End Function

Public Function CVB(ByRef arg As String) As Byte
     
     'Convert from String to Byte
     CopyMemory CVB, ByVal arg, 1
End Function

Public Function CVC(ByRef arg As String) As Currency
     
     'Convert from String to Currency
     CopyMemory CVC, ByVal arg, 8
End Function

Public Function CVD(ByRef arg As String) As Double
     
     'Convert from String to Double
     CopyMemory CVD, ByVal arg, 8
End Function

Public Function CVDt(ByRef arg As String) As Date
     
     'Convert from String to Date
     CopyMemory CVDt, ByVal arg, 8
End Function

Public Function CVI(ByRef arg As String) As Integer
     
     'Convert from String to Integer
     CopyMemory CVI, ByVal arg, 2
End Function

Public Function CVL(ByRef arg As String) As Long
     
     'Convert from String to Long
     CopyMemory CVL, ByVal arg, 4
End Function

Public Function CVS(ByRef arg As String) As Single
     
     'Convert from String to Single
     CopyMemory CVS, ByVal arg, 4
End Function

Public Function CVV(ByRef arg As String) As Variant
     'IMPORTANT: don't use for Variants holding strings, arrays or objects
     
     'Convert from String to Variant
     CopyMemory CVV, ByVal arg, 16
End Function

Public Function ItoL(ByRef Value As Integer) As Long
     
     'Convert from Integer to Long
     ItoL = 0
     CopyMemory ItoL, Value, 2
End Function

Public Function LtoI(ByRef Value As Long) As Integer
     
     'Convert from Long to Integer
     LtoI = 0
     CopyMemory LtoI, Value, 2
End Function

Public Function MKA(ByRef Value) As String
     
     'Convert from Any to String
     MKA = Space$(LenB(Value))
     CopyMemory ByVal MKA, Value, LenB(Value)
End Function

Public Function MKB(ByRef Value As Byte) As String
     
     'Convert from Byte to String
     MKB = Space$(1)
     CopyMemory ByVal MKB, Value, 1
End Function

Public Function MKC(ByRef Value As Currency) As String
     
     'Convert from Currency to String
     MKC = Space$(8)
     CopyMemory ByVal MKC, Value, 8
End Function

Public Function MKD(ByRef Value As Double) As String
     
     'Convert from Double to String
     MKD = Space$(8)
     CopyMemory ByVal MKD, Value, 8
End Function

Public Function MKDt(ByRef Value As Date) As String
     
     'Convert from Date to String
     MKDt = Space$(8)
     CopyMemory ByVal MKDt, Value, 8
End Function

Public Function MKI(ByRef Value As Integer) As String
     
     'Convert from Integer to String
     MKI = Space$(2)
     CopyMemory ByVal MKI, Value, 2
End Function

Public Function MKL(ByRef Value As Long) As String
     
     'Convert from Long to String
     MKL = Space$(4)
     CopyMemory ByVal MKL, Value, 4
End Function

Public Function MKS(ByRef Value As Single) As String
     
     'Convert from Single to String
     MKS = Space$(4)
     CopyMemory ByVal MKS, Value, 4
End Function

Public Function MKV(ByRef Value As Variant) As String
     'IMPORTANT: don't use for Variants holding strings, arrays or objects
     
     'Convert from Variant to String
     MKV = Space$(16)
     CopyMemory ByVal MKV, Value, 16
End Function
