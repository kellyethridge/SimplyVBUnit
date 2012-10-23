Attribute VB_Name = "modArray"
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: SimplyVBUnit.modArray
'
' This module provides functionaly for interacting with arrays.
'
Option Explicit

Public Const SIZEOF_VARIANT     As Long = 16
Public Const ARRAY_DIMENSIONS   As Long = 1


Public Type ArrayProxy
    pVTable     As Long
    This        As IUnknown
    pRelease    As Long
    Data()      As Variant
    SA          As SafeArray1d
End Type


Public Function NewLongs(ByVal Size As Long) As Long()
    SAPtrLong(NewLongs) = SafeArrayCreateVector(vbLong, 0, Size)
End Function

Public Function GetArrayPointer(ByRef Arr As Variant) As Long
    Const BYREF_ARRAY As Long = VT_BYREF Or vbArray
    
    Dim vt  As Integer
    Dim Ptr As Long
    
    vt = VariantType(Arr)
    Select Case vt And BYREF_ARRAY
        Case BYREF_ARRAY:   Ptr = MemLong(MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET))
        Case vbArray:       Ptr = MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET)
        Case Else
            Err.Raise ErrorCode.Argument, , "GetArrayPointer", "Array is required."
    End Select
    
    ' HACK HACK HACK
    '
    ' When an uninitialized array of objects or UDTs is passed into a
    ' function as a ByRef Variant, the array is initialized with just the
    ' SafeArrayDescriptor, at which point, it is a valid array and can
    ' be used by UBound and LBound after the call. So, now we're just
    ' going to assume that any object or UDT array that has just the descriptor
    ' allocated was Null to begin with. That means whenever an Object or UDT
    ' array is passed to any method, it will technically never
    ' be uninitialized, just zero-length.
    Select Case vt And &HFF
        Case vbObject, vbUserDefinedType
            Dim PVDataPtr As Long

            PVDataPtr = MemLong(Ptr + PVDATA_OFFSET)
            If PVDataPtr = vbNullPtr Then
                Ptr = vbNullPtr
            End If
    End Select
    
    GetArrayPointer = Ptr
End Function

Public Function GetArrayRank(ByRef Arr As Variant) As Long
    Dim Ptr As Long
    Ptr = GetArrayPointer(Arr)
    GetArrayRank = SafeArrayGetDim(Ptr)
End Function

Public Function GetArrayElement(ByRef Arr As Variant, ByVal Index As Long) As Variant
    Dim Src As Variant
    Dim SA  As SafeArray1d
    Dim pSA As Long
    pSA = GetArrayPointer(Arr)
    
    On Error GoTo errTrap
    If pSA <> vbNullPtr Then
        CopyMemory SA, ByVal pSA, LenB(SA)
        
        If SA.cDims > 0 Then
            VariantType(Src) = VarType(Arr)
            MemLong(VarPtr(Src) + VARIANTDATA_OFFSET) = VarPtr(SA)
            SA.lLbound = 0
            SA.cElements = GetElementCount(Arr, SA.cDims)
            SA.cDims = 1
            
            VariantCopyInd GetArrayElement, Src(Index)
        End If
    End If

errTrap:
    ZeroMemory Src, SIZEOF_VARIANT
End Function

Public Sub InitArrayProxy(ByRef Proxy As ArrayProxy, ByRef FirstArgument As Variant, ByVal Count As Long)
    FillDescriptor Proxy.SA, FirstArgument, Count
    FillProxy Proxy
End Sub

Public Function IsEmptyOrNullArray(ByRef Arr As Variant) As Boolean
    On Error GoTo errTrap
    
    If UBound(Arr) < LBound(Arr) Then
        IsEmptyOrNullArray = True
    End If
    
    Exit Function
    
errTrap:
    Const SUBSCRIPT_OUT_OF_RANGE As Long = 9
    
    If Err.Number = SUBSCRIPT_OUT_OF_RANGE Then
        IsEmptyOrNullArray = True
    Else
        Err.Raise Err.Number, , Err.Description
    End If
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub FillDescriptor(ByRef Descriptor As SafeArray1d, ByRef FirstArgument As Variant, ByVal Count As Long)
    With Descriptor
        .cbElements = SIZEOF_VARIANT
        .cDims = ARRAY_DIMENSIONS
        .cElements = Count
        .pvData = VarPtr(FirstArgument)
    End With
End Sub

Private Sub FillProxy(ByRef Proxy As ArrayProxy)
    With Proxy
        .pVTable = VarPtr(.pVTable)
        .pRelease = FuncAddr(AddressOf ArrayProxy_Release)
        SAPtr(.Data) = VarPtr(.SA)
        ObjectPtr(.This) = VarPtr(.pVTable)
    End With
End Sub

Private Function GetElementCount(ByRef Arr As Variant, ByVal Dimensions As Long) As Long
    Dim Result As Long
    Result = 1
    
    Dim i As Long
    For i = 1 To Dimensions
        Result = Result * (UBound(Arr, i) - LBound(Arr, i) + 1)
    Next i

    GetElementCount = Result
End Function

Private Function ArrayProxy_Release(ByRef This As ArrayProxy) As Long
    SAPtr(This.Data) = vbNullPtr
    This.SA.pvData = vbNullPtr
End Function


