Attribute VB_Name = "modUtilities"
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: modUtilities
'

''
' This module contains various utility functions used in support of higher functions within the project.
'
Option Explicit

Public Const SIZEOF_VARIANT         As Long = 16
Public Const ARRAY_DIMENSIONS       As Long = 1
Public Const ENUM_MEMBERID          As Long = -4
Public Const LBOUND_OF_COLLECTION   As Long = 1

''
' Structure represents a proxy array to an already existing Variant array.
'
Public Type ArrayProxy
    pVTable     As Long
    This        As IUnknown
    pRelease    As Long
    Data()      As Variant
    SA          As SafeArray1d
End Type


Public Function GetLBound(ByRef Value As Variant) As Long
    Dim Result As Long
    
    If IsArray(Value) Then
        Result = LBound(Value)
    Else
        Result = LBOUND_OF_COLLECTION
    End If
    
    GetLBound = Result
End Function

Public Function GetMissingVariant(Optional ByVal Value As Variant) As Variant
    GetMissingVariant = Value
End Function

Public Function GetEnumerator(ByRef Enumerable As Variant) As IEnumerator
    If IsArray(Enumerable) Then
        Set GetEnumerator = Sim.NewArrayEnumerator(Enumerable)
    Else
        Set GetEnumerator = Sim.NewEnumVariantEnumerator(Enumerable)
    End If
End Function

''
' Returns if the value supports enumeration using the For..Each loop.
'
' @param Value The value to test for enumeration support.
' @return Returns true if the value supports For..Each, false otherwise.
'
Public Function IsEnumerable(ByRef Value As Variant) As Boolean
    Dim Result As Boolean
    
    If IsArray(Value) Then
        Result = True
    ElseIf IsObject(Value) Then
        Result = SupportsEnumeration(Value)
    End If
    
    IsEnumerable = Result
End Function

Private Function SupportsEnumeration(ByVal Obj As Object) As Boolean
    If Obj Is Nothing Then
        Exit Function
    End If
    
    Dim Info As InterfaceInfo
    Set Info = tli.InterfaceInfoFromObject(Obj)
    
    Dim Member As MemberInfo
    For Each Member In Info.Members
        If IsEnumerationMember(Member) Then
            SupportsEnumeration = True
            Exit Function
        End If
    Next
End Function

Private Function IsEnumerationMember(ByVal Member As MemberInfo) As Boolean
    IsEnumerationMember = (Member.MemberId = ENUM_MEMBERID)
End Function



''
' Retrieves the pointer to an array's SafeArray structure.
'
' @param arr The array to retrieve the pointer to.
' @return A pointer to a SafeArray structure or 0 if the array is null.
'
Public Function GetArrayPointer(ByRef Arr As Variant) As Long
    Const BYREF_ARRAY As Long = VT_BYREF Or vbArray
    
    Dim vt  As Integer
    Dim Ptr As Long
    
    vt = VariantType(Arr)
    Select Case vt And BYREF_ARRAY
        Case BYREF_ARRAY:   Ptr = MemLong(MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET))
        Case vbArray:       Ptr = MemLong(VarPtr(Arr) + VARIANTDATA_OFFSET)
        Case Else
            Call Err.Raise(ErrorCode.Argument, , "GetArrayPointer", "Array is required.")
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

''
' Returns the number of dimensions the array is declared with.
'
' @param Arr The array to get the number of dimensions from.
' @return The number of dimensions.
'
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
    If pSA <> 0 Then
        Call CopyMemory(SA, ByVal pSA, LenB(SA))
        
        If SA.cDims > 0 Then
            Dim Count As Long
            Count = 1
            
            Dim i As Long
            For i = 1 To SA.cDims
                Count = Count * (UBound(Arr, i) - LBound(Arr, i) + 1)
            Next i
        
            VariantType(Src) = VarType(Arr)
            MemLong(VarPtr(Src) + VARIANTDATA_OFFSET) = VarPtr(SA)
            
            SA.lLbound = 0
            SA.cElements = Count
            SA.cDims = 1
            
            Call VariantCopyInd(GetArrayElement, Src(Index))
        End If
    End If

errTrap:
    Call ZeroMemory(Src, 16)
End Function


''
' Initializes an ArrayProxy structure with the array elements it will represent.
'
' @param Proxy The structure to be initialized as a proxy to an array of variants.
' @param FirstArgument The ByRef argument of the first element in the array.
' @param Count The number of elements to be in the array.
'
Public Sub InitArrayProxy(ByRef Proxy As ArrayProxy, ByRef FirstArgument As Variant, ByVal Count As Long)
    Call FillDescriptor(Proxy.SA, FirstArgument, Count)
    Call FillProxy(Proxy)
End Sub

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


''
' A helper function to retrieve the address result from using the AddressOf keyword.
'
' @param Addr The address that is passed in by the AddressOf keyword.
' @return The same address as passed in by the AddressOf keyword.
' @remarks Since you cannot directly get the resulting value from the AddressOf
' keyword, you must have it pass the value into a function. This function will simply
' return the address passed in.
'
Public Function FuncAddr(ByVal Addr As Long) As Long
    FuncAddr = Addr
End Function

Private Function ArrayProxy_Release(ByRef This As ArrayProxy) As Long
    SAPtr(This.Data) = vbNullPtr
    This.SA.pvData = vbNullPtr
End Function

Public Function EqualStrings(ByRef String1 As Variant, ByRef String2 As Variant, Optional ByVal IgnoreCase As Boolean) As Boolean
    Dim Method As VbCompareMethod
    If Not IgnoreCase Then
        Method = vbBinaryCompare
    Else
        Method = vbTextCompare
    End If
    
    EqualStrings = (StrComp(String1, String2, Method) = 0)
End Function

Public Function TryGetCount(ByRef Source As Variant, ByRef Result As Long) As Boolean
    On Error GoTo errTrap
    
    If Not IsArray(Source) Then
        Result = Source.Count
        TryGetCount = True
    End If
    
    Exit Function
    
errTrap:
    Result = 0
End Function

Public Function NewLongs(ByVal Size As Long) As Long()
    SAPtrLong(NewLongs) = SafeArrayCreateVector(vbLong, 0, Size)
End Function

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
