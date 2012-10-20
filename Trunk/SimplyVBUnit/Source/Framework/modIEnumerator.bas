Attribute VB_Name = "modIEnumerator"
'
'      Name:    modEnumerator
'
'      Date:    10/17/2004
'
'      Author:  Kelly Ethridge
'
'   This module creates lightweight objects that will wrap
'   a user's object that implements the IEnumerator interface.
'   By using lightweight objects, the IEnumVariant interface
'   can easily be implements, even though it is not VB friendly.
'
'   The lightweight object simply forwards the IEnumVariant calls
'   to the IEnumerable interface implemented in the user enumerator.
'
'   To learn more about lightweight objects, you should refer to
'   classic book:
'       Advanced Visual Basic 6 Power Techniques for Everyday Programs
'       By Matthew Curland

Option Explicit

Private Const IID_IUnknown_Data1        As Long = 0
Private Const IID_IEnumVariant_Data1    As Long = &H20404

Private Type UserEnumWrapper
   pVTable  As Long
   cRefs    As Long
   UserEnum As IEnumerator
End Type

Private Type VTable
   Functions(0 To 6) As Long
End Type

Private mVTable             As VTable
Private mpVTable            As Long
Private IID_IUnknown        As VBGUID
Private IID_IEnumVariant    As VBGUID



Public Function CreateEnumerator(ByVal Obj As IEnumerator) As stdole.IUnknown
    Init
    Obj.Reset
    
    Dim This As Long
    This = AllocateObjectMemory
    
    Dim Wrapper As UserEnumWrapper
    FillWrapper Wrapper, Obj
    CopyWrapperToMemory Wrapper, This
    EraseWrapper Wrapper
    
    ObjectPtr(CreateEnumerator) = This
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Init()
    If mpVTable = vbNullPtr Then
        InitGUIDS
        InitVTable
    End If
End Sub

Private Function AllocateObjectMemory() As Long
    Const OUT_OF_MEMORY             As Long = 7
    Const LEN_OF_USERENUMWRAPPER    As Long = 12
    
    Dim Result As Long
    Result = CoTaskMemAlloc(LEN_OF_USERENUMWRAPPER)
    
    If Result = vbNullPtr Then
        Err.Raise OUT_OF_MEMORY
    End If
    
    AllocateObjectMemory = Result
End Function

Private Sub FillWrapper(ByRef Wrapper As UserEnumWrapper, ByVal Obj As IEnumerator)
    With Wrapper
        Set .UserEnum = Obj
        .cRefs = 1
        .pVTable = mpVTable
    End With
End Sub

Private Sub CopyWrapperToMemory(ByRef Wrapper As UserEnumWrapper, ByVal PtrDestination As Long)
    CopyMemory ByVal PtrDestination, ByVal VarPtr(Wrapper), LenB(Wrapper)
End Sub

Private Sub EraseWrapper(ByRef Wrapper As UserEnumWrapper)
    ZeroMemory ByVal VarPtr(Wrapper), LenB(Wrapper)
End Sub

Private Sub InitGUIDS()
    With IID_IEnumVariant
        .Data1 = &H20404
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Sub

Private Sub InitVTable()
    With mVTable
        .Functions(0) = FuncAddr(AddressOf QueryInterface)
        .Functions(1) = FuncAddr(AddressOf AddRef)
        .Functions(2) = FuncAddr(AddressOf Release)
        .Functions(3) = FuncAddr(AddressOf IEnumVariant_Next)
        .Functions(4) = FuncAddr(AddressOf IEnumVariant_Skip)
        .Functions(5) = FuncAddr(AddressOf IEnumVariant_Reset)
        .Functions(6) = FuncAddr(AddressOf IEnumVariant_Clone)
        
        mpVTable = VarPtr(.Functions(0))
   End With
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  VTable functions in the IEnumVariant and IUnknown interfaces.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' When VB queries the interface, we support only two.
' IUnknown
' IEnumVariant
Private Function QueryInterface(ByRef This As UserEnumWrapper, ByRef riid As VBGUID, ByRef pvObj As Long) As Long
    Dim ok As Long
    
    Select Case riid.Data1
        Case IID_IEnumVariant_Data1
            ok = IsEqualGUID(riid, IID_IEnumVariant)
        Case IID_IUnknown_Data1
            ok = IsEqualGUID(riid, IID_IUnknown)
    End Select
    
    If ok Then
        pvObj = VarPtr(This)
        AddRef This
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function

Private Function AddRef(ByRef This As UserEnumWrapper) As Long
    With This
        .cRefs = .cRefs + 1
        AddRef = .cRefs
    End With
End Function

Private Function Release(ByRef This As UserEnumWrapper) As Long
    With This
        .cRefs = .cRefs - 1
        Release = .cRefs
        
        If .cRefs = 0 Then
            Delete This
        End If
    End With
End Function

Private Sub Delete(ByRef This As UserEnumWrapper)
   Set This.UserEnum = Nothing
   CoTaskMemFree VarPtr(This)
End Sub


Private Function IEnumVariant_Next(ByRef This As UserEnumWrapper, ByVal celt As Long, ByRef prgVar As Variant, ByVal pceltFetched As Long) As Long
    If This.UserEnum.MoveNext Then
        VariantCopyInd prgVar, This.UserEnum.Current
         
        If pceltFetched <> vbNullPtr Then
            MemLong(pceltFetched) = 1
        End If
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function

Private Function IEnumVariant_Skip(ByRef This As UserEnumWrapper, ByVal celt As Long) As Long
    Do While celt > 0
        If This.UserEnum.MoveNext = False Then
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1
    Loop
End Function

Private Function IEnumVariant_Reset(ByRef This As UserEnumWrapper) As Long
    This.UserEnum.Reset
End Function

Private Function IEnumVariant_Clone(ByRef This As UserEnumWrapper, ByRef ppenum As stdole.IUnknown) As Long
    ObjectPtr(ppenum) = VarPtr(This)
    AddRef This
End Function
