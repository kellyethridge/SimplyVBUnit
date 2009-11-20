VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: ArrayEnumerator
'
Option Explicit
Implements IEnumerator

Private Const INDEX_RESET As Long = -1
Private Const EMPTY_ARRAY As Long = -1


Private mArray  As Variant
Private mSA     As SafeArray1d
Private mIndex  As Long


Friend Sub Init(ByRef Arr As Variant)
    Dim pSA As Long
    pSA = GetArrayPointer(Arr)
    
    If ArrayExists(pSA) Then
        Call CreateProxy(Arr, pSA)
    Else
        Call CreateEmptyProxy
    End If
End Sub

Private Function ArrayExists(ByVal Ptr As Long) As Boolean
    ArrayExists = (Ptr <> vbNullPtr)
End Function

Private Sub CreateProxy(ByRef Arr As Variant, ByVal PtrSA As Long)
    Call FillDescriptor(PtrSA)
    
    If HasElements Then
        Call CreateArrayProxy(Arr)
    Else
        Call CreateEmptyProxy
    End If
End Sub

Private Sub FillDescriptor(ByVal Ptr As Long)
    Call CopyMemory(mSA, ByVal Ptr, LenB(mSA))
End Sub

Private Function HasElements() As Boolean
    HasElements = (mSA.cDims > 0)
End Function

Private Sub CreateArrayProxy(ByRef Arr As Variant)
    Dim Count As Long
    Count = 1
    
    Dim i As Long
    For i = 1 To mSA.cDims
        Count = Count * (UBound(Arr, i) - LBound(Arr, i) + 1)
    Next i

    VariantType(mArray) = VarType(Arr)
    MemLong(VarPtr(mArray) + VARIANTDATA_OFFSET) = VarPtr(mSA)
    
    mSA.lLbound = 0
    mSA.cElements = Count
    mSA.cDims = ARRAY_DIMENSIONS
    
    mIndex = INDEX_RESET
End Sub

Private Sub CreateEmptyProxy()
    mSA.cElements = EMPTY_ARRAY
End Sub

Private Sub Class_Terminate()
    Call ZeroMemory(mArray, SIZEOF_VARIANT)
End Sub

Private Property Get IEnumerator_Current() As Variant
    Call VariantCopyInd(IEnumerator_Current, mArray(mIndex))
End Property

Private Function IEnumerator_MoveNext() As Boolean
    mIndex = mIndex + 1
    IEnumerator_MoveNext = (mIndex < mSA.cElements)
End Function

Private Sub IEnumerator_Reset()
    mIndex = INDEX_RESET
End Sub

