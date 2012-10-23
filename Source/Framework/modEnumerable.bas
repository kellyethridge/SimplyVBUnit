Attribute VB_Name = "modEnumerable"
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
' Module: SimplyVBUnit.modEnumerable
'
' Provides functions that deal with Collections and enumerable entities, including arrays.
'
Option Explicit

Public Const ENUM_MEMBERID As Long = -4


Public Function ContainsKey(ByVal Items As Collection, ByRef Key As String) As Boolean
    On Error GoTo errTrap
    
    ContainsKey = IsObject(Items(Key)) Or True
    Exit Function
    
errTrap:
End Function

Public Function GetEnumerator(ByRef Enumerable As Variant) As IEnumerator
    If IsArray(Enumerable) Then
        Set GetEnumerator = Sim.NewArrayEnumerator(Enumerable)
    Else
        Set GetEnumerator = Sim.NewEnumVariantEnumerator(Enumerable)
    End If
End Function

Public Function IsCollection(ByRef Value As Variant) As Boolean
    If IsObject(Value) Then
        If Not Value Is Nothing Then
            IsCollection = TypeOf Value Is Collection
        End If
    End If
End Function

Public Function IsEnumerable(ByRef Value As Variant) As Boolean
    Dim Result As Boolean
    
    If IsArray(Value) Then
        Result = True
    ElseIf IsObject(Value) Then
        Result = SupportsEnumeration(Value)
    End If
    
    IsEnumerable = Result
End Function

Public Function TryGetCount(ByRef Source As Variant, ByRef Result As Long) As Boolean
    On Error GoTo errTrap
    
    If IsObject(Source) Then
        Result = Source.Count
        TryGetCount = True
    End If
    
    Exit Function
    
errTrap:
    Result = 0
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Private Helpers
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
