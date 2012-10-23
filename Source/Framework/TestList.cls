VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Module: SComponent.TestList
'
Option Explicit

Private Const MINIMUM_INDEX         As Long = 0
Private Const DEF_CAPACITY          As Long = 16
Private Const CAPACITY_GROWTH_RATE  As Long = 2
Private Const BASE_ONE_OFFSET       As Long = 1

Private mItems()    As ITest
Private mCount      As Long
Private mComparer   As ITestComparer


Public Property Get Capacity() As Long
    Capacity = UBound(mItems) + 1
End Property

Public Property Get Count() As Long
    Count = mCount
End Property

Public Function Find(ByRef TestName As String) As ITest
    Dim i As Long
    For i = 0 To mCount - 1
        If mItems(i).Name = TestName Then
            Set Find = mItems(i)
            Exit Function
        End If
    Next
End Function

Public Sub Add(ByVal Test As ITest)
    If Test Is Nothing Then _
        Error.CannotBeNothing Me, "Add", "Test"
        
    mCount = mCount + 1
    EnsureCapacity
    Set mItems(mCount - BASE_ONE_OFFSET) = Test
End Sub

Private Sub EnsureCapacity()
    If NeedsMoreCapacity Then
        IncreaseCapacity
    End If
End Sub

Private Function NeedsMoreCapacity() As Boolean
    NeedsMoreCapacity = (CurrentCapacity < RequiredCapacity)
End Function

Private Function CurrentCapacity() As Long
    CurrentCapacity = UBound(mItems) + BASE_ONE_OFFSET
End Function

Private Function RequiredCapacity() As Long
    RequiredCapacity = mCount
End Function

Private Sub IncreaseCapacity()
    Dim NewCapacity As Long
    NewCapacity = CurrentCapacity * CAPACITY_GROWTH_RATE
    
    SetCapacity NewCapacity
End Sub

Private Sub SetCapacity(ByVal Capacity As Long)
    ReDim Preserve mItems(0 To Capacity - BASE_ONE_OFFSET)
End Sub

Public Property Get Item(ByVal Index As Variant) As ITest
Attribute Item.VB_UserMemId = 0
    If VarType(Index) = vbString Then
        Set Item = Find(CStr(Index))
        If Item Is Nothing Then _
            Error.Argument Me, "Item", "Index", "Test name was not found."
    Else
        If (Index < MINIMUM_INDEX) Or (Index >= Me.Count) Then _
            Error.IndexOutOfRange Me, "Item"
            
        Set Item = mItems(Index)
    End If
End Property

Public Function Equals(ByVal Other As TestList) As Boolean
    If SameSize(Other) Then
        Equals = HasEqualContents(Other)
    End If
End Function

Private Function SameSize(ByVal Other As TestList) As Boolean
    SameSize = (mCount = Other.Count)
End Function

Private Function HasEqualContents(ByVal Other As TestList) As Boolean
    Dim i As Long
    For i = 0 To Me.Count - 1
        If Me(i).Name <> Other(i).Name Then
            Exit Function
        End If
    Next
    
    HasEqualContents = True
End Function

Public Sub Sort(Optional ByVal Comparer As ITestComparer)
    AssignComparer Comparer
    QuickSort 0, mCount - BASE_ONE_OFFSET
End Sub

Private Sub AssignComparer(ByVal Comparer As ITestComparer)
    If Comparer Is Nothing Then
        Set mComparer = New TestComparer
    Else
        Set mComparer = Comparer
    End If
End Sub

Private Sub QuickSort(ByVal Left As Long, ByVal Right As Long)
    Dim i       As Long
    Dim j       As Long
    Dim Pivot   As ITest
    
    Do While Left < Right
        i = Left
        j = Right
        Set Pivot = mItems((i + j) \ 2)
        
        Do
            Do While mComparer.Compare(mItems(i), Pivot) < 0
                i = i + 1
            Loop
            Do While mComparer.Compare(mItems(j), Pivot) > 0
                j = j - 1
            Loop
            
            If i > j Then Exit Do
            If i < j Then
                Dim t As Long
                t = ObjectPtr(mItems(i))
                ObjectPtr(mItems(i)) = ObjectPtr(mItems(j))
                ObjectPtr(mItems(j)) = t
            End If
            
            i = i + 1
            j = j - 1
        Loop While i <= j
        
        If j - Left <= Right - i Then
            If Left < j Then QuickSort Left, j
            Left = i
        Else
            If i < Right Then QuickSort i, Right
            Right = j
        End If
    Loop
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Dim Enumerator As TestListEnumerator
    Set Enumerator = Sim.NewTestListEnumerator(Me)
    Set NewEnum = CreateEnumerator(Enumerator)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Class Events
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()
    SetCapacity DEF_CAPACITY
End Sub


