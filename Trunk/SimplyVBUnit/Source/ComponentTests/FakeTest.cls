VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITest

Public Name As String
Public TestCount As Long
Public Parent As ITest
Public Tests As New TestList

Private Property Set ITest_Categories(ByVal RHS As Categorization)

End Property

Private Property Get ITest_Categories() As Categorization

End Property

Private Function ITest_CountTestCases(ByVal Filter As ITestFilter) As Long

End Function

Private Property Get ITest_Fixture() As Object

End Property

Private Property Get ITest_FullName() As String
    ITest_FullName = Parent.FullName & "." & Name
End Property

Private Property Get ITest_HasCategories() As Boolean

End Property

Private Property Get ITest_HasParent() As Boolean

End Property

Private Property Get ITest_IsFixture() As Boolean

End Property

Private Property Get ITest_IsSuite() As Boolean

End Property

Private Property Get ITest_Name() As String
    ITest_Name = Name
End Property

Private Property Set ITest_Parent(ByVal RHS As ITest)
    Set Parent = RHS
End Property

Private Property Get ITest_Parent() As ITest
    Set Parent = Parent
End Property

Private Sub ITest_Reset()

End Sub

Private Function ITest_Run(Optional ByVal Listener As IEventListener, Optional ByVal Filter As ITestFilter) As TestResult
    
End Function

Private Property Let ITest_RunState(ByVal RHS As RunState)

End Property

Private Property Get ITest_RunState() As RunState

End Property

Private Sub ITest_SetRunner(ByVal Runner As ITestRunner)

End Sub

Private Sub ITest_Sort(Optional ByVal Comparer As ITestComparer)

End Sub

Private Property Get ITest_TestClass() As Object

End Property

Private Property Get ITest_TestCount() As Long
    ITest_TestCount = Me.Tests.Count
End Property

Private Property Get ITest_Tests() As TestList
    Set ITest_Tests = Tests
End Property
