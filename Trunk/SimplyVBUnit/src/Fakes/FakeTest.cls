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
Attribute VB_Exposed = True
Option Explicit
Implements ITest

Public Parent As ITest


Private Function ITest_CountTestCases(ByVal Filter As SVBUnit.ITestFilter) As Long

End Function

Private Property Get ITest_HasParent() As Boolean

End Property

Private Property Get ITest_IsFixture() As Boolean

End Property

Private Property Get ITest_IsSuite() As Boolean

End Property

Private Property Get ITest_IsTestCase() As Boolean

End Property

Private Property Get ITest_Parent() As ITest
    Set ITest_Parent = Parent
End Property

Private Property Set ITest_Parent(ByVal RHS As ITest)
    Set Parent = RHS
End Property


Private Property Get ITest_FullName() As String

End Property

Private Property Get ITest_Name() As String

End Property

Private Sub ITest_Reset()

End Sub

Private Function ITest_Run(Optional ByVal Listener As IEventListener, Optional ByVal Filter As ITestFilter) As TestResult

End Function

Private Property Let ITest_RunState(ByVal RHS As SVBUnit.RunState)

End Property

Private Property Get ITest_RunState() As SVBUnit.RunState

End Property

Private Sub ITest_Sort(Optional ByVal Comparer As SVBUnit.ITestComparer)

End Sub

Private Property Get ITest_TestClass() As Object

End Property

Private Property Get ITest_TestCount() As Long

End Property

Private Property Get ITest_Tests() As SVBUnit.TestList

End Property
