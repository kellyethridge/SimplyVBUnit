VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixtureWithTestCases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Sub ITestCaseSource_GetTestCases(ByVal Test As TestCaseCollector)
    Select Case Test.MethodName
        Case "TestMe"
            Call Test.use("hi")
            Call Test.use("hello")
            Call Test.use("Hello")
            Call Test.use("bye")
            Call Test.use("one")
            
        Case "LotsOfTests"
            Dim i As Long
            For i = 1 To 50
                Call Test.use(i)
            Next i
    End Select
End Sub

Public Sub TestMe(ByVal s As String)
    Assert.AreEqual s, s
    TestContext.Out.WriteLine "Hello world [" & s & "]"
End Sub

Public Sub LotsOfTests(ByVal Value As Long)
    Assert.AreEqual Value, Value
End Sub

Public Sub TestStrict()
    Dim a As New Collection
    TestContext.Log.WriteLine "Log - "
    Assert.That a, Iz.EqualTo(Me).Strict
End Sub
