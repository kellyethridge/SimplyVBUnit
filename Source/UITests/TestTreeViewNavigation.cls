VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTreeViewNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITestCaseSource

Private Controller  As TestTreeController
Private Suite       As TestSuite
Private Fixture     As TestFixture
Private View        As TreeView
Private f           As frmTreeView

Public Sub FixtureSetup()
    Set f = New frmTreeView
    Call Load(f)
    Set View = f.TreeView1
    
    Set Fixture = Sim.NewTestFixture(New FixtureWithTests)
    Set Suite = Sim.NewTestSuite("TreeView Tests")
    
    Call Suite.Add(Fixture)
    Set Controller = UI.NewTestTreeController(View, Suite, New EventCastListener)
End Sub

Public Sub FixtureTeardown()
    Set View = Nothing
    Call Unload(f)
    Set f = Nothing
End Sub

Public Function TestSelectedTestIsReturned(ByVal Name As String) As String
    View.Nodes(Name).Selected = True
    
    If Not Controller.SelectedTest Is Nothing Then
        TestSelectedTestIsReturned = Controller.SelectedTest.Name
    End If
End Function


Private Sub ITestCaseSource_GetTestCases(ByVal Test As SimplyVBUnit.TestCaseCollector)
    Select Case Test.MethodName
        Case "TestSelectedTestIsReturned"
            Call Test.use("TreeView Tests.FixtureWithTests.ATest").Expect("ATest")
            Call Test.use("TreeView Tests.FixtureWithTests.AnotherTest").Expect("AnotherTest")
            Call Test.use("TreeView Tests.FixtureWithTests").Expect("FixtureWithTests")
            Call Test.use("TreeView Tests.FixtureWithTests.AThirdTest").Expect("AThirdTest")
    End Select
End Sub
