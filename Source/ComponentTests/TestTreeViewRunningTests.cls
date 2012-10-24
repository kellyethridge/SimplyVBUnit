VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTreeViewRunningTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Controller  As TestTreeController
Private Suite       As TestSuite
Private Fixture     As FakeTest
Private View        As TreeView
Private f           As frmTreeView
Private Listener    As IEventListener


Public Sub Setup()
    Set Listener = New EventCastListener
    
    Set f = New frmTreeView
    Call Load(f)
    Set View = f.TreeView1
    
    Set Fixture = NewTest("Fixture")
    Fixture.Tests.Add NewTest("ATest", Fixture)
    Fixture.Tests.Add NewTest("TestTwo", Fixture)
    
    Set Suite = Sim.NewTestSuite("Root")
    Suite.Add Fixture
    
    Set Controller = UI.NewTestTreeController(View, Suite, Listener)
End Sub

Public Sub Teardown()
    Set View = Nothing
    Call Unload(f)
    Set f = Nothing
End Sub

Public Sub CheckInitialImages()
    Dim Node As Node
    For Each Node In View.Nodes
        Assert.That Node.Image, Iz.EqualTo("NotRun")
    Next Node
End Sub

Public Sub RunSingleTest()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Fixture.Tests(0))
    Result.Success
    Result.AssertCount = 1
    Listener.TestCaseFinished Result
    
    Assert.AreEqual "Passed", View.Nodes("Root").Image
    Assert.AreEqual "Passed", View.Nodes("Root.Fixture").Image
    Assert.AreEqual "Passed", View.Nodes("Root.Fixture.ATest").Image
    Assert.AreEqual "NotRun", View.Nodes("Root.Fixture.TestTwo").Image
End Sub

Public Sub RunSingleFailingTest()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Fixture.Tests(0))
    Result.Failure
    Result.AssertCount = 1
    Listener.TestCaseFinished Result
    
    Assert.AreEqual "Failed", View.Nodes("Root").Image
    Assert.AreEqual "Failed", View.Nodes("Root.Fixture").Image
    Assert.AreEqual "Failed", View.Nodes("Root.Fixture.ATest").Image
    Assert.AreEqual "NotRun", View.Nodes("Root.Fixture.TestTwo").Image
End Sub

Public Sub RunAllWithSingleFailingTest()
    Dim Result As TestResult
    Set Result = Sim.NewTestResult(Fixture.Tests(0))
    Result.Failure
    Result.AssertCount = 1
    Listener.TestCaseFinished Result
    
    Set Result = Sim.NewTestResult(Fixture.Tests(1))
    Result.Success
    Result.AssertCount = 1
    Listener.TestCaseFinished Result
    
    Assert.AreEqual "Failed", View.Nodes("Root").Image
    Assert.AreEqual "Failed", View.Nodes("Root.Fixture").Image
    Assert.AreEqual "Failed", View.Nodes("Root.Fixture.ATest").Image
    Assert.AreEqual "Passed", View.Nodes("Root.Fixture.TestTwo").Image
End Sub

