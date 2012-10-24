VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestTreeViewController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private f           As frmTreeView
Private View        As TreeView
Private Suite       As TestSuite
Private Controller  As TestTreeController
Private Listener    As New EventCastListener

Public Sub Setup()
    Set f = New frmTreeView
    Call Load(f)
    Set View = f.TreeView1
    Set Suite = Sim.NewTestSuite("Root")
    Set Controller = UI.NewTestTreeController(View, Suite, Listener)
End Sub

Public Sub Teardown()
    Set Controller = Nothing
    Set View = Nothing
    Call Unload(f)
    Set f = Nothing
End Sub

Public Sub TestNewController()
    Assert.AreEqual 1, View.Nodes.Count
    Assert.AreEqual App.Path & "\Root", View.Nodes(Suite.FullName).Text
End Sub

Public Sub TestAddEmptyTestFixture()
    Dim Fixture As ITest
    Set Fixture = NewTest("Fixture", Nothing)
    Call Suite.Add(Fixture)
    Call Controller.AddTest(Fixture)
    
    Assert.AreEqual 2, View.Nodes.Count
    
    Dim Node As Node
    Set Node = View.Nodes(Fixture.FullName)
        
    Assert.That Node.Text, Iz.EqualTo("Fixture")
    Assert.That View, modMain.NewTreeNodePathConstraint("Root.Fixture")
End Sub

Public Sub TestAddFixtureWithTests()
    Dim Fixture As FakeTest
    Set Fixture = NewTest("Fixture", Nothing)
    Call Fixture.Tests.Add(NewTest("ATest", Fixture))
    Call Fixture.Tests.Add(NewTest("AnotherTest", Fixture))
    Call Fixture.Tests.Add(NewTest("AThirdTest", Fixture))
    Call Suite.Add(Fixture)
    Call Controller.AddTest(Fixture)

    Assert.AreEqual 5, View.Nodes.Count
    Assert.That View, modMain.NewTreeNodePathConstraint("Root.Fixture.ATest")
    Assert.That View, modMain.NewTreeNodePathConstraint("Root.Fixture.AnotherTest")
    Assert.That View, modMain.NewTreeNodePathConstraint("Root.Fixture.AThirdTest")
End Sub


