Attribute VB_Name = "modMain"
Option Explicit

Public Sub Main()
    Call SetClientInfo(App)
    
    Dim Tests As New TestSuite
    Tests.Add New TestTreeViewController
    Tests.Add New TestTreeViewNavigation
    Tests.Add New TestTreeViewRunningTests
    Tests.Add New TestConfiguration
    

    Dim Runner As TestRunner
    Set Runner = Sim.NewTestRunner(Tests)

    Dim Result As TestResult
    Set Result = Runner.Run(New DebugListener)
End Sub

Public Function NewTreeNodePathConstraint(ByVal ExpectedPath As String) As TreeNodePathConstraint
    Set NewTreeNodePathConstraint = New TreeNodePathConstraint
    Call NewTreeNodePathConstraint.Init(ExpectedPath)
End Function

Public Function NewTest(ByVal Name As String, Optional ByVal Parent As ITest) As FakeTest
    Set NewTest = New FakeTest
    NewTest.Name = Name
    Set NewTest.Parent = Parent
End Function

