VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FixtureWithTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ICategorizable

Public Sub ATest()

End Sub

Public Sub AnotherTest()
    Assert.Ignore "I'm testing here"
End Sub

Public Sub AThirdTest()
    Dim s1 As String
    Dim s2 As String
    
    s1 = String$(100, "a")
    s2 = String$(10, "a")
    Mid$(s1, 40, 1) = "A"
    
    Assert.That s2, Iz.EqualTo(s1)
End Sub

Private Sub ICategorizable_Categorize(ByVal Categories As SimplyVBUnit.CategoryCollector)
    Categories.AddTestCategories "ATest", "Quick Tests"
    Categories.AddTestCategories "AnotherTest", "Quick Tests"
End Sub
