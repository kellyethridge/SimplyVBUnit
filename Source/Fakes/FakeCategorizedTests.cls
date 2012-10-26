VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FakeCategorizedTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements ICategorizable


Private Sub ICategorizable_Categorize(ByVal Cat As CategoryCollector)
    Cat.ForFixture "one", "two"
    Cat.ForTest "Test1", "three"
End Sub

Public Sub Test1()

End Sub

Public Sub Test2()

End Sub
