VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TreeNodePathConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IConstraint

Private mExpected   As String
Private mActual     As String

Public Sub Init(ByVal Expected As String)
    mExpected = Expected
End Sub

Private Function IConstraint_Matches(Actual As Variant) As Boolean
    Dim View As TreeView
    Set View = Actual
    
    Dim Node As Node
    Set Node = View.Nodes(GetPath(1))
    
    If Node Is Nothing Then
        Exit Function
    End If
    
    Dim i As Long
    For i = 2 To GetTotalNodeCount
        Dim Key As String
        Key = GetPath(i)
        
        mActual = Node.Key
        
        Set Node = GetChildNode(Node, Key)
        If Node Is Nothing Then
            Exit Function
        End If
    Next i
    
    IConstraint_Matches = True
End Function

Private Sub IConstraint_WriteActualValueTo(ByVal Writer As TextMessageWriter)

End Sub

Private Sub IConstraint_WriteDescriptionTo(ByVal Writer As TextMessageWriter)
    
End Sub

Private Sub IConstraint_WriteMessageTo(ByVal Writer As TextMessageWriter)
    Call Writer.DisplayDifferences(mExpected, mActual)
End Sub


Private Function GetChildNode(ByVal Parent As Node, ByVal Key As String) As Node
    Dim Result As Node
    Dim Child As Node
    
    If Parent.Children > 0 Then
        Set Child = Parent.Child
        Do While Not Child Is Nothing
            If Child.Key = Key Then
                Set Result = Child
                Exit Do
            End If
            
            Set Child = Child.Next
        Loop
    End If
    
    Set GetChildNode = Result
End Function

Private Function GetPath(ByVal Hops As Long) As String
    Dim i As Long
    i = 1
    Do While Hops > 0
        i = InStr(i + 1, mExpected, ".")
        
        If i = 0 Then
            i = Len(mExpected) + 1
            Exit Do
        End If
        
        Hops = Hops - 1
    Loop
    
    GetPath = Left$(mExpected, i - 1)
End Function

Private Function GetTotalNodeCount() As Long
    Dim s() As String
    s = Split(mExpected, ".")
    GetTotalNodeCount = UBound(s) + 1
End Function
