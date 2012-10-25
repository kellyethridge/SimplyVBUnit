VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringAssertions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' StringAssertions
'
' This module shows many of the ways assertions can be made against string values.
' Not all assertions will pass. This is just to demonstrate how to perform assertions with strings.
' Where possible, examples using straight assertions and constraint assertions are performed to
' demonstrate how to perform assertions with different methods.
'
Option Explicit


Public Sub Equality_WithIdenticalValues_Passes()
    Const Actual    As String = "String Value"
    Const Expected  As String = "String Value"
    
    Assert.That Actual, Iz.EqualTo(Expected)
    Assert.AreEqual Expected, Actual
    Assert.AreEqualStrings Expected, Actual
End Sub

Public Sub Equality_WithNonEqualStrings_Fails()
    Const Actual    As String = "String Value"
    Const Expected  As String = "Not String Value"
    
    Assert.That Actual, Iz.EqualTo(Expected), "This should fail."
End Sub

Public Sub Equality_WithDifferingCase_Fails()
    Const Actual    As String = "string value" ' all lower case
    Const Expected  As String = "String Value"

    Assert.That Actual, Iz.EqualTo(Expected), "This should fail."
End Sub

Public Sub CaseInsensitiveEquality_WithDifferingCase_Passes()
    Const Actual    As String = "string value" ' all lower case
    Const Expected  As String = "String Value"

    Assert.That Actual, Iz.EqualTo(Expected).IgnoreCase, "This should pass."
    Assert.AreEqualStrings Expected, Actual, True, "This should pass."
End Sub

Public Sub StartsWith_WithIdenticalStrings_Passes()
    Const Actual    As String = "String Value"
    Const Expected  As String = "String"

    Assert.That Actual, Iz.StringStarting(Expected)
End Sub

Public Sub StartsWith_WithDifferingCase_Fails()
    Const Actual    As String = "String Value"
    Const Expected  As String = "string"        ' lower case

    Assert.That Actual, Iz.StringStarting(Expected), "This should fail."
End Sub

Public Sub CaseInsensitiveStartsWith_WithDifferingCase_Passes()
    Const Actual    As String = "String Value"
    Const Expected  As String = "string"        ' lower case
    
    Assert.That Actual, Iz.StringStarting(Expected).IgnoreCase, "This should pass."
End Sub

Public Sub Substring_WithStringContainingSubstring_Passes()
    Const Actual    As String = "String Value"
    Const Expected  As String = "ing"

    Assert.That Actual, Iz.StringContaining(Expected)
End Sub

Public Sub IsEmpty_WithEmptyString_Passes()
    Const Actual    As String = ""
    
    Assert.That Actual, Iz.Empty
    Assert.IsEmpty Actual
End Sub

'
' SimplyVBUnit uses the VBScript Regular Expression engine, not the VB Like keyword.
'
Public Sub Matching_WithStringContainingPattern_Passes()
    Const Actual    As String = "abc123def"
    Const Pattern   As String = "\d\d\d"    ' must contain at least 3 consecutive digits.
    
    Assert.That Actual, Iz.StringMatching(Pattern)
End Sub
