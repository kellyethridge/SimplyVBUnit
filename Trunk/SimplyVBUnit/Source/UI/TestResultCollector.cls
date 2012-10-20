VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestResultCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2009 Kelly Ethridge
'
'    This file is part of SimplyVBUnitUI.
'
'    SimplyVBUnitUI is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnitUI is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: TestResultCollector
'

Option Explicit

Private WithEvents mListener As EventCastListener
Attribute mListener.VB_VarHelpID = -1
Private mResults As Collection


Public Property Get Item(ByVal Index As Variant) As TestResult
Attribute Item.VB_UserMemId = 0
    On Error GoTo errTrap
    Set Item = mResults(Index)
errTrap:
End Property

Friend Sub Init(ByVal Listener As EventCastListener)
    Set mListener = Listener
End Sub

Private Sub mListener_RunStarted(ByVal Name As String, ByVal TestCount As Long)
    Set mResults = New Collection
End Sub

Private Sub mListener_TestCaseFinished(ByVal Result As SimplyVBUnit.TestResult)
    Call mResults.Add(Result, Result.Test.FullName)
End Sub

Private Sub mListener_TestFixtureFinished(ByVal Result As SimplyVBUnit.TestResult)
    Call mResults.Add(Result, Result.Test.FullName)
End Sub

Private Sub mListener_TestSuiteFinished(ByVal Result As SimplyVBUnit.TestResult)
    Call mResults.Add(Result, Result.Test.FullName)
End Sub
