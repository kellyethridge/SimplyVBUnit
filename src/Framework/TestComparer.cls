VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'    CopyRight (c) 2008 Kelly Ethridge
'
'    This file is part of SimplyVBUnit.
'
'    SimplyVBUnit is free software; you can redistribute it and/or modify
'    it under the terms of the GNU Library General Public License as published by
'    the Free Software Foundation; either version 2.1 of the License, or
'    (at your option) any later version.
'
'    SimplyVBUnit is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Library General Public License for more details.
'
'    You should have received a copy of the GNU Library General Public License
'    along with Foobar; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'    Module: TestComparer
'

''
' A comparer that will compare two tests by their names.
'
' @remarks The name comparison is not case-sensitive.
'
Option Explicit
Implements ITestComparer


''
' Compares two tests.
'
' @param X One of the tests to be compared.
' @param Y The other test to be compared.
' @return Returns the comparison of the two tests. Negative value means that X is less
' than Y. Positive value means X is greater than Y. Zero means both tests are equal.
'
Public Function Compare(ByVal X As ITest, ByVal Y As ITest) As Long
    Compare = StrComp(X.Name, Y.Name, vbTextCompare)
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ITestComparer Interface
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ITestComparer_Compare(ByVal X As ITest, ByVal Y As ITest) As Long
    ITestComparer_Compare = Compare(X, Y)
End Function
