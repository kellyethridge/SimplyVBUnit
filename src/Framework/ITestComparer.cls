VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITestComparer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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
'    Module: ITestComparer
'

''
' An interface used to implement a comparer for sorting tests.
'
Option Explicit

''
' Requests the two tests to be compared.
'
' @param X The test on the left side of the equation.
' @param Y The test on the right side of the equation.
' @return Returns the comparison result.
' @remarks If the X test is less than the Y test, then a negative value should be returned.
' If the X test is greater than the Y test, then a positive value should be returned.
' If both tests are equal, then 0 is returned.
'
Public Function Compare(ByVal X As ITest, ByVal Y As ITest) As Long: End Function

