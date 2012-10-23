Attribute VB_Name = "modStatics"
'The MIT License (MIT)
'Copyright (c) 2012 Kelly Ethridge
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights to
'use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
'the Software, and to permit persons to whom the Software is furnished to do so,
'subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
'INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
'PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
'FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
'OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'DEALINGS IN THE SOFTWARE.
'
'
' Module: SComponent.modStatics
'
Option Explicit

Public Type NullMethodStatic
    Instance As New NullMethod
End Type

Public Type NullContextStatic
    Instance As New NullContextMethods
End Type

Public Type NullListenerStatic
    Instance As New NullListener
End Type

Public Type TallyStatic
    Zero As New ZeroTally
End Type

Public Type NullTestRunnerStatic
    Instance As New NullTestRunner
End Type

Public Type NullTextWriterStatic
    Instance As New NullTextWriter
End Type

Public NullContext          As NullContextStatic
Public NullListener         As NullListenerStatic
Public NullMethod           As NullMethodStatic
Public Tally                As TallyStatic
Public NullTestRunner       As NullTestRunnerStatic
Public NullTextWriter       As NullTextWriterStatic

Public Sim                  As New SimConstructors
Public Error                As New ErrorHelper
Public TestUtils            As New TestUtils
Public ErrorInfo            As New ErrorInfoStatic
Public Resource             As New ResourceStatic
Public Timing               As New TimingStatic
Public TestCaseBuilder      As New TestCaseBuilder
Public Assert               As New Assertions
Public TestContext          As New TestContextStatic
Public TestContextManager   As New TestContextManager
Public TestFilter           As New TestFilterStatic
Public MsgUtils             As New MsgUtils
Public GlobalSettings       As New GlobalSettings
Public Tolerance            As New ToleranceStatic
Public mHas                 As New HasStatic

Public mIz                  As New IzStatic

Public Property Get Iz() As IzSyntaxHelper
    Set Iz = mIz
End Property

