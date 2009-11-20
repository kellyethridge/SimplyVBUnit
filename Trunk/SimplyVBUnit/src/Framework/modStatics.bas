Attribute VB_Name = "modStatics"
' Copyright 2009 Kelly Ethridge
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' Module: modStatics
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

Public Sim              As New SVBUnitConstructors
Public TestUtils        As New TestUtils
Public ErrorInfo        As New ErrorInfoStatic
Public NullMethod       As NullMethodStatic
Public Resource         As New ResourceStatic
Public Timing           As New TimingStatic
Public TestCaseBuilder  As New TestCaseBuilder
Public Assert           As New Assertions
Public NullContext      As NullContextStatic
Public NullListener     As NullListenerStatic

Private mIz As New IzStatic


Public Property Get Iz() As IzSyntaxHelper
    Set Iz = mIz
End Property

