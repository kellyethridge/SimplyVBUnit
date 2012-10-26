VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MBStaticClasses"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'
' MBStaticClasses
'
Option Explicit

Public Function MB() As MBConstructors
    Set MB = modStaticClasses.MB
End Function

Public Function Category() As CategoryStatic
    Set Category = modStaticClasses.Category
End Function

Public Function Movie() As MovieStatic
    Set Movie = modStaticClasses.Movie
End Function
