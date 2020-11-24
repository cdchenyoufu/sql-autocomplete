VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIntellisenseItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public alias As String         'always lower case! (maybe enforce with property get set)
Public raw_objName As String   'case sensitive
Public methods As String       'space seperated


Function objName(Optional lower As Boolean = True) As String
    If lower Then objName = LCase(raw_objName) Else objName = raw_objName
End Function
