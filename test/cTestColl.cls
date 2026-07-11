VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- test collation: REVCOL compares the reversed strings
Option Explicit

Implements ICollation

Private Property Get ICollation_DefinedNames() As String
    ICollation_DefinedNames = "REVCOL"
End Property

Private Function ICollation_CallbackCollate(ByVal ZeroBasedNameIndex As Long, S1 As String, S2 As String) As Long
    ICollation_CallbackCollate = StrComp(StrReverse(S1), StrReverse(S2), vbBinaryCompare)
End Function
