VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestFuncs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- test scalar UDFs: REV reverses text, TWICE doubles an int, ERRF raises
Option Explicit

Implements IFunction

Private Property Get IFunction_DefinedNames() As String
    IFunction_DefinedNames = "REV,TWICE,ERRF"
End Property

Private Sub IFunction_Callback(ByVal ZeroBasedNameIndex As Long, ByVal ParamCount As Long, UDF As cUDFMethods)
    Select Case ZeroBasedNameIndex
    Case 0
        UDF.SetResultText StrReverse(UDF.GetText(1))
    Case 1
        UDF.SetResultInt32 UDF.GetInt32(1) * 2
    Case 2
        UDF.SetResultError "boom"
    End Select
End Sub
