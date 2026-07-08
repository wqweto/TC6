VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUDFMethods"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get context() As Long
End Property

Public Function GetType(ByVal ParamIdxOneBased As Long) As FieldType
End Function

Public Function GetInt32(ByVal ParamIdxOneBased As Long) As Long
End Function

Public Function GetInt64(ByVal ParamIdxOneBased As Long) As Variant
End Function

Public Function GetDouble(ByVal ParamIdxOneBased As Long) As Double
End Function

Public Function GetBlob(ByVal ParamIdxOneBased As Long) As Byte()
End Function

Public Function GetText(ByVal ParamIdxOneBased As Long) As String
End Function

Public Function GetTextUTF8Ptr(ByVal ParamIdxOneBased As Long) As Long
End Function

Public Function GetUTF8Bytes(ByVal ParamIdxOneBased As Long) As Long
End Function

Public Function GetDateTime(ByVal ParamIdxOneBased As Long) As Date
End Function

Public Sub SetResultError(ErrorString As String)
End Sub

Public Sub SetResultToParam(ByVal ParamIdxOneBased As Long)
End Sub

Public Sub SetResultNull()
End Sub

Public Sub SetResultInt32(ByVal Value As Long)
End Sub

Public Sub SetResultInt64(Value As Variant)
End Sub

Public Sub SetResultDouble(ByVal Value As Double)
End Sub

Public Sub SetResultText(Value As String)
End Sub

Public Sub SetResultUTF8Ptr(ByVal pUTF8 As Long, ByVal nBytes As Long)
End Sub

Public Sub SetResultBlob(B() As Byte)
End Sub

Public Sub SetResultBlobPtr(ByVal pBlob As Long, ByVal nBytes As Long)
End Sub

Public Sub SetResultDate(ByVal D As Date)
End Sub

Public Sub SetResultShortDate(ByVal D As Date)
End Sub

Public Sub SetResultTime(ByVal D As Date)
End Sub
