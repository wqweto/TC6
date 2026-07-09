VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get ActiveConnection() As cConnection
End Property

Public Property Get StepCounter() As Long
End Property

Public Property Get StmtHdl() As Long
End Property

Public Property Get SQL() As String
End Property

Public Property Let SQL(SQL As String)
End Property

Public Property Get NameToIdx(ParamName As String) As Long
End Property

Public Property Get ColCount() As Long
End Property

Public Property Get ColName(ByVal IdxZeroBased As Long) As String
End Property

Public Property Get ColType(ByVal IdxZeroBased As Long) As Long
End Property

Public Property Get ColVal(ByVal IdxZeroBased As Long) As Variant
End Property

Public Function Step() As Boolean
End Function

Public Sub Reset()
End Sub

Public Sub SetText(ByVal ParamIdxOneBased As Long, S As String)
End Sub

Public Sub SetTextPtr(ByVal ParamIdxOneBased As Long, ByVal pS As Long, ByVal LenB_S As Long)
End Sub

Public Sub SetTextUTF8Ptr(ByVal ParamIdxOneBased As Long, ByVal pU As Long, ByVal UByteLen As Long)
End Sub

Public Sub SetBlob(ByVal ParamIdxOneBased As Long, B() As Byte)
End Sub

Public Sub SetBlobPtr(ByVal ParamIdxOneBased As Long, ByVal pb As Long, ByVal BByteLen As Long)
End Sub

Public Sub SetInt32(ByVal ParamIdxOneBased As Long, ByVal Int32 As Long)
End Sub

Public Sub SetInt64(ByVal ParamIdxOneBased As Long, Int64 As Variant)
End Sub

Public Sub SetDouble(ByVal ParamIdxOneBased As Long, ByVal D As Double)
End Sub

Public Sub SetDate(ByVal ParamIdxOneBased As Long, ByVal D As Date)
End Sub

Public Sub SetShortDate(ByVal ParamIdxOneBased As Long, ByVal D As Date)
End Sub

Public Sub SetTime(ByVal ParamIdxOneBased As Long, ByVal D As Date)
End Sub

Public Sub SetBoolean(ByVal ParamIdxOneBased As Long, ByVal B As Boolean)
End Sub

Public Sub SetNull(ByVal ParamIdxOneBased As Long)
End Sub

Public Sub SetAllParamsNull()
End Sub

Public Function IdxToName(ByVal OneBasedIdx As Long) As String
End Function

Public Function ParameterCount() As Long
End Function
