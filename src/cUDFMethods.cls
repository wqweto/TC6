VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
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
'=========================================================================
' cUDFMethods - passed into IFunction/IAggregateFunction callbacks to read
' the invocation arguments (Get*, 1-based) and set the result (SetResult*).
' Wraps the sqlite3_value*/sqlite3_result_* API over the live
' sqlite3_context; only valid for the duration of the callback.
'=========================================================================
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private m_hCtx                      As LongPtr
Private m_lArgc                     As Long
Private m_lpArgv                    As LongPtr

Public Property Get context() As Long
    context = m_hCtx
End Property

Public Function GetType(ByVal ParamIdxOneBased As Long) As FieldType
    GetType = stub_sqlite3_value_type(pvValue(ParamIdxOneBased))
End Function

Public Function GetInt32(ByVal ParamIdxOneBased As Long) As Long
    GetInt32 = stub_sqlite3_value_int(pvValue(ParamIdxOneBased))
End Function

Public Function GetInt64(ByVal ParamIdxOneBased As Long) As Variant
#If Win64 Then
    GetInt64 = Int64Variant(CDec(stub_sqlite3_value_int64(pvValue(ParamIdxOneBased))))
#Else
    '--- recover the true int64: x86 returns raw bits as Currency (value*10000)
    GetInt64 = Int64Variant(CDec(stub_sqlite3_value_int64(pvValue(ParamIdxOneBased))) * CDec(10000))
#End If
End Function

Public Function GetDouble(ByVal ParamIdxOneBased As Long) As Double
    GetDouble = stub_sqlite3_value_double(pvValue(ParamIdxOneBased))
End Function

Public Function GetBlob(ByVal ParamIdxOneBased As Long) As Byte()
    Dim lpValue         As LongPtr
    Dim lLen            As Long
    Dim baRetVal()      As Byte

    lpValue = pvValue(ParamIdxOneBased)
    lLen = stub_sqlite3_value_bytes(lpValue)
    If lLen > 0 Then
        ReDim baRetVal(0 To lLen - 1) As Byte
        Call CopyMemory(baRetVal(0), ByVal stub_sqlite3_value_blob(lpValue), lLen)
    End If
    GetBlob = baRetVal
End Function

Public Function GetText(ByVal ParamIdxOneBased As Long) As String
    GetText = FromUtf8Ptr(stub_sqlite3_value_text(pvValue(ParamIdxOneBased)))
End Function

Public Function GetTextUTF8Ptr(ByVal ParamIdxOneBased As Long) As Long
    GetTextUTF8Ptr = stub_sqlite3_value_text(pvValue(ParamIdxOneBased))
End Function

Public Function GetUTF8Bytes(ByVal ParamIdxOneBased As Long) As Long
    GetUTF8Bytes = stub_sqlite3_value_bytes(pvValue(ParamIdxOneBased))
End Function

Public Function GetDateTime(ByVal ParamIdxOneBased As Long) As Date
    Dim lpValue         As LongPtr

    lpValue = pvValue(ParamIdxOneBased)
    Select Case stub_sqlite3_value_type(lpValue)
    Case sqlite3win32helper.SQLITE_TEXT
        GetDateTime = pvParseIsoDate(FromUtf8Ptr(stub_sqlite3_value_text(lpValue)))
    Case Else
        GetDateTime = CDate(stub_sqlite3_value_double(lpValue))
    End Select
End Function

Public Sub SetResultError(ErrorString As String)
    Call stub_sqlite3_result_error16(m_hCtx, StrPtr(ErrorString), -1)
End Sub

Public Sub SetResultToParam(ByVal ParamIdxOneBased As Long)
    Call stub_sqlite3_result_value(m_hCtx, pvValue(ParamIdxOneBased))
End Sub

Public Sub SetResultNull()
    Call stub_sqlite3_result_null(m_hCtx)
End Sub

Public Sub SetResultInt32(ByVal Value As Long)
    Call stub_sqlite3_result_int(m_hCtx, Value)
End Sub

Public Sub SetResultInt64(Value As Variant)
#If Win64 Then
    Call stub_sqlite3_result_int64(m_hCtx, CLngLng(Value))
#Else
    '--- the x86 declare types the int64 as Currency (raw bits = value*10000)
    Call stub_sqlite3_result_int64(m_hCtx, CCur(CDec(Value) / 10000))
#End If
End Sub

Public Sub SetResultDouble(ByVal Value As Double)
    Call stub_sqlite3_result_double(m_hCtx, Value)
End Sub

Public Sub SetResultText(Value As String)
    Dim baText()        As Byte

    baText = ToUtf8Array(Value)
    If pvArrayLen(baText) > 0 Then
        Call stub_sqlite3_result_text(m_hCtx, VarPtr(baText(0)), pvArrayLen(baText), SQLITE_TRANSIENT)
    Else
        Call stub_sqlite3_result_text(m_hCtx, StrPtr(Value), 0, SQLITE_TRANSIENT)
    End If
End Sub

Public Sub SetResultUTF8Ptr(ByVal pUTF8 As Long, ByVal nBytes As Long)
    Call stub_sqlite3_result_text(m_hCtx, pUTF8, nBytes, SQLITE_TRANSIENT)
End Sub

Public Sub SetResultBlob(B() As Byte)
    If pvArrayLen(B) > 0 Then
        Call stub_sqlite3_result_blob(m_hCtx, VarPtr(B(0)), pvArrayLen(B), SQLITE_TRANSIENT)
    Else
        Call stub_sqlite3_result_zeroblob(m_hCtx, 0)
    End If
End Sub

Public Sub SetResultBlobPtr(ByVal pBlob As Long, ByVal nBytes As Long)
    Call stub_sqlite3_result_blob(m_hCtx, pBlob, nBytes, SQLITE_TRANSIENT)
End Sub

Public Sub SetResultDate(ByVal D As Date)
    SetResultText Format$(D, "yyyy-mm-dd hh:nn:ss")
End Sub

Public Sub SetResultShortDate(ByVal D As Date)
    SetResultText Format$(D, "yyyy-mm-dd")
End Sub

Public Sub SetResultTime(ByVal D As Date)
    SetResultText Format$(D, "hh:nn:ss")
End Sub

Friend Sub frInit(ByVal hCtx As LongPtr, ByVal lArgc As Long, ByVal lpArgv As LongPtr)
    m_hCtx = hCtx
    m_lArgc = lArgc
    m_lpArgv = lpArgv
End Sub

Private Function pvValue(ByVal lIdx As Long) As LongPtr
    Dim lpValue         As LongPtr

    If lIdx < 1 Or lIdx > m_lArgc Then
        Err.Raise 9, "cUDFMethods", "Parameter index out of range"
    End If
    Call CopyMemory(lpValue, ByVal m_lpArgv + (lIdx - 1) * LenB(lpValue), LenB(lpValue))
    pvValue = lpValue
End Function

Private Function pvArrayLen(baBuf() As Byte) As Long
    On Error GoTo QH
    pvArrayLen = UBound(baBuf) - LBound(baBuf) + 1
QH:
End Function

Private Function pvParseIsoDate(sText As String) As Date
    On Error GoTo EH
    '--- ISO formats produced by SetResultDate/GetDateString and friends
    Select Case Len(sText)
    Case 19
        pvParseIsoDate = DateSerial(Val(Left$(sText, 4)), Val(Mid$(sText, 6, 2)), Val(Mid$(sText, 9, 2))) + _
            TimeSerial(Val(Mid$(sText, 12, 2)), Val(Mid$(sText, 15, 2)), Val(Mid$(sText, 18, 2)))
    Case 10
        pvParseIsoDate = DateSerial(Val(Left$(sText, 4)), Val(Mid$(sText, 6, 2)), Val(Mid$(sText, 9, 2)))
    Case 8
        pvParseIsoDate = TimeSerial(Val(Left$(sText, 2)), Val(Mid$(sText, 4, 2)), Val(Mid$(sText, 7, 2)))
    Case Else
        pvParseIsoDate = CDate(sText)
    End Select
    Exit Function
EH:
End Function
