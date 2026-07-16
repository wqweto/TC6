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
'=========================================================================
' cCursor - forward-only low-level cursor over a prepared statement.
' Created by cConnection.CreateCursor; bind parameters with the Set*
' methods, call Step until it returns False, read the current row through
' the Col* accessors, Reset to run again.
'=========================================================================
Option Explicit

Private m_oConnection               As cConnection
Private m_hStmt                     As LongPtr
Private m_sSql                      As String
Private m_lStepCounter              As Long

Public Property Get ActiveConnection() As cConnection
    Set ActiveConnection = m_oConnection
End Property

Public Property Get StepCounter() As Long
    StepCounter = m_lStepCounter
End Property

Public Property Get StmtHdl() As Long
    StmtHdl = m_hStmt
End Property

Public Property Get SQL() As String
    SQL = m_sSql
End Property

Public Property Let SQL(SQL As String)
    pvPrepare SQL
End Property

Public Property Get NameToIdx(ParamName As String) As Long
    NameToIdx = StmtParamIndex(m_hStmt, ParamName)
End Property

Public Property Get ColCount() As Long
    ColCount = stub_sqlite3_column_count(m_hStmt)
End Property

Public Property Get ColName(ByVal IdxZeroBased As Long) As String
    ColName = FromUtf8Ptr(stub_sqlite3_column_name(m_hStmt, IdxZeroBased))
End Property

Public Property Get ColType(ByVal IdxZeroBased As Long) As Long
    ColType = stub_sqlite3_column_type(m_hStmt, IdxZeroBased)
End Property

Public Property Get ColVal(ByVal IdxZeroBased As Long) As Variant
    ColVal = ReadColumnValue(m_hStmt, IdxZeroBased)
End Property

Public Function Step() As Boolean
    Select Case stub_sqlite3_step(m_hStmt)
    Case SQLITE_ROW
        m_lStepCounter = m_lStepCounter + 1
        Step = True
    Case SQLITE_DONE
        '--- end of result set
    Case Else
        Err.Raise vbObjectError, "cCursor.Step", m_oConnection.LastDBError()
    End Select
End Function

Public Sub Reset()
    Call stub_sqlite3_reset(m_hStmt)
    m_lStepCounter = 0
End Sub

Public Sub SetText(ByVal ParamIdxOneBased As Long, S As String)
    pvCheck BindTextValue(m_hStmt, ParamIdxOneBased, S)
End Sub

Public Sub SetTextPtr(ByVal ParamIdxOneBased As Long, ByVal pS As Long, ByVal LenB_S As Long)
    pvCheck stub_sqlite3_bind_text16(m_hStmt, ParamIdxOneBased, pS, LenB_S, SQLITE_TRANSIENT)
End Sub

Public Sub SetTextUTF8Ptr(ByVal ParamIdxOneBased As Long, ByVal pU As Long, ByVal UByteLen As Long)
    pvCheck stub_sqlite3_bind_text(m_hStmt, ParamIdxOneBased, pU, UByteLen, SQLITE_TRANSIENT)
End Sub

Public Sub SetBlob(ByVal ParamIdxOneBased As Long, B() As Byte)
    pvCheck BindBlobValue(m_hStmt, ParamIdxOneBased, B)
End Sub

Public Sub SetBlobPtr(ByVal ParamIdxOneBased As Long, ByVal pb As Long, ByVal BByteLen As Long)
    pvCheck stub_sqlite3_bind_blob(m_hStmt, ParamIdxOneBased, pb, BByteLen, SQLITE_TRANSIENT)
End Sub

Public Sub SetInt32(ByVal ParamIdxOneBased As Long, ByVal Int32 As Long)
    pvCheck stub_sqlite3_bind_int(m_hStmt, ParamIdxOneBased, Int32)
End Sub

Public Sub SetInt64(ByVal ParamIdxOneBased As Long, Int64 As Variant)
    pvCheck BindInt64Value(m_hStmt, ParamIdxOneBased, Int64)
End Sub

Public Sub SetDouble(ByVal ParamIdxOneBased As Long, ByVal D As Double)
    pvCheck stub_sqlite3_bind_double(m_hStmt, ParamIdxOneBased, D)
End Sub

Public Sub SetDate(ByVal ParamIdxOneBased As Long, ByVal D As Date)
    pvCheck BindTextValue(m_hStmt, ParamIdxOneBased, Format$(D, "yyyy-mm-dd hh:nn:ss"))
End Sub

Public Sub SetShortDate(ByVal ParamIdxOneBased As Long, ByVal D As Date)
    pvCheck BindTextValue(m_hStmt, ParamIdxOneBased, Format$(D, "yyyy-mm-dd"))
End Sub

Public Sub SetTime(ByVal ParamIdxOneBased As Long, ByVal D As Date)
    pvCheck BindTextValue(m_hStmt, ParamIdxOneBased, Format$(D, "hh:nn:ss"))
End Sub

Public Sub SetBoolean(ByVal ParamIdxOneBased As Long, ByVal B As Boolean)
    pvCheck stub_sqlite3_bind_int(m_hStmt, ParamIdxOneBased, IIf(B, 1, 0))
End Sub

Public Sub SetNull(ByVal ParamIdxOneBased As Long)
    pvCheck stub_sqlite3_bind_null(m_hStmt, ParamIdxOneBased)
End Sub

Public Sub SetAllParamsNull()
    pvCheck stub_sqlite3_clear_bindings(m_hStmt)
End Sub

Public Function IdxToName(ByVal OneBasedIdx As Long) As String
    IdxToName = StmtParamName(m_hStmt, OneBasedIdx)
End Function

Public Function ParameterCount() As Long
    ParameterCount = stub_sqlite3_bind_parameter_count(m_hStmt)
End Function

Friend Sub frInit(oCnn As cConnection, sSql As String)
    Set m_oConnection = oCnn
    pvPrepare sSql
End Sub

Private Sub pvPrepare(sSql As String)
    pvFinalize
    m_hStmt = PrepareStatement(m_oConnection, sSql, IsSelect:=True)
    m_sSql = sSql
    m_lStepCounter = 0
End Sub

Private Sub pvFinalize()
    If m_hStmt <> 0 Then
        Call stub_sqlite3_finalize(m_hStmt)
        m_hStmt = 0
    End If
End Sub

Private Sub pvCheck(ByVal lRc As Long)
    If lRc <> SQLITE_OK Then
        Err.Raise vbObjectError, "cCursor.SetParam", m_oConnection.LastDBError()
    End If
End Sub

Private Sub Class_Terminate()
    pvFinalize
End Sub
