VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSelectCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cSelectCommand - reusable prepared read statement. Created by
' cConnection.CreateSelectCommand; bind parameters with the Set* methods
' (1-based index, or NameToIdx for :name/@name/$name parameters), then
' Execute returns a disconnected cRecordset. The statement is reset after
' Execute so it can be re-bound and re-run.
'=========================================================================
Option Explicit

Private m_oConnection               As cConnection
Private m_hStmt                     As LongPtr
Private m_sSql                      As String

Public Property Get ActiveConnection() As cConnection
    Set ActiveConnection = m_oConnection
End Property

Public Property Get SQL() As String
    SQL = m_sSql
End Property

Public Property Let SQL(SQL As String)
    pvPrepare SQL
End Property

Public Property Get NameToIdx(ByVal ParamName As String) As Long
Attribute NameToIdx.VB_UserMemId = 0
    NameToIdx = StmtParamIndex(m_hStmt, ParamName)
End Property

Public Sub Save(CommandKey As String)
    pvNotImpl "Save"
End Sub

Public Sub ReplColumnOrTableName(ByVal ParamIdxOneBased As Long, S As String)
    pvNotImpl "ReplColumnOrTableName"
End Sub

Public Sub ReplTextBlock(ByVal ParamIdxOneBased As Long, S As String)
    pvNotImpl "ReplTextBlock"
End Sub

Public Sub SetText(ByVal ParamIdxOneBased As Long, S As String)
    pvCheck BindTextValue(m_hStmt, ParamIdxOneBased, S)
End Sub

Public Sub SetBlob(ByVal ParamIdxOneBased As Long, B() As Byte)
    pvCheck BindBlobValue(m_hStmt, ParamIdxOneBased, B)
End Sub

Public Sub SetInt32(ByVal ParamIdxOneBased As Long, ByVal Int32 As Long)
    pvCheck vbsqlite3_bind_int(m_hStmt, ParamIdxOneBased, Int32)
End Sub

Public Sub SetInt64(ByVal ParamIdxOneBased As Long, Int64 As Variant)
    pvCheck BindInt64Value(m_hStmt, ParamIdxOneBased, Int64)
End Sub

Public Sub SetDouble(ByVal ParamIdxOneBased As Long, ByVal D As Double)
    pvCheck vbsqlite3_bind_double(m_hStmt, ParamIdxOneBased, D)
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
    pvCheck vbsqlite3_bind_int(m_hStmt, ParamIdxOneBased, IIf(B, 1, 0))
End Sub

Public Sub SetNull(ByVal ParamIdxOneBased As Long)
    pvCheck vbsqlite3_bind_null(m_hStmt, ParamIdxOneBased)
End Sub

Public Function Execute(Optional ByVal ReadOnly As Boolean) As cRecordset
    Dim oRs             As cRecordset

    Set oRs = New cRecordset
    On Error GoTo EH
    oRs.frLoadStmt m_oConnection, m_sSql, m_hStmt
    Call vbsqlite3_reset(m_hStmt)
    Set Execute = oRs
    Exit Function
EH:
    '--- reset even when materialising fails, so the statement stays reusable
    Call vbsqlite3_reset(m_hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function IdxToName(ByVal OneBasedIdx As Long) As String
    IdxToName = StmtParamName(m_hStmt, OneBasedIdx)
End Function

Public Function ParameterCount() As Long
    ParameterCount = vbsqlite3_bind_parameter_count(m_hStmt)
End Function

Friend Sub frInit(oCnn As cConnection, sSql As String)
    Set m_oConnection = oCnn
    pvPrepare sSql
End Sub

Private Sub pvPrepare(sSql As String)
    pvFinalize
    m_hStmt = PrepareStatement(m_oConnection, sSql)
    m_sSql = sSql
End Sub

Private Sub pvFinalize()
    If m_hStmt <> 0 Then
        Call vbsqlite3_finalize(m_hStmt)
        m_hStmt = 0
    End If
End Sub

Private Sub pvCheck(ByVal lRc As Long)
    If lRc <> SQLITE_OK Then
        Err.Raise vbObjectError, "cSelectCommand", m_oConnection.LastDBError()
    End If
End Sub

Private Sub pvNotImpl(sName As String)
    Err.Raise vbObjectError, "cSelectCommand." & sName, "Not implemented yet"
End Sub

Private Sub Class_Terminate()
    pvFinalize
End Sub
