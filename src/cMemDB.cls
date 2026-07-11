VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cMemDB - convenience wrapper around an in-memory database. Creates its
' own :memory: cConnection on construction (replaceable via Set Cnn) and
' delegates data access to it, adding table-oriented helpers (GetTable,
' GetSum/GetAvg/GetMin/GetMax/GetCount, CreateTableFromRs).
'=========================================================================
Option Explicit

Private m_oCnn                      As cConnection

Public Property Get Cnn() As cConnection
    Set Cnn = m_oCnn
End Property

Public Property Set Cnn(RHS As cConnection)
    Set m_oCnn = RHS
End Property

Public Function GetRs(ByVal SQL As String, ParamArray P() As Variant) As cRecordset
    Dim aParams()       As Variant

    aParams = P
    Set GetRs = m_oCnn.frGetRs(SQL, aParams)
End Function

Public Sub CreateTable(TableName As String, Optional FieldDefs As Collection)
    m_oCnn.CreateTable TableName, FieldDefs
End Sub

Public Function NewFieldDefs() As Collection
    Set NewFieldDefs = New Collection
End Function

Public Function GetTable(TableName As String, Optional ByVal WhereClause As String, Optional ByVal OrderByClause As String) As cRecordset
    Set GetTable = m_oCnn.OpenRecordset(pvTableSql("*", TableName, WhereClause, OrderByClause))
End Function

Public Function GetSingleVal(SQL As String) As Variant
    Dim oRs             As cRecordset

    Set oRs = m_oCnn.OpenRecordset(SQL)
    If oRs.RecordCount > 0 Then
        GetSingleVal = oRs.ValueMatrix(0, 0)
    Else
        GetSingleVal = Null
    End If
End Function

Public Function GetSum(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
    GetSum = GetSingleVal(pvTableSql("SUM(" & QuoteIdentifier(Field) & ")", TableName, WhereClause, vbNullString))
End Function

Public Function GetAvg(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
    GetAvg = GetSingleVal(pvTableSql("AVG(" & QuoteIdentifier(Field) & ")", TableName, WhereClause, vbNullString))
End Function

Public Function GetMin(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
    GetMin = GetSingleVal(pvTableSql("MIN(" & QuoteIdentifier(Field) & ")", TableName, WhereClause, vbNullString))
End Function

Public Function GetMax(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
    GetMax = GetSingleVal(pvTableSql("MAX(" & QuoteIdentifier(Field) & ")", TableName, WhereClause, vbNullString))
End Function

Public Function GetCount(TableName As String, Optional ByVal Field As String, Optional ByVal WhereClause As String) As Variant
    Dim sExpr           As String

    If Len(Field) = 0 Then
        sExpr = "COUNT(*)"
    Else
        sExpr = "COUNT(" & QuoteIdentifier(Field) & ")"
    End If
    GetCount = GetSingleVal(pvTableSql(sExpr, TableName, WhereClause, vbNullString))
End Function

Public Sub Exec(SQL As String)
    m_oCnn.Execute SQL
End Sub

Public Sub ExecCmd(ByVal SQL As String, ParamArray P() As Variant)
    Dim aParams()       As Variant

    aParams = P
    m_oCnn.frExecCmd SQL, aParams
End Sub

Public Function CreateCommand(SQL As String) As cCommand
    Set CreateCommand = m_oCnn.CreateCommand(SQL)
End Function

Public Function CreateSelectCommand(SQL As String) As cSelectCommand
    Set CreateSelectCommand = m_oCnn.CreateSelectCommand(SQL)
End Function

Public Sub BeginTrans(Optional SavePointName As String)
    m_oCnn.BeginTrans SavePointName
End Sub

Public Sub CommitTrans(Optional SavePointName As String)
    m_oCnn.CommitTrans SavePointName
End Sub

Public Sub RollbackTrans(Optional ToSavePointName As String, Optional ByVal ReUseSavePointForNewAttempts As Boolean)
    m_oCnn.RollbackTrans ToSavePointName, ReUseSavePointForNewAttempts
End Sub

Public Sub CreateTableFromRs(AdoOrSQLiteRsOrByteContent As Variant, ByVal DstTableName As String, Optional ByVal WithPrimaryKeys As Boolean)
    Dim oRs             As cRecordset
    Dim baContent()     As Byte

    If IsObject(AdoOrSQLiteRsOrByteContent) Then
        If TypeOf AdoOrSQLiteRsOrByteContent Is cRecordset Then
            CreateTableFromRecordset m_oCnn, AdoOrSQLiteRsOrByteContent, DstTableName, False, WithPrimaryKeys
            Exit Sub
        End If
    ElseIf IsArray(AdoOrSQLiteRsOrByteContent) Then
        baContent = AdoOrSQLiteRsOrByteContent
        Set oRs = New cRecordset
        oRs.Content = baContent
        CreateTableFromRecordset m_oCnn, oRs, DstTableName, False, WithPrimaryKeys
        Exit Sub
    End If
    Err.Raise vbObjectError, "cMemDB.CreateTableFromRs", "Only cRecordset and Content byte-array sources are implemented yet"
End Sub

Public Function CreateCursor(SQL As String) As cCursor
    Set CreateCursor = m_oCnn.CreateCursor(SQL)
End Function

Private Function pvTableSql(sSelect As String, sTable As String, sWhere As String, sOrderBy As String) As String
    pvTableSql = "SELECT " & sSelect & " FROM " & QuoteIdentifier(sTable)
    If Len(sWhere) > 0 Then
        pvTableSql = pvTableSql & " WHERE " & sWhere
    End If
    If Len(sOrderBy) > 0 Then
        pvTableSql = pvTableSql & " ORDER BY " & sOrderBy
    End If
End Function

Private Sub Class_Initialize()
    Set m_oCnn = New cConnection
    m_oCnn.CreateNewDB ":memory:"
End Sub
