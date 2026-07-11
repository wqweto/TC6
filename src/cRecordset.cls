VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cRecordset - disconnected recordset over a SQLite result set.
'
' Implemented: load-all-rows on open (prepare/step/column into an in-memory
' matrix), forward/backward navigation, RecordCount/BOF/EOF/AbsolutePosition/
' Bookmark, Fields (cFields/cField), ValueMatrix, GetRows, ReQuery, and the
' batched write surface (field writes, AddNew/Delete/UpdateBatch/
' ResetChanges) - changes are tracked in memory and written back by
' UpdateBatch inside a savepoint, addressing rows by the origin table's
' PK columns (see pvAnalyzeUpdatable). Sort/SortRefresh reorder the matrix
' in memory (multi-key, SQLite type ordering, cursor follows its row) and
' Find* scan it with an ADO-style criterion (see pvParseCriterion).
' Content serialization, ADO interop and JSON export are not implemented
' yet and raise; unimplemented read-only members return defaults.
'
' NOTE: a facade-over-cRowset split (strong references, orphaned fields
' stay readable) was considered and REJECTED - not going to implement. The
' original RC6 raises error 91 on every recordset-dependent cField member
' once the recordset is gone (verified against RC6.dll 3.42.0), which is
' exactly what the weak-ref design here does (see cField.pvRs).
'=========================================================================
Option Explicit

Public Enum ResetEnum
    ResetAll = 0
    ResetDeletesOnly = 1
    ResetInsertsOnly = 2
    ResetUpdatesOnly = 3
End Enum

Public Event FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
Public Event Move(ByVal NewRowIdxZeroBased As Long)
Public Event AddNew(ByVal NewRowIdxZeroBased As Long)
Public Event Delete(ByVal NewRowIdxZeroBased As Long)
Public Event QueryFinished()
Public Event Reset()

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Enum eRowState
    ROW_CLEAN = 0
    ROW_MODIFIED = 1
    ROW_NEW = 2
End Enum

Private Enum eFindOp
    OP_EQ
    OP_NE
    OP_LT
    OP_LE
    OP_GT
    OP_GE
    OP_LIKE
    OP_ISNULL
    OP_ISNOTNULL
End Enum

Private m_oConnection               As cConnection
Private m_oFields                   As cFields
Private m_sSql                      As String
Private m_sSort                     As String
Private m_bMapNullToEmpty           As Boolean
Private m_bReadOnly                 As Boolean
Private m_bAutoCreateUniqueID64     As Boolean
Private m_lFieldCount               As Long
Private m_lRecordCount              As Long
Private m_lCurRow                   As Long
Private m_aFieldName()              As String
Private m_aData()                   As Variant
'--- per-column origin metadata (captured on every open, see cField)
Private m_aOrigColumn()             As String
Private m_aOrigTable()              As String
Private m_aOrigDb()                 As String
Private m_aDeclType()               As String
Private m_cColMeta                  As VBA.Collection
'--- per-column schema carried by a loaded Content blob (see frColSchema)
Private m_aColSchema()              As cColumn
Private m_bHasColSchema             As Boolean
'--- updatability analysis (per open)
Private m_bUpdatable                As Boolean
Private m_sOrigDb                   As String
Private m_sOrigTable                As String
Private m_aPkCols()                 As Long
Private m_lPkCount                  As Long
Private m_lAutoPkCol                As Long
'--- pending-changes tracking (lazy, see pvEnsureWriteState)
Private m_bWriteState               As Boolean
Private m_aOriginal()               As Variant
Private m_aDirty()                  As Boolean
Private m_aRowState()               As Long
Private m_cDeleted                  As VBA.Collection

Public Property Get ContainsChanges() As Boolean
    Dim lRow            As Long

    If Not m_cDeleted Is Nothing Then
        If m_cDeleted.Count > 0 Then
            ContainsChanges = True
            Exit Property
        End If
    End If
    If m_bWriteState Then
        For lRow = 0 To m_lRecordCount - 1
            If m_aRowState(lRow) <> ROW_CLEAN Then
                ContainsChanges = True
                Exit Property
            End If
        Next
    End If
End Property

Public Property Get AutoCreateUniqueID64() As Boolean
    AutoCreateUniqueID64 = m_bAutoCreateUniqueID64
End Property

Public Property Let AutoCreateUniqueID64(ByVal RHS As Boolean)
    m_bAutoCreateUniqueID64 = RHS
End Property

Public Property Get RecordCount() As Long
    RecordCount = m_lRecordCount
End Property

Public Property Get Updatable() As Boolean
    Updatable = m_bUpdatable
End Property

Public Property Get UniqueID64() As Variant
    UniqueID64 = CreateUniqueID64()
End Property

Public Property Get BOF() As Boolean
    BOF = (m_lRecordCount = 0) Or (m_lCurRow < 0)
End Property

Public Property Get EOF() As Boolean
    EOF = (m_lRecordCount = 0) Or (m_lCurRow >= m_lRecordCount)
End Property

Public Property Get AbsolutePosition() As Long
    AbsolutePosition = m_lCurRow
End Property

Public Property Let AbsolutePosition(ByVal RHS As Long)
    If RHS < 0 Then
        m_lCurRow = -1
    ElseIf RHS >= m_lRecordCount Then
        m_lCurRow = m_lRecordCount
    Else
        m_lCurRow = RHS
    End If
    RaiseEvent Move(m_lCurRow)
End Property

Public Property Get Bookmark() As Long
    Bookmark = m_lCurRow
End Property

Public Property Let Bookmark(ByVal RHS As Long)
    AbsolutePosition = RHS
End Property

Public Property Get Sort() As String
    Sort = m_sSort
End Property

Public Property Let Sort(ByVal RHS As String)
    m_sSort = RHS
    pvApplySort
End Property

Public Property Get Fields() As cFields
Attribute Fields.VB_UserMemId = 0
    Set Fields = m_oFields
End Property

Public Property Get SQL() As String
    SQL = m_sSql
End Property

Public Property Let SQL(SQL As String)
    m_sSql = SQL
End Property

Public Property Get ValueMatrix(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long) As Variant
    ValueMatrix = frCellValueAt(RowIdxZeroBased, ColIdxZeroBased)
End Property

Public Property Let ValueMatrix(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long, RHS As Variant)
    frSetCellValueAt RowIdxZeroBased, ColIdxZeroBased, RHS
End Property

Public Property Get ActiveConnection() As cConnection
    Set ActiveConnection = m_oConnection
End Property

Public Property Set ActiveConnection(RHS As cConnection)
    Set m_oConnection = RHS
End Property

Public Property Get ContentChangesOnly() As Byte()
End Property

Public Property Get Content() As Byte()
    Content = pvContentWrite()
End Property

Public Property Let Content(RHS() As Byte)
    pvContentRead RHS
End Property

Public Function UniqueID64ToVBDate(UniqueID64 As Variant, Optional ReturnFracSeconds As Double) As Date
    Dim decSerial       As Variant
    Dim decSeconds      As Variant

    '--- inverse of CreateUniqueID64: value / 10^14 is the VB date serial
    decSerial = CDec(UniqueID64) / CDec("100000000000000")
    UniqueID64ToVBDate = CDate(CDbl(decSerial))
    decSeconds = decSerial * 86400
    ReturnFracSeconds = CDbl(decSeconds - Int(decSeconds))
End Function

Public Sub AddNew()
    Dim lCol            As Long

    pvCheckUpdatable
    pvEnsureWriteState
    pvEnsureRowCapacity m_lRecordCount + 1
    For lCol = 0 To m_lFieldCount - 1
        m_aData(lCol, m_lRecordCount) = pvNullCell()
        m_aDirty(lCol, m_lRecordCount) = False
    Next
    '--- RC6 semantics: fill the INTEGER PK with a time-based unique id
    If m_bAutoCreateUniqueID64 And m_lAutoPkCol >= 0 Then
        m_aData(m_lAutoPkCol, m_lRecordCount) = CreateUniqueID64()
    End If
    m_aRowState(m_lRecordCount) = ROW_NEW
    m_lRecordCount = m_lRecordCount + 1
    m_lCurRow = m_lRecordCount - 1
    RaiseEvent AddNew(m_lCurRow)
End Sub

Public Sub Delete()
    Dim aEntry()        As Variant
    Dim aRow()          As Variant
    Dim aOrig()         As Variant
    Dim aDirty()        As Boolean
    Dim aPk()           As Variant
    Dim lCol            As Long
    Dim lPkIdx          As Long

    pvCheckUpdatable
    If m_lCurRow < 0 Or m_lCurRow >= m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "No current row (BOF or EOF)"
    End If
    pvEnsureWriteState
    '--- a pending insert just vanishes; existing rows queue a DELETE and keep
    '--- a full snapshot so ResetChanges can restore them faithfully
    If m_aRowState(m_lCurRow) <> ROW_NEW Then
        ReDim aEntry(0 To 4)
        ReDim aRow(0 To m_lFieldCount - 1)
        ReDim aOrig(0 To m_lFieldCount - 1)
        ReDim aDirty(0 To m_lFieldCount - 1)
        For lCol = 0 To m_lFieldCount - 1
            aRow(lCol) = m_aData(lCol, m_lCurRow)
            aOrig(lCol) = m_aOriginal(lCol, m_lCurRow)
            aDirty(lCol) = m_aDirty(lCol, m_lCurRow)
        Next
        ReDim aPk(0 To m_lPkCount - 1)
        For lPkIdx = 0 To m_lPkCount - 1
            aPk(lPkIdx) = pvPkWhereValue(m_lCurRow, lPkIdx)
        Next
        aEntry(0) = aRow
        aEntry(1) = aPk
        aEntry(2) = m_aRowState(m_lCurRow)
        aEntry(3) = aDirty
        aEntry(4) = aOrig
        m_cDeleted.Add aEntry
    End If
    pvRemoveRow m_lCurRow
    RaiseEvent Delete(m_lCurRow)
End Sub

Public Sub UpdateBatch()
    Dim vElem           As Variant
    Dim lRow            As Long

    If Not ContainsChanges Then
        Exit Sub
    End If
    pvCheckUpdatable
    m_oConnection.BeginTrans "tc6_updatebatch"
    On Error GoTo EH
    For Each vElem In m_cDeleted
        pvBatchDelete vElem
    Next
    For lRow = 0 To m_lRecordCount - 1
        Select Case m_aRowState(lRow)
        Case ROW_MODIFIED
            pvBatchUpdate lRow
        Case ROW_NEW
            pvBatchInsert lRow
        End Select
    Next
    m_oConnection.CommitTrans
    pvClearWriteState
    Exit Sub
EH:
    m_oConnection.RollbackTrans "tc6_updatebatch"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub MoveFirst()
    If m_lRecordCount > 0 Then
        m_lCurRow = 0
    End If
    RaiseEvent Move(m_lCurRow)
End Sub

Public Sub MoveLast()
    If m_lRecordCount > 0 Then
        m_lCurRow = m_lRecordCount - 1
    End If
    RaiseEvent Move(m_lCurRow)
End Sub

Public Sub MoveNext()
    If m_lCurRow < m_lRecordCount Then
        m_lCurRow = m_lCurRow + 1
    End If
    RaiseEvent Move(m_lCurRow)
End Sub

Public Sub MovePrevious()
    If m_lCurRow >= 0 Then
        m_lCurRow = m_lCurRow - 1
    End If
    RaiseEvent Move(m_lCurRow)
End Sub

Public Sub SortRefresh()
    pvApplySort
End Sub

Public Function FindFirst(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    FindFirst = pvFindMove(0, 1, Criterion, DistinctNullValues)
End Function

Public Function FindNext(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    FindNext = pvFindMove(m_lCurRow + 1, 1, Criterion, DistinctNullValues)
End Function

Public Function FindPrevious(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    FindPrevious = pvFindMove(m_lCurRow - 1, -1, Criterion, DistinctNullValues)
End Function

Public Function FindLast(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    FindLast = pvFindMove(m_lRecordCount - 1, -1, Criterion, DistinctNullValues)
End Function

Public Sub OpenRecordset(Optional SQL As String, Optional Cnn As Variant, Optional ByVal ReadOnly As Boolean)
    If Not IsMissing(Cnn) Then
        If IsObject(Cnn) Then
            Set m_oConnection = Cnn
        End If
    End If
    If Len(SQL) > 0 Then
        m_sSql = SQL
    End If
    m_bReadOnly = ReadOnly
    pvLoad
End Sub

Public Function GetRows(Optional ByVal RowCount As Long = -1, Optional ByVal StartIdxZeroBased As Long, Optional FieldListCommaSeparated As String, Optional ByVal TransPosed As Boolean) As Variant()
    Dim aCol()          As Long
    Dim lNumCols        As Long
    Dim lFirst          As Long
    Dim lLast           As Long
    Dim lNumRows        As Long
    Dim aOut()          As Variant
    Dim lRow            As Long
    Dim lCol            As Long

    aCol = pvResolveColumns(FieldListCommaSeparated, lNumCols)
    lFirst = StartIdxZeroBased
    If lFirst < 0 Then
        lFirst = 0
    End If
    If RowCount < 0 Then
        lLast = m_lRecordCount - 1
    Else
        lLast = lFirst + RowCount - 1
    End If
    If lLast > m_lRecordCount - 1 Then
        lLast = m_lRecordCount - 1
    End If
    lNumRows = lLast - lFirst + 1
    If lNumRows <= 0 Or lNumCols = 0 Then
        Exit Function
    End If
    '--- default orientation matches ADO GetRows: (fieldIdx, rowIdx)
    If TransPosed Then
        ReDim aOut(0 To lNumRows - 1, 0 To lNumCols - 1)
        For lRow = 0 To lNumRows - 1
            For lCol = 0 To lNumCols - 1
                aOut(lRow, lCol) = m_aData(aCol(lCol), lFirst + lRow)
            Next
        Next
    Else
        ReDim aOut(0 To lNumCols - 1, 0 To lNumRows - 1)
        For lCol = 0 To lNumCols - 1
            For lRow = 0 To lNumRows - 1
                aOut(lCol, lRow) = m_aData(aCol(lCol), lFirst + lRow)
            Next
        Next
    End If
    GetRows = aOut
End Function

Public Function GetRowsWithHeaders(Optional ByVal RowCount As Long = -1, Optional ByVal StartIdxZeroBased As Long, Optional FieldListCommaSeparated As String, Optional ByVal TransPosed As Boolean, Optional ByVal HeaderAtIdxMinus1 As Boolean) As Variant()
    pvNotImpl "GetRowsWithHeaders"
End Function

Public Sub ReQuery()
    pvLoad
End Sub

Public Sub ResetChanges(Optional ByVal WhichChanges As ResetEnum = ResetAll)
    Dim lRow            As Long
    Dim lCol            As Long

    If Not m_bWriteState Then
        Exit Sub
    End If
    '--- pending inserts: drop the appended rows
    If WhichChanges = ResetAll Or WhichChanges = ResetInsertsOnly Then
        For lRow = m_lRecordCount - 1 To 0 Step -1
            If m_aRowState(lRow) = ROW_NEW Then
                pvRemoveRow lRow
            End If
        Next
        If m_lCurRow > m_lRecordCount Then
            m_lCurRow = m_lRecordCount
        End If
    End If
    '--- pending deletes: re-append the snapshotted rows (before the updates
    '--- pass, so on ResetAll their pending modifications get reset too)
    If WhichChanges = ResetAll Or WhichChanges = ResetDeletesOnly Then
        pvRestoreDeleted
    End If
    '--- pending updates: restore each dirty cell's original
    If WhichChanges = ResetAll Or WhichChanges = ResetUpdatesOnly Then
        For lRow = 0 To m_lRecordCount - 1
            If m_aRowState(lRow) = ROW_MODIFIED Then
                For lCol = 0 To m_lFieldCount - 1
                    If m_aDirty(lCol, lRow) Then
                        m_aData(lCol, lRow) = m_aOriginal(lCol, lRow)
                        m_aDirty(lCol, lRow) = False
                        m_aOriginal(lCol, lRow) = Empty
                    End If
                Next
                m_aRowState(lRow) = ROW_CLEAN
            End If
        Next
    End If
    RaiseEvent Reset
End Sub

Public Function DataSource() As Object
End Function

Public Function GetADORsFromContent() As Object
End Function

Public Function ToJSONUTF8(Optional ByVal ValueArrAsColsRows As Boolean, Optional ByVal Indent As Byte, Optional ByVal UniEscaping As Boolean, Optional ByVal LfChar As Byte) As Byte()
    pvNotImpl "ToJSONUTF8"
End Function

Friend Sub frOpen(oCnn As cConnection, sSql As String, ByVal bReadOnly As Boolean)
    Set m_oConnection = oCnn
    m_sSql = sSql
    m_bReadOnly = bReadOnly
    m_bMapNullToEmpty = oCnn.MapDbNullToEmpty
    pvLoad
End Sub

Friend Sub frOpenParams(oCnn As cConnection, sSql As String, aParams() As Variant)
    Dim hStmt           As LongPtr
    Dim lIdx            As Long

    Set m_oConnection = oCnn
    m_sSql = sSql
    m_bMapNullToEmpty = oCnn.MapDbNullToEmpty
    hStmt = pvPrepare(sSql)
    On Error GoTo EH
    For lIdx = LBound(aParams) To UBound(aParams)
        If BindVariant(hStmt, lIdx - LBound(aParams) + 1, aParams(lIdx)) <> SQLITE_OK Then
            Err.Raise vbObjectError, "cRecordset", m_oConnection.LastDBError()
        End If
    Next
    '--- pvMaterialize finalizes the statement on all paths incl. raising
    On Error GoTo 0
    pvMaterialize hStmt
    Exit Sub
EH:
    '--- a bind error must not leak the statement
    Call stub_sqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Friend Sub frLoadStmt(oCnn As cConnection, sSql As String, ByVal hStmt As LongPtr)
    '--- materialise from a statement owned by the caller (cSelectCommand):
    '--- the statement is read to completion but NOT finalized here
    Set m_oConnection = oCnn
    m_sSql = sSql
    m_bMapNullToEmpty = oCnn.MapDbNullToEmpty
    pvClear
    pvMaterialize hStmt, False
End Sub

Friend Function frFieldCount() As Long
    frFieldCount = m_lFieldCount
End Function

Friend Function frFieldName(ByVal lCol As Long) As String
    If lCol >= 0 And lCol < m_lFieldCount Then
        frFieldName = m_aFieldName(lCol)
    End If
End Function

Friend Function frFieldIndex(sName As String) As Long
    Dim lIdx            As Long

    frFieldIndex = -1
    For lIdx = 0 To m_lFieldCount - 1
        If StrComp(m_aFieldName(lIdx), sName, vbTextCompare) = 0 Then
            frFieldIndex = lIdx
            Exit Function
        End If
    Next
End Function

Friend Function frCellValue(ByVal lCol As Long) As Variant
    If m_lCurRow < 0 Or m_lCurRow >= m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "No current row (BOF or EOF)"
    End If
    frCellValue = m_aData(lCol, m_lCurRow)
End Function

Friend Function frCellValueAt(ByVal lRow As Long, ByVal lCol As Long) As Variant
    If lRow < 0 Or lRow >= m_lRecordCount Or lCol < 0 Or lCol >= m_lFieldCount Then
        Err.Raise 9
    End If
    frCellValueAt = m_aData(lCol, lRow)
End Function

Friend Sub frSetCellValue(ByVal lCol As Long, vValue As Variant)
    If m_lCurRow < 0 Or m_lCurRow >= m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "No current row (BOF or EOF)"
    End If
    frSetCellValueAt m_lCurRow, lCol, vValue
End Sub

Friend Sub frSetCellValueAt(ByVal lRow As Long, ByVal lCol As Long, vValue As Variant)
    If lRow < 0 Or lRow >= m_lRecordCount Or lCol < 0 Or lCol >= m_lFieldCount Then
        Err.Raise 9
    End If
    pvCheckWritable lCol
    pvEnsureWriteState
    '--- first change to an existing row's cell captures the original value
    If m_aRowState(lRow) <> ROW_NEW Then
        If Not m_aDirty(lCol, lRow) Then
            m_aOriginal(lCol, lRow) = m_aData(lCol, lRow)
            m_aDirty(lCol, lRow) = True
        End If
        m_aRowState(lRow) = ROW_MODIFIED
    End If
    m_aData(lCol, lRow) = vValue
    RaiseEvent FieldChange(lRow, lCol)
End Sub

Friend Function frCellDirty(ByVal lCol As Long) As Boolean
    If m_bWriteState And m_lCurRow >= 0 And m_lCurRow < m_lRecordCount Then
        If lCol >= 0 And lCol < m_lFieldCount Then
            frCellDirty = m_aDirty(lCol, m_lCurRow)
        End If
    End If
End Function

Friend Function frCellOriginal(ByVal lCol As Long) As Variant
    If frCellDirty(lCol) Then
        frCellOriginal = m_aOriginal(lCol, m_lCurRow)
    Else
        frCellOriginal = frCellValue(lCol)
    End If
End Function

Friend Function frColumnUpdatable(ByVal lCol As Long) As Boolean
    If m_bUpdatable And lCol >= 0 And lCol < m_lFieldCount Then
        frColumnUpdatable = Len(m_aOrigColumn(lCol)) > 0
    End If
End Function

Friend Function frColOriginName(ByVal lCol As Long) As String
    If lCol >= 0 And lCol < m_lFieldCount Then
        frColOriginName = m_aOrigColumn(lCol)
    End If
End Function

Friend Function frColOriginTable(ByVal lCol As Long) As String
    If lCol >= 0 And lCol < m_lFieldCount Then
        frColOriginTable = m_aOrigTable(lCol)
    End If
End Function

Friend Function frColOriginDb(ByVal lCol As Long) As String
    If lCol >= 0 And lCol < m_lFieldCount Then
        frColOriginDb = m_aOrigDb(lCol)
    End If
End Function

Friend Function frColDeclType(ByVal lCol As Long) As String
    If lCol >= 0 And lCol < m_lFieldCount Then
        frColDeclType = m_aDeclType(lCol)
    End If
End Function

Friend Function frColSchema(ByVal lCol As Long) As cColumn
    Dim sKey            As String
    Dim oColumns        As cColumns

    '--- lazily build (and cache per origin table) the cColumns enrichment
    '--- so cField can expose the pragma-backed schema verdicts
    If lCol < 0 Or lCol >= m_lFieldCount Then
        Exit Function
    End If
    '--- a Content-loaded recordset carries its schema in the blob
    If m_bHasColSchema Then
        Set frColSchema = m_aColSchema(lCol)
        Exit Function
    End If
    If Len(m_aOrigColumn(lCol)) = 0 Or Len(m_aOrigTable(lCol)) = 0 Then
        Exit Function
    End If
    If m_cColMeta Is Nothing Then
        Set m_cColMeta = New VBA.Collection
    End If
    sKey = m_aOrigDb(lCol) & "." & m_aOrigTable(lCol)
    On Error Resume Next
    Set oColumns = m_cColMeta.Item(sKey)
    On Error GoTo 0
    If oColumns Is Nothing Then
        Set oColumns = New cColumns
        oColumns.frInit m_oConnection, m_aOrigDb(lCol), m_aOrigTable(lCol)
        m_cColMeta.Add oColumns, sKey
    End If
    On Error Resume Next
    Set frColSchema = oColumns.Item(m_aOrigColumn(lCol))
    On Error GoTo 0
End Function

Private Sub pvLoad()
    pvMaterialize pvPrepare(m_sSql)
End Sub

Private Function pvPrepare(sSql As String) As LongPtr
    pvClear
    pvPrepare = PrepareStatement(m_oConnection, sSql)
End Function

Private Sub pvMaterialize(ByVal hStmt As LongPtr, Optional ByVal bFinalize As Boolean = True)
    Dim lCol            As Long
    Dim lRow            As Long
    Dim lCap            As Long
    Dim lRc             As Long

    m_lFieldCount = stub_sqlite3_column_count(hStmt)
    lCap = 16
    If m_lFieldCount > 0 Then
        ReDim m_aFieldName(0 To m_lFieldCount - 1)
        For lCol = 0 To m_lFieldCount - 1
            m_aFieldName(lCol) = FromUtf8Ptr(stub_sqlite3_column_name(hStmt, lCol))
        Next
        ReDim m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1)
    End If
    pvCaptureColumnMeta hStmt
    pvAnalyzeUpdatable
    lRow = 0
    Do
        lRc = stub_sqlite3_step(hStmt)
        If lRc <> SQLITE_ROW Then
            Exit Do
        End If
        If m_lFieldCount > 0 Then
            If lRow > UBound(m_aData, 2) Then
                lCap = lCap * 2
                ReDim Preserve m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1)
            End If
            For lCol = 0 To m_lFieldCount - 1
                m_aData(lCol, lRow) = ReadColumnValue(hStmt, lCol)
                '--- RC6 semantics: DB NULLs surface as Empty by default
                If m_bMapNullToEmpty Then
                    If IsNull(m_aData(lCol, lRow)) Then
                        m_aData(lCol, lRow) = Empty
                    End If
                End If
            Next
        End If
        lRow = lRow + 1
    Loop
    If bFinalize Then
        Call stub_sqlite3_finalize(hStmt)
    End If
    '--- a step error (e.g. a UDF calling SetResultError) fails the load
    If lRc <> SQLITE_DONE Then
        Err.Raise vbObjectError, "cRecordset", m_oConnection.LastDBError()
    End If
    m_lRecordCount = lRow
    m_lCurRow = -1
    If m_lRecordCount > 0 Then
        m_lCurRow = 0
    End If
    pvEnsureFields
    RaiseEvent QueryFinished
End Sub

Private Sub pvClear()
    m_lFieldCount = 0
    m_lRecordCount = 0
    m_lCurRow = -1
    Erase m_aData
    Erase m_aFieldName
    '--- column metadata + updatability + pending-changes state
    m_bUpdatable = False
    m_sOrigDb = vbNullString
    m_sOrigTable = vbNullString
    m_lPkCount = 0
    m_lAutoPkCol = -1
    Erase m_aOrigColumn
    Erase m_aOrigTable
    Erase m_aOrigDb
    Erase m_aDeclType
    Set m_cColMeta = Nothing
    Erase m_aColSchema
    m_bHasColSchema = False
    Erase m_aPkCols
    m_bWriteState = False
    Erase m_aOriginal
    Erase m_aDirty
    Erase m_aRowState
    Set m_cDeleted = Nothing
    pvReleaseFields
End Sub

Private Sub pvEnsureFields()
    Set m_oFields = New cFields
    m_oFields.frInit Me
End Sub

Private Sub pvReleaseFields()
    '--- invalidate the current fields (on ReQuery or terminate) so any
    '--- externally-held cField fails safely rather than reading the next
    '--- result set or, after terminate, freed memory
    If Not m_oFields Is Nothing Then
        m_oFields.frTerminate
        Set m_oFields = Nothing
    End If
End Sub

Private Function pvResolveColumns(sList As String, lNumCols As Long) As Long()
    Dim aRet()          As Long
    Dim aName()         As String
    Dim lIdx            As Long
    Dim lCol            As Long

    lNumCols = 0
    If Len(Trim$(sList)) = 0 Then
        If m_lFieldCount = 0 Then
            Exit Function
        End If
        ReDim aRet(0 To m_lFieldCount - 1)
        For lIdx = 0 To m_lFieldCount - 1
            aRet(lIdx) = lIdx
        Next
        lNumCols = m_lFieldCount
    Else
        aName = Split(sList, ",")
        ReDim aRet(0 To UBound(aName))
        For lIdx = 0 To UBound(aName)
            lCol = frFieldIndex(Trim$(aName(lIdx)))
            If lCol >= 0 Then
                aRet(lNumCols) = lCol
                lNumCols = lNumCols + 1
            End If
        Next
        If lNumCols = 0 Then
            Exit Function
        End If
        ReDim Preserve aRet(0 To lNumCols - 1)
    End If
    pvResolveColumns = aRet
End Function

Private Sub pvCaptureColumnMeta(ByVal hStmt As LongPtr)
    Dim lCol            As Long

    If m_lFieldCount = 0 Then
        Exit Sub
    End If
    ReDim m_aOrigColumn(0 To m_lFieldCount - 1)
    ReDim m_aOrigTable(0 To m_lFieldCount - 1)
    ReDim m_aOrigDb(0 To m_lFieldCount - 1)
    ReDim m_aDeclType(0 To m_lFieldCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        m_aOrigColumn(lCol) = FromUtf8Ptr(stub_sqlite3_column_origin_name(hStmt, lCol))
        m_aOrigTable(lCol) = FromUtf8Ptr(stub_sqlite3_column_table_name(hStmt, lCol))
        m_aOrigDb(lCol) = FromUtf8Ptr(stub_sqlite3_column_database_name(hStmt, lCol))
        m_aDeclType(lCol) = FromUtf8Ptr(stub_sqlite3_column_decltype(hStmt, lCol))
    Next
End Sub

Private Sub pvAnalyzeUpdatable()
    Dim lCol            As Long
    Dim aPkName()       As String
    Dim sPkType         As String
    Dim lPkIdx          As Long
    Dim bFound          As Boolean

    If m_bReadOnly Or m_lFieldCount = 0 Then
        Exit Sub
    End If
    '--- updatable iff every table-backed column originates in one table and
    '--- all of that table's PK columns are present in the result set
    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigTable(lCol)) > 0 Then
            If Len(m_sOrigTable) = 0 Then
                m_sOrigTable = m_aOrigTable(lCol)
                m_sOrigDb = m_aOrigDb(lCol)
            ElseIf StrComp(m_sOrigTable, m_aOrigTable(lCol), vbTextCompare) <> 0 Then
                Exit Sub
            End If
        End If
    Next
    If Len(m_sOrigTable) = 0 Then
        Exit Sub
    End If
    If Len(m_sOrigDb) = 0 Then
        m_sOrigDb = "main"
    End If
    aPkName = pvGetPkColumns(sPkType)
    If m_lPkCount = 0 Then
        Exit Sub
    End If
    ReDim m_aPkCols(0 To m_lPkCount - 1)
    For lPkIdx = 0 To m_lPkCount - 1
        bFound = False
        For lCol = 0 To m_lFieldCount - 1
            If StrComp(m_aOrigColumn(lCol), aPkName(lPkIdx), vbTextCompare) = 0 Then
                m_aPkCols(lPkIdx) = lCol
                bFound = True
                Exit For
            End If
        Next
        If Not bFound Then
            Exit Sub
        End If
    Next
    '--- a single INTEGER PRIMARY KEY is the rowid alias: SQLite auto-assigns
    '--- it on NULL insert, so pvBatchInsert can backfill the cell
    If m_lPkCount = 1 And StrComp(sPkType, "INTEGER", vbTextCompare) = 0 Then
        m_lAutoPkCol = m_aPkCols(0)
    End If
    m_bUpdatable = True
End Sub

Private Function pvGetPkColumns(sSinglePkType As String) As String()
    Dim hStmt           As LongPtr
    Dim aName()         As String
    Dim aPos()          As Long
    Dim aRet()          As String
    Dim lCount          As Long
    Dim lPk             As Long
    Dim lIdx            As Long

    ReDim aName(0 To 3)
    ReDim aPos(0 To 3)
    hStmt = PrepareStatement(m_oConnection, "PRAGMA " & QuoteIdentifier(m_sOrigDb) & ".table_info(" & QuoteIdentifier(m_sOrigTable) & ")")
    Do While stub_sqlite3_step(hStmt) = SQLITE_ROW
        '--- table_info columns: cid, name, type, notnull, dflt_value, pk
        lPk = stub_sqlite3_column_int(hStmt, 5)
        If lPk > 0 Then
            If lCount > UBound(aName) Then
                ReDim Preserve aName(0 To 2 * lCount)
                ReDim Preserve aPos(0 To 2 * lCount)
            End If
            aName(lCount) = FromUtf8Ptr(stub_sqlite3_column_text(hStmt, 1))
            aPos(lCount) = lPk
            If lPk = 1 Then
                sSinglePkType = FromUtf8Ptr(stub_sqlite3_column_text(hStmt, 2))
            End If
            lCount = lCount + 1
        End If
    Loop
    Call stub_sqlite3_finalize(hStmt)
    m_lPkCount = lCount
    If lCount > 0 Then
        ReDim aRet(0 To lCount - 1)
        For lIdx = 0 To lCount - 1
            aRet(aPos(lIdx) - 1) = aName(lIdx)
        Next
        pvGetPkColumns = aRet
    End If
End Function

Private Sub pvEnsureWriteState()
    If Not m_bWriteState Then
        ReDim m_aOriginal(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2))
        ReDim m_aDirty(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2))
        ReDim m_aRowState(0 To UBound(m_aData, 2))
        Set m_cDeleted = New VBA.Collection
        m_bWriteState = True
    End If
End Sub

Private Sub pvEnsureRowCapacity(ByVal lRows As Long)
    Dim lCap            As Long

    lCap = UBound(m_aData, 2) + 1
    If lRows > lCap Then
        Do
            lCap = lCap * 2
        Loop While lRows > lCap
        ReDim Preserve m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1)
        If m_bWriteState Then
            ReDim Preserve m_aOriginal(0 To m_lFieldCount - 1, 0 To lCap - 1)
            ReDim Preserve m_aDirty(0 To m_lFieldCount - 1, 0 To lCap - 1)
            ReDim Preserve m_aRowState(0 To lCap - 1)
        End If
    End If
End Sub

Private Sub pvCheckUpdatable()
    If Not m_bUpdatable Then
        Err.Raise vbObjectError, "cRecordset", "Recordset is not updatable"
    End If
End Sub

Private Sub pvCheckWritable(ByVal lCol As Long)
    pvCheckUpdatable
    If Len(m_aOrigColumn(lCol)) = 0 Then
        Err.Raise vbObjectError, "cRecordset", "Column is an expression and cannot be updated"
    End If
End Sub

Private Function pvPkWhereValue(ByVal lRow As Long, ByVal lPkIdx As Long) As Variant
    Dim lCol            As Long

    '--- WHERE must address the row as stored in the DB: original value when
    '--- the PK cell itself has a pending change
    lCol = m_aPkCols(lPkIdx)
    If m_bWriteState Then
        If m_aDirty(lCol, lRow) Then
            pvPkWhereValue = m_aOriginal(lCol, lRow)
            Exit Function
        End If
    End If
    pvPkWhereValue = m_aData(lCol, lRow)
End Function

Private Function pvPkWhereClause() As String
    Dim lPkIdx          As Long
    Dim sSql            As String

    For lPkIdx = 0 To m_lPkCount - 1
        If lPkIdx > 0 Then
            sSql = sSql & " AND "
        End If
        sSql = sSql & QuoteIdentifier(m_aOrigColumn(m_aPkCols(lPkIdx))) & " = ?"
    Next
    pvPkWhereClause = sSql
End Function

Private Function pvQualifiedTable() As String
    pvQualifiedTable = QuoteIdentifier(m_sOrigDb) & "." & QuoteIdentifier(m_sOrigTable)
End Function

Private Sub pvBatchDelete(vEntry As Variant)
    Dim aParams()       As Variant

    aParams = vEntry(1)
    pvExecParams "DELETE FROM " & pvQualifiedTable() & " WHERE " & pvPkWhereClause(), aParams
End Sub

Private Sub pvBatchUpdate(ByVal lRow As Long)
    Dim sSet            As String
    Dim aParams()       As Variant
    Dim lCount          As Long
    Dim lCol            As Long
    Dim lPkIdx          As Long

    ReDim aParams(0 To m_lFieldCount + m_lPkCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        If m_aDirty(lCol, lRow) And Len(m_aOrigColumn(lCol)) > 0 Then
            If Len(sSet) > 0 Then
                sSet = sSet & ", "
            End If
            sSet = sSet & QuoteIdentifier(m_aOrigColumn(lCol)) & " = ?"
            aParams(lCount) = m_aData(lCol, lRow)
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        Exit Sub
    End If
    For lPkIdx = 0 To m_lPkCount - 1
        aParams(lCount) = pvPkWhereValue(lRow, lPkIdx)
        lCount = lCount + 1
    Next
    ReDim Preserve aParams(0 To lCount - 1)
    pvExecParams "UPDATE " & pvQualifiedTable() & " SET " & sSet & " WHERE " & pvPkWhereClause(), aParams
End Sub

Private Sub pvBatchInsert(ByVal lRow As Long)
    Dim sCols           As String
    Dim sVals           As String
    Dim aParams()       As Variant
    Dim lCount          As Long
    Dim lCol            As Long

    ReDim aParams(0 To m_lFieldCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigColumn(lCol)) > 0 Then
            If lCount > 0 Then
                sCols = sCols & ", "
                sVals = sVals & ", "
            End If
            sCols = sCols & QuoteIdentifier(m_aOrigColumn(lCol))
            sVals = sVals & "?"
            aParams(lCount) = m_aData(lCol, lRow)
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        Exit Sub
    End If
    ReDim Preserve aParams(0 To lCount - 1)
    pvExecParams "INSERT INTO " & pvQualifiedTable() & " (" & sCols & ") VALUES (" & sVals & ")", aParams
    '--- INTEGER PRIMARY KEY auto-assigned on NULL insert: backfill the cell
    If m_lAutoPkCol >= 0 Then
        If pvIsNullCell(m_aData(m_lAutoPkCol, lRow)) Then
            m_aData(m_lAutoPkCol, lRow) = m_oConnection.LastInsertAutoID
        End If
    End If
End Sub

Private Sub pvExecParams(sSql As String, aParams() As Variant)
    Dim hStmt           As LongPtr
    Dim lIdx            As Long
    Dim lRc             As Long

    hStmt = PrepareStatement(m_oConnection, sSql)
    On Error GoTo EH
    For lIdx = 0 To UBound(aParams)
        If BindVariant(hStmt, lIdx + 1, aParams(lIdx)) <> SQLITE_OK Then
            Err.Raise vbObjectError, "cRecordset.UpdateBatch", m_oConnection.LastDBError()
        End If
    Next
    Do
        lRc = stub_sqlite3_step(hStmt)
    Loop While lRc = SQLITE_ROW
    If lRc <> SQLITE_DONE Then
        Err.Raise vbObjectError, "cRecordset.UpdateBatch", m_oConnection.LastDBError()
    End If
    Call stub_sqlite3_finalize(hStmt)
    Exit Sub
EH:
    Call stub_sqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub pvRemoveRow(ByVal lRow As Long)
    Dim lCol            As Long
    Dim lIdx            As Long

    For lIdx = lRow To m_lRecordCount - 2
        For lCol = 0 To m_lFieldCount - 1
            m_aData(lCol, lIdx) = m_aData(lCol, lIdx + 1)
            m_aOriginal(lCol, lIdx) = m_aOriginal(lCol, lIdx + 1)
            m_aDirty(lCol, lIdx) = m_aDirty(lCol, lIdx + 1)
        Next
        m_aRowState(lIdx) = m_aRowState(lIdx + 1)
    Next
    m_lRecordCount = m_lRecordCount - 1
End Sub

Private Sub pvRestoreDeleted()
    Dim vEntry          As Variant
    Dim aRow()          As Variant
    Dim aOrig()         As Variant
    Dim aDirty()        As Boolean
    Dim lCol            As Long
    Dim lNew            As Long

    Do While m_cDeleted.Count > 0
        vEntry = m_cDeleted.Item(1)
        m_cDeleted.Remove 1
        pvEnsureRowCapacity m_lRecordCount + 1
        lNew = m_lRecordCount
        aRow = vEntry(0)
        aDirty = vEntry(3)
        aOrig = vEntry(4)
        For lCol = 0 To m_lFieldCount - 1
            m_aData(lCol, lNew) = aRow(lCol)
            m_aOriginal(lCol, lNew) = aOrig(lCol)
            m_aDirty(lCol, lNew) = aDirty(lCol)
        Next
        m_aRowState(lNew) = vEntry(2)
        m_lRecordCount = m_lRecordCount + 1
    Loop
End Sub

Private Sub pvClearWriteState()
    Dim lCol            As Long
    Dim lRow            As Long

    If m_bWriteState Then
        For lRow = 0 To m_lRecordCount - 1
            m_aRowState(lRow) = ROW_CLEAN
            For lCol = 0 To m_lFieldCount - 1
                m_aDirty(lCol, lRow) = False
                m_aOriginal(lCol, lRow) = Empty
            Next
        Next
        Set m_cDeleted = New VBA.Collection
    End If
End Sub

Private Sub pvApplySort()
    Dim aCols()         As Long
    Dim aDesc()         As Boolean
    Dim lKeys           As Long
    Dim aIdx()          As Long
    Dim lRow            As Long

    If Len(Trim$(m_sSort)) = 0 Or m_lRecordCount = 0 Then
        Exit Sub
    End If
    pvParseSort aCols, aDesc, lKeys
    If lKeys = 0 Then
        Exit Sub
    End If
    ReDim aIdx(0 To m_lRecordCount - 1)
    For lRow = 0 To m_lRecordCount - 1
        aIdx(lRow) = lRow
    Next
    pvQuickSortRows aIdx, 0, m_lRecordCount - 1, aCols, aDesc, lKeys
    pvPermuteRows aIdx
End Sub

Private Sub pvParseSort(aCols() As Long, aDesc() As Boolean, lKeys As Long)
    Dim aTok()          As String
    Dim lIdx            As Long
    Dim sName           As String
    Dim sDir            As String
    Dim lPos            As Long

    aTok = Split(m_sSort, ",")
    ReDim aCols(0 To UBound(aTok))
    ReDim aDesc(0 To UBound(aTok))
    lKeys = 0
    For lIdx = 0 To UBound(aTok)
        sName = Trim$(aTok(lIdx))
        If Len(sName) > 0 Then
            sDir = vbNullString
            If Left$(sName, 1) = "[" Then
                lPos = InStr(sName, "]")
                If lPos = 0 Then
                    Err.Raise 5, "cRecordset.Sort", "Invalid sort expression"
                End If
                sDir = Trim$(Mid$(sName, lPos + 1))
                sName = Mid$(sName, 2, lPos - 2)
            Else
                lPos = InStr(sName, " ")
                If lPos > 0 Then
                    sDir = Trim$(Mid$(sName, lPos + 1))
                    sName = Left$(sName, lPos - 1)
                End If
            End If
            aCols(lKeys) = frFieldIndex(sName)
            If aCols(lKeys) < 0 Then
                Err.Raise vbObjectError, "cRecordset.Sort", "Unknown field '" & sName & "'"
            End If
            aDesc(lKeys) = (StrComp(sDir, "DESC", vbTextCompare) = 0)
            lKeys = lKeys + 1
        End If
    Next
End Sub

Private Sub pvQuickSortRows(aIdx() As Long, ByVal lLo As Long, ByVal lHi As Long, aCols() As Long, aDesc() As Boolean, ByVal lKeys As Long)
    Dim lLeft           As Long
    Dim lRight          As Long
    Dim lPivot          As Long
    Dim lTemp           As Long

    lLeft = lLo
    lRight = lHi
    lPivot = aIdx((lLo + lHi) \ 2)
    Do While lLeft <= lRight
        Do While pvCompareRows(aIdx(lLeft), lPivot, aCols, aDesc, lKeys) < 0
            lLeft = lLeft + 1
        Loop
        Do While pvCompareRows(aIdx(lRight), lPivot, aCols, aDesc, lKeys) > 0
            lRight = lRight - 1
        Loop
        If lLeft <= lRight Then
            lTemp = aIdx(lLeft)
            aIdx(lLeft) = aIdx(lRight)
            aIdx(lRight) = lTemp
            lLeft = lLeft + 1
            lRight = lRight - 1
        End If
    Loop
    If lLo < lRight Then
        pvQuickSortRows aIdx, lLo, lRight, aCols, aDesc, lKeys
    End If
    If lLeft < lHi Then
        pvQuickSortRows aIdx, lLeft, lHi, aCols, aDesc, lKeys
    End If
End Sub

Private Function pvCompareRows(ByVal lRow1 As Long, ByVal lRow2 As Long, aCols() As Long, aDesc() As Boolean, ByVal lKeys As Long) As Long
    Dim lIdx            As Long
    Dim lCmp            As Long

    For lIdx = 0 To lKeys - 1
        lCmp = pvCompareValues(m_aData(aCols(lIdx), lRow1), m_aData(aCols(lIdx), lRow2))
        If aDesc(lIdx) Then
            lCmp = -lCmp
        End If
        If lCmp <> 0 Then
            pvCompareRows = lCmp
            Exit Function
        End If
    Next
    '--- tie-break on the original row order to keep the sort stable
    pvCompareRows = Sgn(lRow1 - lRow2)
End Function

Private Sub pvPermuteRows(aIdx() As Long)
    Dim aDataNew()      As Variant
    Dim aOriginalNew()  As Variant
    Dim aDirtyNew()     As Boolean
    Dim aStateNew()     As Long
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lNewCur         As Long

    ReDim aDataNew(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2))
    lNewCur = m_lCurRow
    For lRow = 0 To m_lRecordCount - 1
        For lCol = 0 To m_lFieldCount - 1
            aDataNew(lCol, lRow) = m_aData(lCol, aIdx(lRow))
        Next
        If aIdx(lRow) = m_lCurRow Then
            lNewCur = lRow
        End If
    Next
    m_aData = aDataNew
    If m_bWriteState Then
        ReDim aOriginalNew(0 To m_lFieldCount - 1, 0 To UBound(m_aOriginal, 2))
        ReDim aDirtyNew(0 To m_lFieldCount - 1, 0 To UBound(m_aDirty, 2))
        ReDim aStateNew(0 To UBound(m_aRowState))
        For lRow = 0 To m_lRecordCount - 1
            For lCol = 0 To m_lFieldCount - 1
                aOriginalNew(lCol, lRow) = m_aOriginal(lCol, aIdx(lRow))
                aDirtyNew(lCol, lRow) = m_aDirty(lCol, aIdx(lRow))
            Next
            aStateNew(lRow) = m_aRowState(aIdx(lRow))
        Next
        m_aOriginal = aOriginalNew
        m_aDirty = aDirtyNew
        m_aRowState = aStateNew
    End If
    '--- the cursor follows its row to the row's new position
    If m_lCurRow >= 0 And m_lCurRow < m_lRecordCount Then
        If lNewCur <> m_lCurRow Then
            m_lCurRow = lNewCur
            RaiseEvent Move(m_lCurRow)
        End If
    End If
End Sub

Private Function pvFindMove(ByVal lFrom As Long, ByVal lStep As Long, sCriterion As String, ByVal bDistinctNull As Boolean) As Boolean
    Dim lCol            As Long
    Dim lOp             As Long
    Dim vValue          As Variant
    Dim lRow            As Long

    If Not pvParseCriterion(sCriterion, lCol, lOp, vValue) Then
        Err.Raise 5, "cRecordset.Find", "Invalid criterion '" & sCriterion & "'"
    End If
    lRow = lFrom
    Do While lRow >= 0 And lRow < m_lRecordCount
        If pvMatchValue(m_aData(lCol, lRow), lOp, vValue, bDistinctNull) Then
            '--- found: move there; not found leaves the position unchanged
            AbsolutePosition = lRow
            pvFindMove = True
            Exit Function
        End If
        lRow = lRow + lStep
    Loop
End Function

Private Function pvParseCriterion(sCriterion As String, lCol As Long, lOp As Long, vValue As Variant) As Boolean
    Dim sText           As String
    Dim sName           As String
    Dim lPos            As Long

    sText = Trim$(sCriterion)
    '--- field name: [bracketed] or up to the first space/operator char
    If Left$(sText, 1) = "[" Then
        lPos = InStr(sText, "]")
        If lPos = 0 Then
            Exit Function
        End If
        sName = Mid$(sText, 2, lPos - 2)
        sText = LTrim$(Mid$(sText, lPos + 1))
    Else
        lPos = 1
        Do While lPos <= Len(sText)
            If InStr(" =<>", Mid$(sText, lPos, 1)) > 0 Then
                Exit Do
            End If
            lPos = lPos + 1
        Loop
        sName = Left$(sText, lPos - 1)
        sText = LTrim$(Mid$(sText, lPos))
    End If
    lCol = frFieldIndex(sName)
    If lCol < 0 Then
        Exit Function
    End If
    '--- operator
    If Left$(sText, 2) = ">=" Then
        lOp = OP_GE
        sText = Mid$(sText, 3)
    ElseIf Left$(sText, 2) = "<=" Then
        lOp = OP_LE
        sText = Mid$(sText, 3)
    ElseIf Left$(sText, 2) = "<>" Then
        lOp = OP_NE
        sText = Mid$(sText, 3)
    ElseIf Left$(sText, 1) = "=" Then
        lOp = OP_EQ
        sText = Mid$(sText, 2)
    ElseIf Left$(sText, 1) = ">" Then
        lOp = OP_GT
        sText = Mid$(sText, 2)
    ElseIf Left$(sText, 1) = "<" Then
        lOp = OP_LT
        sText = Mid$(sText, 2)
    ElseIf StrComp(Left$(sText, 5), "LIKE ", vbTextCompare) = 0 Then
        lOp = OP_LIKE
        sText = Mid$(sText, 6)
    ElseIf StrComp(Left$(sText, 3), "IS ", vbTextCompare) = 0 Then
        sText = Trim$(Mid$(sText, 4))
        If StrComp(sText, "NULL", vbTextCompare) = 0 Then
            lOp = OP_ISNULL
        ElseIf StrComp(sText, "NOT NULL", vbTextCompare) = 0 Then
            lOp = OP_ISNOTNULL
        Else
            Exit Function
        End If
        pvParseCriterion = True
        Exit Function
    Else
        Exit Function
    End If
    '--- literal: 'string' ('' escapes a quote), NULL, or a number
    sText = Trim$(sText)
    If Left$(sText, 1) = "'" Then
        If Len(sText) < 2 Or Right$(sText, 1) <> "'" Then
            Exit Function
        End If
        vValue = Replace(Mid$(sText, 2, Len(sText) - 2), "''", "'")
    ElseIf StrComp(sText, "NULL", vbTextCompare) = 0 Then
        vValue = Null
    ElseIf IsNumeric(sText) Then
        vValue = Val(sText)
    Else
        Exit Function
    End If
    pvParseCriterion = True
End Function

Private Function pvMatchValue(vCell As Variant, ByVal lOp As Long, vValue As Variant, ByVal bDistinctNull As Boolean) As Boolean
    Dim lCmp            As Long

    Select Case lOp
    Case OP_ISNULL
        pvMatchValue = pvIsNullCell(vCell)
    Case OP_ISNOTNULL
        pvMatchValue = Not pvIsNullCell(vCell)
    Case OP_LIKE
        If Not pvIsNullCell(vCell) Then
            pvMatchValue = LCase$(CStr(vCell)) Like pvLikePattern(LCase$(CStr(vValue)))
        End If
    Case Else
        If pvIsNullCell(vCell) Or IsNull(vValue) Then
            '--- SQL three-valued logic (never matches) unless
            '--- DistinctNullValues=False treats Null as a regular value
            If Not bDistinctNull Then
                If pvIsNullCell(vCell) And IsNull(vValue) Then
                    pvMatchValue = (lOp = OP_EQ Or lOp = OP_LE Or lOp = OP_GE)
                Else
                    pvMatchValue = (lOp = OP_NE)
                End If
            End If
        Else
            lCmp = pvCompareValues(vCell, vValue)
            Select Case lOp
            Case OP_EQ
                pvMatchValue = (lCmp = 0)
            Case OP_NE
                pvMatchValue = (lCmp <> 0)
            Case OP_LT
                pvMatchValue = (lCmp < 0)
            Case OP_LE
                pvMatchValue = (lCmp <= 0)
            Case OP_GT
                pvMatchValue = (lCmp > 0)
            Case OP_GE
                pvMatchValue = (lCmp >= 0)
            End Select
        End If
    End Select
End Function

Private Function pvLikePattern(sPattern As String) As String
    Dim lIdx            As Long
    Dim sChar           As String

    '--- SQLite wildcards (% _) to VB Like (* ?), escaping VB specials
    For lIdx = 1 To Len(sPattern)
        sChar = Mid$(sPattern, lIdx, 1)
        Select Case sChar
        Case "%"
            pvLikePattern = pvLikePattern & "*"
        Case "_"
            pvLikePattern = pvLikePattern & "?"
        Case "*", "?", "#", "["
            pvLikePattern = pvLikePattern & "[" & sChar & "]"
        Case Else
            pvLikePattern = pvLikePattern & sChar
        End Select
    Next
End Function

Private Function pvCompareValues(v1 As Variant, v2 As Variant) As Long
    Dim lRank1          As Long
    Dim lRank2          As Long

    '--- SQLite type ordering: NULL < numeric < text < blob
    lRank1 = pvTypeRank(v1)
    lRank2 = pvTypeRank(v2)
    If lRank1 <> lRank2 Then
        pvCompareValues = Sgn(lRank1 - lRank2)
        Exit Function
    End If
    Select Case lRank1
    Case 1
        If CDec(v1) < CDec(v2) Then
            pvCompareValues = -1
        ElseIf CDec(v1) > CDec(v2) Then
            pvCompareValues = 1
        End If
    Case 2
        pvCompareValues = StrComp(CStr(v1), CStr(v2), vbBinaryCompare)
    Case 3
        pvCompareValues = pvCompareBlobs(v1, v2)
    End Select
End Function

Private Function pvTypeRank(vValue As Variant) As Long
    Select Case VarType(vValue)
    Case vbNull, vbEmpty
        pvTypeRank = 0
    Case vbString
        pvTypeRank = 2
    Case vbByte + vbArray
        pvTypeRank = 3
    Case Else
        pvTypeRank = 1
    End Select
End Function

Private Function pvCompareBlobs(v1 As Variant, v2 As Variant) As Long
    Dim baBlob1()       As Byte
    Dim baBlob2()       As Byte
    Dim lLen1           As Long
    Dim lLen2           As Long
    Dim lIdx            As Long

    baBlob1 = v1
    baBlob2 = v2
    On Error Resume Next
    lLen1 = UBound(baBlob1) + 1
    lLen2 = UBound(baBlob2) + 1
    On Error GoTo 0
    For lIdx = 0 To IIf(lLen1 < lLen2, lLen1, lLen2) - 1
        If baBlob1(lIdx) <> baBlob2(lIdx) Then
            pvCompareBlobs = Sgn(CLng(baBlob1(lIdx)) - baBlob2(lIdx))
            Exit Function
        End If
    Next
    pvCompareBlobs = Sgn(lLen1 - lLen2)
End Function

'--- Content serialization (RC6-compatible blob, see NOTES for the format)

Private Function pvContentWrite() As Byte()
    Dim baOut()         As Byte
    Dim lPos            As Long
    Dim lRows           As Long
    Dim lCol            As Long
    Dim lRow            As Long
    Dim lType           As Long
    Dim lWidth          As Long
    Dim oColumn         As cColumn
    Dim sText           As String
    Dim baCell()        As Byte
    Dim lTotal          As Long
    Dim lOffs           As Long
    Dim aKeyCols()      As Long
    Dim aTabCol()       As Long
    Dim lTables         As Long

    lRows = m_lRecordCount
    ReDim baOut(0 To 1023)
    '--- header
    pvPutLong baOut, lPos, 0                        '--- cbTotal (patched below)
    pvPutLong baOut, lPos, lRows
    pvPutLong baOut, lPos, lRows
    pvPutLong baOut, lPos, lRows
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, IIf(lRows = 0, -1, m_lCurRow)
    pvPutLong baOut, lPos, 4 * (lRows + 1)
    pvPutLong baOut, lPos, 0                        '--- ptr slot (ignored)
    pvPutLong baOut, lPos, lRows
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, 0                        '--- ptr slot (ignored)
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, 0
    '--- row-order chain: one Long per row (next row index, 0-terminated)
    For lRow = 1 To lRows - 1
        pvPutLong baOut, lPos, lRow
    Next
    If lRows > 0 Then
        pvPutLong baOut, lPos, 0
    End If
    If m_oConnection Is Nothing Then
        pvPutStr baOut, lPos, vbNullString
    Else
        pvPutStr baOut, lPos, m_oConnection.FileName
    End If
    pvPutStr baOut, lPos, m_sSql
    '--- distinct origin tables in first-appearance order (aTabCol holds the
    '--- first column index of each table; used for the per-field table index
    '--- and the DML-template tail)
    ReDim aTabCol(0 To m_lFieldCount)
    lTables = 0
    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigTable(lCol)) > 0 Then
            For lRow = 0 To lTables - 1
                If StrComp(m_aOrigTable(aTabCol(lRow)), m_aOrigTable(lCol), vbTextCompare) = 0 Then
                    Exit For
                End If
            Next
            If lRow = lTables Then
                aTabCol(lTables) = lCol
                lTables = lTables + 1
            End If
        End If
    Next
    '--- fields
    pvPutLong baOut, lPos, m_lFieldCount
    For lCol = 0 To m_lFieldCount - 1
        Set oColumn = frColSchema(lCol)
        If Len(m_aOrigTable(lCol)) > 0 Then
            pvPutStr baOut, lPos, m_aOrigDb(lCol)
            pvPutStr baOut, lPos, "[" & m_aOrigTable(lCol) & "]"
            pvPutStr baOut, lPos, "[" & m_aOrigColumn(lCol) & "]"
            pvPutStr baOut, lPos, m_aDeclType(lCol)
            sText = vbNullString
            If Not oColumn Is Nothing Then
                sText = pvStripQuotes(oColumn.DefaultValue)
            End If
            If Len(sText) = 0 Then
                sText = "NULL"
            End If
            pvPutStr baOut, lPos, sText
            sText = vbNullString
            If Not oColumn Is Nothing Then
                sText = oColumn.Collate
            End If
            If Len(sText) = 0 Then
                sText = "BINARY"
            End If
            pvPutStr baOut, lPos, sText
        Else
            '--- expression column: all six metadata strings are empty
            pvPutStr baOut, lPos, vbNullString
            pvPutStr baOut, lPos, vbNullString
            pvPutStr baOut, lPos, vbNullString
            pvPutStr baOut, lPos, vbNullString
            pvPutStr baOut, lPos, vbNullString
            pvPutStr baOut, lPos, vbNullString
        End If
        pvPutStr baOut, lPos, m_aFieldName(lCol)
        '--- fixed block: origin table index + column ordinal + flags
        lTotal = 0
        For lRow = 0 To lTables - 1
            If StrComp(m_aOrigTable(aTabCol(lRow)), m_aOrigTable(lCol), vbTextCompare) = 0 Then
                lTotal = lRow
                Exit For
            End If
        Next
        pvPutLong baOut, lPos, lTotal
        pvPutLong baOut, lPos, lCol                 '--- column ordinal
        If oColumn Is Nothing Then
            pvPutByte baOut, lPos, 0
            pvPutByte baOut, lPos, 0
            pvPutByte baOut, lPos, 0
        Else
            '--- RC6 flags: {notnull, pk, autoinc} (UNIQUE is not serialized)
            pvPutByte baOut, lPos, IIf(oColumn.NotNullConstraint, 1, 0)
            pvPutByte baOut, lPos, IIf(oColumn.PrimaryKey, 1, 0)
            pvPutByte baOut, lPos, IIf(oColumn.PrimaryAutoIncrement, 1, 0)
        End If
        lType = pvDeclTypeCode(m_aDeclType(lCol))
        If Len(m_aDeclType(lCol)) = 0 Then
            '--- expression columns type by value storage class (RC6: text
            '--- stays text, blobs stay blobs, all numerics become doubles)
            lType = pvValueTypeCode(lCol)
        End If
        pvPutLong baOut, lPos, lType
        pvPutByte baOut, lPos, 0
        lWidth = 0
        If lType = 1 Then
            lWidth = pvIntWidth(lCol)
        End If
        pvPutLong baOut, lPos, lWidth
        '--- per-row not-null flags (one spare element)
        pvPutLong baOut, lPos, lRows + 1
        For lRow = 0 To lRows - 1
            pvPutByte baOut, lPos, IIf(pvIsNullCell(m_aData(lCol, lRow)), 0, 1)
        Next
        pvPutByte baOut, lPos, 0
        '--- cell data
        Select Case lType
        Case 1
            pvPutLong baOut, lPos, (lRows + 1) * lWidth
            For lRow = 0 To lRows - 1
                pvPutIntCell baOut, lPos, m_aData(lCol, lRow), lWidth
            Next
            pvPutZeros baOut, lPos, lWidth
            pvPutLong baOut, lPos, 0
        Case 2
            pvPutLong baOut, lPos, (lRows + 1) * 8
            For lRow = 0 To lRows - 1
                pvPutDouble baOut, lPos, pvCellDbl(m_aData(lCol, lRow))
            Next
            pvPutZeros baOut, lPos, 8
            pvPutLong baOut, lPos, 0
        Case Else
            '--- TEXT/BLOB: concatenated content + rows+1 start offsets
            lTotal = 0
            For lRow = 0 To lRows - 1
                lTotal = lTotal + pvCellBytesLen(m_aData(lCol, lRow), lType)
            Next
            pvPutLong baOut, lPos, lTotal
            For lRow = 0 To lRows - 1
                baCell = pvCellBytes(m_aData(lCol, lRow), lType)
                If pvBytesLen(baCell) > 0 Then
                    pvPutBytes baOut, lPos, baCell
                End If
            Next
            pvPutLong baOut, lPos, (lRows + 1) * 4
            lOffs = 0
            For lRow = 0 To lRows - 1
                pvPutLong baOut, lPos, lOffs
                lOffs = lOffs + pvCellBytesLen(m_aData(lCol, lRow), lType)
            Next
            pvPutLong baOut, lPos, lOffs
        End Select
    Next
    '--- tail: one DML set per distinct origin table (first-appearance
    '--- order): count of sets, then per table 3 DML templates + key column
    '--- ordinals (the table's pk columns in the select, or all of its
    '--- select columns when no pk), then 42 zero bytes; the whole blob is
    '--- padded to even length
    pvPutLong baOut, lPos, lTables
    For lRow = 0 To lTables - 1
        lOffs = aTabCol(lRow)
        sText = m_aOrigTable(lOffs)
        pvPutStr baOut, lPos, "DELETE FROM [" & sText & "] "
        pvPutStr baOut, lPos, "UPDATE [" & sText & "] SET "
        pvPutStr baOut, lPos, "INSERT INTO [" & sText & "] ("
        lTotal = pvContentKeyCols(sText, m_aOrigDb(lOffs), aKeyCols)
        pvPutLong baOut, lPos, lTotal
        For lCol = 0 To lTotal - 1
            pvPutLong baOut, lPos, aKeyCols(lCol)
        Next
    Next
    pvPutZeros baOut, lPos, 42
    If (lPos And 1) <> 0 Then
        pvPutZeros baOut, lPos, 1
    End If
    ReDim Preserve baOut(0 To lPos - 1)
    Call CopyMemory(baOut(0), lPos, 4)
    pvContentWrite = baOut
End Function

Private Sub pvContentRead(baContent() As Byte)
    Dim lPos            As Long
    Dim lRows           As Long
    Dim lCol            As Long
    Dim lRow            As Long
    Dim lCurRow         As Long
    Dim lType           As Long
    Dim lWidth          As Long
    Dim sDefault        As String
    Dim sCollate        As String
    Dim lNotNull        As Long
    Dim lUnique         As Long
    Dim lPk             As Long
    Dim lAutoInc        As Long
    Dim aFlags()        As Byte
    Dim lCount          As Long
    Dim baText()        As Byte
    Dim aOffs()         As Long
    Dim lUpdatable      As Long
    Dim sPkType         As String

    pvClear
    lPos = 4                                        '--- cbTotal
    lRows = pvGetLong(baContent, lPos)
    pvGetLong baContent, lPos
    pvGetLong baContent, lPos
    pvGetLong baContent, lPos
    lCurRow = pvGetLong(baContent, lPos)
    lPos = lPos + 4 * 10                            '--- array cb + ptr slots + zeros
    lPos = lPos + 4 * lRows                         '--- row-order chain (natural order assumed)
    pvGetStr baContent, lPos                        '--- source DB filename
    m_sSql = pvGetStr(baContent, lPos)
    m_lFieldCount = pvGetLong(baContent, lPos)
    If m_lFieldCount > 0 Then
        ReDim m_aFieldName(0 To m_lFieldCount - 1)
        ReDim m_aOrigColumn(0 To m_lFieldCount - 1)
        ReDim m_aOrigTable(0 To m_lFieldCount - 1)
        ReDim m_aOrigDb(0 To m_lFieldCount - 1)
        ReDim m_aDeclType(0 To m_lFieldCount - 1)
        ReDim m_aColSchema(0 To m_lFieldCount - 1)
        ReDim m_aData(0 To m_lFieldCount - 1, 0 To IIf(lRows > 16, lRows, 16) - 1)
    End If
    For lCol = 0 To m_lFieldCount - 1
        m_aOrigDb(lCol) = pvGetStr(baContent, lPos)
        m_aOrigTable(lCol) = pvUnbracket(pvGetStr(baContent, lPos))
        m_aOrigColumn(lCol) = pvUnbracket(pvGetStr(baContent, lPos))
        m_aDeclType(lCol) = pvGetStr(baContent, lPos)
        sDefault = pvGetStr(baContent, lPos)
        sCollate = pvGetStr(baContent, lPos)
        m_aFieldName(lCol) = pvGetStr(baContent, lPos)
        pvGetLong baContent, lPos
        pvGetLong baContent, lPos                   '--- column ordinal
        lUnique = 0                                 '--- not carried by the blob
        lNotNull = pvGetByte(baContent, lPos)       '--- notnull-or-autoinc
        lPk = pvGetByte(baContent, lPos)
        lAutoInc = pvGetByte(baContent, lPos)
        lType = pvGetLong(baContent, lPos)
        pvGetByte baContent, lPos
        lWidth = pvGetLong(baContent, lPos)
        If Len(m_aOrigTable(lCol)) > 0 Then
            Set m_aColSchema(lCol) = New cColumn
            If StrComp(sDefault, "NULL", vbBinaryCompare) = 0 Then
                sDefault = vbNullString
            End If
            m_aColSchema(lCol).frInit m_aOrigColumn(lCol), m_aDeclType(lCol), sDefault, _
                lNotNull <> 0, lPk, lAutoInc <> 0, sCollate
            If lUnique <> 0 Then
                m_aColSchema(lCol).frMarkUnique
            End If
        End If
        '--- per-row not-null flags
        lCount = pvGetLong(baContent, lPos)
        ReDim aFlags(0 To IIf(lCount > 0, lCount, 1) - 1)
        For lRow = 0 To lCount - 1
            aFlags(lRow) = pvGetByte(baContent, lPos)
        Next
        '--- cell data
        Select Case lType
        Case 1
            lCount = pvGetLong(baContent, lPos)
            For lRow = 0 To lRows - 1
                If aFlags(lRow) = 0 Then
                    m_aData(lCol, lRow) = pvNullCell()
                    lPos = lPos + lWidth
                Else
                    m_aData(lCol, lRow) = pvGetIntCell(baContent, lPos, lWidth)
                End If
            Next
            lPos = lPos + lCount - lRows * lWidth   '--- spare element
            pvGetLong baContent, lPos
        Case 2
            lCount = pvGetLong(baContent, lPos)
            For lRow = 0 To lRows - 1
                If aFlags(lRow) = 0 Then
                    m_aData(lCol, lRow) = pvNullCell()
                    lPos = lPos + 8
                Else
                    m_aData(lCol, lRow) = pvGetDouble(baContent, lPos)
                End If
            Next
            lPos = lPos + lCount - lRows * 8
            pvGetLong baContent, lPos
        Case Else
            lCount = pvGetLong(baContent, lPos)
            ReDim baText(0 To IIf(lCount > 0, lCount, 1) - 1)
            For lRow = 0 To lCount - 1
                baText(lRow) = pvGetByte(baContent, lPos)
            Next
            lCount = pvGetLong(baContent, lPos) \ 4
            ReDim aOffs(0 To IIf(lCount > 0, lCount, 1) - 1)
            For lRow = 0 To lCount - 1
                aOffs(lRow) = pvGetLong(baContent, lPos)
            Next
            For lRow = 0 To lRows - 1
                If aFlags(lRow) = 0 Then
                    m_aData(lCol, lRow) = pvNullCell()
                Else
                    m_aData(lCol, lRow) = pvSliceCell(baText, aOffs(lRow), aOffs(lRow + 1) - aOffs(lRow), lType)
                End If
            Next
        End Select
    Next
    lUpdatable = pvGetLong(baContent, lPos)         '--- count of DML sets
    For lCol = 1 To lUpdatable
        pvGetStr baContent, lPos
        pvGetStr baContent, lPos
        pvGetStr baContent, lPos
        lCount = pvGetLong(baContent, lPos)
        lPos = lPos + 4 * lCount
    Next
    m_bHasColSchema = True
    m_lRecordCount = lRows
    If lCurRow < -1 Then
        lCurRow = -1
    ElseIf lCurRow > lRows Then
        lCurRow = lRows
    End If
    m_lCurRow = lCurRow
    '--- rebuild the updatability verdict from the carried schema: a single
    '--- origin table whose PK columns are all present (rowid-only
    '--- updatability is not supported by the PK-addressed UpdateBatch)
    If lUpdatable <> 0 Then
        pvAnalyzeFromSchema
    End If
    '--- no QueryFinished here: RC6 raises no events on Content load
    pvEnsureFields
End Sub

Private Sub pvAnalyzeFromSchema()
    Dim lCol            As Long

    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigTable(lCol)) > 0 Then
            If Len(m_sOrigTable) = 0 Then
                m_sOrigTable = m_aOrigTable(lCol)
                m_sOrigDb = m_aOrigDb(lCol)
            ElseIf StrComp(m_sOrigTable, m_aOrigTable(lCol), vbTextCompare) <> 0 Then
                Exit Sub
            End If
        End If
    Next
    If Len(m_sOrigTable) = 0 Then
        Exit Sub
    End If
    If Len(m_sOrigDb) = 0 Then
        m_sOrigDb = "main"
    End If
    ReDim m_aPkCols(0 To m_lFieldCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        If Not m_aColSchema(lCol) Is Nothing Then
            If m_aColSchema(lCol).PrimaryKey Then
                m_aPkCols(m_lPkCount) = lCol
                m_lPkCount = m_lPkCount + 1
            End If
        End If
    Next
    If m_lPkCount = 0 Then
        Exit Sub
    End If
    ReDim Preserve m_aPkCols(0 To m_lPkCount - 1)
    If m_lPkCount = 1 Then
        If pvDeclTypeCode(m_aDeclType(m_aPkCols(0))) = 1 Then
            m_lAutoPkCol = m_aPkCols(0)
        End If
    End If
    m_bUpdatable = True
End Sub

Private Function pvDeclTypeCode(sDeclType As String) As Long
    Dim sUpper          As String

    '--- SQLite affinity rules collapsed to the four RC6 type codes; empty
    '--- decltype (expression columns) serializes as FLOAT (RC6 behavior)
    sUpper = UCase$(sDeclType)
    If InStr(sUpper, "INT") > 0 Then
        pvDeclTypeCode = 1
    ElseIf InStr(sUpper, "CHAR") > 0 Or InStr(sUpper, "CLOB") > 0 Or InStr(sUpper, "TEXT") > 0 Then
        pvDeclTypeCode = 3
    ElseIf InStr(sUpper, "BLOB") > 0 Then
        pvDeclTypeCode = 4
    Else
        pvDeclTypeCode = 2
    End If
End Function

Private Function pvContentKeyCols(sTable As String, sDb As String, aKey() As Long) As Long
    Dim lCol            As Long
    Dim lCount          As Long
    Dim cPkNames        As VBA.Collection
    Dim hStmt           As LongPtr
    Dim bIsPk           As Boolean
    Dim oColumn         As cColumn

    ReDim aKey(0 To m_lFieldCount - 1)
    '--- collect pk column names of the origin table
    Set cPkNames = New VBA.Collection
    If Not m_oConnection Is Nothing Then
        hStmt = PrepareStatement(m_oConnection, "PRAGMA " & QuoteIdentifier(IIf(Len(sDb) > 0, sDb, "main")) & ".table_info(" & QuoteIdentifier(sTable) & ")")
        Do While stub_sqlite3_step(hStmt) = SQLITE_ROW
            If stub_sqlite3_column_int(hStmt, 5) > 0 Then
                cPkNames.Add Empty, "#" & FromUtf8Ptr(stub_sqlite3_column_text(hStmt, 1))
            End If
        Loop
        Call stub_sqlite3_finalize(hStmt)
    End If
    '--- key columns: pk ordinals in select order, else all columns
    For lCol = 0 To m_lFieldCount - 1
        bIsPk = False
        If StrComp(m_aOrigTable(lCol), sTable, vbTextCompare) = 0 Then
            If Not m_oConnection Is Nothing Then
                bIsPk = pvCollectionHasKey(cPkNames, "#" & m_aOrigColumn(lCol))
            Else
                Set oColumn = frColSchema(lCol)
                If Not oColumn Is Nothing Then
                    bIsPk = oColumn.PrimaryKey
                End If
            End If
        End If
        If bIsPk Then
            aKey(lCount) = lCol
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        For lCol = 0 To m_lFieldCount - 1
            If StrComp(m_aOrigTable(lCol), sTable, vbTextCompare) = 0 Then
                aKey(lCount) = lCol
                lCount = lCount + 1
            End If
        Next
    End If
    pvContentKeyCols = lCount
End Function

Private Function pvCollectionHasKey(cColl As VBA.Collection, sKey As String) As Boolean
    On Error GoTo QH
    cColl.Item sKey
    pvCollectionHasKey = True
QH:
End Function

Private Function pvValueTypeCode(ByVal lCol As Long) As Long
    Dim lRow            As Long

    pvValueTypeCode = 2
    For lRow = 0 To m_lRecordCount - 1
        If Not pvIsNullCell(m_aData(lCol, lRow)) Then
            Select Case VarType(m_aData(lCol, lRow))
            Case vbString
                pvValueTypeCode = 3
            Case vbByte + vbArray
                pvValueTypeCode = 4
            End Select
            Exit Function
        End If
    Next
End Function

Private Function pvIntWidth(ByVal lCol As Long) As Long
    Dim lRow            As Long
    Dim decValue        As Variant

    '--- smallest RC6 width class holding every value: 1 (unsigned byte),
    '--- 4 (int32) or 8 (int64); the PK column is always at least 4
    If lCol = m_lAutoPkCol Then
        pvIntWidth = 4
    Else
        pvIntWidth = 1
    End If
    For lRow = 0 To m_lRecordCount - 1
        If Not pvIsNullCell(m_aData(lCol, lRow)) Then
            decValue = CDec(m_aData(lCol, lRow))
            If decValue < 0 Or decValue > 255 Then
                If decValue < -2147483648# Or decValue > 2147483647# Then
                    pvIntWidth = 8
                    Exit Function
                End If
                pvIntWidth = 4
            End If
        End If
    Next
End Function

Private Function pvNullCell() As Variant
    '--- RC6 semantics: DB NULLs surface as Empty unless MapDbNullToEmpty
    '--- was disabled on the source connection
    If Not m_bMapNullToEmpty Then
        pvNullCell = Null
    End If
End Function

Private Function pvIsNullCell(vCell As Variant) As Boolean
    pvIsNullCell = IsNull(vCell) Or IsEmpty(vCell)
End Function

Private Function pvCellDbl(vCell As Variant) As Double
    If Not pvIsNullCell(vCell) Then
        pvCellDbl = CDbl(vCell)
    End If
End Function

Private Function pvCellBytes(vCell As Variant, ByVal lType As Long) As Byte()
    Dim baEmpty()       As Byte

    If pvIsNullCell(vCell) Then
        pvCellBytes = baEmpty
    ElseIf lType = 4 And VarType(vCell) = vbByte + vbArray Then
        pvCellBytes = vCell
    Else
        pvCellBytes = ToUtf8Array(CStr(vCell))
    End If
End Function

Private Function pvCellBytesLen(vCell As Variant, ByVal lType As Long) As Long
    Dim baCell()        As Byte

    baCell = pvCellBytes(vCell, lType)
    pvCellBytesLen = pvBytesLen(baCell)
End Function

Private Function pvBytesLen(baBuf() As Byte) As Long
    On Error GoTo QH
    pvBytesLen = UBound(baBuf) - LBound(baBuf) + 1
QH:
End Function

Private Function pvSliceCell(baText() As Byte, ByVal lStart As Long, ByVal lLen As Long, ByVal lType As Long) As Variant
    Dim baCell()        As Byte

    If lLen > 0 Then
        ReDim baCell(0 To lLen - 1)
        Call CopyMemory(baCell(0), baText(lStart), lLen)
    End If
    If lType = 4 Then
        pvSliceCell = baCell
    ElseIf lLen = 0 Then
        pvSliceCell = vbNullString
    Else
        pvSliceCell = FromUtf8Array(baCell)
    End If
End Function

Private Function pvStripQuotes(sText As String) As String
    If Len(sText) >= 2 And Left$(sText, 1) = "'" And Right$(sText, 1) = "'" Then
        pvStripQuotes = Replace(Mid$(sText, 2, Len(sText) - 2), "''", "'")
    Else
        pvStripQuotes = sText
    End If
End Function

Private Function pvUnbracket(sName As String) As String
    If Len(sName) >= 2 And Left$(sName, 1) = "[" And Right$(sName, 1) = "]" Then
        pvUnbracket = Mid$(sName, 2, Len(sName) - 2)
    Else
        pvUnbracket = sName
    End If
End Function

Private Sub pvPutIntCell(baBuf() As Byte, lPos As Long, vCell As Variant, ByVal lWidth As Long)
    Dim cyRaw           As Currency
    Dim baTmp(0 To 7)   As Byte

    If Not pvIsNullCell(vCell) Then
        '--- int64 raw bits via the Currency carrier (value scaled by 10000)
        cyRaw = CCur(CDec(vCell) / 10000)
        Call CopyMemory(baTmp(0), cyRaw, 8)
    End If
    pvEnsureBuf baBuf, lPos + lWidth
    Call CopyMemory(baBuf(lPos), baTmp(0), lWidth)
    lPos = lPos + lWidth
End Sub

Private Function pvGetIntCell(baBuf() As Byte, lPos As Long, ByVal lWidth As Long) As Variant
    Dim baTmp(0 To 7)   As Byte
    Dim cyRaw           As Currency
    Dim decValue        As Variant
    Dim lIdx            As Long

    Call CopyMemory(baTmp(0), baBuf(lPos), lWidth)
    lPos = lPos + lWidth
    '--- sign-extend int32/int64 widths; width 1 is unsigned
    If lWidth > 1 And lWidth < 8 Then
        If (baTmp(lWidth - 1) And &H80) <> 0 Then
            For lIdx = lWidth To 7
                baTmp(lIdx) = &HFF
            Next
        End If
    End If
    Call CopyMemory(cyRaw, baTmp(0), 8)
    decValue = CDec(cyRaw) * CDec(10000)
    If decValue >= -2147483648# And decValue <= 2147483647# Then
        pvGetIntCell = CLng(decValue)
    Else
        pvGetIntCell = decValue
    End If
End Function

Private Sub pvPutDouble(baBuf() As Byte, lPos As Long, ByVal dblValue As Double)
    pvEnsureBuf baBuf, lPos + 8
    Call CopyMemory(baBuf(lPos), dblValue, 8)
    lPos = lPos + 8
End Sub

Private Function pvGetDouble(baBuf() As Byte, lPos As Long) As Double
    Call CopyMemory(pvGetDouble, baBuf(lPos), 8)
    lPos = lPos + 8
End Function

Private Sub pvPutLong(baBuf() As Byte, lPos As Long, ByVal lValue As Long)
    pvEnsureBuf baBuf, lPos + 4
    Call CopyMemory(baBuf(lPos), lValue, 4)
    lPos = lPos + 4
End Sub

Private Function pvGetLong(baBuf() As Byte, lPos As Long) As Long
    Call CopyMemory(pvGetLong, baBuf(lPos), 4)
    lPos = lPos + 4
End Function

Private Sub pvPutByte(baBuf() As Byte, lPos As Long, ByVal bytValue As Byte)
    pvEnsureBuf baBuf, lPos + 1
    baBuf(lPos) = bytValue
    lPos = lPos + 1
End Sub

Private Function pvGetByte(baBuf() As Byte, lPos As Long) As Byte
    pvGetByte = baBuf(lPos)
    lPos = lPos + 1
End Function

Private Sub pvPutStr(baBuf() As Byte, lPos As Long, sText As String)
    pvPutLong baBuf, lPos, LenB(sText)
    If LenB(sText) > 0 Then
        pvEnsureBuf baBuf, lPos + LenB(sText)
        Call CopyMemory(baBuf(lPos), ByVal StrPtr(sText), LenB(sText))
        lPos = lPos + LenB(sText)
    End If
End Sub

Private Function pvGetStr(baBuf() As Byte, lPos As Long) As String
    Dim lSize           As Long

    lSize = pvGetLong(baBuf, lPos)
    If lSize > 0 Then
        pvGetStr = String$(lSize \ 2, 0)
        Call CopyMemory(ByVal StrPtr(pvGetStr), baBuf(lPos), lSize)
        lPos = lPos + lSize
    End If
End Function

Private Sub pvPutBytes(baBuf() As Byte, lPos As Long, baData() As Byte)
    Dim lSize           As Long

    lSize = pvBytesLen(baData)
    If lSize > 0 Then
        pvEnsureBuf baBuf, lPos + lSize
        Call CopyMemory(baBuf(lPos), baData(LBound(baData)), lSize)
        lPos = lPos + lSize
    End If
End Sub

Private Sub pvPutZeros(baBuf() As Byte, lPos As Long, ByVal lCount As Long)
    pvEnsureBuf baBuf, lPos + lCount
    lPos = lPos + lCount
End Sub

Private Sub pvEnsureBuf(baBuf() As Byte, ByVal lNeeded As Long)
    Dim lCap            As Long

    lCap = UBound(baBuf) + 1
    If lNeeded > lCap Then
        Do
            lCap = lCap * 2
        Loop While lNeeded > lCap
        ReDim Preserve baBuf(0 To lCap - 1)
    End If
End Sub

Private Sub pvNotImpl(sName As String)
    Err.Raise vbObjectError, "cRecordset." & sName, "Not implemented yet"
End Sub

Private Sub Class_Initialize()
    m_bMapNullToEmpty = True
    g_lLiveRecordsets = g_lLiveRecordsets + 1
End Sub

Private Sub Class_Terminate()
    pvReleaseFields
    g_lLiveRecordsets = g_lLiveRecordsets - 1
End Sub
