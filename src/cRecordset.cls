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
Private m_bReadOnly                 As Boolean
Private m_bAutoCreateUniqueID64     As Boolean
Private m_lFieldCount               As Long
Private m_lRecordCount              As Long
Private m_lCurRow                   As Long
Private m_aFieldName()              As String
Private m_vData()                   As Variant
'--- per-column origin metadata (captured on every open, see cField)
Private m_aOrigColumn()             As String
Private m_aOrigTable()              As String
Private m_aOrigDb()                 As String
Private m_aDeclType()               As String
Private m_cColMeta                  As VBA.Collection
'--- updatability analysis (per open)
Private m_bUpdatable                As Boolean
Private m_sOrigDb                   As String
Private m_sOrigTable                As String
Private m_alPkCols()                As Long
Private m_lPkCount                  As Long
Private m_lAutoPkCol                As Long
'--- pending-changes tracking (lazy, see pvEnsureWriteState)
Private m_bWriteState               As Boolean
Private m_vOrig()                   As Variant
Private m_abDirty()                 As Boolean
Private m_alRowState()              As Long
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
            If m_alRowState(lRow) <> ROW_CLEAN Then
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
End Property

Public Property Let Content(RHS() As Byte)
    pvNotImpl "Content (write)"
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
        m_vData(lCol, m_lRecordCount) = Null
        m_abDirty(lCol, m_lRecordCount) = False
    Next
    '--- RC6 semantics: fill the INTEGER PK with a time-based unique id
    If m_bAutoCreateUniqueID64 And m_lAutoPkCol >= 0 Then
        m_vData(m_lAutoPkCol, m_lRecordCount) = CreateUniqueID64()
    End If
    m_alRowState(m_lRecordCount) = ROW_NEW
    m_lRecordCount = m_lRecordCount + 1
    m_lCurRow = m_lRecordCount - 1
    RaiseEvent AddNew(m_lCurRow)
End Sub

Public Sub Delete()
    Dim vEntry()        As Variant
    Dim vRow()          As Variant
    Dim vOrig()         As Variant
    Dim abDirty()       As Boolean
    Dim vPk()           As Variant
    Dim lCol            As Long
    Dim lPkIdx          As Long

    pvCheckUpdatable
    If m_lCurRow < 0 Or m_lCurRow >= m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "No current row (BOF or EOF)"
    End If
    pvEnsureWriteState
    '--- a pending insert just vanishes; existing rows queue a DELETE and keep
    '--- a full snapshot so ResetChanges can restore them faithfully
    If m_alRowState(m_lCurRow) <> ROW_NEW Then
        ReDim vEntry(0 To 4)
        ReDim vRow(0 To m_lFieldCount - 1)
        ReDim vOrig(0 To m_lFieldCount - 1)
        ReDim abDirty(0 To m_lFieldCount - 1)
        For lCol = 0 To m_lFieldCount - 1
            vRow(lCol) = m_vData(lCol, m_lCurRow)
            vOrig(lCol) = m_vOrig(lCol, m_lCurRow)
            abDirty(lCol) = m_abDirty(lCol, m_lCurRow)
        Next
        ReDim vPk(0 To m_lPkCount - 1)
        For lPkIdx = 0 To m_lPkCount - 1
            vPk(lPkIdx) = pvPkWhereValue(m_lCurRow, lPkIdx)
        Next
        vEntry(0) = vRow
        vEntry(1) = vPk
        vEntry(2) = m_alRowState(m_lCurRow)
        vEntry(3) = abDirty
        vEntry(4) = vOrig
        m_cDeleted.Add vEntry
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
        Select Case m_alRowState(lRow)
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
                aOut(lRow, lCol) = m_vData(aCol(lCol), lFirst + lRow)
            Next
        Next
    Else
        ReDim aOut(0 To lNumCols - 1, 0 To lNumRows - 1)
        For lCol = 0 To lNumCols - 1
            For lRow = 0 To lNumRows - 1
                aOut(lCol, lRow) = m_vData(aCol(lCol), lFirst + lRow)
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
            If m_alRowState(lRow) = ROW_NEW Then
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
            If m_alRowState(lRow) = ROW_MODIFIED Then
                For lCol = 0 To m_lFieldCount - 1
                    If m_abDirty(lCol, lRow) Then
                        m_vData(lCol, lRow) = m_vOrig(lCol, lRow)
                        m_abDirty(lCol, lRow) = False
                        m_vOrig(lCol, lRow) = Empty
                    End If
                Next
                m_alRowState(lRow) = ROW_CLEAN
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
    pvLoad
End Sub

Friend Sub frOpenParams(oCnn As cConnection, sSql As String, vParams() As Variant)
    Dim hStmt           As LongPtr
    Dim lIdx            As Long

    Set m_oConnection = oCnn
    m_sSql = sSql
    hStmt = pvPrepare(sSql)
    On Error GoTo EH
    For lIdx = LBound(vParams) To UBound(vParams)
        If BindVariant(hStmt, lIdx - LBound(vParams) + 1, vParams(lIdx)) <> SQLITE_OK Then
            Err.Raise vbObjectError, "cRecordset", m_oConnection.LastDBError()
        End If
    Next
    pvMaterialize hStmt
    Exit Sub
EH:
    '--- pvMaterialize would have finalized; a bind error must not leak it
    Call vbsqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Friend Sub frLoadStmt(oCnn As cConnection, sSql As String, ByVal hStmt As LongPtr)
    '--- materialise from a statement owned by the caller (cSelectCommand):
    '--- the statement is read to completion but NOT finalized here
    Set m_oConnection = oCnn
    m_sSql = sSql
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
    frCellValue = m_vData(lCol, m_lCurRow)
End Function

Friend Function frCellValueAt(ByVal lRow As Long, ByVal lCol As Long) As Variant
    If lRow < 0 Or lRow >= m_lRecordCount Or lCol < 0 Or lCol >= m_lFieldCount Then
        Err.Raise 9
    End If
    frCellValueAt = m_vData(lCol, lRow)
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
    If m_alRowState(lRow) <> ROW_NEW Then
        If Not m_abDirty(lCol, lRow) Then
            m_vOrig(lCol, lRow) = m_vData(lCol, lRow)
            m_abDirty(lCol, lRow) = True
        End If
        m_alRowState(lRow) = ROW_MODIFIED
    End If
    m_vData(lCol, lRow) = vValue
    RaiseEvent FieldChange(lRow, lCol)
End Sub

Friend Function frCellDirty(ByVal lCol As Long) As Boolean
    If m_bWriteState And m_lCurRow >= 0 And m_lCurRow < m_lRecordCount Then
        If lCol >= 0 And lCol < m_lFieldCount Then
            frCellDirty = m_abDirty(lCol, m_lCurRow)
        End If
    End If
End Function

Friend Function frCellOriginal(ByVal lCol As Long) As Variant
    If frCellDirty(lCol) Then
        frCellOriginal = m_vOrig(lCol, m_lCurRow)
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

    m_lFieldCount = vbsqlite3_column_count(hStmt)
    lCap = 16
    If m_lFieldCount > 0 Then
        ReDim m_aFieldName(0 To m_lFieldCount - 1)
        For lCol = 0 To m_lFieldCount - 1
            m_aFieldName(lCol) = FromUtf8Ptr(vbsqlite3_column_name(hStmt, lCol))
        Next
        ReDim m_vData(0 To m_lFieldCount - 1, 0 To lCap - 1)
    End If
    pvCaptureColumnMeta hStmt
    pvAnalyzeUpdatable
    lRow = 0
    Do While vbsqlite3_step(hStmt) = SQLITE_ROW
        If m_lFieldCount > 0 Then
            If lRow > UBound(m_vData, 2) Then
                lCap = lCap * 2
                ReDim Preserve m_vData(0 To m_lFieldCount - 1, 0 To lCap - 1)
            End If
            For lCol = 0 To m_lFieldCount - 1
                m_vData(lCol, lRow) = ReadColumnValue(hStmt, lCol)
            Next
        End If
        lRow = lRow + 1
    Loop
    If bFinalize Then
        Call vbsqlite3_finalize(hStmt)
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
    Erase m_vData
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
    Erase m_alPkCols
    m_bWriteState = False
    Erase m_vOrig
    Erase m_abDirty
    Erase m_alRowState
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
    Dim asName()        As String
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
        asName = Split(sList, ",")
        ReDim aRet(0 To UBound(asName))
        For lIdx = 0 To UBound(asName)
            lCol = frFieldIndex(Trim$(asName(lIdx)))
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
        m_aOrigColumn(lCol) = FromUtf8Ptr(vbsqlite3_column_origin_name(hStmt, lCol))
        m_aOrigTable(lCol) = FromUtf8Ptr(vbsqlite3_column_table_name(hStmt, lCol))
        m_aOrigDb(lCol) = FromUtf8Ptr(vbsqlite3_column_database_name(hStmt, lCol))
        m_aDeclType(lCol) = FromUtf8Ptr(vbsqlite3_column_decltype(hStmt, lCol))
    Next
End Sub

Private Sub pvAnalyzeUpdatable()
    Dim lCol            As Long
    Dim asPkName()      As String
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
    asPkName = pvGetPkColumns(sPkType)
    If m_lPkCount = 0 Then
        Exit Sub
    End If
    ReDim m_alPkCols(0 To m_lPkCount - 1)
    For lPkIdx = 0 To m_lPkCount - 1
        bFound = False
        For lCol = 0 To m_lFieldCount - 1
            If StrComp(m_aOrigColumn(lCol), asPkName(lPkIdx), vbTextCompare) = 0 Then
                m_alPkCols(lPkIdx) = lCol
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
        m_lAutoPkCol = m_alPkCols(0)
    End If
    m_bUpdatable = True
End Sub

Private Function pvGetPkColumns(sSinglePkType As String) As String()
    Dim hStmt           As LongPtr
    Dim asName()        As String
    Dim alPos()         As Long
    Dim asRet()         As String
    Dim lCount          As Long
    Dim lPk             As Long
    Dim lIdx            As Long

    ReDim asName(0 To 3)
    ReDim alPos(0 To 3)
    hStmt = PrepareStatement(m_oConnection, "PRAGMA " & QuoteIdentifier(m_sOrigDb) & ".table_info(" & QuoteIdentifier(m_sOrigTable) & ")")
    Do While vbsqlite3_step(hStmt) = SQLITE_ROW
        '--- table_info columns: cid, name, type, notnull, dflt_value, pk
        lPk = vbsqlite3_column_int(hStmt, 5)
        If lPk > 0 Then
            If lCount > UBound(asName) Then
                ReDim Preserve asName(0 To 2 * lCount)
                ReDim Preserve alPos(0 To 2 * lCount)
            End If
            asName(lCount) = FromUtf8Ptr(vbsqlite3_column_text(hStmt, 1))
            alPos(lCount) = lPk
            If lPk = 1 Then
                sSinglePkType = FromUtf8Ptr(vbsqlite3_column_text(hStmt, 2))
            End If
            lCount = lCount + 1
        End If
    Loop
    Call vbsqlite3_finalize(hStmt)
    m_lPkCount = lCount
    If lCount > 0 Then
        ReDim asRet(0 To lCount - 1)
        For lIdx = 0 To lCount - 1
            asRet(alPos(lIdx) - 1) = asName(lIdx)
        Next
        pvGetPkColumns = asRet
    End If
End Function

Private Sub pvEnsureWriteState()
    If Not m_bWriteState Then
        ReDim m_vOrig(0 To m_lFieldCount - 1, 0 To UBound(m_vData, 2))
        ReDim m_abDirty(0 To m_lFieldCount - 1, 0 To UBound(m_vData, 2))
        ReDim m_alRowState(0 To UBound(m_vData, 2))
        Set m_cDeleted = New VBA.Collection
        m_bWriteState = True
    End If
End Sub

Private Sub pvEnsureRowCapacity(ByVal lRows As Long)
    Dim lCap            As Long

    lCap = UBound(m_vData, 2) + 1
    If lRows > lCap Then
        Do
            lCap = lCap * 2
        Loop While lRows > lCap
        ReDim Preserve m_vData(0 To m_lFieldCount - 1, 0 To lCap - 1)
        If m_bWriteState Then
            ReDim Preserve m_vOrig(0 To m_lFieldCount - 1, 0 To lCap - 1)
            ReDim Preserve m_abDirty(0 To m_lFieldCount - 1, 0 To lCap - 1)
            ReDim Preserve m_alRowState(0 To lCap - 1)
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
    lCol = m_alPkCols(lPkIdx)
    If m_bWriteState Then
        If m_abDirty(lCol, lRow) Then
            pvPkWhereValue = m_vOrig(lCol, lRow)
            Exit Function
        End If
    End If
    pvPkWhereValue = m_vData(lCol, lRow)
End Function

Private Function pvPkWhereClause() As String
    Dim lPkIdx          As Long
    Dim sSql            As String

    For lPkIdx = 0 To m_lPkCount - 1
        If lPkIdx > 0 Then
            sSql = sSql & " AND "
        End If
        sSql = sSql & QuoteIdentifier(m_aOrigColumn(m_alPkCols(lPkIdx))) & " = ?"
    Next
    pvPkWhereClause = sSql
End Function

Private Function pvQualifiedTable() As String
    pvQualifiedTable = QuoteIdentifier(m_sOrigDb) & "." & QuoteIdentifier(m_sOrigTable)
End Function

Private Sub pvBatchDelete(vEntry As Variant)
    Dim vParams()       As Variant

    vParams = vEntry(1)
    pvExecParams "DELETE FROM " & pvQualifiedTable() & " WHERE " & pvPkWhereClause(), vParams
End Sub

Private Sub pvBatchUpdate(ByVal lRow As Long)
    Dim sSet            As String
    Dim vParams()       As Variant
    Dim lCount          As Long
    Dim lCol            As Long
    Dim lPkIdx          As Long

    ReDim vParams(0 To m_lFieldCount + m_lPkCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        If m_abDirty(lCol, lRow) And Len(m_aOrigColumn(lCol)) > 0 Then
            If Len(sSet) > 0 Then
                sSet = sSet & ", "
            End If
            sSet = sSet & QuoteIdentifier(m_aOrigColumn(lCol)) & " = ?"
            vParams(lCount) = m_vData(lCol, lRow)
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        Exit Sub
    End If
    For lPkIdx = 0 To m_lPkCount - 1
        vParams(lCount) = pvPkWhereValue(lRow, lPkIdx)
        lCount = lCount + 1
    Next
    ReDim Preserve vParams(0 To lCount - 1)
    pvExecParams "UPDATE " & pvQualifiedTable() & " SET " & sSet & " WHERE " & pvPkWhereClause(), vParams
End Sub

Private Sub pvBatchInsert(ByVal lRow As Long)
    Dim sCols           As String
    Dim sVals           As String
    Dim vParams()       As Variant
    Dim lCount          As Long
    Dim lCol            As Long

    ReDim vParams(0 To m_lFieldCount - 1)
    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigColumn(lCol)) > 0 Then
            If lCount > 0 Then
                sCols = sCols & ", "
                sVals = sVals & ", "
            End If
            sCols = sCols & QuoteIdentifier(m_aOrigColumn(lCol))
            sVals = sVals & "?"
            vParams(lCount) = m_vData(lCol, lRow)
            lCount = lCount + 1
        End If
    Next
    If lCount = 0 Then
        Exit Sub
    End If
    ReDim Preserve vParams(0 To lCount - 1)
    pvExecParams "INSERT INTO " & pvQualifiedTable() & " (" & sCols & ") VALUES (" & sVals & ")", vParams
    '--- INTEGER PRIMARY KEY auto-assigned on NULL insert: backfill the cell
    If m_lAutoPkCol >= 0 Then
        If IsNull(m_vData(m_lAutoPkCol, lRow)) Then
            m_vData(m_lAutoPkCol, lRow) = m_oConnection.LastInsertAutoID
        End If
    End If
End Sub

Private Sub pvExecParams(sSql As String, vParams() As Variant)
    Dim hStmt           As LongPtr
    Dim lIdx            As Long
    Dim lRc             As Long

    hStmt = PrepareStatement(m_oConnection, sSql)
    On Error GoTo EH
    For lIdx = 0 To UBound(vParams)
        If BindVariant(hStmt, lIdx + 1, vParams(lIdx)) <> SQLITE_OK Then
            Err.Raise vbObjectError, "cRecordset.UpdateBatch", m_oConnection.LastDBError()
        End If
    Next
    Do
        lRc = vbsqlite3_step(hStmt)
    Loop While lRc = SQLITE_ROW
    If lRc <> SQLITE_DONE Then
        Err.Raise vbObjectError, "cRecordset.UpdateBatch", m_oConnection.LastDBError()
    End If
    Call vbsqlite3_finalize(hStmt)
    Exit Sub
EH:
    Call vbsqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Sub pvRemoveRow(ByVal lRow As Long)
    Dim lCol            As Long
    Dim lIdx            As Long

    For lIdx = lRow To m_lRecordCount - 2
        For lCol = 0 To m_lFieldCount - 1
            m_vData(lCol, lIdx) = m_vData(lCol, lIdx + 1)
            m_vOrig(lCol, lIdx) = m_vOrig(lCol, lIdx + 1)
            m_abDirty(lCol, lIdx) = m_abDirty(lCol, lIdx + 1)
        Next
        m_alRowState(lIdx) = m_alRowState(lIdx + 1)
    Next
    m_lRecordCount = m_lRecordCount - 1
End Sub

Private Sub pvRestoreDeleted()
    Dim vEntry          As Variant
    Dim vRow            As Variant
    Dim vOrig           As Variant
    Dim vDirty          As Variant
    Dim lCol            As Long
    Dim lNew            As Long

    Do While m_cDeleted.Count > 0
        vEntry = m_cDeleted.Item(1)
        m_cDeleted.Remove 1
        pvEnsureRowCapacity m_lRecordCount + 1
        lNew = m_lRecordCount
        vRow = vEntry(0)
        vDirty = vEntry(3)
        vOrig = vEntry(4)
        For lCol = 0 To m_lFieldCount - 1
            m_vData(lCol, lNew) = vRow(lCol)
            m_vOrig(lCol, lNew) = vOrig(lCol)
            m_abDirty(lCol, lNew) = vDirty(lCol)
        Next
        m_alRowState(lNew) = vEntry(2)
        m_lRecordCount = m_lRecordCount + 1
    Loop
End Sub

Private Sub pvClearWriteState()
    Dim lCol            As Long
    Dim lRow            As Long

    If m_bWriteState Then
        For lRow = 0 To m_lRecordCount - 1
            m_alRowState(lRow) = ROW_CLEAN
            For lCol = 0 To m_lFieldCount - 1
                m_abDirty(lCol, lRow) = False
                m_vOrig(lCol, lRow) = Empty
            Next
        Next
        Set m_cDeleted = New VBA.Collection
    End If
End Sub

Private Sub pvApplySort()
    Dim alCols()        As Long
    Dim abDesc()        As Boolean
    Dim lKeys           As Long
    Dim alIdx()         As Long
    Dim lRow            As Long

    If Len(Trim$(m_sSort)) = 0 Or m_lRecordCount = 0 Then
        Exit Sub
    End If
    pvParseSort alCols, abDesc, lKeys
    If lKeys = 0 Then
        Exit Sub
    End If
    ReDim alIdx(0 To m_lRecordCount - 1)
    For lRow = 0 To m_lRecordCount - 1
        alIdx(lRow) = lRow
    Next
    pvQuickSortRows alIdx, 0, m_lRecordCount - 1, alCols, abDesc, lKeys
    pvPermuteRows alIdx
End Sub

Private Sub pvParseSort(alCols() As Long, abDesc() As Boolean, lKeys As Long)
    Dim asTok()         As String
    Dim lIdx            As Long
    Dim sName           As String
    Dim sDir            As String
    Dim lPos            As Long

    asTok = Split(m_sSort, ",")
    ReDim alCols(0 To UBound(asTok))
    ReDim abDesc(0 To UBound(asTok))
    lKeys = 0
    For lIdx = 0 To UBound(asTok)
        sName = Trim$(asTok(lIdx))
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
            alCols(lKeys) = frFieldIndex(sName)
            If alCols(lKeys) < 0 Then
                Err.Raise vbObjectError, "cRecordset.Sort", "Unknown field '" & sName & "'"
            End If
            abDesc(lKeys) = (StrComp(sDir, "DESC", vbTextCompare) = 0)
            lKeys = lKeys + 1
        End If
    Next
End Sub

Private Sub pvQuickSortRows(alIdx() As Long, ByVal lLo As Long, ByVal lHi As Long, alCols() As Long, abDesc() As Boolean, ByVal lKeys As Long)
    Dim lLeft           As Long
    Dim lRight          As Long
    Dim lPivot          As Long
    Dim lTemp           As Long

    lLeft = lLo
    lRight = lHi
    lPivot = alIdx((lLo + lHi) \ 2)
    Do While lLeft <= lRight
        Do While pvCompareRows(alIdx(lLeft), lPivot, alCols, abDesc, lKeys) < 0
            lLeft = lLeft + 1
        Loop
        Do While pvCompareRows(alIdx(lRight), lPivot, alCols, abDesc, lKeys) > 0
            lRight = lRight - 1
        Loop
        If lLeft <= lRight Then
            lTemp = alIdx(lLeft)
            alIdx(lLeft) = alIdx(lRight)
            alIdx(lRight) = lTemp
            lLeft = lLeft + 1
            lRight = lRight - 1
        End If
    Loop
    If lLo < lRight Then
        pvQuickSortRows alIdx, lLo, lRight, alCols, abDesc, lKeys
    End If
    If lLeft < lHi Then
        pvQuickSortRows alIdx, lLeft, lHi, alCols, abDesc, lKeys
    End If
End Sub

Private Function pvCompareRows(ByVal lRow1 As Long, ByVal lRow2 As Long, alCols() As Long, abDesc() As Boolean, ByVal lKeys As Long) As Long
    Dim lIdx            As Long
    Dim lCmp            As Long

    For lIdx = 0 To lKeys - 1
        lCmp = pvCompareValues(m_vData(alCols(lIdx), lRow1), m_vData(alCols(lIdx), lRow2))
        If abDesc(lIdx) Then
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

Private Sub pvPermuteRows(alIdx() As Long)
    Dim vNew()          As Variant
    Dim vOrigNew()      As Variant
    Dim abDirtyNew()    As Boolean
    Dim alStateNew()    As Long
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lNewCur         As Long

    ReDim vNew(0 To m_lFieldCount - 1, 0 To UBound(m_vData, 2))
    lNewCur = m_lCurRow
    For lRow = 0 To m_lRecordCount - 1
        For lCol = 0 To m_lFieldCount - 1
            vNew(lCol, lRow) = m_vData(lCol, alIdx(lRow))
        Next
        If alIdx(lRow) = m_lCurRow Then
            lNewCur = lRow
        End If
    Next
    m_vData = vNew
    If m_bWriteState Then
        ReDim vOrigNew(0 To m_lFieldCount - 1, 0 To UBound(m_vOrig, 2))
        ReDim abDirtyNew(0 To m_lFieldCount - 1, 0 To UBound(m_abDirty, 2))
        ReDim alStateNew(0 To UBound(m_alRowState))
        For lRow = 0 To m_lRecordCount - 1
            For lCol = 0 To m_lFieldCount - 1
                vOrigNew(lCol, lRow) = m_vOrig(lCol, alIdx(lRow))
                abDirtyNew(lCol, lRow) = m_abDirty(lCol, alIdx(lRow))
            Next
            alStateNew(lRow) = m_alRowState(alIdx(lRow))
        Next
        m_vOrig = vOrigNew
        m_abDirty = abDirtyNew
        m_alRowState = alStateNew
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
        If pvMatchValue(m_vData(lCol, lRow), lOp, vValue, bDistinctNull) Then
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
        pvMatchValue = IsNull(vCell)
    Case OP_ISNOTNULL
        pvMatchValue = Not IsNull(vCell)
    Case OP_LIKE
        If Not IsNull(vCell) Then
            pvMatchValue = LCase$(CStr(vCell)) Like pvLikePattern(LCase$(CStr(vValue)))
        End If
    Case Else
        If IsNull(vCell) Or IsNull(vValue) Then
            '--- SQL three-valued logic (never matches) unless
            '--- DistinctNullValues=False treats Null as a regular value
            If Not bDistinctNull Then
                If IsNull(vCell) And IsNull(vValue) Then
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

Private Sub pvNotImpl(sName As String)
    Err.Raise vbObjectError, "cRecordset." & sName, "Not implemented yet"
End Sub

Private Sub Class_Initialize()
    g_lLiveRecordsets = g_lLiveRecordsets + 1
End Sub

Private Sub Class_Terminate()
    pvReleaseFields
    g_lLiveRecordsets = g_lLiveRecordsets - 1
End Sub
