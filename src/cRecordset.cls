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
' once the recordset is gone (verified against RC6.dll 6.0.15), which is
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

'--- int64 VARIANT type (no VB6 vbLongLong constant)
Private Const VT_I8                         As Integer = 20

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (pbBinary As Any, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As LongPtr, pcchString As Long) As Long

Private Enum eRowState
    ROW_CLEAN = 0
    ROW_MODIFIED = 1
    ROW_NEW = 2
End Enum

'--- RC6 Content blob column type codes (see NOTES)
Private Enum eDeclType
    TYPE_INTEGER = 1
    TYPE_FLOAT = 2
    TYPE_TEXT = 3
    TYPE_BLOB = 4
    TYPE_DATE = 6
    TYPE_BOOLEAN = 7
End Enum

'--- access-time value coercion kinds (RC6 reader semantics, see NOTES)
Private Enum eCoerceKind
    COERCE_NONE = 0
    COERCE_DATE = 1
    COERCE_BOOLEAN = 2
    COERCE_BYTE = 3
    COERCE_INT64 = 4
End Enum

'--- ContentChangesOnly pending-op kinds
Private Enum eChangeOp
    CHG_DELETE = 0
    CHG_UPDATE = 1
    CHG_INSERT = 2
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
'--- stable physical row labels (RC6 slot ids: survive deletes/sorts, new
'--- rows get fresh labels past every allocated one)
Private m_aRowLabel()               As Long
Private m_lNextLabel                As Long
Private m_lLoadedRows               As Long
'--- per-column access-time coercion kind (RC6 semantics, see NOTES):
'--- 0 = none, 1 = date, 2 = boolean, 3 = byte width, 4 = int64 width
Private m_aColCoerce()              As eCoerceKind
'--- pending ops loaded from a ContentChangesOnly blob (see NOTES format)
Private m_lChgCellCount             As Long
Private m_aChgCellRow()             As Long
Private m_aChgCellCol()             As Long
Private m_aChgCellValue()           As Variant
Private m_lChgRowCount              As Long
Private m_aChgRowIdx()              As Long
Private m_aChgRowPk()               As Variant
Private m_aChgRowDelete()           As Boolean
Private m_lChgKeyCol                As Long

Public Property Get ContainsChanges() As Boolean
    Dim lRow            As Long

    If m_lChgCellCount > 0 Then
        ContainsChanges = True
        Exit Property
    End If
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
    '--- RC6 semantics: 1-based; -1 = empty recordset, -2 = BOF, -3 = EOF
    If m_lRecordCount = 0 Then
        AbsolutePosition = -1
    ElseIf m_lCurRow < 0 Then
        AbsolutePosition = -2
    ElseIf m_lCurRow >= m_lRecordCount Then
        AbsolutePosition = -3
    Else
        AbsolutePosition = m_lCurRow + 1
    End If
End Property

Public Property Let AbsolutePosition(ByVal RHS As Long)
    If RHS < 1 Or RHS > m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "Invalid position"
    End If
    If m_lCurRow <> RHS - 1 Then
        m_lCurRow = RHS - 1
        RaiseEvent Move(m_lCurRow)
    End If
End Property

Public Property Get Bookmark() As Long
    Bookmark = AbsolutePosition
End Property

Public Property Let Bookmark(ByVal RHS As Long)
    AbsolutePosition = RHS
End Property

Public Property Get Sort() As String
    Sort = m_sSort
End Property

Public Property Let Sort(ByVal RHS As String)
    m_sSort = RHS
    pvApplySort False
    '--- RC6 semantics: setting Sort rewinds to the first row and always
    '--- raises Move(0)
    If m_lRecordCount > 0 Then
        m_lCurRow = 0
        RaiseEvent Move(0)
    End If
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
    ContentChangesOnly = pvContentWriteChanges()
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
        m_aDirty(m_lAutoPkCol, m_lRecordCount) = True
    End If
    m_aRowState(m_lRecordCount) = ROW_NEW
    m_aRowLabel(m_lRecordCount) = m_lNextLabel
    m_lNextLabel = m_lNextLabel + 1
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
        ReDim aEntry(0 To 5) As Variant
        ReDim aRow(0 To m_lFieldCount - 1) As Variant
        ReDim aOrig(0 To m_lFieldCount - 1) As Variant
        ReDim aDirty(0 To m_lFieldCount - 1) As Boolean
        For lCol = 0 To m_lFieldCount - 1
            aRow(lCol) = m_aData(lCol, m_lCurRow)
            aOrig(lCol) = m_aOriginal(lCol, m_lCurRow)
            aDirty(lCol) = m_aDirty(lCol, m_lCurRow)
        Next
        ReDim aPk(0 To m_lPkCount - 1) As Variant
        For lPkIdx = 0 To m_lPkCount - 1
            aPk(lPkIdx) = pvPkWhereValue(m_lCurRow, lPkIdx)
        Next
        aEntry(0) = aRow
        aEntry(1) = aPk
        aEntry(2) = m_aRowState(m_lCurRow)
        aEntry(3) = aDirty
        aEntry(4) = aOrig
        aEntry(5) = m_aRowLabel(m_lCurRow)
        m_cDeleted.Add aEntry
    End If
    pvRemoveRow m_lCurRow
    '--- RC6 semantics: the cursor stays put (next row slides in) or clamps
    '--- to the new last row; the event carries the new position (-1 when
    '--- the recordset just became empty)
    If m_lCurRow > m_lRecordCount - 1 Then
        m_lCurRow = m_lRecordCount - 1
    End If
    RaiseEvent Delete(pvEventPos())
End Sub

Public Sub UpdateBatch()
    Dim vElem           As Variant
    Dim lRow            As Long

    If m_lChgCellCount > 0 Then
        pvApplyChangesOnly
        Exit Sub
    End If
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
    '--- RC6 semantics: silent no-op on an empty recordset; Move raised only
    '--- when the position actually changes
    If m_lRecordCount = 0 Then
        Exit Sub
    End If
    If m_lCurRow <> 0 Then
        m_lCurRow = 0
        RaiseEvent Move(0)
    End If
End Sub

Public Sub MoveLast()
    If m_lRecordCount = 0 Then
        Exit Sub
    End If
    If m_lCurRow <> m_lRecordCount - 1 Then
        m_lCurRow = m_lRecordCount - 1
        RaiseEvent Move(m_lCurRow)
    End If
End Sub

Public Sub MoveNext()
    '--- RC6 semantics: moving past the last row lands on EOF (Move(-2));
    '--- another MoveNext raises an error
    If m_lRecordCount = 0 Or m_lCurRow >= m_lRecordCount Then
        Err.Raise vbObjectError, "cRecordset", "Already at EOF"
    End If
    m_lCurRow = m_lCurRow + 1
    RaiseEvent Move(pvEventPos())
End Sub

Public Sub MovePrevious()
    If m_lRecordCount = 0 Or m_lCurRow < 0 Then
        Err.Raise vbObjectError, "cRecordset", "Already at BOF"
    End If
    m_lCurRow = m_lCurRow - 1
    RaiseEvent Move(pvEventPos())
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
        ReDim aOut(0 To lNumRows - 1, 0 To lNumCols - 1) As Variant
        For lRow = 0 To lNumRows - 1
            For lCol = 0 To lNumCols - 1
                aOut(lRow, lCol) = pvCoerceCell(aCol(lCol), m_aData(aCol(lCol), lFirst + lRow))
            Next
        Next
    Else
        ReDim aOut(0 To lNumCols - 1, 0 To lNumRows - 1) As Variant
        For lCol = 0 To lNumCols - 1
            For lRow = 0 To lNumRows - 1
                aOut(lCol, lRow) = pvCoerceCell(aCol(lCol), m_aData(aCol(lCol), lFirst + lRow))
            Next
        Next
    End If
    GetRows = aOut
End Function

Public Function GetRowsWithHeaders(Optional ByVal RowCount As Long = -1, Optional ByVal StartIdxZeroBased As Long, Optional FieldListCommaSeparated As String, Optional ByVal TransPosed As Boolean, Optional ByVal HeaderAtIdxMinus1 As Boolean) As Variant()
    Dim aCol()          As Long
    Dim lNumCols        As Long
    Dim lFirst          As Long
    Dim lLast           As Long
    Dim lNumRows        As Long
    Dim aOut()          As Variant
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lBase           As Long

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
    If lNumRows < 0 Then
        lNumRows = 0
    End If
    If lNumCols = 0 Then
        Exit Function
    End If
    '--- RC6 semantics: header row/col of field names precedes the data,
    '--- at index -1 instead of 0 with HeaderAtIdxMinus1
    lBase = IIf(HeaderAtIdxMinus1, -1, 0)
    If TransPosed Then
        ReDim aOut(lBase To lBase + lNumRows, 0 To lNumCols - 1) As Variant
        For lCol = 0 To lNumCols - 1
            aOut(lBase, lCol) = m_aFieldName(aCol(lCol))
            For lRow = 0 To lNumRows - 1
                aOut(lBase + 1 + lRow, lCol) = pvCoerceCell(aCol(lCol), m_aData(aCol(lCol), lFirst + lRow))
            Next
        Next
    Else
        ReDim aOut(0 To lNumCols - 1, lBase To lBase + lNumRows) As Variant
        For lCol = 0 To lNumCols - 1
            aOut(lCol, lBase) = m_aFieldName(aCol(lCol))
            For lRow = 0 To lNumRows - 1
                aOut(lCol, lBase + 1 + lRow) = pvCoerceCell(aCol(lCol), m_aData(aCol(lCol), lFirst + lRow))
            Next
        Next
    End If
    GetRowsWithHeaders = aOut
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
    '--- reverting sort-key edits may invalidate the current order (RC6
    '--- re-sorts too, with its own Move storm)
    If Len(Trim$(m_sSort)) > 0 Then
        pvApplySort
    End If
    RaiseEvent Reset
End Sub

Public Function DataSource() As Object
    Set DataSource = GetADORsFromContent()
End Function

Public Function GetADORsFromContent() As Object
    Dim oArs            As Object
    Dim lCol            As Long
    Dim lRow            As Long
    Dim eType           As eDeclType
    Dim eAdoType        As eAdoDataType
    Dim eAttrib         As eAdoFieldAttrib
    Dim oColumn         As cColumn
    Dim vValue          As Variant

    Set oArs = CreateObject("ADODB.Recordset")
    For lCol = 0 To m_lFieldCount - 1
        eType = pvDeclTypeCode(m_aDeclType(lCol))
        If Len(m_aDeclType(lCol)) = 0 Then
            eType = pvValueTypeCode(lCol)
        End If
        Set oColumn = frColSchema(lCol)
        eAttrib = adFldUpdatable Or adFldIsNullable
        eAdoType = 0
        Select Case eType
        Case TYPE_INTEGER
            If Not oColumn Is Nothing Then
                If oColumn.PrimaryKey Then
                    eAdoType = adInteger
                    eAttrib = eAttrib Or adFldFixed Or adFldKeyColumn
                End If
            End If
            If eAdoType = 0 Then
                eAdoType = adDecimal
                eAttrib = eAttrib Or adFldFixed
            End If
        Case TYPE_FLOAT
            eAdoType = adDouble
            eAttrib = eAttrib Or adFldFixed
        Case TYPE_BLOB
            eAdoType = adLongVarBinary
            eAttrib = eAttrib Or adFldLong
        Case Else
            eAdoType = adLongVarWChar
            eAttrib = eAttrib Or adFldLong
        End Select
        If eAdoType = adLongVarWChar Or eAdoType = adLongVarBinary Then
            oArs.Fields.Append m_aFieldName(lCol), eAdoType, -1, eAttrib
        Else
            oArs.Fields.Append m_aFieldName(lCol), eAdoType, 0, eAttrib
        End If
    Next
    oArs.Open
    For lRow = 0 To m_lRecordCount - 1
        oArs.AddNew
        For lCol = 0 To m_lFieldCount - 1
            vValue = m_aData(lCol, lRow)
            If Not pvIsNullCell(vValue) Then
                If VarType(vValue) = VT_I8 Then
                    vValue = CDec(vValue)
                End If
                oArs.Fields(lCol).Value = vValue
            End If
        Next
        oArs.Update
    Next
    If m_lRecordCount > 0 Then
        oArs.MoveFirst
    End If
    Set GetADORsFromContent = oArs
End Function

Public Function ToJSONUTF8(Optional ByVal ValueArrAsColsRows As Boolean, Optional ByVal Indent As Byte, Optional ByVal UniEscaping As Boolean, Optional ByVal LfChar As Byte) As Byte()
    Dim sLf             As String
    Dim sI1             As String
    Dim sI2             As String
    Dim sI3             As String
    Dim aFields()       As String
    Dim aPairs(0 To 4)  As String
    Dim aParts()        As String
    Dim aElems()        As String
    Dim lCol            As Long
    Dim lRow            As Long
    Dim oColumn         As cColumn
    Dim sText           As String
    Dim sJson           As String
    Dim sKey            As String
    Dim lCount          As Long
    Dim lIdx            As Long

    If Indent > 0 Then
        If LfChar = 0 Then
            sLf = vbCrLf
        Else
            sLf = Chr$(LfChar)
        End If
        sI1 = Space$(Indent)
        sI2 = Space$(2 * CLng(Indent))
        '--- continuation lines align under the "{ "/"[ " opener
        sI3 = Space$(2 * CLng(Indent) + 2)
    End If
    '--- field descriptors (expression columns: empty Type/DefaultValue)
    ReDim aFields(0 To m_lFieldCount) As String
    For lCol = 0 To m_lFieldCount - 1
        Set oColumn = frColSchema(lCol)
        aPairs(0) = """Name"": " & pvJsonString(m_aFieldName(lCol), UniEscaping)
        aPairs(1) = """Type"": " & pvJsonString(m_aDeclType(lCol), UniEscaping)
        If oColumn Is Nothing Then
            aPairs(2) = """PrimaryKey"": false"
            aPairs(3) = """Nullable"": true"
        Else
            aPairs(2) = """PrimaryKey"": " & IIf(oColumn.PrimaryKey, "true", "false")
            aPairs(3) = """Nullable"": " & IIf(oColumn.NotNullConstraint, "false", "true")
        End If
        sText = vbNullString
        If Len(m_aOrigTable(lCol)) > 0 Then
            If Not oColumn Is Nothing Then
                sText = pvStripQuotes(oColumn.DefaultValue)
            End If
            If Len(sText) = 0 Then
                sText = "NULL"
            End If
        End If
        aPairs(4) = """DefaultValue"": " & pvJsonString(sText, UniEscaping)
        If Indent = 0 Then
            aFields(lCol) = "{ " & Join(aPairs, ", ") & "}"
        Else
            aFields(lCol) = sI2 & "{ " & aPairs(0)
            For lIdx = 1 To UBound(aPairs)
                aFields(lCol) = aFields(lCol) & "," & sLf & sI3 & aPairs(lIdx)
            Next
            aFields(lCol) = aFields(lCol) & sLf & sI2 & "}"
        End If
    Next
    ReDim Preserve aFields(0 To IIf(m_lFieldCount > 0, m_lFieldCount - 1, 0)) As String
    '--- value matrix: rows of cells, or columns of row-values (an empty
    '--- recordset serializes as [] in both modes)
    If ValueArrAsColsRows Then
        sKey = "ColsRows"
        lCount = m_lFieldCount
    Else
        sKey = "RowsCols"
        lCount = m_lRecordCount
    End If
    If m_lRecordCount = 0 Then
        lCount = 0
    End If
    ReDim aParts(0 To lCount) As String
    For lIdx = 0 To lCount - 1
        If ValueArrAsColsRows Then
            ReDim aElems(0 To m_lRecordCount) As String
            For lRow = 0 To m_lRecordCount - 1
                aElems(lRow) = pvJsonValue(lIdx, m_aData(lIdx, lRow), UniEscaping)
            Next
            ReDim Preserve aElems(0 To IIf(m_lRecordCount > 0, m_lRecordCount - 1, 0)) As String
        Else
            ReDim aElems(0 To m_lFieldCount) As String
            For lCol = 0 To m_lFieldCount - 1
                aElems(lCol) = pvJsonValue(lCol, m_aData(lCol, lIdx), UniEscaping)
            Next
            ReDim Preserve aElems(0 To IIf(m_lFieldCount > 0, m_lFieldCount - 1, 0)) As String
        End If
        If Indent = 0 Then
            aParts(lIdx) = "[ " & Join(aElems, ",  ") & "]"
        Else
            aParts(lIdx) = sI2 & "[ " & aElems(0)
            For lCol = 1 To UBound(aElems)
                aParts(lIdx) = aParts(lIdx) & "," & sLf & sI3 & aElems(lCol)
            Next
            aParts(lIdx) = aParts(lIdx) & sLf & sI2 & "]"
        End If
    Next
    ReDim Preserve aParts(0 To IIf(lCount > 0, lCount - 1, 0)) As String
    If Indent = 0 Then
        sJson = "{""RecordCount"": " & m_lRecordCount & ",""Fields"": [" & Join(aFields, ",") & "],""" & sKey & """: [" & Join(aParts, ",") & "]}"
    ElseIf lCount = 0 Then
        sJson = "{" & sLf & sI1 & """RecordCount"": " & m_lRecordCount & "," & sLf & sI1 & """Fields"": [" & sLf & Join(aFields, "," & sLf) & sLf & sI1 & "]," & sLf & sI1 & """" & sKey & """: []" & sLf & "}"
    Else
        sJson = "{" & sLf & sI1 & """RecordCount"": " & m_lRecordCount & "," & sLf & sI1 & """Fields"": [" & sLf & Join(aFields, "," & sLf) & sLf & sI1 & "]," & sLf & sI1 & """" & sKey & """: [" & sLf & Join(aParts, "," & sLf) & sLf & sI1 & "]" & sLf & "}"
    End If
    ToJSONUTF8 = ToUtf8Array(sJson)
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
    frCellValue = pvCoerceCell(lCol, m_aData(lCol, m_lCurRow))
End Function

Friend Function frCellValueAt(ByVal lRow As Long, ByVal lCol As Long) As Variant
    If lRow < 0 Or lRow >= m_lRecordCount Or lCol < 0 Or lCol >= m_lFieldCount Then
        Err.Raise 9
    End If
    frCellValueAt = pvCoerceCell(lCol, m_aData(lCol, lRow))
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
    '--- RC6 semantics: assigning the current value is a complete no-op (no
    '--- dirty flag, no FieldChange)
    If pvCellValueEquals(m_aData(lCol, lRow), vValue) Then
        Exit Sub
    End If
    pvEnsureWriteState
    '--- first change to an existing row's cell captures the original value;
    '--- new-row assignments only mark the cell (unassigned cells stay off
    '--- the ContentChangesOnly blob like RC6's)
    If m_aRowState(lRow) <> ROW_NEW Then
        If Not m_aDirty(lCol, lRow) Then
            m_aOriginal(lCol, lRow) = m_aData(lCol, lRow)
            m_aDirty(lCol, lRow) = True
        End If
        m_aRowState(lRow) = ROW_MODIFIED
    Else
        m_aDirty(lCol, lRow) = True
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
        ReDim m_aFieldName(0 To m_lFieldCount - 1) As String
        For lCol = 0 To m_lFieldCount - 1
            m_aFieldName(lCol) = FromUtf8Ptr(stub_sqlite3_column_name(hStmt, lCol))
        Next
        ReDim m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1) As Variant
        ReDim m_aRowLabel(0 To lCap - 1) As Long
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
                ReDim Preserve m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1) As Variant
                ReDim Preserve m_aRowLabel(0 To lCap - 1) As Long
            End If
            m_aRowLabel(lRow) = lRow
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
    m_lNextLabel = lRow
    m_lLoadedRows = lRow
    pvComputeCoercion
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
    Erase m_aRowLabel
    m_lNextLabel = 0
    m_lLoadedRows = 0
    m_lChgCellCount = 0
    m_lChgRowCount = 0
    m_lChgKeyCol = -1
    Erase m_aChgCellRow
    Erase m_aChgCellCol
    Erase m_aChgCellValue
    Erase m_aChgRowIdx
    Erase m_aChgRowPk
    Erase m_aChgRowDelete
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
        ReDim aRet(0 To m_lFieldCount - 1) As Long
        For lIdx = 0 To m_lFieldCount - 1
            aRet(lIdx) = lIdx
        Next
        lNumCols = m_lFieldCount
    Else
        aName = Split(sList, ",")
        ReDim aRet(0 To UBound(aName)) As Long
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
        ReDim Preserve aRet(0 To lNumCols - 1) As Long
    End If
    pvResolveColumns = aRet
End Function

Private Sub pvCaptureColumnMeta(ByVal hStmt As LongPtr)
    Dim lCol            As Long

    If m_lFieldCount = 0 Then
        Exit Sub
    End If
    ReDim m_aOrigColumn(0 To m_lFieldCount - 1) As String
    ReDim m_aOrigTable(0 To m_lFieldCount - 1) As String
    ReDim m_aOrigDb(0 To m_lFieldCount - 1) As String
    ReDim m_aDeclType(0 To m_lFieldCount - 1) As String
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
    ReDim m_aPkCols(0 To m_lPkCount - 1) As Long
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

    ReDim aName(0 To 3) As String
    ReDim aPos(0 To 3) As Long
    hStmt = PrepareStatement(m_oConnection, "PRAGMA " & QuoteIdentifier(m_sOrigDb) & ".table_info(" & QuoteIdentifier(m_sOrigTable) & ")")
    Do While stub_sqlite3_step(hStmt) = SQLITE_ROW
        '--- table_info columns: cid, name, type, notnull, dflt_value, pk
        lPk = stub_sqlite3_column_int(hStmt, 5)
        If lPk > 0 Then
            If lCount > UBound(aName) Then
                ReDim Preserve aName(0 To 2 * lCount) As String
                ReDim Preserve aPos(0 To 2 * lCount) As Long
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
        ReDim aRet(0 To lCount - 1) As String
        For lIdx = 0 To lCount - 1
            aRet(aPos(lIdx) - 1) = aName(lIdx)
        Next
        pvGetPkColumns = aRet
    End If
End Function

Private Sub pvEnsureWriteState()
    If Not m_bWriteState Then
        ReDim m_aOriginal(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2)) As Variant
        ReDim m_aDirty(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2)) As Boolean
        ReDim m_aRowState(0 To UBound(m_aData, 2)) As Long
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
        ReDim Preserve m_aData(0 To m_lFieldCount - 1, 0 To lCap - 1) As Variant
        ReDim Preserve m_aRowLabel(0 To lCap - 1) As Long
        If m_bWriteState Then
            ReDim Preserve m_aOriginal(0 To m_lFieldCount - 1, 0 To lCap - 1) As Variant
            ReDim Preserve m_aDirty(0 To m_lFieldCount - 1, 0 To lCap - 1) As Boolean
            ReDim Preserve m_aRowState(0 To lCap - 1) As Long
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

    ReDim aParams(0 To m_lFieldCount + m_lPkCount - 1) As Variant
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
    ReDim Preserve aParams(0 To lCount - 1) As Variant
    pvExecParams "UPDATE " & pvQualifiedTable() & " SET " & sSet & " WHERE " & pvPkWhereClause(), aParams
End Sub

Private Sub pvBatchInsert(ByVal lRow As Long)
    Dim sCols           As String
    Dim sVals           As String
    Dim aParams()       As Variant
    Dim lCount          As Long
    Dim lCol            As Long

    ReDim aParams(0 To m_lFieldCount - 1) As Variant
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
    ReDim Preserve aParams(0 To lCount - 1) As Variant
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
        m_aRowLabel(lIdx) = m_aRowLabel(lIdx + 1)
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
        m_aRowLabel(lNew) = vEntry(5)
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

Private Sub pvApplySort(Optional ByVal bRaiseMove As Boolean = True)
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
    ReDim aIdx(0 To m_lRecordCount - 1) As Long
    For lRow = 0 To m_lRecordCount - 1
        aIdx(lRow) = lRow
    Next
    pvQuickSortRows aIdx, 0, m_lRecordCount - 1, aCols, aDesc, lKeys
    pvPermuteRows aIdx, bRaiseMove
End Sub

Private Sub pvParseSort(aCols() As Long, aDesc() As Boolean, lKeys As Long)
    Dim aTok()          As String
    Dim lIdx            As Long
    Dim sName           As String
    Dim sDir            As String
    Dim lPos            As Long

    aTok = Split(m_sSort, ",")
    ReDim aCols(0 To UBound(aTok)) As Long
    ReDim aDesc(0 To UBound(aTok)) As Boolean
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

Private Sub pvPermuteRows(aIdx() As Long, Optional ByVal bRaiseMove As Boolean = True)
    Dim aDataNew()      As Variant
    Dim aOriginalNew()  As Variant
    Dim aDirtyNew()     As Boolean
    Dim aStateNew()     As Long
    Dim aLabelNew()     As Long
    Dim lRow            As Long
    Dim lCol            As Long
    Dim lNewCur         As Long

    ReDim aDataNew(0 To m_lFieldCount - 1, 0 To UBound(m_aData, 2)) As Variant
    ReDim aLabelNew(0 To UBound(m_aRowLabel)) As Long
    lNewCur = m_lCurRow
    For lRow = 0 To m_lRecordCount - 1
        For lCol = 0 To m_lFieldCount - 1
            aDataNew(lCol, lRow) = m_aData(lCol, aIdx(lRow))
        Next
        aLabelNew(lRow) = m_aRowLabel(aIdx(lRow))
        If aIdx(lRow) = m_lCurRow Then
            lNewCur = lRow
        End If
    Next
    m_aData = aDataNew
    m_aRowLabel = aLabelNew
    If m_bWriteState Then
        ReDim aOriginalNew(0 To m_lFieldCount - 1, 0 To UBound(m_aOriginal, 2)) As Variant
        ReDim aDirtyNew(0 To m_lFieldCount - 1, 0 To UBound(m_aDirty, 2)) As Boolean
        ReDim aStateNew(0 To UBound(m_aRowState)) As Long
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
            If bRaiseMove Then
                RaiseEvent Move(m_lCurRow)
            End If
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
            If m_lCurRow <> lRow Then
                m_lCurRow = lRow
                RaiseEvent Move(lRow)
            End If
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
    Dim eType           As eDeclType
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
    ReDim baOut(0 To 1023) As Byte
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
    lTables = pvCollectTables(aTabCol)
    '--- fields
    pvPutLong baOut, lPos, m_lFieldCount
    For lCol = 0 To m_lFieldCount - 1
        pvWriteFieldMeta baOut, lPos, lCol, aTabCol, lTables, eType, lWidth
        '--- per-row not-null flags (one spare element)
        pvPutLong baOut, lPos, lRows + 1
        For lRow = 0 To lRows - 1
            pvPutByte baOut, lPos, IIf(pvIsNullCell(m_aData(lCol, lRow)), 0, 1)
        Next
        pvPutByte baOut, lPos, 0
        '--- cell data
        Select Case eType
        Case TYPE_INTEGER, TYPE_BOOLEAN
            pvPutLong baOut, lPos, (lRows + 1) * lWidth
            For lRow = 0 To lRows - 1
                pvPutIntCell baOut, lPos, m_aData(lCol, lRow), lWidth
            Next
            pvPutZeros baOut, lPos, lWidth
            pvPutLong baOut, lPos, 0
        Case TYPE_FLOAT
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
                lTotal = lTotal + pvCellBytesLen(m_aData(lCol, lRow), eType)
            Next
            pvPutLong baOut, lPos, lTotal
            For lRow = 0 To lRows - 1
                baCell = pvCellBytes(m_aData(lCol, lRow), eType)
                If pvBytesLen(baCell) > 0 Then
                    pvPutBytes baOut, lPos, baCell
                End If
            Next
            pvPutLong baOut, lPos, (lRows + 1) * 4
            lOffs = 0
            For lRow = 0 To lRows - 1
                pvPutLong baOut, lPos, lOffs
                lOffs = lOffs + pvCellBytesLen(m_aData(lCol, lRow), eType)
            Next
            pvPutLong baOut, lPos, lOffs
        End Select
    Next
    pvWriteDmlSets baOut, lPos, aTabCol, lTables
    pvPutZeros baOut, lPos, 42
    If (lPos And 1) <> 0 Then
        pvPutZeros baOut, lPos, 1
    End If
    ReDim Preserve baOut(0 To lPos - 1) As Byte
    Call CopyMemory(baOut(0), lPos, 4)
    pvContentWrite = baOut
End Function

Private Function pvCollectTables(aTabCol() As Long) As Long
    Dim lCol            As Long
    Dim lIdx            As Long
    Dim lTables         As Long

    '--- distinct origin tables in first-appearance order (aTabCol holds the
    '--- first column index of each table)
    ReDim aTabCol(0 To m_lFieldCount) As Long
    For lCol = 0 To m_lFieldCount - 1
        If Len(m_aOrigTable(lCol)) > 0 Then
            For lIdx = 0 To lTables - 1
                If StrComp(m_aOrigTable(aTabCol(lIdx)), m_aOrigTable(lCol), vbTextCompare) = 0 Then
                    Exit For
                End If
            Next
            If lIdx = lTables Then
                aTabCol(lTables) = lCol
                lTables = lTables + 1
            End If
        End If
    Next
    pvCollectTables = lTables
End Function

Private Sub pvWriteFieldMeta(baOut() As Byte, lPos As Long, ByVal lCol As Long, aTabCol() As Long, ByVal lTables As Long, eType As eDeclType, lWidth As Long, Optional ByVal bLoadedOnly As Boolean)
    Dim oColumn         As cColumn
    Dim sText           As String
    Dim lIdx            As Long
    Dim lTabIdx         As Long

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
    lTabIdx = 0
    For lIdx = 0 To lTables - 1
        If StrComp(m_aOrigTable(aTabCol(lIdx)), m_aOrigTable(lCol), vbTextCompare) = 0 Then
            lTabIdx = lIdx
            Exit For
        End If
    Next
    pvPutLong baOut, lPos, lTabIdx
    pvPutLong baOut, lPos, lCol                     '--- column ordinal
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
    eType = pvDeclTypeCode(m_aDeclType(lCol))
    If Len(m_aDeclType(lCol)) = 0 Then
        '--- expression columns type by value storage class (RC6: text
        '--- stays text, blobs stay blobs, all numerics become doubles)
        eType = pvValueTypeCode(lCol)
    End If
    pvPutLong baOut, lPos, eType
    pvPutByte baOut, lPos, 0
    lWidth = 0
    If eType = TYPE_INTEGER Or eType = TYPE_BOOLEAN Then
        lWidth = pvIntWidth(lCol, bLoadedOnly)
    End If
    pvPutLong baOut, lPos, lWidth
End Sub

Private Sub pvWriteDmlSets(baOut() As Byte, lPos As Long, aTabCol() As Long, ByVal lTables As Long)
    Dim lIdx            As Long
    Dim lCol            As Long
    Dim lCount          As Long
    Dim sText           As String
    Dim aKeyCols()      As Long

    '--- one DML set per distinct origin table: 3 DML templates + key column
    '--- ordinals (the table's pk columns in the select, or all of its
    '--- select columns when no pk)
    pvPutLong baOut, lPos, lTables
    For lIdx = 0 To lTables - 1
        sText = m_aOrigTable(aTabCol(lIdx))
        pvPutStr baOut, lPos, "DELETE FROM [" & sText & "] "
        pvPutStr baOut, lPos, "UPDATE [" & sText & "] SET "
        pvPutStr baOut, lPos, "INSERT INTO [" & sText & "] ("
        lCount = pvContentKeyCols(sText, m_aOrigDb(aTabCol(lIdx)), aKeyCols)
        pvPutLong baOut, lPos, lCount
        For lCol = 0 To lCount - 1
            pvPutLong baOut, lPos, aKeyCols(lCol)
        Next
    Next
End Sub

Private Function pvContentWriteChanges() As Byte()
    Const DEL_MAGIC_LO  As Long = &HE9B32082
    Const DEF_MAGIC_LO  As Long = &HE9B318C6
    Const DEL_MAGIC_HI  As Long = &H2316A9
    Dim baOut()         As Byte
    Dim lPos            As Long
    Dim lCol            As Long
    Dim lRow            As Long
    Dim eType           As eDeclType
    Dim lWidth          As Long
    Dim aTabCol()       As Long
    Dim lTables         As Long
    Dim lCells          As Long
    Dim aSlot()         As Long
    Dim aVt()           As Long
    Dim aLo()           As Long
    Dim aHi()           As Long
    Dim baStream()      As Byte
    Dim lStream         As Long
    Dim lRowsC          As Long
    Dim aRowIdx()       As Long
    Dim aRowPk()        As Variant
    Dim vEntry          As Variant
    Dim aPk()           As Variant
    Dim lIdx            As Long
    Dim lKeyCol         As Long
    Dim bTextPk         As Boolean
    Dim lMax            As Long
    Dim baPkStream()    As Byte
    Dim lPkStream       As Long
    Dim lVt             As Long
    Dim lLo             As Long
    Dim lHi             As Long
    Dim bPointer        As Boolean
    Dim aOpLabel()      As Long
    Dim aOpKind()       As eChangeOp
    Dim aOpRow()        As Long
    Dim aOpPk()         As Variant
    Dim lOps            As Long

    If Not m_bUpdatable Or m_lPkCount <> 1 Then
        Err.Raise vbObjectError, "cRecordset", "Recordset is not updatable"
    End If
    lKeyCol = m_aPkCols(0)
    bTextPk = pvIsTextishCol(lKeyCol)
    lMax = m_lRecordCount * m_lFieldCount + 16
    If Not m_cDeleted Is Nothing Then
        lMax = lMax + m_cDeleted.Count
    End If
    ReDim aSlot(0 To lMax) As Long
    ReDim aVt(0 To lMax) As Long
    ReDim aLo(0 To lMax) As Long
    ReDim aHi(0 To lMax) As Long
    ReDim aRowIdx(0 To lMax) As Long
    ReDim aRowPk(0 To lMax) As Variant
    ReDim baStream(0 To 255) As Byte
    ReDim baPkStream(0 To 255) As Byte
    '--- collect delete/modify/insert ops sorted by physical row label (RC6
    '--- emission order); deletes carry their snapshotted label
    ReDim aOpLabel(0 To lMax) As Long
    ReDim aOpKind(0 To lMax) As eChangeOp
    ReDim aOpRow(0 To lMax) As Long
    ReDim aOpPk(0 To lMax) As Variant
    If Not m_cDeleted Is Nothing Then
        For Each vEntry In m_cDeleted
            aPk = vEntry(1)
            pvChgAddOp aOpLabel, aOpKind, aOpRow, aOpPk, lOps, CLng(vEntry(5)), CHG_DELETE, 0, aPk(0)
        Next
    End If
    If m_bWriteState Then
        For lRow = 0 To m_lRecordCount - 1
            Select Case m_aRowState(lRow)
            Case ROW_MODIFIED
                pvChgAddOp aOpLabel, aOpKind, aOpRow, aOpPk, lOps, m_aRowLabel(lRow), CHG_UPDATE, lRow, pvPkWhereValue(lRow, 0)
            Case ROW_NEW
                pvChgAddOp aOpLabel, aOpKind, aOpRow, aOpPk, lOps, m_aRowLabel(lRow), CHG_INSERT, lRow, Empty
            End Select
        Next
    End If
    For lIdx = 0 To lOps - 1
        lRow = aOpRow(lIdx)
        Select Case aOpKind(lIdx)
        Case CHG_DELETE
            aSlot(lCells) = aOpLabel(lIdx) * 512
            aVt(lCells) = vbCurrency
            aLo(lCells) = DEL_MAGIC_LO
            aHi(lCells) = DEL_MAGIC_HI
            lCells = lCells + 1
            aRowIdx(lRowsC) = aOpLabel(lIdx)
            aRowPk(lRowsC) = aOpPk(lIdx)
            lRowsC = lRowsC + 1
        Case CHG_UPDATE
            For lCol = 0 To m_lFieldCount - 1
                If m_aDirty(lCol, lRow) Then
                    pvChgPutCell m_aData(lCol, lRow), aOpLabel(lIdx), lCol, lKeyCol, aSlot, aVt, aLo, aHi, lCells, baStream, lStream
                End If
            Next
            aRowIdx(lRowsC) = aOpLabel(lIdx)
            aRowPk(lRowsC) = aOpPk(lIdx)
            lRowsC = lRowsC + 1
        Case CHG_INSERT
            '--- inserts carry only the assigned cells; an unassigned pk gets
            '--- the "use default" magic so the loader lets SQLite assign it
            For lCol = 0 To m_lFieldCount - 1
                If m_aDirty(lCol, lRow) Then
                    pvChgPutCell m_aData(lCol, lRow), aOpLabel(lIdx), lCol, lKeyCol, aSlot, aVt, aLo, aHi, lCells, baStream, lStream
                ElseIf lCol = lKeyCol Then
                    aSlot(lCells) = aOpLabel(lIdx) * 512 + lCol
                    aVt(lCells) = vbCurrency
                    aLo(lCells) = DEF_MAGIC_LO
                    aHi(lCells) = DEL_MAGIC_HI
                    lCells = lCells + 1
                End If
            Next
        End Select
    Next
    ReDim baOut(0 To 1023) As Byte
    pvPutLong baOut, lPos, 0
    pvPutLong baOut, lPos, m_lLoadedRows
    If m_oConnection Is Nothing Then
        pvPutStr baOut, lPos, vbNullString
    Else
        pvPutStr baOut, lPos, m_oConnection.FileName
    End If
    pvPutStr baOut, lPos, m_sSql
    lTables = pvCollectTables(aTabCol)
    pvPutLong baOut, lPos, m_lFieldCount
    For lCol = 0 To m_lFieldCount - 1
        pvWriteFieldMeta baOut, lPos, lCol, aTabCol, lTables, eType, lWidth, True
    Next
    pvWriteDmlSets baOut, lPos, aTabCol, lTables
    '--- cell slot ids + records + shared text stream
    pvPutLong baOut, lPos, lCells
    For lIdx = 0 To lCells - 1
        pvPutLong baOut, lPos, aSlot(lIdx)
    Next
    pvPutZeros baOut, lPos, 16
    For lIdx = 0 To lCells - 1
        pvPutLong baOut, lPos, aVt(lIdx)
        pvPutLong baOut, lPos, IIf(aVt(lIdx) = vbCurrency, 2, 0)
        pvPutLong baOut, lPos, aLo(lIdx)
        pvPutLong baOut, lPos, aHi(lIdx)
    Next
    For lIdx = 0 To lStream - 1
        pvPutByte baOut, lPos, baStream(lIdx)
    Next
    '--- rows needing a WHERE (updates + deletes): row labels + pk records
    pvPutByte baOut, lPos, 0
    pvPutLong baOut, lPos, lRowsC
    '--- flag byte: set when a text-pk WHERE row carries the label-0 handle
    lVt = 0
    If bTextPk Then
        For lIdx = 0 To lRowsC - 1
            If aRowIdx(lIdx) = 0 Then
                lVt = 1
            End If
        Next
    End If
    pvPutByte baOut, lPos, lVt
    For lIdx = 0 To lRowsC - 1
        If Not bTextPk Then
            pvPutLong baOut, lPos, 2 * aRowIdx(lIdx)
        ElseIf aRowIdx(lIdx) = 0 Then
            pvPutLong baOut, lPos, &H800000
        Else
            pvPutLong baOut, lPos, &H1000000 - 2 * aRowIdx(lIdx)
        End If
    Next
    pvPutZeros baOut, lPos, 15
    For lIdx = 0 To lRowsC - 1
        pvChgEncodeValue aRowPk(lIdx), lVt, lLo, lHi, baPkStream, lPkStream, bPointer
        pvPutLong baOut, lPos, lVt
        pvPutLong baOut, lPos, 1
        pvPutLong baOut, lPos, lLo
        pvPutLong baOut, lPos, lHi
    Next
    For lIdx = 0 To lPkStream - 1
        pvPutByte baOut, lPos, baPkStream(lIdx)
    Next
    pvPutByte baOut, lPos, 0
    If (lPos And 1) <> 0 Then
        pvPutZeros baOut, lPos, 1
    End If
    ReDim Preserve baOut(0 To lPos - 1) As Byte
    lIdx = -lPos
    Call CopyMemory(baOut(0), lIdx, 4)
    pvContentWriteChanges = baOut
End Function

Private Sub pvChgAddOp(aOpLabel() As Long, aOpKind() As eChangeOp, aOpRow() As Long, aOpPk() As Variant, lOps As Long, ByVal lLabel As Long, ByVal eKind As eChangeOp, ByVal lRow As Long, vPk As Variant)
    Dim lIdx            As Long

    lIdx = lOps
    Do While lIdx > 0
        If aOpLabel(lIdx - 1) <= lLabel Then
            Exit Do
        End If
        aOpLabel(lIdx) = aOpLabel(lIdx - 1)
        aOpKind(lIdx) = aOpKind(lIdx - 1)
        aOpRow(lIdx) = aOpRow(lIdx - 1)
        aOpPk(lIdx) = aOpPk(lIdx - 1)
        lIdx = lIdx - 1
    Loop
    aOpLabel(lIdx) = lLabel
    aOpKind(lIdx) = eKind
    aOpRow(lIdx) = lRow
    aOpPk(lIdx) = vPk
    lOps = lOps + 1
End Sub

Private Sub pvChgPutCell(vValue As Variant, ByVal lRow As Long, ByVal lCol As Long, ByVal lKeyCol As Long, aSlot() As Long, aVt() As Long, aLo() As Long, aHi() As Long, lCells As Long, baStream() As Byte, lStream As Long)
    Dim lVt             As Long
    Dim lLo             As Long
    Dim lHi             As Long
    Dim bPointer        As Boolean

    pvChgEncodeValue vValue, lVt, lLo, lHi, baStream, lStream, bPointer
    If bPointer Then
        aSlot(lCells) = Not (lRow * 512& + pvTextColIndex(lCol, lKeyCol))
    Else
        aSlot(lCells) = lRow * 512& + lCol
    End If
    aVt(lCells) = lVt
    aLo(lCells) = lLo
    aHi(lCells) = lHi
    lCells = lCells + 1
End Sub

Private Sub pvChgEncodeValue(vValue As Variant, lVt As Long, lLo As Long, lHi As Long, baStream() As Byte, lStream As Long, bPointer As Boolean)
    Dim baCell()        As Byte
    Dim sText           As String
    Dim dblValue        As Double

    lVt = 0
    lLo = 0
    lHi = 0
    bPointer = False
    If pvIsNullCell(vValue) Then
        Exit Sub
    End If
    Select Case VarType(vValue)
    Case vbByte Or vbArray
        baCell = vValue
        lVt = vbLong
        lLo = pvBytesLen(baCell)
        pvChgAppendBytes baStream, lStream, baCell
        bPointer = True
    Case vbString
        sText = vValue
        lVt = vbLong
        lLo = -LenB(sText)
        baCell = sText
        pvChgAppendBytes baStream, lStream, baCell
        bPointer = True
    Case vbInteger
        lVt = vbInteger
        lLo = CLng(vValue) And &HFFFF&
    Case vbLong, vbByte, vbBoolean
        lVt = vbLong
        lLo = CLng(vValue)
    Case vbCurrency, vbDecimal, VT_I8
        lVt = vbDecimal
        pvChgLow8 vValue, lLo, lHi
    Case Else
        dblValue = CDbl(vValue)
        lVt = vbDouble
        Call CopyMemory(lLo, dblValue, 4)
        Call CopyMemory(lHi, ByVal VarPtr(dblValue) + 4, 4)
    End Select
End Sub

Private Sub pvChgLow8(vValue As Variant, lLo As Long, lHi As Long)
    Dim vDec            As Variant

    If VarType(vValue) = VT_I8 Then
        Call CopyMemory(lLo, ByVal VarPtr(vValue) + 8, 4)
        Call CopyMemory(lHi, ByVal VarPtr(vValue) + 12, 4)
    Else
        vDec = CDec(vValue)
        Call CopyMemory(lLo, ByVal VarPtr(vDec) + 8, 4)
        Call CopyMemory(lHi, ByVal VarPtr(vDec) + 12, 4)
    End If
End Sub

Private Sub pvChgAppendBytes(baStream() As Byte, lStream As Long, baCell() As Byte)
    Dim lSize           As Long

    lSize = pvBytesLen(baCell)
    If lSize = 0 Then
        Exit Sub
    End If
    Do While lStream + lSize > UBound(baStream) + 1
        ReDim Preserve baStream(0 To 2 * (UBound(baStream) + 1) - 1) As Byte
    Loop
    Call CopyMemory(baStream(lStream), baCell(0), lSize)
    lStream = lStream + lSize
End Sub

Private Function pvTextColIndex(ByVal lCol As Long, ByVal lKeyCol As Long) As Long
    Dim lIdx            As Long

    For lIdx = 0 To lCol - 1
        If lIdx <> lKeyCol Then
            If pvIsTextishCol(lIdx) Then
                pvTextColIndex = pvTextColIndex + 1
            End If
        End If
    Next
End Function

Private Function pvIsTextishCol(ByVal lCol As Long) As Boolean
    Select Case pvDeclTypeCode(m_aDeclType(lCol))
    Case TYPE_TEXT, TYPE_BLOB, TYPE_DATE
        pvIsTextishCol = True
    End Select
End Function

Private Sub pvContentReadChanges(baContent() As Byte)
    Const DEL_MAGIC_LO  As Long = &HE9B32082
    Const DEF_MAGIC_LO  As Long = &HE9B318C6
    Const DEL_MAGIC_HI  As Long = &H2316A9
    Dim lPos            As Long
    Dim lCol            As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lCount          As Long
    Dim sDefault        As String
    Dim sCollate        As String
    Dim lNotNull        As Long
    Dim lPk             As Long
    Dim lAutoInc        As Long
    Dim lTables         As Long
    Dim lCells          As Long
    Dim aSlot()         As Long
    Dim aVt()           As Long
    Dim aLo()           As Long
    Dim aHi()           As Long
    Dim lRowsC          As Long
    Dim lVal            As Long
    Dim bTextPk         As Boolean
    Dim nVal            As Integer
    Dim sngVal          As Single
    Dim dblVal          As Double
    Dim cyVal           As Currency

    pvClear
    lPos = 4
    pvGetLong baContent, lPos                       '--- source RecordCount
    pvGetStr baContent, lPos                        '--- source DB filename
    m_sSql = pvGetStr(baContent, lPos)
    m_lFieldCount = pvGetLong(baContent, lPos)
    If m_lFieldCount > 0 Then
        ReDim m_aFieldName(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigColumn(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigTable(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigDb(0 To m_lFieldCount - 1) As String
        ReDim m_aDeclType(0 To m_lFieldCount - 1) As String
        ReDim m_aColSchema(0 To m_lFieldCount - 1) As cColumn
        ReDim m_aData(0 To m_lFieldCount - 1, 0 To 15) As Variant
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
        lNotNull = pvGetByte(baContent, lPos)
        lPk = pvGetByte(baContent, lPos)
        lAutoInc = pvGetByte(baContent, lPos)
        pvGetLong baContent, lPos                   '--- type code
        pvGetByte baContent, lPos
        pvGetLong baContent, lPos                   '--- int width class
        If Len(m_aOrigTable(lCol)) > 0 Then
            Set m_aColSchema(lCol) = New cColumn
            If StrComp(sDefault, "NULL", vbBinaryCompare) = 0 Then
                sDefault = vbNullString
            End If
            m_aColSchema(lCol).frInit m_aOrigColumn(lCol), m_aDeclType(lCol), sDefault, lNotNull <> 0, lPk, lAutoInc <> 0, sCollate
        End If
    Next
    '--- DML sets: capture the first table's single key column ordinal
    m_lChgKeyCol = -1
    lTables = pvGetLong(baContent, lPos)
    For lIdx = 1 To lTables
        pvGetStr baContent, lPos
        pvGetStr baContent, lPos
        pvGetStr baContent, lPos
        lCount = pvGetLong(baContent, lPos)
        For lJdx = 1 To lCount
            lVal = pvGetLong(baContent, lPos)
            If lIdx = 1 And lJdx = 1 And lCount = 1 Then
                m_lChgKeyCol = lVal
            End If
        Next
    Next
    '--- cell slots + records, then the shared text stream in cell order
    lCells = pvGetLong(baContent, lPos)
    m_lChgCellCount = lCells
    If lCells > 0 Then
        ReDim aSlot(0 To lCells - 1) As Long
        ReDim aVt(0 To lCells - 1) As Long
        ReDim aLo(0 To lCells - 1) As Long
        ReDim aHi(0 To lCells - 1) As Long
        ReDim m_aChgCellRow(0 To lCells - 1) As Long
        ReDim m_aChgCellCol(0 To lCells - 1) As Long
        ReDim m_aChgCellValue(0 To lCells - 1) As Variant
    End If
    For lIdx = 0 To lCells - 1
        aSlot(lIdx) = pvGetLong(baContent, lPos)
    Next
    lPos = lPos + 16
    For lIdx = 0 To lCells - 1
        aVt(lIdx) = pvGetLong(baContent, lPos)
        pvGetLong baContent, lPos
        aLo(lIdx) = pvGetLong(baContent, lPos)
        aHi(lIdx) = pvGetLong(baContent, lPos)
    Next
    For lIdx = 0 To lCells - 1
        If aSlot(lIdx) < 0 Then
            lVal = Not aSlot(lIdx)
            m_aChgCellRow(lIdx) = lVal \ 512
            m_aChgCellCol(lIdx) = pvColFromTextIndex(lVal Mod 512)
            '--- positive size = raw blob bytes, negative = UTF-16 text
            m_aChgCellValue(lIdx) = pvChgSlice(baContent, lPos, Abs(aLo(lIdx)), aLo(lIdx) > 0)
        Else
            m_aChgCellRow(lIdx) = aSlot(lIdx) \ 512
            m_aChgCellCol(lIdx) = aSlot(lIdx) Mod 512
            Select Case aVt(lIdx)
            Case vbEmpty, vbNull
                m_aChgCellValue(lIdx) = pvNullCell()
            Case vbInteger
                Call CopyMemory(nVal, aLo(lIdx), 2)
                m_aChgCellValue(lIdx) = CLng(nVal)
            Case vbLong
                m_aChgCellValue(lIdx) = aLo(lIdx)
            Case vbSingle
                Call CopyMemory(sngVal, aLo(lIdx), 4)
                m_aChgCellValue(lIdx) = sngVal
            Case vbDouble, vbDate
                Call CopyMemory(dblVal, aLo(lIdx), 4)
                Call CopyMemory(ByVal VarPtr(dblVal) + 4, aHi(lIdx), 4)
                m_aChgCellValue(lIdx) = dblVal
            Case vbCurrency
                If aLo(lIdx) = DEL_MAGIC_LO And aHi(lIdx) = DEL_MAGIC_HI Then
                    m_aChgCellCol(lIdx) = -1
                ElseIf aLo(lIdx) = DEF_MAGIC_LO And aHi(lIdx) = DEL_MAGIC_HI Then
                    '--- "use default" pk cell: excluded from the INSERT
                    m_aChgCellCol(lIdx) = -2
                Else
                    Call CopyMemory(cyVal, aLo(lIdx), 4)
                    Call CopyMemory(ByVal VarPtr(cyVal) + 4, aHi(lIdx), 4)
                    m_aChgCellValue(lIdx) = cyVal
                End If
            Case Else
                m_aChgCellValue(lIdx) = pvChgInt64(aLo(lIdx), aHi(lIdx))
            End Select
        End If
    Next
    '--- rows with WHERE records
    pvGetByte baContent, lPos
    lRowsC = pvGetLong(baContent, lPos)
    pvGetByte baContent, lPos
    m_lChgRowCount = lRowsC
    If lRowsC > 0 Then
        ReDim m_aChgRowIdx(0 To lRowsC - 1) As Long
        ReDim m_aChgRowPk(0 To lRowsC - 1) As Variant
        ReDim m_aChgRowDelete(0 To lRowsC - 1) As Boolean
        ReDim aVt(0 To lRowsC - 1) As Long
        ReDim aLo(0 To lRowsC - 1) As Long
        ReDim aHi(0 To lRowsC - 1) As Long
    End If
    If m_lChgKeyCol >= 0 Then
        bTextPk = pvIsTextishCol(m_lChgKeyCol)
    End If
    For lIdx = 0 To lRowsC - 1
        lVal = pvGetLong(baContent, lPos)
        If Not bTextPk Then
            m_aChgRowIdx(lIdx) = lVal \ 2
        ElseIf lVal = &H800000 Then
            m_aChgRowIdx(lIdx) = 0
        Else
            m_aChgRowIdx(lIdx) = (&H1000000 - lVal) \ 2
        End If
    Next
    lPos = lPos + 15
    For lIdx = 0 To lRowsC - 1
        aVt(lIdx) = pvGetLong(baContent, lPos)
        pvGetLong baContent, lPos
        aLo(lIdx) = pvGetLong(baContent, lPos)
        aHi(lIdx) = pvGetLong(baContent, lPos)
    Next
    For lIdx = 0 To lRowsC - 1
        If bTextPk Then
            m_aChgRowPk(lIdx) = pvChgSlice(baContent, lPos, Abs(aLo(lIdx)), aLo(lIdx) > 0)
        ElseIf aVt(lIdx) = vbDecimal Or aVt(lIdx) = VT_I8 Then
            m_aChgRowPk(lIdx) = pvChgInt64(aLo(lIdx), aHi(lIdx))
        Else
            m_aChgRowPk(lIdx) = aLo(lIdx)
        End If
        For lJdx = 0 To lCells - 1
            If m_aChgCellCol(lJdx) = -1 And m_aChgCellRow(lJdx) = m_aChgRowIdx(lIdx) Then
                m_aChgRowDelete(lIdx) = True
            End If
        Next
    Next
    m_lRecordCount = 0
    m_lCurRow = -1
    m_bUpdatable = False
    m_bHasColSchema = True
    pvComputeCoercion
    pvEnsureFields
End Sub

Private Function pvColFromTextIndex(ByVal lTextIdx As Long) As Long
    Dim lCol            As Long
    Dim lSeen           As Long

    For lCol = 0 To m_lFieldCount - 1
        If lCol <> m_lChgKeyCol Then
            If pvIsTextishCol(lCol) Then
                If lSeen = lTextIdx Then
                    pvColFromTextIndex = lCol
                    Exit Function
                End If
                lSeen = lSeen + 1
            End If
        End If
    Next
    pvColFromTextIndex = 0
End Function

Private Function pvChgSlice(baContent() As Byte, lPos As Long, ByVal lSize As Long, ByVal bBlob As Boolean) As Variant
    Dim baCell()        As Byte
    Dim sText           As String

    If bBlob Then
        baCell = vbNullString
        If lSize > 0 Then
            ReDim baCell(0 To lSize - 1) As Byte
            Call CopyMemory(baCell(0), baContent(lPos), lSize)
        End If
        pvChgSlice = baCell
    Else
        sText = String$(lSize \ 2, 0)
        If lSize > 0 Then
            Call CopyMemory(ByVal StrPtr(sText), baContent(lPos), lSize)
        End If
        pvChgSlice = sText
    End If
    lPos = lPos + lSize
End Function

Private Function pvChgInt64(ByVal lLo As Long, ByVal lHi As Long) As Variant
    Dim vRet            As Variant
    Dim nVt             As Integer

    nVt = VT_I8
    Call CopyMemory(vRet, nVt, 2)
    Call CopyMemory(ByVal VarPtr(vRet) + 8, lLo, 4)
    Call CopyMemory(ByVal VarPtr(vRet) + 12, lHi, 4)
    pvChgInt64 = vRet
End Function

Private Sub pvApplyChangesOnly()
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim sSql            As String
    Dim sCols           As String
    Dim aParams()       As Variant
    Dim lCount          As Long
    Dim sTable          As String
    Dim sKey            As String
    Dim aDone()         As Boolean

    If m_oConnection Is Nothing Then
        Err.Raise vbObjectError, "cRecordset", "No active connection"
    End If
    If m_lChgKeyCol < 0 Or m_lChgKeyCol >= m_lFieldCount Then
        Err.Raise vbObjectError, "cRecordset", "Recordset is not updatable"
    End If
    sTable = QuoteIdentifier(IIf(Len(m_aOrigDb(m_lChgKeyCol)) > 0, m_aOrigDb(m_lChgKeyCol), "main")) & "." & QuoteIdentifier(m_aOrigTable(m_lChgKeyCol))
    sKey = QuoteIdentifier(m_aOrigColumn(m_lChgKeyCol))
    m_oConnection.BeginTrans "tc6_applychanges"
    On Error GoTo EH
    For lIdx = 0 To m_lChgRowCount - 1
        If m_aChgRowDelete(lIdx) Then
            ReDim aParams(0 To 0) As Variant
            aParams(0) = m_aChgRowPk(lIdx)
            pvExecParams "DELETE FROM " & sTable & " WHERE " & sKey & " = ?", aParams
        Else
            sSql = vbNullString
            ReDim aParams(0 To m_lChgCellCount) As Variant
            lCount = 0
            For lJdx = 0 To m_lChgCellCount - 1
                If m_aChgCellRow(lJdx) = m_aChgRowIdx(lIdx) And m_aChgCellCol(lJdx) >= 0 Then
                    If lCount > 0 Then
                        sSql = sSql & ", "
                    End If
                    sSql = sSql & QuoteIdentifier(m_aOrigColumn(m_aChgCellCol(lJdx))) & " = ?"
                    aParams(lCount) = m_aChgCellValue(lJdx)
                    lCount = lCount + 1
                End If
            Next
            If lCount > 0 Then
                aParams(lCount) = m_aChgRowPk(lIdx)
                ReDim Preserve aParams(0 To lCount) As Variant
                pvExecParams "UPDATE " & sTable & " SET " & sSql & " WHERE " & sKey & " = ?", aParams
            End If
        End If
    Next
    '--- cell rows without a WHERE record are pending inserts
    ReDim aDone(0 To m_lChgCellCount) As Boolean
    For lIdx = 0 To m_lChgCellCount - 1
        If m_aChgCellCol(lIdx) >= 0 And Not aDone(lIdx) And Not pvChgRowHasWhere(m_aChgCellRow(lIdx)) Then
            sSql = vbNullString
            sCols = vbNullString
            ReDim aParams(0 To m_lChgCellCount) As Variant
            lCount = 0
            For lJdx = lIdx To m_lChgCellCount - 1
                If m_aChgCellRow(lJdx) = m_aChgCellRow(lIdx) And m_aChgCellCol(lJdx) >= 0 Then
                    If lCount > 0 Then
                        sCols = sCols & ", "
                        sSql = sSql & ", "
                    End If
                    sCols = sCols & QuoteIdentifier(m_aOrigColumn(m_aChgCellCol(lJdx)))
                    sSql = sSql & "?"
                    aParams(lCount) = m_aChgCellValue(lJdx)
                    lCount = lCount + 1
                    aDone(lJdx) = True
                End If
            Next
            ReDim Preserve aParams(0 To lCount - 1) As Variant
            pvExecParams "INSERT INTO " & sTable & " (" & sCols & ") VALUES (" & sSql & ")", aParams
        End If
    Next
    m_oConnection.CommitTrans "tc6_applychanges"
    m_lChgCellCount = 0
    m_lChgRowCount = 0
    Exit Sub
EH:
    m_oConnection.RollbackTrans "tc6_applychanges"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function pvChgRowHasWhere(ByVal lRow As Long) As Boolean
    Dim lIdx            As Long

    For lIdx = 0 To m_lChgRowCount - 1
        If m_aChgRowIdx(lIdx) = lRow Then
            pvChgRowHasWhere = True
            Exit Function
        End If
    Next
End Function

Private Sub pvContentRead(baContent() As Byte)
    Dim lPos            As Long
    Dim lRows           As Long
    Dim lCol            As Long
    Dim lRow            As Long
    Dim lCurRow         As Long
    Dim eType           As eDeclType
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

    lPos = 0
    If pvGetLong(baContent, lPos) < 0 Then
        pvContentReadChanges baContent
        Exit Sub
    End If
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
        ReDim m_aFieldName(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigColumn(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigTable(0 To m_lFieldCount - 1) As String
        ReDim m_aOrigDb(0 To m_lFieldCount - 1) As String
        ReDim m_aDeclType(0 To m_lFieldCount - 1) As String
        ReDim m_aColSchema(0 To m_lFieldCount - 1) As cColumn
        ReDim m_aData(0 To m_lFieldCount - 1, 0 To IIf(lRows > 16, lRows, 16) - 1) As Variant
        ReDim m_aRowLabel(0 To IIf(lRows > 16, lRows, 16) - 1) As Long
        For lRow = 0 To lRows - 1
            m_aRowLabel(lRow) = lRow
        Next
        m_lNextLabel = lRows
        m_lLoadedRows = lRows
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
        eType = pvGetLong(baContent, lPos)
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
        ReDim aFlags(0 To IIf(lCount > 0, lCount, 1) - 1) As Byte
        For lRow = 0 To lCount - 1
            aFlags(lRow) = pvGetByte(baContent, lPos)
        Next
        '--- cell data
        Select Case eType
        Case TYPE_INTEGER, TYPE_BOOLEAN
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
        Case TYPE_FLOAT
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
            ReDim baText(0 To IIf(lCount > 0, lCount, 1) - 1) As Byte
            For lRow = 0 To lCount - 1
                baText(lRow) = pvGetByte(baContent, lPos)
            Next
            lCount = pvGetLong(baContent, lPos) \ 4
            ReDim aOffs(0 To IIf(lCount > 0, lCount, 1) - 1) As Long
            For lRow = 0 To lCount - 1
                aOffs(lRow) = pvGetLong(baContent, lPos)
            Next
            For lRow = 0 To lRows - 1
                If aFlags(lRow) = 0 Then
                    m_aData(lCol, lRow) = pvNullCell()
                Else
                    m_aData(lCol, lRow) = pvSliceCell(baText, aOffs(lRow), aOffs(lRow + 1) - aOffs(lRow), eType)
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
    pvComputeCoercion
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
    ReDim m_aPkCols(0 To m_lFieldCount - 1) As Long
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
    ReDim Preserve m_aPkCols(0 To m_lPkCount - 1) As Long
    If m_lPkCount = 1 Then
        If pvDeclTypeCode(m_aDeclType(m_aPkCols(0))) = TYPE_INTEGER Then
            m_lAutoPkCol = m_aPkCols(0)
        End If
    End If
    m_bUpdatable = True
End Sub

Private Function pvDeclTypeCode(sDeclType As String) As eDeclType
    Dim sUpper          As String

    '--- SQLite affinity rules collapsed to the RC6 type codes (date cells
    '--- store as raw-value text, boolean cells as width-class integers);
    '--- empty decltype (expression columns) serializes as FLOAT
    sUpper = UCase$(sDeclType)
    If InStr(sUpper, "DATE") > 0 Or InStr(sUpper, "TIME") > 0 Then
        pvDeclTypeCode = TYPE_DATE
    ElseIf InStr(sUpper, "BIT") > 0 Or InStr(sUpper, "BOOL") > 0 Then
        pvDeclTypeCode = TYPE_BOOLEAN
    ElseIf InStr(sUpper, "INT") > 0 Then
        pvDeclTypeCode = TYPE_INTEGER
    ElseIf InStr(sUpper, "CHAR") > 0 Or InStr(sUpper, "CLOB") > 0 Or InStr(sUpper, "TEXT") > 0 Then
        pvDeclTypeCode = TYPE_TEXT
    ElseIf InStr(sUpper, "BLOB") > 0 Then
        pvDeclTypeCode = TYPE_BLOB
    Else
        pvDeclTypeCode = TYPE_FLOAT
    End If
End Function

Private Function pvContentKeyCols(sTable As String, sDb As String, aKey() As Long) As Long
    Dim lCol            As Long
    Dim lCount          As Long
    Dim cPkNames        As VBA.Collection
    Dim hStmt           As LongPtr
    Dim bIsPk           As Boolean
    Dim oColumn         As cColumn

    ReDim aKey(0 To m_lFieldCount - 1) As Long
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

Private Function pvValueTypeCode(ByVal lCol As Long) As eDeclType
    Dim lRow            As Long

    pvValueTypeCode = TYPE_FLOAT
    For lRow = 0 To m_lRecordCount - 1
        If Not pvIsNullCell(m_aData(lCol, lRow)) Then
            Select Case VarType(m_aData(lCol, lRow))
            Case vbString
                pvValueTypeCode = TYPE_TEXT
            Case vbByte + vbArray
                pvValueTypeCode = TYPE_BLOB
            End Select
            Exit Function
        End If
    Next
End Function

Private Function pvIntWidth(ByVal lCol As Long, Optional ByVal bLoadedOnly As Boolean) As Long
    Dim lRow            As Long
    Dim decValue        As Variant
    Dim vValue          As Variant

    '--- smallest RC6 width class holding every value: 1 (unsigned byte),
    '--- 4 (int32) or 8 (int64); the PK column is always at least 4; the
    '--- changes blob keeps the load-time class (RC6 never widens it)
    If lCol = m_lAutoPkCol Then
        pvIntWidth = 4
    Else
        pvIntWidth = 1
    End If
    For lRow = 0 To m_lRecordCount - 1
        vValue = m_aData(lCol, lRow)
        If bLoadedOnly And m_bWriteState Then
            If m_aRowState(lRow) = ROW_NEW Then
                vValue = Empty
            ElseIf m_aDirty(lCol, lRow) Then
                vValue = m_aOriginal(lCol, lRow)
            End If
        End If
        If Not pvIsNullCell(vValue) And IsNumeric(vValue) Then
            decValue = CDec(vValue)
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

Private Function pvJsonValue(ByVal lCol As Long, vValue As Variant, ByVal bUniEscaping As Boolean) As String
    Dim baCell()        As Byte

    If pvIsNullCell(vValue) Then
        pvJsonValue = "null"
    ElseIf m_aColCoerce(lCol) = COERCE_BOOLEAN Then
        '--- RC6: BIT/BOOL columns serialize as JSON booleans
        pvJsonValue = IIf(pvCoerceCell(lCol, vValue), "true", "false")
    ElseIf VarType(vValue) = vbString Then
        pvJsonValue = pvJsonString(CStr(vValue), bUniEscaping)
    ElseIf m_aColCoerce(lCol) = COERCE_DATE Then
        '--- RC6: numeric-stored dates serialize as quoted raw values
        pvJsonValue = """" & pvJsonNumber(vValue) & """"
    ElseIf IsArray(vValue) Then
        baCell = vValue
        pvJsonValue = """" & pvBase64(baCell) & """"
    Else
        pvJsonValue = pvJsonNumber(vValue)
    End If
End Function

Private Function pvJsonNumber(vValue As Variant) As String
    '--- RC6 formats via Decimal (0.000015 not 1.5E-05); out-of-range
    '--- doubles fall back to scientific notation (1E+300)
    On Error GoTo EH
    pvJsonNumber = CStr(CDec(vValue))
    Exit Function
EH:
    pvJsonNumber = CStr(vValue)
End Function

Private Function pvJsonString(sText As String, ByVal bUniEscaping As Boolean) As String
    Dim lIdx            As Long
    Dim lChar           As Long
    Dim sChunk          As String

    pvJsonString = """"
    For lIdx = 1 To Len(sText)
        lChar = AscW(Mid$(sText, lIdx, 1)) And &HFFFF&
        Select Case lChar
        Case 34
            sChunk = "\"""
        Case 92
            sChunk = "\\"
        Case 8
            sChunk = "\b"
        Case 9
            sChunk = "\t"
        Case 10
            sChunk = "\n"
        Case 12
            sChunk = "\f"
        Case 13
            sChunk = "\r"
        Case Is < 32
            sChunk = "\u" & LCase$(Right$("000" & Hex$(lChar), 4))
        Case Is > 127
            If bUniEscaping Then
                sChunk = "\u" & LCase$(Right$("000" & Hex$(lChar), 4))
            Else
                sChunk = Mid$(sText, lIdx, 1)
            End If
        Case Else
            sChunk = Mid$(sText, lIdx, 1)
        End Select
        pvJsonString = pvJsonString & sChunk
    Next
    pvJsonString = pvJsonString & """"
End Function

Private Function pvBase64(baBytes() As Byte) As String
    Const CRYPT_STRING_BASE64 As Long = 1
    Const CRYPT_STRING_NOCRLF As Long = &H40000000
    Dim lSize           As Long

    If pvBytesLen(baBytes) = 0 Then
        Exit Function
    End If
    If CryptBinaryToString(baBytes(0), pvBytesLen(baBytes), CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, 0, lSize) = 0 Then
        Exit Function
    End If
    pvBase64 = String$(lSize, 0)
    Call CryptBinaryToString(baBytes(0), pvBytesLen(baBytes), CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF, StrPtr(pvBase64), lSize)
    pvBase64 = Left$(pvBase64, lSize)
    If Right$(pvBase64, 1) = vbNullChar Then
        pvBase64 = Left$(pvBase64, Len(pvBase64) - 1)
    End If
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

Private Function pvEventPos() As Long
    '--- RC6 event position codes: row index, -1 = BOF/empty, -2 = EOF
    If m_lRecordCount = 0 Then
        pvEventPos = -1
    ElseIf m_lCurRow < 0 Then
        pvEventPos = -1
    ElseIf m_lCurRow >= m_lRecordCount Then
        pvEventPos = -2
    Else
        pvEventPos = m_lCurRow
    End If
End Function

Private Sub pvComputeCoercion()
    Dim lCol            As Long

    If m_lFieldCount = 0 Then
        Exit Sub
    End If
    ReDim m_aColCoerce(0 To m_lFieldCount - 1) As eCoerceKind
    For lCol = 0 To m_lFieldCount - 1
        Select Case pvDeclTypeCode(m_aDeclType(lCol))
        Case TYPE_DATE
            m_aColCoerce(lCol) = COERCE_DATE
        Case TYPE_BOOLEAN
            m_aColCoerce(lCol) = COERCE_BOOLEAN
        Case TYPE_INTEGER
            '--- INTEGER columns type by their load-time width class
            Select Case pvIntWidth(lCol)
            Case 1
                m_aColCoerce(lCol) = COERCE_BYTE
            Case 8
                m_aColCoerce(lCol) = COERCE_INT64
            End Select
        End Select
    Next
End Sub

Private Function pvCoerceCell(ByVal lCol As Long, vValue As Variant) As Variant
    Dim dblValue        As Double

    pvCoerceCell = vValue
    If pvIsNullCell(vValue) Then
        Exit Function
    End If
    On Error GoTo EH
    Select Case m_aColCoerce(lCol)
    Case COERCE_DATE
        '--- numeric date storage is a VB date serial (not julian)
        If VarType(vValue) = vbString Then
            pvCoerceCell = CDate(vValue)
        Else
            pvCoerceCell = CDate(CDbl(vValue))
        End If
    Case COERCE_BOOLEAN
        If VarType(vValue) = vbString Then
            dblValue = Val(vValue)
        Else
            dblValue = CDbl(vValue)
        End If
        pvCoerceCell = (dblValue >= 1)
    Case COERCE_BYTE
        If vValue >= 0 And vValue <= 255 Then
            pvCoerceCell = CByte(vValue)
        End If
    Case COERCE_INT64
        pvCoerceCell = Int64Variant(vValue)
    End Select
    Exit Function
EH:
    pvCoerceCell = Empty
End Function

Private Function pvCellValueEquals(vOld As Variant, vNew As Variant) As Boolean
    Dim baOld()         As Byte
    Dim baNew()         As Byte
    Dim lIdx            As Long

    On Error GoTo QH
    If pvIsNullCell(vOld) And pvIsNullCell(vNew) Then
        pvCellValueEquals = True
    ElseIf pvIsNullCell(vOld) Or pvIsNullCell(vNew) Then
        '--- not equal
    ElseIf IsArray(vOld) And IsArray(vNew) Then
        baOld = vOld
        baNew = vNew
        If pvBytesLen(baOld) <> pvBytesLen(baNew) Then
            Exit Function
        End If
        For lIdx = 0 To pvBytesLen(baOld) - 1
            If baOld(lIdx) <> baNew(lIdx) Then
                Exit Function
            End If
        Next
        pvCellValueEquals = True
    ElseIf Not IsArray(vOld) And Not IsArray(vNew) Then
        pvCellValueEquals = (vOld = vNew)
    End If
QH:
End Function

Private Function pvCellDbl(vCell As Variant) As Double
    If Not pvIsNullCell(vCell) Then
        pvCellDbl = CDbl(vCell)
    End If
End Function

Private Function pvCellBytes(vCell As Variant, ByVal eType As eDeclType) As Byte()
    Dim baEmpty()       As Byte

    If pvIsNullCell(vCell) Then
        pvCellBytes = baEmpty
    ElseIf eType = TYPE_BLOB And VarType(vCell) = vbByte + vbArray Then
        pvCellBytes = vCell
    ElseIf VarType(vCell) = vbString Then
        pvCellBytes = ToUtf8Array(CStr(vCell))
    Else
        '--- numeric cells in text/date columns keep the invariant "." form
        pvCellBytes = ToUtf8Array(pvJsonNumber(vCell))
    End If
End Function

Private Function pvCellBytesLen(vCell As Variant, ByVal eType As eDeclType) As Long
    Dim baCell()        As Byte

    baCell = pvCellBytes(vCell, eType)
    pvCellBytesLen = pvBytesLen(baCell)
End Function

Private Function pvBytesLen(baBuf() As Byte) As Long
    On Error GoTo QH
    pvBytesLen = UBound(baBuf) - LBound(baBuf) + 1
QH:
End Function

Private Function pvSliceCell(baText() As Byte, ByVal lStart As Long, ByVal lLen As Long, ByVal eType As eDeclType) As Variant
    Dim baCell()        As Byte

    If lLen > 0 Then
        ReDim baCell(0 To lLen - 1) As Byte
        Call CopyMemory(baCell(0), baText(lStart), lLen)
    End If
    If eType = TYPE_BLOB Then
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
        ReDim Preserve baBuf(0 To lCap - 1) As Byte
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
