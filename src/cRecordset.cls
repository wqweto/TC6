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
' Bookmark, Fields (cFields/cField), ValueMatrix (read), GetRows, ReQuery.
' The updatable surface (AddNew/Delete/UpdateBatch/ResetChanges), Sort/Find,
' Content serialization, ADO interop and JSON export are not implemented yet
' and raise; unimplemented read-only members return defaults.
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

Public Property Get ContainsChanges() As Boolean
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
End Property

Public Property Get UniqueID64() As Variant
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
    pvNotImpl "ValueMatrix (write)"
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
End Function

Public Sub AddNew()
    pvNotImpl "AddNew"
End Sub

Public Sub Delete()
    pvNotImpl "Delete"
End Sub

Public Sub UpdateBatch()
    pvNotImpl "UpdateBatch"
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
    pvNotImpl "SortRefresh"
End Sub

Public Function FindFirst(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    pvNotImpl "FindFirst"
End Function

Public Function FindNext(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    pvNotImpl "FindNext"
End Function

Public Function FindPrevious(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    pvNotImpl "FindPrevious"
End Function

Public Function FindLast(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
    pvNotImpl "FindLast"
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
    pvNotImpl "ResetChanges"
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
