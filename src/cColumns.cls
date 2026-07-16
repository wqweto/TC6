VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cColumns - the columns of one table, snapshotted at creation from
' PRAGMA table_info enriched with sqlite3_table_column_metadata (collation
' + autoincrement) and single-column UNIQUE indexes (PRAGMA index_list/
' index_info). Item accepts a 0-based index or a column name (keyed).
'=========================================================================
Option Explicit

Private m_cItems                    As VBA.Collection

Public Property Get Count() As Long
    Count = m_cItems.Count
End Property

Public Property Get Item(Idx_Or_ColumnName As Variant) As cColumn
Attribute Item.VB_UserMemId = 0
    Dim lIdx            As Long

    If IsNumeric(Idx_Or_ColumnName) Then
        lIdx = CLng(Idx_Or_ColumnName)
        If lIdx < 0 Or lIdx >= m_cItems.Count Then
            Err.Raise 9
        End If
        Set Item = m_cItems.Item(lIdx + 1)
    Else
        Set Item = pvItemByName(CStr(Idx_Or_ColumnName))
    End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Friend Sub frInit(oCnn As cConnection, sDb As String, sTable As String)
    Dim oRs             As cRecordset
    Dim oIdxRs          As cRecordset
    Dim oColRs          As cRecordset
    Dim oColumn         As cColumn
    Dim lRow            As Long
    Dim lIdx            As Long
    Dim sName           As String
    Dim sDefault        As String
    Dim sCollate        As String
    Dim bAutoInc        As Boolean

    Set m_cItems = New VBA.Collection
    '--- table_info columns: cid, name, type, notnull, dflt_value, pk
    Set oRs = oCnn.GetRs("PRAGMA " & QuoteIdentifier(sDb) & ".table_info(" & QuoteIdentifier(sTable) & ")")
    For lRow = 0 To oRs.RecordCount - 1
        sName = CStr(oRs.ValueMatrix(lRow, 1))
        sDefault = vbNullString
        If Not IsNull(oRs.ValueMatrix(lRow, 4)) Then
            sDefault = CStr(oRs.ValueMatrix(lRow, 4))
        End If
        pvColumnMetadata oCnn, sDb, sTable, sName, sCollate, bAutoInc
        Set oColumn = New cColumn
        oColumn.frInit sName, CStr(oRs.ValueMatrix(lRow, 2)), sDefault, _
            CLng(oRs.ValueMatrix(lRow, 3)) <> 0, CLng(oRs.ValueMatrix(lRow, 5)), bAutoInc, sCollate
        m_cItems.Add oColumn, sName
    Next
    '--- a single-column UNIQUE index marks its column as unique-constrained
    Set oIdxRs = oCnn.GetRs("PRAGMA " & QuoteIdentifier(sDb) & ".index_list(" & QuoteIdentifier(sTable) & ")")
    For lIdx = 0 To oIdxRs.RecordCount - 1
        '--- index_list columns: seq, name, unique, origin, partial
        If CLng(oIdxRs.ValueMatrix(lIdx, 2)) <> 0 Then
            Set oColRs = oCnn.GetRs("PRAGMA " & QuoteIdentifier(sDb) & ".index_info(" & QuoteIdentifier(CStr(oIdxRs.ValueMatrix(lIdx, 1))) & ")")
            If oColRs.RecordCount = 1 Then
                If Not IsNull(oColRs.ValueMatrix(0, 2)) Then
                    pvMarkUnique CStr(oColRs.ValueMatrix(0, 2))
                End If
            End If
        End If
    Next
End Sub

Private Sub pvColumnMetadata(oCnn As cConnection, sDb As String, sTable As String, sColumn As String, sCollate As String, bAutoInc As Boolean)
    Dim baDb()          As Byte
    Dim baTable()       As Byte
    Dim baColumn()      As Byte
    Dim lpCollSeq       As LongPtr
    Dim lNotNull        As Long
    Dim lPk             As Long
    Dim lAutoInc        As Long

    sCollate = vbNullString
    bAutoInc = False
    baDb = ToUtf8Array(sDb & vbNullChar)
    baTable = ToUtf8Array(sTable & vbNullChar)
    baColumn = ToUtf8Array(sColumn & vbNullChar)
    If stub_sqlite3_table_column_metadata(oCnn.frDbHandle, VarPtr(baDb(0)), VarPtr(baTable(0)), VarPtr(baColumn(0)), _
            0, VarPtr(lpCollSeq), VarPtr(lNotNull), VarPtr(lPk), VarPtr(lAutoInc)) = SQLITE_OK Then
        sCollate = FromUtf8Ptr(lpCollSeq)
        bAutoInc = (lAutoInc <> 0)
    End If
End Sub

Private Sub pvMarkUnique(sName As String)
    Dim oColumn         As cColumn

    On Error GoTo QH
    Set oColumn = m_cItems.Item(sName)
    oColumn.frMarkUnique
QH:
End Sub

Private Function pvItemByName(sName As String) As cColumn
    On Error GoTo EH
    Set pvItemByName = m_cItems.Item(sName)
    Exit Function
EH:
    '--- clear the pending collection error so 9 gets its default description
    Err.Clear
    Err.Raise 9
End Function
