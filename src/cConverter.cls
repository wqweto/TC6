VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cConverter - ADO/OLEDB database -> SQLite migration (see NOTES)
'=========================================================================
Option Explicit

Public Event SchemaProgress(ByVal TableName As String, ByVal TablesCount As Long, ByVal TablesCreated As Long)
Public Event InsertProgress(ByVal TableName As String, ByVal RecordCount As Long, ByVal RecordsInserted As Long)
Public Event IndexProgress(ByVal TableName As String, ByVal IndexName As String, ByVal IndexesCount As Long, ByVal IndexesCreated As Long)

Private m_bNoCase                   As Boolean

Public Property Get CaseInSensitive_TextColumns() As Boolean
    CaseInSensitive_TextColumns = m_bNoCase
End Property

Public Property Let CaseInSensitive_TextColumns(ByVal RHS As Boolean)
    m_bNoCase = RHS
End Property

Public Sub ConvertDatabase(Cnn As Object, sCnn As cConnection)
    Const adSchemaColumns                   As Long = 4
    Const adSchemaTables                    As Long = 20
    Const adSchemaPrimaryKeys               As Long = 28
    Const adOpenStatic                      As Long = 3
    Const adLockReadOnly                    As Long = 1
    Dim oSchema         As Object
    Dim cTables         As VBA.Collection
    Dim cPkNames        As VBA.Collection
    Dim vName           As Variant
    Dim oAdoRs          As Object
    Dim lDone           As Long
    Dim aTypes()        As String
    Dim lIdx            As Long
    Dim lMaxLen         As Long

    Set cTables = New VBA.Collection
    Set oSchema = Cnn.OpenSchema(adSchemaTables)
    Do While Not oSchema.EOF
        If StrComp(CStr(oSchema.Fields("TABLE_TYPE").Value), "TABLE", vbTextCompare) = 0 Then
            cTables.Add CStr(oSchema.Fields("TABLE_NAME").Value)
        End If
        oSchema.MoveNext
    Loop
    sCnn.BeginTrans "tc6_convertdatabase"
    On Error GoTo EH
    For Each vName In cTables
        Set cPkNames = New VBA.Collection
        Set oSchema = Cnn.OpenSchema(adSchemaPrimaryKeys, Array(Empty, Empty, vName))
        Do While Not oSchema.EOF
            cPkNames.Add CStr(oSchema.Fields("COLUMN_NAME").Value)
            oSchema.MoveNext
        Loop
        '--- decltypes from the columns schema rowset (carries the sizes the
        '--- recordset fields lack)
        Set oAdoRs = CreateObject("ADODB.Recordset")
        oAdoRs.Open "SELECT * FROM [" & vName & "]", Cnn, adOpenStatic, adLockReadOnly
        ReDim aTypes(0 To oAdoRs.Fields.Count) As String
        Set oSchema = Cnn.OpenSchema(adSchemaColumns, Array(Empty, Empty, vName))
        Do While Not oSchema.EOF
            lIdx = oSchema.Fields("ORDINAL_POSITION").Value - 1
            If lIdx >= 0 And lIdx < oAdoRs.Fields.Count Then
                lMaxLen = 0
                If Not IsNull(oSchema.Fields("CHARACTER_MAXIMUM_LENGTH").Value) Then
                    lMaxLen = oSchema.Fields("CHARACTER_MAXIMUM_LENGTH").Value
                End If
                aTypes(lIdx) = AdoTypeToDeclType(oSchema.Fields("DATA_TYPE").Value, lMaxLen)
                If Not CBool(oSchema.Fields("IS_NULLABLE").Value) Then
                    If Not pvContains(cPkNames, CStr(oSchema.Fields("COLUMN_NAME").Value)) Then
                        aTypes(lIdx) = aTypes(lIdx) & " NOT NULL"
                    End If
                End If
            End If
            oSchema.MoveNext
        Loop
        CopyAdoRsToTable sCnn, CStr(vName), oAdoRs, aTypes, m_bNoCase, cPkNames, Me
        oAdoRs.Close
        lDone = lDone + 1
        RaiseEvent SchemaProgress(CStr(vName), cTables.Count, lDone)
    Next
    sCnn.CommitTrans "tc6_convertdatabase"
    Exit Sub
EH:
    sCnn.RollbackTrans "tc6_convertdatabase"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ConvertIndexes(Cnn As Object, sCnn As cConnection)
    Const adSchemaIndexes                   As Long = 12
    Dim oSchema         As Object
    Dim cNames          As VBA.Collection
    Dim cTabs           As VBA.Collection
    Dim cCols           As VBA.Collection
    Dim cUnique         As VBA.Collection
    Dim sKey            As String
    Dim lIdx            As Long
    Dim lDone           As Long
    Dim sCols           As String

    '--- group index columns per (table, index), skipping primary keys
    Set cNames = New VBA.Collection
    Set cTabs = New VBA.Collection
    Set cCols = New VBA.Collection
    Set cUnique = New VBA.Collection
    Set oSchema = Cnn.OpenSchema(adSchemaIndexes)
    Do While Not oSchema.EOF
        If Not CBool(oSchema.Fields("PRIMARY_KEY").Value) Then
            sKey = "#" & oSchema.Fields("TABLE_NAME").Value & "#" & oSchema.Fields("INDEX_NAME").Value
            On Error Resume Next
            sCols = vbNullString
            sCols = cCols.Item(sKey)
            On Error GoTo 0
            If Len(sCols) = 0 Then
                cNames.Add CStr(oSchema.Fields("INDEX_NAME").Value), sKey
                cTabs.Add CStr(oSchema.Fields("TABLE_NAME").Value), sKey
                cUnique.Add CBool(oSchema.Fields("UNIQUE").Value), sKey
            Else
                cCols.Remove sKey
                sCols = sCols & ", "
            End If
            cCols.Add sCols & "[" & oSchema.Fields("COLUMN_NAME").Value & "]", sKey
        End If
        oSchema.MoveNext
    Loop
    sCnn.BeginTrans "tc6_convertindexes"
    On Error GoTo EH
    For lIdx = 1 To cNames.Count
        sCnn.Execute "CREATE " & IIf(cUnique.Item(lIdx), "UNIQUE ", vbNullString) & "INDEX [idx_" & _
            cNames.Item(lIdx) & "] ON [" & cTabs.Item(lIdx) & "] (" & cCols.Item(lIdx) & ")"
        lDone = lDone + 1
        RaiseEvent IndexProgress(cTabs.Item(lIdx), cNames.Item(lIdx), cNames.Count, lDone)
    Next
    sCnn.CommitTrans "tc6_convertindexes"
    Exit Sub
EH:
    sCnn.RollbackTrans "tc6_convertindexes"
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Friend Sub frInsertProgress(sTable As String, ByVal lCount As Long, ByVal lDone As Long)
    RaiseEvent InsertProgress(sTable, lCount, lDone)
End Sub

Private Function pvContains(cItems As VBA.Collection, sName As String) As Boolean
    Dim vElem           As Variant

    For Each vElem In cItems
        If StrComp(CStr(vElem), sName, vbTextCompare) = 0 Then
            pvContains = True
            Exit Function
        End If
    Next
End Function
