Attribute VB_Name = "mdSchemaTests"
'=========================================================================
' mdSchemaTests - tests for the schema objects (cDataBases/cDataBase/
' cTables/cTable/cColumns/cColumn/cIndexes/cTriggers/cViews), run via
' mdTestRunner
'=========================================================================
Option Explicit

Public Sub RunSchemaTests()
    Test_DataBases
    Test_Tables
    Test_Columns
    Test_Indexes
    Test_TriggersAndViews
End Sub

Private Function pvSchemaDb() As cConnection
    Dim oCnn            As cConnection

    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t1(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL DEFAULT 'x' COLLATE NOCASE, email TEXT UNIQUE, score REAL)"
    oCnn.Execute "CREATE TABLE t2(k TEXT, v TEXT)"
    oCnn.Execute "CREATE INDEX ix_t1_name ON t1(name)"
    oCnn.Execute "CREATE TRIGGER tr_t1_ins AFTER INSERT ON t1 BEGIN UPDATE t1 SET score = 0 WHERE score IS NULL; END"
    oCnn.Execute "CREATE VIEW v_t1 AS SELECT id, name FROM t1 WHERE score > 1"
    Set pvSchemaDb = oCnn
End Function

Private Sub Test_DataBases()
    Dim oCnn            As cConnection
    Dim oDbs            As cDataBases
    Dim oDb             As cDataBase
    Dim sNames          As String

    If Not TestBegin("cDataBases.Basic") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSchemaDb()
    Set oDbs = oCnn.DataBases
    AssertEqLng oDbs.Count, 1, "one database after open"
    AssertEqStr oDbs.Item(0).Name, "main", "Item(0) is main"
    AssertEqStr oDbs.Item("main").NameInBrackets, "[main]", "Item by name + NameInBrackets"
    oDbs.AttachDataBase ":memory:", "aux1"
    AssertEqLng oDbs.Count, 2, "two databases after attach"
    For Each oDb In oDbs
        sNames = sNames & oDb.Name & ";"
    Next
    AssertEqStr sNames, "main;aux1;", "For Each enumerates in attach order"
    oDbs.DetachDataBase "aux1"
    AssertEqLng oDbs.Count, 1, "one database after detach"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Tables()
    Dim oCnn            As cConnection
    Dim oTables         As cTables
    Dim oTable          As cTable
    Dim sNames          As String

    If Not TestBegin("cTables.Basic") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSchemaDb()
    Set oTables = oCnn.DataBases("main").Tables
    '--- sqlite_sequence (from AUTOINCREMENT) is internal and excluded
    AssertEqLng oTables.Count, 2, "two user tables (internal sqlite_* excluded)"
    For Each oTable In oTables
        sNames = sNames & oTable.Name & ";"
    Next
    AssertEqStr sNames, "t1;t2;", "For Each enumerates tables by name"
    AssertEqStr oTables.Item("t1").NameInBrackets, "[t1]", "Item by name + NameInBrackets"
    AssertTrue InStr(1, oTables.Item("t1").SQLForCreate, "CREATE TABLE", vbTextCompare) = 1, "SQLForCreate holds the DDL"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Columns()
    Dim oCnn            As cConnection
    Dim oColumns        As cColumns

    If Not TestBegin("cColumns.Basic") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSchemaDb()
    Set oColumns = oCnn.DataBases("main").Tables("t1").Columns
    AssertEqLng oColumns.Count, 4, "four columns"
    AssertEqStr oColumns.Item(3).Name, "score", "Item by 0-based index"
    AssertTrue oColumns.Item("id").PrimaryKey, "id is PK"
    AssertTrue oColumns.Item("id").PrimaryAutoIncrement, "id is AUTOINCREMENT"
    AssertEqStr oColumns.Item("id").ColumnType, "INTEGER", "id declared type"
    AssertTrue oColumns.Item("name").NotNullConstraint, "name NOT NULL"
    AssertEqStr oColumns.Item("name").DefaultValue, "'x'", "name default (verbatim)"
    AssertEqStr oColumns.Item("name").Collate, "NOCASE", "name collation"
    AssertTrue oColumns.Item("email").UniqueConstraint, "email UNIQUE (single-column unique index)"
    AssertTrue Not oColumns.Item("score").UniqueConstraint, "score not unique"
    AssertTrue Not oColumns.Item("score").PrimaryKey, "score not PK"
    AssertEqStr oColumns.Item("score").ColumnType, "REAL", "score declared type"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Indexes()
    Dim oCnn            As cConnection
    Dim oIndexes        As cIndexes

    If Not TestBegin("cIndexes.Basic") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSchemaDb()
    Set oIndexes = oCnn.DataBases("main").Tables("t1").Indexes
    '--- ix_t1_name + the implicit unique index for email
    AssertEqLng oIndexes.Count, 2, "explicit + implicit unique index"
    AssertTrue InStr(1, oIndexes.Item("ix_t1_name").SQL, "CREATE INDEX", vbTextCompare) = 1, "explicit index DDL"
    AssertEqLng Len(oIndexes.Item(1).SQL), 0, "implicit autoindex has no DDL"
    AssertEqLng oCnn.DataBases("main").Tables("t2").Indexes.Count, 0, "t2 has no indexes"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_TriggersAndViews()
    Dim oCnn            As cConnection
    Dim oTriggers       As cTriggers
    Dim oViews          As cViews

    If Not TestBegin("cTriggersAndViews.Basic") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSchemaDb()
    Set oTriggers = oCnn.DataBases("main").Tables("t1").Triggers
    AssertEqLng oTriggers.Count, 1, "one trigger on t1"
    AssertEqStr oTriggers.Item(0).Name, "tr_t1_ins", "trigger name"
    AssertTrue InStr(1, oTriggers.Item("tr_t1_ins").SQL, "CREATE TRIGGER", vbTextCompare) = 1, "trigger DDL"
    AssertEqLng oCnn.DataBases("main").Tables("t2").Triggers.Count, 0, "t2 has no triggers"
    Set oViews = oCnn.DataBases("main").Views
    AssertEqLng oViews.Count, 1, "one view"
    AssertEqStr oViews.Item("v_t1").Name, "v_t1", "view name"
    AssertTrue InStr(1, oViews.Item(0).SQLForCreate, "CREATE VIEW", vbTextCompare) = 1, "view DDL"
    AssertEqStr Left$(oViews.Item(0).SQL, 6), "SELECT", "view SQL is the SELECT body"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
