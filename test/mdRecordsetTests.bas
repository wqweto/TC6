Attribute VB_Name = "mdRecordsetTests"
'=========================================================================
' mdRecordsetTests - tests for cRecordset (run via mdTestRunner)
'=========================================================================
Option Explicit

Public Sub RunRecordsetTests()
    Test_BasicSelect
    Test_Navigation
    Test_TypesAndValueMatrix
    Test_Blob
    Test_GetRows
    Test_DuplicateFieldNames
    Test_Empty
    Test_OpenSchema
    Test_NoReferenceCycle
    Test_FieldOutlivesRecordset
    Test_FieldInvalidAfterReQuery
    Test_UpdatableAnalysis
    Test_UpdateBatchModify
    Test_AddNewInsert
    Test_DeleteBatch
    Test_ResetChanges
    Test_Sort
    Test_Find
    Test_FieldMetadata
    Test_AutoCreateUniqueID64
End Sub

Private Function pvSeededDb() As cConnection
    Dim oCnn            As cConnection

    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5)"
    oCnn.Execute "INSERT INTO t VALUES(2, 'beta', 2.5)"
    oCnn.Execute "INSERT INTO t VALUES(3, 'gamma', 3.5)"
    Set pvSeededDb = oCnn
End Function

Private Sub Test_BasicSelect()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.BasicSelect") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    AssertEqLng oRs.RecordCount, 3, "RecordCount"
    AssertEqLng CLng(oRs.Fields.Count), 3, "Fields.Count"
    AssertTrue Not oRs.BOF And Not oRs.EOF, "positioned on first row after open"
    AssertEqStr oRs.Fields(1).Name, "name", "Fields(1).Name"
    AssertEqLng CLng(oRs.Fields("id").Value), 1, "Fields(""id"").Value on row 0"
    AssertEqStr CStr(oRs.Fields("name").Value), "alpha", "Fields(""name"").Value on row 0"
    AssertTrue oRs.Fields("score").Value = 1.5, "Fields(""score"").Value on row 0"
    AssertTrue oRs.Fields.Exists("score"), "Fields.Exists(""score"")"
    AssertTrue Not oRs.Fields.Exists("nope"), "Not Fields.Exists(""nope"")"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Navigation()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.Navigation") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t ORDER BY id")
    oRs.MoveFirst
    AssertEqLng oRs.AbsolutePosition, 0, "AbsolutePosition after MoveFirst"
    AssertEqLng CLng(oRs.Fields(0).Value), 1, "id on first row"
    oRs.MoveNext
    AssertEqLng CLng(oRs.Fields(0).Value), 2, "id after MoveNext"
    oRs.MoveLast
    AssertEqLng CLng(oRs.Fields(0).Value), 3, "id after MoveLast"
    AssertTrue Not oRs.EOF, "not EOF while on last row"
    oRs.MoveNext
    AssertTrue oRs.EOF, "EOF after moving past last row"
    oRs.MovePrevious
    AssertEqLng CLng(oRs.Fields(0).Value), 3, "back on last row after MovePrevious"
    oRs.AbsolutePosition = 1
    AssertEqLng CLng(oRs.Fields(0).Value), 2, "row after AbsolutePosition = 1"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_TypesAndValueMatrix()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.TypesAndValueMatrix") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE v(i INTEGER, r REAL, s TEXT, n INTEGER)"
    oCnn.Execute "INSERT INTO v VALUES(42, 3.5, 'hi', NULL)"
    Set oRs = oCnn.OpenRecordset("SELECT i, r, s, n FROM v")
    AssertEqLng oRs.RecordCount, 1, "one row"
    AssertEqLng CLng(oRs.ValueMatrix(0, 0)), 42, "integer via ValueMatrix"
    AssertTrue oRs.ValueMatrix(0, 1) = 3.5, "real via ValueMatrix"
    AssertEqStr CStr(oRs.ValueMatrix(0, 2)), "hi", "text via ValueMatrix"
    AssertTrue IsNull(oRs.ValueMatrix(0, 3)), "NULL cell is Null"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Blob()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim baBlob()        As Byte

    If Not TestBegin("cRecordset.Blob") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE b(data BLOB)"
    oCnn.Execute "INSERT INTO b VALUES(X'01FF80')"
    Set oRs = oCnn.OpenRecordset("SELECT data FROM b")
    baBlob = oRs.Fields("data").Value
    AssertEqLng UBound(baBlob) - LBound(baBlob) + 1, 3, "blob length"
    AssertEqLng CLng(baBlob(0)), 1, "blob byte 0"
    AssertEqLng CLng(baBlob(1)), 255, "blob byte 1"
    AssertEqLng CLng(baBlob(2)), 128, "blob byte 2"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_GetRows()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim vRows           As Variant

    If Not TestBegin("cRecordset.GetRows") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
    vRows = oRs.GetRows()
    AssertEqLng UBound(vRows, 1), 1, "GetRows column upper bound"
    AssertEqLng UBound(vRows, 2), 2, "GetRows row upper bound"
    AssertEqLng CLng(vRows(0, 0)), 1, "GetRows(col0,row0) = first id"
    AssertEqStr CStr(vRows(1, 2)), "gamma", "GetRows(col1,row2) = last name"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_DuplicateFieldNames()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.DuplicateFieldNames") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    '--- two result columns both named "id" (name aliased to id)
    Set oRs = oCnn.OpenRecordset("SELECT id, name AS id FROM t ORDER BY id")
    AssertEqLng CLng(oRs.Fields.Count), 2, "two fields, both named id"
    AssertEqLng CLng(oRs.Fields("id").Value), 1, "Fields(""id"") returns the first occurrence"
    AssertEqStr CStr(oRs.Fields(1).Value), "alpha", "Fields(1) reaches the duplicate by index"
    AssertTrue oRs.Fields.Exists("id"), "Exists(""id"") true despite duplicate"
    AssertTrue Not oRs.Fields.Exists("name"), "Exists(""name"") false (aliased away)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Empty()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.Empty") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t WHERE id > 100")
    AssertEqLng oRs.RecordCount, 0, "empty RecordCount"
    AssertTrue oRs.BOF, "empty recordset BOF"
    AssertTrue oRs.EOF, "empty recordset EOF"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_NoReferenceCycle()
    Dim lBefore         As Long

    If Not TestBegin("cRecordset.NoReferenceCycle") Then Exit Sub
    On Error GoTo EH
    lBefore = g_lLiveRecordsets
    pvMakeAndDropRecordset
    '--- if cField/cFields held strong back-references the recordset would
    '--- never terminate and the live count would stay elevated
    AssertEqLng g_lLiveRecordsets, lBefore, "recordset instance freed after use (no cycle)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub pvMakeAndDropRecordset()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oField          As cField
    Dim vValue          As Variant

    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
    '--- materialise cFields + cField and exercise the weak-ref deref path
    Set oField = oRs.Fields(0)
    vValue = oField.Value
    '--- all locals released on return; the recordset must terminate here
End Sub

Private Sub Test_UpdatableAnalysis()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim bRaised         As Boolean

    If Not TestBegin("cRecordset.UpdatableAnalysis") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t")
    AssertTrue oRs.Updatable, "single-table select with PK is updatable"
    AssertTrue oRs.Fields("name").Updateable, "table column is updateable"
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t", ReadOnly:=True)
    AssertTrue Not oRs.Updatable, "ReadOnly open is not updatable"
    Set oRs = oCnn.OpenRecordset("SELECT COUNT(*) AS n FROM t")
    AssertTrue Not oRs.Updatable, "expression-only select is not updatable"
    On Error Resume Next
    oRs.Fields("n").Value = 42
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "writing a non-updatable recordset raises"
    Set oRs = oCnn.OpenRecordset("SELECT name FROM t")
    AssertTrue Not oRs.Updatable, "select without the PK column is not updatable"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_UpdateBatchModify()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.UpdateBatchModify") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    AssertTrue Not oRs.ContainsChanges, "no changes after open"
    oRs.Fields("name").Value = "ALPHA"
    AssertTrue oRs.ContainsChanges, "ContainsChanges after field write"
    AssertTrue oRs.Fields("name").Changed, "field reports Changed"
    AssertEqStr CStr(oRs.Fields("name").OriginalValue), "alpha", "OriginalValue keeps pre-change value"
    AssertEqStr CStr(oRs.Fields("name").Value), "ALPHA", "Value returns pending value"
    '--- second row via ValueMatrix write
    oRs.ValueMatrix(1, 2) = 9.5
    oRs.UpdateBatch
    AssertTrue Not oRs.ContainsChanges, "no changes after UpdateBatch"
    AssertTrue Not oRs.Fields("name").Changed, "field clean after UpdateBatch"
    '--- verify persisted independently
    Set oRs = oCnn.GetRs("SELECT name, score FROM t WHERE id = ?", 1)
    AssertEqStr CStr(oRs.Fields("name").Value), "ALPHA", "update persisted"
    Set oRs = oCnn.GetRs("SELECT score FROM t WHERE id = ?", 2)
    AssertTrue oRs.Fields("score").Value = 9.5, "ValueMatrix write persisted"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_AddNewInsert()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.AddNewInsert") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.AddNew
    AssertEqLng oRs.AbsolutePosition, 3, "positioned on the new row"
    AssertEqLng oRs.RecordCount, 4, "RecordCount includes pending insert"
    oRs.Fields("name").Value = "delta"
    oRs.Fields("score").Value = 4.5
    oRs.UpdateBatch
    AssertEqLng CLng(oRs.Fields("id").Value), 4, "INTEGER PK backfilled from last rowid"
    Set oRs = oCnn.GetRs("SELECT name FROM t WHERE id = ?", 4)
    AssertEqLng oRs.RecordCount, 1, "inserted row present in DB"
    AssertEqStr CStr(oRs.Fields("name").Value), "delta", "inserted values persisted"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_DeleteBatch()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.DeleteBatch") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
    oRs.MoveFirst
    oRs.Delete
    AssertEqLng oRs.RecordCount, 2, "RecordCount drops on Delete"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "positioned on the following row"
    AssertTrue oRs.ContainsChanges, "delete is a pending change"
    oRs.UpdateBatch
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM t WHERE id = ?", 1)
    AssertEqLng CLng(oRs.Fields("n").Value), 0, "row deleted from DB"
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM t")
    AssertEqLng CLng(oRs.Fields("n").Value), 2, "other rows intact"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ResetChanges()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.ResetChanges") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
    '--- queue one of each pending change, then discard them all
    oRs.Fields("name").Value = "changed"
    oRs.AddNew
    oRs.Fields("name").Value = "pending"
    oRs.MoveLast
    oRs.MoveFirst
    oRs.MoveNext
    oRs.Delete
    AssertTrue oRs.ContainsChanges, "changes pending before reset"
    oRs.ResetChanges
    AssertTrue Not oRs.ContainsChanges, "no changes after ResetChanges"
    AssertEqLng oRs.RecordCount, 3, "RecordCount restored (insert dropped, delete restored)"
    oRs.MoveFirst
    AssertEqStr CStr(oRs.Fields("name").Value), "alpha", "modified value restored"
    '--- DB never touched
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM t")
    AssertEqLng CLng(oRs.Fields("n").Value), 3, "DB unchanged by discarded edits"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Sort()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.Sort") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.MoveFirst
    AssertEqStr CStr(oRs.Fields("name").Value), "alpha", "positioned on alpha before sort"
    oRs.Sort = "name DESC"
    AssertEqStr CStr(oRs.ValueMatrix(0, 1)), "gamma", "first row after DESC sort"
    AssertEqStr CStr(oRs.ValueMatrix(2, 1)), "alpha", "last row after DESC sort"
    AssertEqLng oRs.AbsolutePosition, 2, "cursor followed its row"
    AssertEqStr CStr(oRs.Fields("name").Value), "alpha", "still on alpha after sort"
    '--- pending changes travel with their rows through a sort
    oRs.Fields("score").Value = 9.9
    oRs.Sort = "score"
    oRs.UpdateBatch
    Set oRs = oCnn.GetRs("SELECT score FROM t WHERE name = ?", "alpha")
    AssertTrue oRs.Fields("score").Value = 9.9, "modified cell persisted after sort"
    '--- NULLs sort first; SortRefresh re-applies after data changes
    oCnn.ExecCmd "INSERT INTO t(name, score) VALUES(NULL, 0.5)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t")
    oRs.Sort = "name"
    AssertTrue IsNull(oRs.ValueMatrix(0, 1)), "NULL sorts first ascending"
    oRs.Fields(1).Value = "zzz"
    oRs.SortRefresh
    AssertEqStr CStr(oRs.ValueMatrix(oRs.RecordCount - 1, 1)), "zzz", "SortRefresh re-applies current sort"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Find()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.Find") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    AssertTrue oRs.FindFirst("name = 'beta'"), "FindFirst equality hit"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "positioned on beta"
    AssertTrue Not oRs.FindFirst("name = 'zzz'"), "FindFirst miss returns False"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "miss leaves the position unchanged"
    AssertTrue oRs.FindFirst("score > 1.5"), "FindFirst numeric comparison"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "first score above 1.5"
    AssertTrue oRs.FindNext("score > 1.5"), "FindNext continues past current"
    AssertEqLng CLng(oRs.Fields("id").Value), 3, "next score above 1.5"
    AssertTrue Not oRs.FindNext("score > 1.5"), "FindNext exhausts"
    AssertTrue oRs.FindLast("score < 3"), "FindLast scans backwards"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "last score below 3"
    AssertTrue oRs.FindPrevious("score < 2"), "FindPrevious from current"
    AssertEqLng CLng(oRs.Fields("id").Value), 1, "previous score below 2"
    AssertTrue oRs.FindFirst("name LIKE 'g%'"), "LIKE with SQLite wildcards"
    AssertEqLng CLng(oRs.Fields("id").Value), 3, "LIKE matched gamma"
    AssertTrue oRs.FindFirst("[name] = 'alpha'"), "bracketed field name"
    '--- NULL semantics: IS NULL always works, = NULL needs non-distinct nulls
    oCnn.ExecCmd "INSERT INTO t(name, score) VALUES(NULL, 7)"
    oRs.ReQuery
    AssertTrue oRs.FindFirst("name IS NULL"), "IS NULL matches the null row"
    AssertEqLng CLng(oRs.Fields("id").Value), 4, "positioned on the null row"
    AssertTrue Not oRs.FindFirst("name = NULL"), "= NULL never matches with distinct nulls (default)"
    AssertTrue oRs.FindFirst("name = NULL", False), "= NULL matches with DistinctNullValues=False"
    AssertTrue oRs.FindFirst("name <> 'alpha'", False), "<> treats null as a regular value when non-distinct"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "first non-alpha row"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_FieldMetadata()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.FieldMetadata") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE meta1(id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR(50) NOT NULL DEFAULT 'x' COLLATE NOCASE, email TEXT UNIQUE, score REAL)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name AS alias_name, email, score, score * 2 AS expr FROM meta1")
    AssertEqStr oRs.Fields("alias_name").OriginalColumnName, "name", "OriginalColumnName behind alias"
    AssertEqStr oRs.Fields("alias_name").OriginalTableName, "meta1", "OriginalTableName"
    AssertEqStr oRs.Fields("alias_name").OriginalDataBaseName, "main", "OriginalDataBaseName"
    AssertEqStr oRs.Fields("alias_name").OriginalDataType, "VARCHAR(50)", "OriginalDataType"
    AssertEqLng oRs.Fields("alias_name").DefinedSize, 50, "DefinedSize from declared type"
    AssertEqStr oRs.Fields("alias_name").DefaultValue, "'x'", "DefaultValue verbatim"
    AssertTrue oRs.Fields("alias_name").NotNullConstraint, "NotNullConstraint"
    AssertEqStr oRs.Fields("alias_name").CollationSequence, "NOCASE", "CollationSequence"
    AssertTrue oRs.Fields("id").PrimaryKey, "PrimaryKey"
    AssertTrue oRs.Fields("id").AutoIncrement, "AutoIncrement"
    AssertTrue oRs.Fields("email").UniqueConstraint, "UniqueConstraint via implicit index"
    AssertEqStr oRs.Fields("expr").OriginalColumnName, "", "expression column has no origin"
    AssertTrue Not oRs.Fields("expr").PrimaryKey, "expression column is not PK"
    Set oRs = oCnn.OpenRecordset("SELECT id FROM meta1", ReadOnly:=True)
    AssertTrue oRs.Fields("id").PrimaryKey, "metadata works on a ReadOnly open"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_AutoCreateUniqueID64()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.AutoCreateUniqueID64") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t")
    AssertTrue Not oRs.AutoCreateUniqueID64, "off by default (matches RC6)"
    oRs.AutoCreateUniqueID64 = True
    oRs.AddNew
    AssertTrue Not IsNull(oRs.Fields("id").Value), "INTEGER PK auto-filled on AddNew"
    AssertTrue CDec(oRs.Fields("id").Value) > CDec("4000000000000000000"), "id has time-based magnitude"
    oRs.Fields("name").Value = "auto"
    oRs.UpdateBatch
    Set oRs = oCnn.GetRs("SELECT id FROM t WHERE name = ?", "auto")
    AssertEqLng oRs.RecordCount, 1, "auto-id row persisted"
    AssertTrue CDec(oRs.Fields("id").Value) > CDec("4000000000000000000"), "persisted id keeps full precision"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

'--- these two tests pin the FINAL contract, matching original RC6 behavior
'--- (verified against RC6.dll 3.42.0): an orphaned/stale cField raises
'--- error 91 on every recordset-dependent member (a cRowset split that
'--- would keep it readable was considered and rejected)
Private Sub Test_FieldOutlivesRecordset()
    Dim oField          As cField
    Dim vValue          As Variant
    Dim lErr            As Long

    If Not TestBegin("cRecordset.FieldOutlivesRecordset") Then Exit Sub
    On Error GoTo EH
    '--- oField survives the recordset; frTerminate zeroed its weak pointer,
    '--- so pvRs bails out with error 91 before any dereference (no AV)
    Set oField = pvOrphanField()
    On Error Resume Next
    vValue = oField.Value
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, 91, "field access after its recordset is freed raises 91 (no crash)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Function pvOrphanField() As cField
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t ORDER BY id")
    Set pvOrphanField = oRs.Fields(0)
    '--- oRs (and oCnn) released on return -> cRecordset.Class_Terminate ->
    '--- cFields.frTerminate -> cField.frTerminate clears the weak pointer
End Function

Private Sub Test_FieldInvalidAfterReQuery()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oOldField       As cField
    Dim vValue          As Variant
    Dim lErr            As Long

    If Not TestBegin("cRecordset.FieldInvalidAfterReQuery") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t ORDER BY id")
    Set oOldField = oRs.Fields(0)
    AssertEqLng CLng(oOldField.Value), 1, "field reads its cell before ReQuery"
    oRs.ReQuery
    '--- ReQuery rebuilds the fields; the field held from the prior load must
    '--- be invalidated, not silently bound to the new result set
    On Error Resume Next
    vValue = oOldField.Value
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, 91, "field held across ReQuery is invalidated (raises 91)"
    AssertEqLng CLng(oRs.Fields(0).Value), 1, "fresh field after ReQuery works"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_OpenSchema()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cRecordset.OpenSchema") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenSchema()
    AssertTrue oRs.RecordCount > 0, "OpenSchema returns at least one object"
    AssertTrue oRs.Fields.Exists("name"), "schema recordset has a name column"
    AssertTrue oRs.Fields.Exists("sql"), "schema recordset has a sql column"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
