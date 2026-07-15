Attribute VB_Name = "mdRecordsetTests"
'=========================================================================
' mdRecordsetTests - tests for cRecordset (run via mdTestRunner)
'=========================================================================
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

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
    Test_ContentRoundTrip
    Test_CreateTableFromRsContent
    Test_ContentRC6Compat
    Test_ContentChangesOnly
    Test_ContentChangesOnlyRC6
    Test_ChangesOnlyUpdatesMultiRows
    Test_ChangesOnlyUpdateToNull
    Test_ChangesOnlyUpdateBlob
    Test_ChangesOnlyUpdateInt64
    Test_ChangesOnlyUpdatePkValue
    Test_ChangesOnlyDeleteMultiple
    Test_ChangesOnlyInsertMultiple
    Test_ChangesOnlyInsertAutoPk
    Test_ChangesOnlyNoChanges
    Test_ChangesOnlyTextPkOps
    Test_ChangesOnlyNotUpdatable
    Test_ChangesOnlyRC6Compat
    Test_GetRowsWithHeaders
    Test_ToJSONUTF8
    Test_JsonRC6Compat
    Test_ReaderCoercions
End Sub

Private Sub Test_ReaderCoercions()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRs2            As cRecordset
    Dim baContent()     As Byte
    Dim sJson           As String
    Dim vRows           As Variant

    If Not TestBegin("cRecordset.ReaderCoercions") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    '--- INTEGER columns type by their load-time width class (RC6-probed)
    oCnn.Execute "CREATE TABLE w(a INTEGER, b INTEGER, c INTEGER, id INTEGER PRIMARY KEY)"
    oCnn.Execute "INSERT INTO w VALUES(1, 1, 1, 1), (200, 70000, 9007199254740993, 2)"
    Set oRs = oCnn.OpenRecordset("SELECT a, b, c, id FROM w ORDER BY id")
    AssertEqStr TypeName(oRs.Fields("a").Value), "Byte", "width-1 column reads Byte"
    AssertEqStr TypeName(oRs.Fields("b").Value), "Long", "width-4 column reads Long"
    AssertEqLng VarType(oRs.Fields("c").Value), 20, "width-8 column reads VT_I8 even for small values"
    AssertEqStr TypeName(oRs.Fields("id").Value), "Long", "pk column is at least width 4"
    AssertTrue CDec(oRs.ValueMatrix(1, 2)) = CDec("9007199254740993"), "int64 keeps precision"
    '--- DATE flavors: text and VB-serial numeric storage, bogus text -> Empty
    oCnn.Execute "CREATE TABLE d(d1 DATE, d2 DATETIME, d3 TIMESTAMP, b1 BIT, b2 BOOLEAN)"
    oCnn.Execute "INSERT INTO d VALUES('2020-02-01', '2020-02-01 10:30:00', 44562.4375, 1, 0)"
    oCnn.Execute "INSERT INTO d VALUES('bogus', NULL, NULL, 2, -1)"
    Set oRs = oCnn.OpenRecordset("SELECT * FROM d")
    AssertTrue oRs.Fields("d1").Value = DateSerial(2020, 2, 1), "DATE text coerces to Date"
    AssertTrue oRs.Fields("d2").Value = DateSerial(2020, 2, 1) + TimeSerial(10, 30, 0), "DATETIME text coerces to Date"
    AssertTrue oRs.Fields("d3").Value = CDate(44562.4375), "numeric date storage is a VB serial"
    AssertEqStr TypeName(oRs.Fields("b1").Value), "Boolean", "BIT coerces to Boolean"
    AssertTrue oRs.Fields("b1").Value, "BIT 1 is True"
    AssertTrue Not oRs.Fields("b2").Value, "BOOLEAN 0 is False"
    oRs.MoveNext
    AssertTrue IsEmpty(oRs.Fields("d1").Value), "unparsable date text reads Empty"
    AssertTrue oRs.Fields("b1").Value, "BIT 2 is True"
    AssertTrue Not oRs.Fields("b2").Value, "BIT -1 is False (RC6: value >= 1)"
    '--- JSON: booleans as literals, dates as quoted raw values
    Set oRs = oCnn.OpenRecordset("SELECT d2, d3, b1 FROM d WHERE b1 = 1")
    sJson = FromUtf8Array(oRs.ToJSONUTF8())
    AssertTrue InStr(sJson, """2020-02-01 10:30:00"",  ""44562.4375"",  true") > 0, "JSON keeps raw dates and boolean literals"
    '--- Content round-trip preserves the coercions
    baContent = oRs.Content
    Set oRs2 = New cRecordset
    oRs2.Content = baContent
    AssertTrue oRs2.Fields("d2").Value = DateSerial(2020, 2, 1) + TimeSerial(10, 30, 0), "date survives Content round-trip"
    AssertEqStr TypeName(oRs2.Fields("b1").Value), "Boolean", "boolean survives Content round-trip"
    '--- GetRows carries the coerced values
    vRows = oRs.GetRows()
    AssertEqStr TypeName(vRows(2, 0)), "Boolean", "GetRows returns coerced cells"
    TestEnd
    Exit Sub
EH:
    TestErr
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
    AssertEqLng oRs.AbsolutePosition, 1, "AbsolutePosition after MoveFirst (1-based)"
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
    oRs.AbsolutePosition = 2
    AssertEqLng CLng(oRs.Fields(0).Value), 2, "row after AbsolutePosition = 2 (1-based)"
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
    AssertTrue IsEmpty(oRs.ValueMatrix(0, 3)), "NULL cell maps to Empty (RC6 default)"
    '--- MapDbNullToEmpty = False surfaces true Nulls instead
    oCnn.MapDbNullToEmpty = False
    Set oRs = oCnn.OpenRecordset("SELECT n FROM v")
    AssertTrue IsNull(oRs.ValueMatrix(0, 0)), "NULL cell stays Null with MapDbNullToEmpty=False"
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
    AssertEqLng oRs.AbsolutePosition, 4, "positioned on the new row (1-based)"
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
    AssertEqLng oRs.AbsolutePosition, 1, "cursor rewound to first row (RC6 semantics)"
    AssertEqStr CStr(oRs.Fields("name").Value), "gamma", "on gamma after sort rewind"
    '--- pending changes travel with their rows through a sort
    oRs.MoveLast
    AssertEqStr CStr(oRs.Fields("name").Value), "alpha", "back on alpha via MoveLast"
    oRs.Fields("score").Value = 9.9
    oRs.Sort = "score"
    oRs.UpdateBatch
    Set oRs = oCnn.GetRs("SELECT score FROM t WHERE name = ?", "alpha")
    AssertTrue oRs.Fields("score").Value = 9.9, "modified cell persisted after sort"
    '--- NULLs sort first; SortRefresh re-applies after data changes
    oCnn.ExecCmd "INSERT INTO t(name, score) VALUES(NULL, 0.5)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t")
    oRs.Sort = "name"
    AssertTrue IsEmpty(oRs.ValueMatrix(0, 1)), "NULL sorts first ascending"
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
    AssertTrue Not IsEmpty(oRs.Fields("id").Value), "INTEGER PK auto-filled on AddNew"
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

Private Sub Test_ContentRoundTrip()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRs2            As cRecordset
    Dim baContent()     As Byte
    Dim baBlob()        As Byte

    If Not TestBegin("cRecordset.ContentRoundTrip") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5, NULL)"
    oCnn.Execute "INSERT INTO t VALUES(2, NULL, NULL, X'010203')"
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score, data FROM t ORDER BY id")
    baContent = oRs.Content
    AssertTrue UBound(baContent) > 100, "Content produces a blob"
    Set oRs2 = New cRecordset
    oRs2.Content = baContent
    AssertEqLng oRs2.RecordCount, 2, "round-trip RecordCount"
    AssertEqLng CLng(oRs2.Fields.Count), 4, "round-trip field count"
    AssertTrue oRs2.Updatable, "round-trip stays updatable"
    oRs2.MoveFirst
    AssertEqLng CLng(oRs2.Fields("id").Value), 1, "row0 id"
    AssertEqStr CStr(oRs2.Fields("name").Value), "alpha", "row0 name"
    AssertTrue oRs2.Fields("score").Value = 1.5, "row0 score"
    oRs2.MoveNext
    AssertEqLng CLng(oRs2.Fields("id").Value), 2, "row1 id"
    AssertTrue IsEmpty(oRs2.Fields("name").Value), "row1 NULL maps to Empty (RC6 default)"
    baBlob = oRs2.Fields("data").Value
    AssertEqLng UBound(baBlob) - LBound(baBlob) + 1, 3, "row1 blob length"
    AssertEqLng CLng(baBlob(2)), 3, "row1 blob bytes"
    AssertEqStr oRs2.Fields("id").OriginalTableName, "t", "metadata carried: origin table"
    AssertEqStr oRs2.Fields("id").OriginalDataType, "INTEGER", "metadata carried: decltype"
    AssertTrue oRs2.Fields("id").PrimaryKey, "metadata carried: PK"
    AssertEqStr oRs2.SQL, "SELECT id, name, score, data FROM t ORDER BY id", "SQL carried"
    '--- re-attach a connection: the disconnected blob is updatable again
    Set oRs2.ActiveConnection = oCnn
    oRs2.MoveFirst
    oRs2.Fields("name").Value = "ALPHA"
    oRs2.UpdateBatch
    AssertEqStr CStr(oCnn.GetRs("SELECT name FROM t WHERE id = ?", 1).Fields("name").Value), "ALPHA", "UpdateBatch after Content re-attach"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CreateTableFromRsContent()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim baContent()     As Byte

    If Not TestBegin("cRecordset.CreateTableFromRsContent") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    baContent = oRs.Content
    oCnn.CreateTableFromRsContent baContent, "t_copy", False, True
    AssertEqLng CLng(oCnn.GetRs("SELECT COUNT(*) AS n FROM t_copy").Fields("n").Value), 3, "content rows copied"
    AssertEqStr CStr(oCnn.GetRs("SELECT name FROM t_copy WHERE id = 2").Fields("name").Value), "beta", "content values copied"
    AssertTrue oCnn.DataBases("main").Tables("t_copy").Columns("id").PrimaryKey, "WithPrimaryKeys honored"
    '--- default table name comes from the blob (fresh connection)
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.CreateTableFromRsContent baContent
    AssertEqLng CLng(oCnn.GetRs("SELECT COUNT(*) AS n FROM t").Fields("n").Value), 3, "default name from blob origin table"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ContentRC6Compat()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRc6Rs          As Object
    Dim baContent()     As Byte

    If Not TestBegin("cRecordset.ContentRC6Compat") Then Exit Sub
    '--- requires the original RC6.dll registered on this machine; each case
    '--- builds identical data in both engines and asserts byte-identity +
    '--- cross-loading in both directions
    On Error Resume Next
    Set oRc6Rs = CreateObject("RC6.cConnection")
    If oRc6Rs Is Nothing Then
        TestSkipCurrent "RC6.dll not registered"
        Exit Sub
    End If
    Set oRc6Rs = Nothing
    On Error GoTo EH
    pvCompatCase "basic", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, NULL, 2.5)", "SELECT id, name, score FROM t ORDER BY id"
    pvCompatCase "int64", "CREATE TABLE t(v INTEGER)", _
        "INSERT INTO t VALUES(9007199254740993), (-9007199254740993), (0)", "SELECT v FROM t"
    pvCompatCase "bytes_width1", "CREATE TABLE t(v INTEGER)", _
        "INSERT INTO t VALUES(0), (255), (128)", "SELECT v FROM t"
    pvCompatCase "int32_width4", "CREATE TABLE t(v INTEGER)", _
        "INSERT INTO t VALUES(-1), (100), (70000)", "SELECT v FROM t"
    '--- NB: values like 1e-300 parse 1 ULP apart in SQLite 3.42 (RC6) vs
    '--- newer engines, so only binary-exact literals are compared here
    pvCompatCase "reals", "CREATE TABLE t(v REAL)", _
        "INSERT INTO t VALUES(0.0), (-1.5), (1e300), (0.125)", "SELECT v FROM t"
    pvCompatCase "text_utf8", "CREATE TABLE t(v TEXT)", _
        "INSERT INTO t VALUES(CAST(X'D0A2D0B5D181D182' AS TEXT)), ('ab''cd'), ('')", "SELECT v FROM t"
    pvCompatCase "text_long", "CREATE TABLE t(v TEXT)", _
        "INSERT INTO t VALUES(replace(hex(zeroblob(200)), '00', 'xy'))", "SELECT v FROM t"
    pvCompatCase "text_null_mix", "CREATE TABLE t(v TEXT)", _
        "INSERT INTO t VALUES('a'), (NULL), (''), ('b')", "SELECT v FROM t"
    pvCompatCase "blob_mix", "CREATE TABLE t(v BLOB)", _
        "INSERT INTO t VALUES(X''), (NULL), (X'00FF10')", "SELECT v FROM t"
    pvCompatCase "blob_big", "CREATE TABLE t(v BLOB)", _
        "INSERT INTO t VALUES(zeroblob(100)), (X'01')", "SELECT v FROM t"
    pvCompatCase "many_rows", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        "INSERT INTO t(v) WITH RECURSIVE c(x) AS (SELECT 1 UNION ALL SELECT x + 1 FROM c WHERE x < 20) SELECT 'r' || x FROM c", _
        "SELECT id, v FROM t ORDER BY id"
    pvCompatCase "expr_num", "CREATE TABLE t(i INTEGER)", _
        "INSERT INTO t VALUES(1), (2)", "SELECT i + 1 AS e, i * 2.5 AS f FROM t"
    pvCompatCase "expr_text", "CREATE TABLE t(i INTEGER)", _
        "INSERT INTO t VALUES(1)", "SELECT 'x' || i AS e FROM t"
    pvCompatCase "mixed_expr", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        "INSERT INTO t(v) VALUES('a')", "SELECT id, v, length(v) AS len FROM t"
    pvCompatCase "join_2tables", "CREATE TABLE a(i INTEGER); CREATE TABLE b(j INTEGER)", _
        "INSERT INTO a VALUES(1); INSERT INTO b VALUES(2)", "SELECT i, j FROM a, b"
    pvCompatCase "no_pk", "CREATE TABLE t(v TEXT)", _
        "INSERT INTO t VALUES('x')", "SELECT v FROM t"
    pvCompatCase "constraints", "CREATE TABLE t(id INTEGER PRIMARY KEY AUTOINCREMENT, u TEXT UNIQUE, d TEXT DEFAULT 'dv', n TEXT NOT NULL DEFAULT 'x' COLLATE NOCASE)", _
        "INSERT INTO t(u, n) VALUES('a', 'b')", "SELECT id, u, d, n FROM t"
    pvCompatCase "exotic_names", "CREATE TABLE [my tab]([my col] TEXT, [select] INTEGER)", _
        "INSERT INTO [my tab] VALUES('v', 1)", "SELECT [my col], [select] FROM [my tab]"
    pvCompatCase "empty_rs", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        vbNullString, "SELECT id, v FROM t"
    pvCompatCase "num_default", "CREATE TABLE t(id INTEGER PRIMARY KEY, q REAL DEFAULT 2.5)", _
        "INSERT INTO t DEFAULT VALUES", "SELECT id, q FROM t"
    pvCompatCase "agg_minmax", "CREATE TABLE t(v INTEGER)", _
        "INSERT INTO t VALUES(5), (1), (9)", "SELECT MAX(v) AS mx, MIN(v) AS mn, COUNT(*) AS cnt FROM t"
    pvCompatCase "agg_group", "CREATE TABLE t(g TEXT, v REAL)", _
        "INSERT INTO t VALUES('a', 1.5), ('a', 2.5), ('b', 10)", "SELECT g, SUM(v) AS s, AVG(v) AS av FROM t GROUP BY g ORDER BY g"
    pvCompatCase "date_bit", "CREATE TABLE t(id INTEGER PRIMARY KEY, d DATETIME, b BIT)", _
        "INSERT INTO t VALUES(1, '2020-02-01 10:30:00', 1), (2, 44562.4375, 0), (3, NULL, NULL)", "SELECT id, d, b FROM t ORDER BY id"
    '--- metadata survives an RC6 load of a TC6 blob
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha')"
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t")
    baContent = oRs.Content
    Set oRc6Rs = CreateObject("RC6.cRecordset")
    oRc6Rs.Content = baContent
    AssertTrue oRc6Rs.Updatable, "RC6 loads TC6 blob: Updatable"
    AssertTrue oRc6Rs.Fields("id").PrimaryKey, "RC6 loads TC6 blob: PK metadata"
    AssertEqStr CStr(oRc6Rs.Fields("id").OriginalDataType), "INTEGER", "RC6 loads TC6 blob: decltype"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub pvCompatCase(sCase As String, sDdl As String, sIns As String, sSel As String)
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRs2            As cRecordset
    Dim oRc6Cnn         As Object
    Dim oRc6Rs          As Object
    Dim baTc6()         As Byte
    Dim baRc6()         As Byte
    Dim lIdx            As Long
    Dim lRow            As Long
    Dim lCol            As Long
    Dim bSame           As Boolean

    '--- identical data in both engines
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute sDdl
    If Len(sIns) > 0 Then
        oCnn.Execute sIns
    End If
    Set oRs = oCnn.OpenRecordset(sSel)
    baTc6 = oRs.Content
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    oRc6Cnn.CreateNewDB ":memory:"
    oRc6Cnn.Execute sDdl
    If Len(sIns) > 0 Then
        oRc6Cnn.Execute sIns
    End If
    Set oRc6Rs = oRc6Cnn.OpenRecordset(sSel)
    baRc6 = oRc6Rs.Content
    '--- byte-identical outside the two ignored heap-pointer slots
    bSame = (UBound(baTc6) = UBound(baRc6))
    If bSame Then
        For lIdx = 0 To UBound(baTc6)
            If baTc6(lIdx) <> baRc6(lIdx) Then
                If Not ((lIdx >= &H1C And lIdx <= &H1F) Or (lIdx >= &H28 And lIdx <= &H2B)) Then
                    bSame = False
                    Exit For
                End If
            End If
        Next
    Else
        lIdx = -1
    End If
    AssertTrue bSame, sCase & ": blob byte-identical to RC6 (size " & (UBound(baTc6) + 1) & " vs " & (UBound(baRc6) + 1) & ", diff at &H" & Hex$(lIdx) & ")"
    '--- RC6 loads the TC6 blob
    Set oRc6Rs = CreateObject("RC6.cRecordset")
    oRc6Rs.Content = baTc6
    AssertEqLng oRc6Rs.RecordCount, oRs.RecordCount, sCase & ": RC6 loads TC6 blob"
    '--- TC6 loads the RC6 blob and every cell matches the original
    Set oRs2 = New cRecordset
    oRs2.Content = baRc6
    bSame = (oRs2.RecordCount = oRs.RecordCount And oRs2.Fields.Count = oRs.Fields.Count)
    If bSame Then
        For lRow = 0 To oRs.RecordCount - 1
            For lCol = 0 To oRs.Fields.Count - 1
                If Not pvCellsEqual(oRs.ValueMatrix(lRow, lCol), oRs2.ValueMatrix(lRow, lCol)) Then
                    bSame = False
                End If
            Next
        Next
    End If
    AssertTrue bSame, sCase & ": TC6 round-trips the RC6 blob cell-for-cell"
End Sub

Private Function pvCellsEqual(vLeft As Variant, vRight As Variant) As Boolean
    Dim baLeft()        As Byte
    Dim baRight()       As Byte
    Dim lIdx            As Long

    If IsNull(vLeft) Or IsEmpty(vLeft) Or IsNull(vRight) Or IsEmpty(vRight) Then
        pvCellsEqual = (IsNull(vLeft) Or IsEmpty(vLeft)) And (IsNull(vRight) Or IsEmpty(vRight))
    ElseIf VarType(vLeft) = vbByte + vbArray Or VarType(vRight) = vbByte + vbArray Then
        If VarType(vLeft) <> VarType(vRight) Then
            Exit Function
        End If
        baLeft = vLeft
        baRight = vRight
        If pvArrLen(baLeft) <> pvArrLen(baRight) Then
            Exit Function
        End If
        For lIdx = 0 To pvArrLen(baLeft) - 1
            If baLeft(lIdx) <> baRight(lIdx) Then
                Exit Function
            End If
        Next
        pvCellsEqual = True
    Else
        On Error Resume Next
        pvCellsEqual = (vLeft = vRight)
    End If
End Function

Private Function pvArrLen(baBuf() As Byte) As Long
    On Error GoTo QH
    pvArrLen = UBound(baBuf) - LBound(baBuf) + 1
QH:
End Function

Private Sub Test_ContentChangesOnly()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim oRs2            As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ContentChangesOnly") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Fields("name").Value = "ALPHA2"
    oRs.MoveNext
    oRs.Fields("score").Value = 9.25
    oRs.MoveLast
    oRs.Delete
    oRs.AddNew
    oRs.Fields("id").Value = 9
    oRs.Fields("name").Value = "inserted"
    oRs.Fields("score").Value = 4.5
    baChanges = oRs.ContentChangesOnly
    AssertTrue UBound(baChanges) > 100, "changes blob produced"
    '--- apply to a second connection holding identical data
    Set oCnn2 = pvSeededDb()
    Set oRs2 = oCnn2.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs2.Content = baChanges
    AssertEqLng oRs2.RecordCount, 0, "changes blob loads with no visible rows"
    AssertTrue oRs2.ContainsChanges, "pending changes present after load"
    AssertTrue Not oRs2.Updatable, "changes-only recordset is not updatable"
    oRs2.UpdateBatch
    AssertTrue Not oRs2.ContainsChanges, "changes cleared after apply"
    Set oRs2 = oCnn2.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    AssertEqLng oRs2.RecordCount, 3, "row count after modify+delete+insert"
    AssertEqStr CStr(oRs2.Fields("name").Value), "ALPHA2", "modified text applied"
    oRs2.MoveNext
    AssertTrue oRs2.Fields("score").Value = 9.25, "modified real applied"
    oRs2.MoveNext
    AssertEqLng CLng(oRs2.Fields("id").Value), 9, "pending insert applied"
    AssertEqStr CStr(oRs2.Fields("name").Value), "inserted", "inserted text applied"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub pvApplyChanges(oCnn As cConnection, baChanges() As Byte)
    Dim oRs             As cRecordset
    Dim baTemp()        As Byte

    '--- VB6 bug: assigning a ByRef array parameter straight to a Property
    '--- Let crashes at runtime - copy to a local first
    baTemp = baChanges
    Set oRs = New cRecordset
    Set oRs.ActiveConnection = oCnn
    oRs.Content = baTemp
    oRs.UpdateBatch
End Sub

Private Sub Test_ChangesOnlyUpdatesMultiRows()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyUpdatesMultiRows") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Fields("name").Value = "N0"
    oRs.MoveNext
    oRs.Fields("name").Value = "N1"
    oRs.MoveNext
    oRs.Fields("score").Value = 7.5
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT name, score FROM t ORDER BY id")
    AssertEqStr CStr(oRs.Fields("name").Value), "N0", "row0 name updated"
    oRs.MoveNext
    AssertEqStr CStr(oRs.Fields("name").Value), "N1", "row1 name updated"
    oRs.MoveNext
    AssertEqStr CStr(oRs.Fields("name").Value), "gamma", "row2 name untouched"
    AssertTrue oRs.Fields("score").Value = 7.5, "row2 score updated"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyUpdateToNull()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyUpdateToNull") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Fields("name").Value = Null
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM t WHERE name IS NULL AND id = 1")
    AssertEqLng CLng(oRs.Fields(0).Value), 1, "NULL update applied in the DB"
    Set oRs = oCnn2.OpenRecordset("SELECT name FROM t WHERE id = 1")
    AssertTrue IsEmpty(oRs.Fields(0).Value), "NULL surfaces as Empty after apply"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyUpdateBlob()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte
    Dim baBlob()        As Byte

    If Not TestBegin("cRecordset.ChangesOnlyUpdateBlob") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE b(id INTEGER PRIMARY KEY, d BLOB)"
    oCnn.Execute "INSERT INTO b VALUES(1, X'AABB')"
    Set oRs = oCnn.OpenRecordset("SELECT id, d FROM b")
    ReDim baBlob(0 To 2) As Byte
    baBlob(0) = 1
    baBlob(1) = 0
    baBlob(2) = 255
    oRs.Fields("d").Value = baBlob
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = New cConnection
    oCnn2.CreateNewDB ":memory:"
    oCnn2.Execute "CREATE TABLE b(id INTEGER PRIMARY KEY, d BLOB)"
    oCnn2.Execute "INSERT INTO b VALUES(1, X'AABB')"
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT d FROM b WHERE id = 1")
    baBlob = oRs.Fields(0).Value
    AssertEqLng UBound(baBlob) - LBound(baBlob) + 1, 3, "blob length updated"
    AssertEqLng CLng(baBlob(0)), 1, "blob byte 0"
    AssertEqLng CLng(baBlob(1)), 0, "blob byte 1 (embedded zero)"
    AssertEqLng CLng(baBlob(2)), 255, "blob byte 2"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyUpdateInt64()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyUpdateInt64") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE w(id INTEGER PRIMARY KEY, v INTEGER)"
    oCnn.Execute "INSERT INTO w VALUES(1, 10)"
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM w")
    oRs.Fields("v").Value = CDec("9007199254740993")
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = New cConnection
    oCnn2.CreateNewDB ":memory:"
    oCnn2.Execute "CREATE TABLE w(id INTEGER PRIMARY KEY, v INTEGER)"
    oCnn2.Execute "INSERT INTO w VALUES(1, 10)"
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT v FROM w WHERE id = 1")
    AssertTrue CDec(oRs.Fields(0).Value) = CDec("9007199254740993"), "int64 value survives full precision"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyUpdatePkValue()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyUpdatePkValue") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.MoveNext
    oRs.Fields("id").Value = 7
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM t WHERE id = 2")
    AssertEqLng CLng(oRs.Fields(0).Value), 0, "WHERE used the original pk value"
    Set oRs = oCnn2.OpenRecordset("SELECT name FROM t WHERE id = 7")
    AssertEqStr CStr(oRs.Fields(0).Value), "beta", "row re-keyed to the new pk"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyDeleteMultiple()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyDeleteMultiple") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Delete
    oRs.Delete
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT id, name FROM t")
    AssertEqLng oRs.RecordCount, 1, "two deletes applied"
    AssertEqStr CStr(oRs.Fields("name").Value), "gamma", "surviving row"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyInsertMultiple()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyInsertMultiple") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.AddNew
    oRs.Fields("id").Value = 9
    oRs.Fields("name").Value = "nine"
    oRs.Fields("score").Value = 9.5
    oRs.AddNew
    oRs.Fields("id").Value = 10
    oRs.Fields("score").Value = 10.5
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM t")
    AssertEqLng CLng(oRs.Fields(0).Value), 5, "both inserts applied"
    Set oRs = oCnn2.OpenRecordset("SELECT name, score FROM t WHERE id = 10")
    AssertTrue IsEmpty(oRs.Fields("name").Value), "unassigned cell inserted as NULL"
    AssertTrue oRs.Fields("score").Value = 10.5, "second insert score"
    Set oRs = oCnn2.OpenRecordset("SELECT name FROM t WHERE id = 9")
    AssertEqStr CStr(oRs.Fields(0).Value), "nine", "first insert name"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyInsertAutoPk()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyInsertAutoPk") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.AddNew
    oRs.Fields("name").Value = "auto"
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = pvSeededDb()
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT id FROM t WHERE name = 'auto'")
    AssertEqLng oRs.RecordCount, 1, "auto-pk insert applied"
    AssertEqLng CLng(oRs.Fields(0).Value), 4, "INTEGER PK auto-assigned on NULL insert"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyNoChanges()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim oRs2            As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyNoChanges") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    baChanges = oRs.ContentChangesOnly
    AssertTrue UBound(baChanges) > 100, "clean recordset still yields a blob"
    Set oCnn2 = pvSeededDb()
    Set oRs2 = New cRecordset
    Set oRs2.ActiveConnection = oCnn2
    oRs2.Content = baChanges
    AssertEqLng oRs2.RecordCount, 0, "no visible rows"
    AssertTrue Not oRs2.ContainsChanges, "no pending changes"
    oRs2.UpdateBatch
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM t")
    AssertEqLng CLng(oRs.Fields(0).Value), 3, "no-op apply leaves the table intact"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyTextPkOps()
    Dim oCnn            As cConnection
    Dim oCnn2           As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ChangesOnlyTextPkOps") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE p(k TEXT PRIMARY KEY, v TEXT)"
    oCnn.Execute "INSERT INTO p VALUES('k1', 'a'), ('k2', 'b'), ('k3', 'c')"
    Set oRs = oCnn.OpenRecordset("SELECT k, v FROM p ORDER BY k")
    oRs.Fields("v").Value = "A2"
    oRs.MoveLast
    oRs.Delete
    baChanges = oRs.ContentChangesOnly
    Set oCnn2 = New cConnection
    oCnn2.CreateNewDB ":memory:"
    oCnn2.Execute "CREATE TABLE p(k TEXT PRIMARY KEY, v TEXT)"
    oCnn2.Execute "INSERT INTO p VALUES('k1', 'a'), ('k2', 'b'), ('k3', 'c')"
    pvApplyChanges oCnn2, baChanges
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM p")
    AssertEqLng CLng(oRs.Fields(0).Value), 2, "text-pk delete applied"
    Set oRs = oCnn2.OpenRecordset("SELECT v FROM p WHERE k = 'k1'")
    AssertEqStr CStr(oRs.Fields(0).Value), "A2", "text-pk update applied"
    Set oRs = oCnn2.OpenRecordset("SELECT COUNT(*) FROM p WHERE k = 'k3'")
    AssertEqLng CLng(oRs.Fields(0).Value), 0, "deleted key gone"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyNotUpdatable()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim baChanges()     As Byte
    Dim lErr            As Long

    If Not TestBegin("cRecordset.ChangesOnlyNotUpdatable") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id + 1 AS e FROM t")
    lErr = 0
    On Error Resume Next
    baChanges = oRs.ContentChangesOnly
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, vbObjectError, "expression recordset raises on ContentChangesOnly"
    Set oRs = oCnn.OpenRecordset("SELECT name FROM t")
    lErr = 0
    On Error Resume Next
    baChanges = oRs.ContentChangesOnly
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, vbObjectError, "recordset without its pk raises on ContentChangesOnly"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ChangesOnlyRC6Compat()
    Dim oProbe          As Object

    If Not TestBegin("cRecordset.ChangesOnlyRC6Compat") Then Exit Sub
    On Error Resume Next
    Set oProbe = CreateObject("RC6.cConnection")
    If oProbe Is Nothing Then
        TestSkipCurrent "RC6.dll not registered"
        Exit Sub
    End If
    On Error GoTo EH
    pvChangesCompatCase "upd_multi", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("upd", 1, "name", "N0"), Array("upd", 2, "name", "N1"), Array("upd", 3, "score", 7.5))
    pvChangesCompatCase "upd_null", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("upd", 1, "name", Null))
    pvChangesCompatCase "upd_blob", "CREATE TABLE b(id INTEGER PRIMARY KEY, d BLOB)", _
        "INSERT INTO b VALUES(1, X'AABB')", "SELECT id, d FROM b", _
        Array(Array("upd", 1, "d", pvTestBlob()))
    pvChangesCompatCase "upd_int64", "CREATE TABLE w(id INTEGER PRIMARY KEY, v INTEGER)", _
        "INSERT INTO w VALUES(1, 10)", "SELECT id, v FROM w", _
        Array(Array("upd", 1, "v", CDec("9007199254740993")))
    pvChangesCompatCase "upd_pk", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("upd", 2, "id", 7&))
    pvChangesCompatCase "del_multi", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("del", 1), Array("del", 1))
    pvChangesCompatCase "ins_multi", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("ins", Array("id", 9&, "name", "nine", "score", 9.5)), Array("ins", Array("id", 10&, "score", 10.5)))
    pvChangesCompatCase "ins_autopk", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("ins", Array("name", "auto")))
    pvChangesCompatCase "mixed", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("upd", 1, "name", "X"), Array("del", 3), Array("ins", Array("id", 9&, "name", "x", "score", 9.5)))
    pvChangesCompatCase "textpk", "CREATE TABLE p(k TEXT PRIMARY KEY, v TEXT)", _
        "INSERT INTO p VALUES('k1', 'a'), ('k2', 'b'), ('k3', 'c')", "SELECT k, v FROM p ORDER BY k", _
        Array(Array("upd", 1, "v", "A2"), Array("del", 3))
    pvChangesCompatCase "sort_mod", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)", "SELECT id, name, score FROM t ORDER BY id", _
        Array(Array("sort", "name DESC"), Array("upd", 1, "score", 6.5))
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Function pvTestBlob() As Byte()
    Dim baBlob()        As Byte

    ReDim baBlob(0 To 2) As Byte
    baBlob(0) = 1
    baBlob(1) = 0
    baBlob(2) = 255
    pvTestBlob = baBlob
End Function

Private Sub pvApplyOps(oRs As Object, vOps As Variant)
    Dim vOp             As Variant
    Dim vPairs          As Variant
    Dim lIdx            As Long

    For Each vOp In vOps
        Select Case CStr(vOp(0))
        Case "upd"
            oRs.AbsolutePosition = vOp(1)
            oRs.Fields(CStr(vOp(2))).Value = vOp(3)
        Case "del"
            oRs.AbsolutePosition = vOp(1)
            oRs.Delete
        Case "ins"
            oRs.AddNew
            vPairs = vOp(1)
            For lIdx = 0 To UBound(vPairs) Step 2
                oRs.Fields(CStr(vPairs(lIdx))).Value = vPairs(lIdx + 1)
            Next
        Case "sort"
            oRs.Sort = CStr(vOp(1))
        End Select
    Next
End Sub

Private Sub pvChangesCompatCase(sCase As String, sDdl As String, sIns As String, sSel As String, vOps As Variant)
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRc6Cnn         As Object
    Dim oRc6Rs          As Object
    Dim baTc6()         As Byte
    Dim baRc6()         As Byte
    Dim sReason         As String
    Dim oCnn2           As cConnection
    Dim oRc6Cnn2        As Object
    Dim oRs2            As cRecordset
    Dim lRow            As Long
    Dim lCol            As Long
    Dim bSame           As Boolean

    '--- identical data + identical pending ops in both engines
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute sDdl
    oCnn.Execute sIns
    Set oRs = oCnn.OpenRecordset(sSel)
    pvApplyOps oRs, vOps
    baTc6 = oRs.ContentChangesOnly
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    oRc6Cnn.CreateNewDB ":memory:"
    oRc6Cnn.Execute sDdl
    oRc6Cnn.Execute sIns
    Set oRc6Rs = oRc6Cnn.OpenRecordset(sSel)
    pvApplyOps oRc6Rs, vOps
    baRc6 = oRc6Rs.ContentChangesOnly
    '--- binary-compatible outside pointer/junk fields
    sReason = pvChangesBlobsDiff(baTc6, baRc6)
    AssertTrue Len(sReason) = 0, sCase & ": changes blob binary-compatible with RC6" & IIf(Len(sReason) > 0, " (" & sReason & ")", vbNullString)
    '--- cross-apply: RC6 applies the TC6 blob, TC6 applies the RC6 blob and
    '--- both tables end up cell-for-cell identical
    Set oRc6Cnn2 = CreateObject("RC6.cConnection")
    oRc6Cnn2.CreateNewDB ":memory:"
    oRc6Cnn2.Execute sDdl
    oRc6Cnn2.Execute sIns
    Set oRc6Rs = oRc6Cnn2.OpenRecordset(sSel)
    oRc6Rs.Content = baTc6
    oRc6Rs.UpdateBatch
    Set oCnn2 = New cConnection
    oCnn2.CreateNewDB ":memory:"
    oCnn2.Execute sDdl
    oCnn2.Execute sIns
    pvApplyChanges oCnn2, baRc6
    Set oRs2 = oCnn2.OpenRecordset(sSel)
    Set oRc6Rs = oRc6Cnn2.OpenRecordset(sSel)
    bSame = (oRs2.RecordCount = CLng(oRc6Rs.RecordCount) And CLng(oRs2.Fields.Count) = CLng(oRc6Rs.Fields.Count))
    If bSame Then
        For lRow = 0 To oRs2.RecordCount - 1
            For lCol = 0 To oRs2.Fields.Count - 1
                If Not pvCellsEqual(oRs2.ValueMatrix(lRow, lCol), oRc6Rs.ValueMatrix(lRow, lCol)) Then
                    bSame = False
                End If
            Next
        Next
    End If
    AssertTrue bSame, sCase & ": cross-applied tables match cell-for-cell"
End Sub

Private Function pvBLong(baBuf() As Byte, lPos As Long) As Long
    Call CopyMemory(pvBLong, baBuf(lPos), 4)
    lPos = lPos + 4
End Function

Private Sub pvBSkipStr(baBuf() As Byte, lPos As Long)
    Dim lSize           As Long

    lSize = pvBLong(baBuf, lPos)
    lPos = lPos + lSize
End Sub

Private Function pvChangesBlobsDiff(baA() As Byte, baB() As Byte) As String
    Dim lPos            As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim lFields         As Long
    Dim lTables         As Long
    Dim lCount          As Long
    Dim lCells          As Long
    Dim aSlot()         As Long
    Dim aVtA()          As Long
    Dim aLoA()          As Long
    Dim lStream         As Long
    Dim lRowsC          As Long
    Dim lTmp            As Long

    If UBound(baA) <> UBound(baB) Then
        pvChangesBlobsDiff = "size " & (UBound(baA) + 1) & " vs " & (UBound(baB) + 1)
        Exit Function
    End If
    '--- walk blob A to find the cells offset; the whole head is
    '--- deterministic so it must be byte-identical
    lPos = 4
    pvBLong baA, lPos
    pvBSkipStr baA, lPos
    pvBSkipStr baA, lPos
    lFields = pvBLong(baA, lPos)
    For lIdx = 1 To lFields
        For lJdx = 1 To 7
            pvBSkipStr baA, lPos
        Next
        lPos = lPos + 20
    Next
    lTables = pvBLong(baA, lPos)
    For lIdx = 1 To lTables
        pvBSkipStr baA, lPos
        pvBSkipStr baA, lPos
        pvBSkipStr baA, lPos
        lCount = pvBLong(baA, lPos)
        lPos = lPos + 4 * lCount
    Next
    lTmp = pvChgCmpRange(baA, baB, 4, lPos - 4)
    If lTmp >= 0 Then
        pvChangesBlobsDiff = "head diff at &H" & Hex$(lTmp)
        Exit Function
    End If
    '--- cells: count + slots exact, records vt exact, unions masked by kind
    lCells = pvBLong(baA, lPos)
    lTmp = pvChgCmpRange(baA, baB, lPos - 4, 4 + 4 * lCells)
    If lTmp >= 0 Then
        pvChangesBlobsDiff = "cell slots diff at &H" & Hex$(lTmp)
        Exit Function
    End If
    If lCells > 0 Then
        ReDim aSlot(0 To lCells - 1) As Long
        ReDim aVtA(0 To lCells - 1) As Long
        ReDim aLoA(0 To lCells - 1) As Long
    End If
    For lIdx = 0 To lCells - 1
        aSlot(lIdx) = pvBLong(baA, lPos)
    Next
    lPos = lPos + 16
    For lIdx = 0 To lCells - 1
        aVtA(lIdx) = pvBLong(baA, lPos)
        If pvChgCmpRange(baA, baB, lPos - 4, 4) >= 0 Then
            pvChangesBlobsDiff = "cell " & lIdx & " vt " & aVtA(lIdx) & " vs &H" & Hex$(pvBLongAt(baB, lPos - 4))
            Exit Function
        End If
        pvBLong baA, lPos
        aLoA(lIdx) = pvBLongAt(baA, lPos)
        lTmp = -1
        If aSlot(lIdx) < 0 Then
            lTmp = pvChgCmpRange(baA, baB, lPos, 4)
        Else
            Select Case aVtA(lIdx)
            Case 0, 1
            Case 2
                lTmp = pvChgCmpRange(baA, baB, lPos, 2)
            Case 3
                lTmp = pvChgCmpRange(baA, baB, lPos, 4)
            Case Else
                lTmp = pvChgCmpRange(baA, baB, lPos, 8)
            End Select
        End If
        If lTmp >= 0 Then
            pvChangesBlobsDiff = "cell " & lIdx & " union diff at &H" & Hex$(lTmp)
            Exit Function
        End If
        lPos = lPos + 8
    Next
    '--- shared text stream
    lStream = 0
    For lIdx = 0 To lCells - 1
        If aSlot(lIdx) < 0 Then
            lStream = lStream + Abs(aLoA(lIdx))
        End If
    Next
    lTmp = pvChgCmpRange(baA, baB, lPos, lStream + 1)
    If lTmp >= 0 Then
        pvChangesBlobsDiff = "text stream diff at &H" & Hex$(lTmp)
        Exit Function
    End If
    lPos = lPos + lStream + 1
    '--- WHERE rows: labels exact, pk records masked like cells
    lRowsC = pvBLong(baA, lPos)
    lTmp = pvChgCmpRange(baA, baB, lPos - 4, 4 + 1 + 4 * lRowsC)
    If lTmp >= 0 Then
        pvChangesBlobsDiff = "row labels diff at &H" & Hex$(lTmp)
        Exit Function
    End If
    lPos = lPos + 1 + 4 * lRowsC + 15
    For lIdx = 0 To lRowsC - 1
        If pvChgCmpRange(baA, baB, lPos, 4) >= 0 Then
            pvChangesBlobsDiff = "pk " & lIdx & " vt diff"
            Exit Function
        End If
        lCount = pvBLongAt(baA, lPos)
        lTmp = -1
        Select Case lCount
        Case 5, 14, 20
            lTmp = pvChgCmpRange(baA, baB, lPos + 8, 8)
        Case Else
            lTmp = pvChgCmpRange(baA, baB, lPos + 8, 4)
        End Select
        If lTmp >= 0 Then
            pvChangesBlobsDiff = "pk " & lIdx & " value diff at &H" & Hex$(lTmp)
            Exit Function
        End If
        lPos = lPos + 16
    Next
    '--- pk text stream + trailing byte + padding are deterministic
    lTmp = pvChgCmpRange(baA, baB, lPos, UBound(baA) + 1 - lPos)
    If lTmp >= 0 Then
        pvChangesBlobsDiff = "tail diff at &H" & Hex$(lTmp)
    End If
End Function

Private Function pvBLongAt(baBuf() As Byte, ByVal lPos As Long) As Long
    Call CopyMemory(pvBLongAt, baBuf(lPos), 4)
End Function

Private Function pvChgCmpRange(baA() As Byte, baB() As Byte, ByVal lFrom As Long, ByVal lLen As Long) As Long
    Dim lIdx            As Long

    pvChgCmpRange = -1
    For lIdx = lFrom To lFrom + lLen - 1
        If baA(lIdx) <> baB(lIdx) Then
            pvChgCmpRange = lIdx
            Exit Function
        End If
    Next
End Function

Private Sub Test_GetRowsWithHeaders()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim vRows           As Variant

    If Not TestBegin("cRecordset.GetRowsWithHeaders") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
    vRows = oRs.GetRowsWithHeaders()
    AssertEqLng UBound(vRows, 1), 1, "cols dimension"
    AssertEqLng UBound(vRows, 2), 3, "rows dimension incl header"
    AssertEqStr CStr(vRows(0, 0)), "id", "header at row 0"
    AssertEqStr CStr(vRows(1, 0)), "name", "second header"
    AssertEqLng CLng(vRows(0, 1)), 1, "first data row shifted"
    AssertEqStr CStr(vRows(1, 3)), "gamma", "last data row"
    vRows = oRs.GetRowsWithHeaders(-1, 0, vbNullString, False, True)
    AssertEqLng LBound(vRows, 2), -1, "HeaderAtIdxMinus1 lower bound"
    AssertEqStr CStr(vRows(1, -1)), "name", "header at index -1"
    AssertEqLng CLng(vRows(0, 0)), 1, "data at index 0"
    vRows = oRs.GetRowsWithHeaders(-1, 0, vbNullString, True)
    AssertEqLng UBound(vRows, 1), 3, "transposed rows dimension incl header"
    AssertEqLng UBound(vRows, 2), 1, "transposed cols dimension"
    AssertEqStr CStr(vRows(0, 1)), "name", "transposed header"
    AssertEqStr CStr(vRows(2, 1)), "beta", "transposed data"
    vRows = oRs.GetRowsWithHeaders(2, 1, "name")
    AssertEqLng UBound(vRows, 1), 0, "field list restricts columns"
    AssertEqLng UBound(vRows, 2), 2, "RowCount counts data rows"
    AssertEqStr CStr(vRows(0, 0)) & "," & CStr(vRows(0, 1)) & "," & CStr(vRows(0, 2)), "name,beta,gamma", "subset values"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ToJSONUTF8()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim sJson           As String

    If Not TestBegin("cRecordset.ToJSONUTF8") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5, X'01FF'), (2, NULL, NULL, NULL)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score, data FROM t ORDER BY id")
    sJson = FromUtf8Array(oRs.ToJSONUTF8())
    AssertEqStr sJson, "{""RecordCount"": 2,""Fields"": [{ ""Name"": ""id"", ""Type"": ""INTEGER"", ""PrimaryKey"": true, ""Nullable"": true, ""DefaultValue"": ""NULL""},{ ""Name"": ""name"", ""Type"": ""TEXT"", ""PrimaryKey"": false, ""Nullable"": true, ""DefaultValue"": ""NULL""},{ ""Name"": ""score"", ""Type"": ""REAL"", ""PrimaryKey"": false, ""Nullable"": true, ""DefaultValue"": ""NULL""},{ ""Name"": ""data"", ""Type"": ""BLOB"", ""PrimaryKey"": false, ""Nullable"": true, ""DefaultValue"": ""NULL""}],""RowsCols"": [[ 1,  ""alpha"",  1.5,  ""Af8=""],[ 2,  null,  null,  null]]}", "compact JSON"
    sJson = FromUtf8Array(oRs.ToJSONUTF8(True))
    AssertTrue InStr(sJson, """ColsRows"": [[ 1,  2],[ ""alpha"",  null]") > 0, "ColsRows transposed values"
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t WHERE 0")
    sJson = FromUtf8Array(oRs.ToJSONUTF8())
    AssertTrue InStr(sJson, """RowsCols"": []}") > 0, "empty recordset serializes []"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub pvJsonCompatCase(sCase As String, sDdl As String, sIns As String, sSel As String, ByVal bColsRows As Boolean, ByVal bytIndent As Byte, ByVal bUniEscaping As Boolean, ByVal bytLfChar As Byte)
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRc6Cnn         As Object
    Dim oRc6Rs          As Object
    Dim baTc6()         As Byte
    Dim baRc6()         As Byte
    Dim lIdx            As Long

    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute sDdl
    If Len(sIns) > 0 Then
        oCnn.Execute sIns
    End If
    Set oRs = oCnn.OpenRecordset(sSel)
    baTc6 = oRs.ToJSONUTF8(bColsRows, bytIndent, bUniEscaping, bytLfChar)
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    oRc6Cnn.CreateNewDB ":memory:"
    oRc6Cnn.Execute sDdl
    If Len(sIns) > 0 Then
        oRc6Cnn.Execute sIns
    End If
    Set oRc6Rs = oRc6Cnn.OpenRecordset(sSel)
    baRc6 = oRc6Rs.ToJSONUTF8(bColsRows, bytIndent, bUniEscaping, bytLfChar)
    lIdx = -1
    If pvArrLen(baTc6) = pvArrLen(baRc6) Then
        lIdx = pvChgCmpRange(baTc6, baRc6, 0, pvArrLen(baTc6))
    End If
    AssertTrue pvArrLen(baTc6) = pvArrLen(baRc6) And lIdx < 0, sCase & ": JSON byte-identical to RC6 (size " & pvArrLen(baTc6) & " vs " & pvArrLen(baRc6) & ", diff at " & lIdx & ")"
End Sub

Private Sub Test_JsonRC6Compat()
    Dim oProbe          As Object

    If Not TestBegin("cRecordset.JsonRC6Compat") Then Exit Sub
    On Error Resume Next
    Set oProbe = CreateObject("RC6.cConnection")
    If oProbe Is Nothing Then
        TestSkipCurrent "RC6.dll not registered"
        Exit Sub
    End If
    On Error GoTo EH
    pvJsonCompatCase "mixed_types", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5, X'01FF'), (2, NULL, NULL, NULL), (3, 'q""uo' || char(10) || 'te', -0.25, X'')", _
        "SELECT id, name, score, data FROM t ORDER BY id", False, 0, False, 0
    pvJsonCompatCase "colsrows", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        "INSERT INTO t VALUES(1, 'a'), (2, 'b')", "SELECT id, v FROM t ORDER BY id", True, 0, False, 0
    pvJsonCompatCase "indent2", "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB)", _
        "INSERT INTO t VALUES(1, 'alpha', 1.5, X'01FF'), (2, NULL, NULL, NULL)", _
        "SELECT id, name, score, data FROM t ORDER BY id", False, 2, False, 0
    pvJsonCompatCase "indent3_lf10", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        "INSERT INTO t VALUES(1, 'a')", "SELECT id, v FROM t", False, 3, False, 10
    pvJsonCompatCase "unicode_raw", "CREATE TABLE u(v TEXT)", _
        "INSERT INTO u VALUES(CAST(X'D0A2D0B5D181D182' AS TEXT))", "SELECT v FROM u", False, 0, False, 0
    pvJsonCompatCase "unicode_esc", "CREATE TABLE u(v TEXT)", _
        "INSERT INTO u VALUES(CAST(X'D0A2D0B5D181D182' AS TEXT) || 'x' || char(255))", "SELECT v FROM u", False, 0, True, 0
    pvJsonCompatCase "escapes", "CREATE TABLE u(v TEXT)", _
        "INSERT INTO u VALUES('tab' || char(9) || 'bs\' || char(1) || char(31) || 'cr' || char(13) || 'ff' || char(12) || 'end')", _
        "SELECT v FROM u", False, 0, False, 0
    pvJsonCompatCase "reals_dec", "CREATE TABLE r(v REAL, w REAL NOT NULL DEFAULT 3, x TEXT DEFAULT 'dv')", _
        "INSERT INTO r VALUES(0.1, 1e300, 'a'), (1.5E-5, -0.5, 'b')", "SELECT v, w, x FROM r", False, 0, False, 0
    pvJsonCompatCase "int64", "CREATE TABLE w(v INTEGER)", _
        "INSERT INTO w VALUES(9007199254740993), (-9007199254740993)", "SELECT v FROM w", False, 0, False, 0
    pvJsonCompatCase "expressions", "CREATE TABLE r(v REAL, x TEXT)", _
        "INSERT INTO r VALUES(0.1, 'a'), (2.5, 'b')", "SELECT v + 1 AS e, 'x' || x AS f FROM r", False, 0, False, 0
    pvJsonCompatCase "empty_rs", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        vbNullString, "SELECT id, v FROM t", False, 0, False, 0
    pvJsonCompatCase "empty_indent", "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)", _
        vbNullString, "SELECT id, v FROM t", True, 2, False, 0
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ContentChangesOnlyRC6()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oRc6Cnn         As Object
    Dim oRc6Rs          As Object
    Dim baChanges()     As Byte

    If Not TestBegin("cRecordset.ContentChangesOnlyRC6") Then Exit Sub
    On Error Resume Next
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    If oRc6Cnn Is Nothing Then
        TestSkipCurrent "RC6.dll not registered"
        Exit Sub
    End If
    On Error GoTo EH
    '--- RC6 applies a TC6 changes blob (modify + delete + insert)
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Fields("name").Value = "ALPHA2"
    oRs.MoveLast
    oRs.Delete
    oRs.AddNew
    oRs.Fields("id").Value = 9
    oRs.Fields("name").Value = "ins"
    oRs.Fields("score").Value = 4.5
    baChanges = oRs.ContentChangesOnly
    oRc6Cnn.CreateNewDB ":memory:"
    oRc6Cnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)"
    oRc6Cnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)"
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRc6Rs.Content = baChanges
    oRc6Rs.UpdateBatch
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT COUNT(*), SUM(id) FROM t")
    AssertEqLng CLng(oRc6Rs.Fields(0).Value), 3, "RC6 applied: row count"
    AssertEqLng CLng(oRc6Rs.Fields(1).Value), 12, "RC6 applied: delete + insert (ids 1+2+9)"
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT name FROM t WHERE id = 1")
    AssertEqStr CStr(oRc6Rs.Fields(0).Value), "ALPHA2", "RC6 applied: modified text"
    '--- TC6 applies an RC6 changes blob
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRc6Rs.Fields("name").Value = "RCMOD"
    oRc6Rs.MoveLast
    oRc6Rs.Delete
    baChanges = oRc6Rs.ContentChangesOnly
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'ALPHA2', 1.5), (2, 'beta', 2.5), (9, 'ins', 4.5)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    oRs.Content = baChanges
    AssertTrue oRs.ContainsChanges, "TC6 loads RC6 changes blob with pending changes"
    oRs.UpdateBatch
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    AssertEqLng oRs.RecordCount, 2, "TC6 applied: RC6 delete"
    AssertEqStr CStr(oRs.Fields("name").Value), "RCMOD", "TC6 applied: RC6 modified text"
    '--- text pk labels survive the RC6 loader validation
    oCnn.Execute "CREATE TABLE p(k TEXT PRIMARY KEY, v TEXT)"
    oCnn.Execute "INSERT INTO p VALUES('k1', 'a'), ('k2', 'b')"
    Set oRs = oCnn.OpenRecordset("SELECT k, v FROM p ORDER BY k")
    oRs.MoveNext
    oRs.Fields("v").Value = "B2"
    baChanges = oRs.ContentChangesOnly
    oRc6Cnn.Execute "CREATE TABLE p(k TEXT PRIMARY KEY, v TEXT)"
    oRc6Cnn.Execute "INSERT INTO p VALUES('k1', 'a'), ('k2', 'b')"
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT k, v FROM p ORDER BY k")
    oRc6Rs.Content = baChanges
    oRc6Rs.UpdateBatch
    Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT v FROM p WHERE k = 'k2'")
    AssertEqStr CStr(oRc6Rs.Fields(0).Value), "B2", "RC6 applied: TC6 text-pk modify"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

'--- these two tests pin the FINAL contract, matching original RC6 behavior
'--- (verified against RC6.dll 6.0.15): an orphaned/stale cField raises
'--- error 91 on every recordset-dependent member (a cRowset split that
'--- would keep it readable was considered and rejected)
Private Sub Test_FieldOutlivesRecordset()
    Dim oField          As cField
    Dim vValue          As Variant
    Dim lErr            As Long

    If Not TestBegin("cRecordset.FieldOutlivesRecordset") Then Exit Sub
    On Error GoTo EH
    '--- oField survives the recordset; frTerminate zeroed its weak ref,
    '--- so member access raises error 91 before any dereference (no AV)
    Set oField = pvOrphanField()
    On Error Resume Next
    vValue = oField.Value
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, 91, "field access after its recordset is freed raises 91 (no crash)"
    On Error Resume Next
    vValue = oField.OriginalValue
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, 91, "OriginalValue on orphan raises 91 (no crash)"
    On Error Resume Next
    vValue = oField.UnderlyingValue
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, 91, "UnderlyingValue on orphan raises 91 (no crash)"
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
