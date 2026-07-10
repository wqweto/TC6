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

Private Sub Test_FieldOutlivesRecordset()
    Dim oField          As cField
    Dim vValue          As Variant
    Dim bRaised         As Boolean

    If Not TestBegin("cRecordset.FieldOutlivesRecordset") Then Exit Sub
    On Error GoTo EH
    '--- oField survives the recordset; frTerminate must have zeroed its weak
    '--- pointer so this fails safely instead of dereferencing freed memory
    Set oField = pvOrphanField()
    On Error Resume Next
    vValue = oField.Value
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "field access after its recordset is freed raises (no crash)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_FieldInvalidAfterReQuery()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oOldField       As cField
    Dim vValue          As Variant
    Dim bRaised         As Boolean

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
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "field held across ReQuery is invalidated (raises)"
    AssertEqLng CLng(oRs.Fields(0).Value), 1, "fresh field after ReQuery works"
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
