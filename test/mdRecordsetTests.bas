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
    Test_Empty
    Test_OpenSchema
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
