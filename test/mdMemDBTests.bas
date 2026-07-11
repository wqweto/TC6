Attribute VB_Name = "mdMemDBTests"
'=========================================================================
' mdMemDBTests - tests for cMemDB (run via mdTestRunner)
'=========================================================================
Option Explicit

Public Sub RunMemDBTests()
    Test_MemDBBasics
    Test_MemDBAggregates
    Test_MemDBCreateTableFromRs
End Sub

Private Function pvSeededMemDB() As cMemDB
    Dim oMem            As cMemDB
    Dim oFds            As Collection

    Set oMem = New cMemDB
    Set oFds = oMem.NewFieldDefs()
    oFds.Add "k TEXT"
    oFds.Add "v INTEGER"
    oMem.CreateTable "kv", oFds
    oMem.ExecCmd "INSERT INTO kv VALUES(?, ?)", "a", 1
    oMem.ExecCmd "INSERT INTO kv VALUES(?, ?)", "b", 2
    oMem.Exec "INSERT INTO kv VALUES('c', 3)"
    Set pvSeededMemDB = oMem
End Function

Private Sub Test_MemDBBasics()
    Dim oMem            As cMemDB
    Dim oRs             As cRecordset
    Dim oCmd            As cSelectCommand

    If Not TestBegin("cMemDB.Basics") Then Exit Sub
    On Error GoTo EH
    Set oMem = pvSeededMemDB()
    AssertTrue Not oMem.Cnn Is Nothing, "own :memory: connection on create"
    AssertEqLng CLng(oMem.GetRs("SELECT COUNT(*) AS n FROM kv").Fields("n").Value), 3, "GetRs over seeded rows"
    Set oRs = oMem.GetTable("kv", "v >= 2", "v DESC")
    AssertEqLng oRs.RecordCount, 2, "GetTable with WHERE"
    AssertEqStr CStr(oRs.Fields("k").Value), "c", "GetTable ORDER BY applied"
    AssertEqLng CLng(oMem.GetSingleVal("SELECT v FROM kv WHERE k = 'b'")), 2, "GetSingleVal"
    AssertTrue IsNull(oMem.GetSingleVal("SELECT v FROM kv WHERE k = 'zzz'")), "GetSingleVal on empty result is Null"
    '--- transactions delegate to the inner connection
    oMem.BeginTrans
    oMem.Exec "DELETE FROM kv"
    oMem.RollbackTrans
    AssertEqLng CLng(oMem.GetCount("kv")), 3, "rollback restored the rows"
    '--- prepared statements delegate too
    Set oCmd = oMem.CreateSelectCommand("SELECT COUNT(*) AS n FROM kv WHERE v >= ?")
    oCmd.SetInt32 1, 2
    AssertEqLng CLng(oCmd.Execute().Fields("n").Value), 2, "CreateSelectCommand works"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_MemDBAggregates()
    Dim oMem            As cMemDB

    If Not TestBegin("cMemDB.Aggregates") Then Exit Sub
    On Error GoTo EH
    Set oMem = pvSeededMemDB()
    AssertEqLng CLng(oMem.GetSum("kv", "v")), 6, "GetSum"
    AssertTrue oMem.GetAvg("kv", "v") = 2, "GetAvg"
    AssertEqLng CLng(oMem.GetMin("kv", "v")), 1, "GetMin"
    AssertEqLng CLng(oMem.GetMax("kv", "v")), 3, "GetMax"
    AssertEqLng CLng(oMem.GetCount("kv")), 3, "GetCount(*)"
    AssertEqLng CLng(oMem.GetCount("kv", "v", "v >= 2")), 2, "GetCount with field + WHERE"
    AssertEqLng CLng(oMem.GetSum("kv", "v", "k <> 'a'")), 5, "GetSum with WHERE"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_MemDBCreateTableFromRs()
    Dim oCnn            As cConnection
    Dim oMem            As cMemDB
    Dim oRs             As cRecordset

    If Not TestBegin("cMemDB.CreateTableFromRs") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5), (3, 'gamma', 3.5)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score FROM t ORDER BY id")
    Set oMem = New cMemDB
    oMem.CreateTableFromRs oRs, "t_copy", True
    AssertEqLng CLng(oMem.GetCount("t_copy")), 3, "all rows copied"
    AssertEqStr CStr(oMem.GetSingleVal("SELECT name FROM t_copy WHERE id = 2")), "beta", "values copied"
    AssertTrue oMem.Cnn.DataBases("main").Tables("t_copy").Columns("id").PrimaryKey, "WithPrimaryKeys kept the PK"
    AssertEqStr oMem.Cnn.DataBases("main").Tables("t_copy").Columns("score").ColumnType, "REAL", "declared types carried over"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
