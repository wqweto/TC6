Attribute VB_Name = "mdConnectionTests"
'=========================================================================
' mdConnectionTests - tests for cConnection (run via mdTestRunner)
'=========================================================================
Option Explicit

Public Sub RunConnectionTests()
    Test_Version
    Test_CreateAndInsert
    Test_Pragmas
    Test_Transactions
    Test_Savepoints
    Test_ExecuteRaisesOnBadSql
    Test_ErrorInfo
    Test_DateHelpers
    Test_GetRs
    Test_ExecCmd
    Test_ExecCmdBlob
    Test_BindInt64AndDate
    Test_UniqueID64
    Test_CopyDatabase
    Test_CreateTable
End Sub

Private Sub Test_Version()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.Version") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue Len(oCnn.Version()) > 0, "Version is not empty"
    AssertEqStr Left$(oCnn.Version(), 2), "3.", "Version starts with major 3"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CreateAndInsert()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.CreateAndInsert") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    AssertTrue oCnn.DBHdl <> 0, "DBHdl is non-zero after open"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)"
    oCnn.Execute "INSERT INTO t(v) VALUES('a')"
    AssertEqLng oCnn.AffectedRows, 1, "AffectedRows after single insert"
    oCnn.Execute "INSERT INTO t(v) VALUES('b')"
    oCnn.Execute "INSERT INTO t(v) VALUES('c')"
    AssertEqLng CLng(oCnn.LastInsertAutoID), 3, "LastInsertAutoID after 3 inserts"
    AssertEqStr oCnn.CheckIntegrity(), "ok", "integrity_check ok"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Pragmas()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.Pragmas") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    AssertTrue oCnn.PageSize > 0, "PageSize is positive"
    oCnn.BusyTimeOutSeconds = 2.5
    AssertTrue oCnn.BusyTimeOutSeconds = 2.5, "BusyTimeOutSeconds round-trips"
    '--- setting synchronous must not raise (value not observable on :memory:)
    oCnn.Synchronous = SynchronousNormal
    AssertTrue True, "Synchronous let does not raise"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Transactions()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.Transactions") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    oCnn.Execute "CREATE TABLE t(v INTEGER)"
    AssertEqLng oCnn.TransactionStackCounter, 0, "stack empty at start"
    oCnn.BeginTrans
    AssertEqLng oCnn.TransactionStackCounter, 1, "stack depth 1 after BeginTrans"
    oCnn.Execute "INSERT INTO t VALUES(1)"
    oCnn.CommitTrans
    AssertEqLng oCnn.TransactionStackCounter, 0, "stack empty after CommitTrans"
    oCnn.BeginTrans
    oCnn.Execute "INSERT INTO t VALUES(2)"
    oCnn.RollbackTrans
    AssertEqLng oCnn.TransactionStackCounter, 0, "stack empty after RollbackTrans"
    AssertEqStr oCnn.CheckIntegrity(), "ok", "integrity_check ok after txns"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Savepoints()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.Savepoints") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    oCnn.Execute "CREATE TABLE t(v INTEGER)"
    oCnn.EnableNestedTransactions = True
    oCnn.BeginTrans
    oCnn.BeginTrans "sp1"
    AssertEqLng oCnn.TransactionStackCounter, 2, "depth 2 after outer + savepoint"
    oCnn.Execute "INSERT INTO t VALUES(99)"
    oCnn.RollbackTrans "sp1"
    AssertEqLng oCnn.TransactionStackCounter, 1, "depth 1 after rollback to savepoint"
    oCnn.CommitTrans
    AssertEqLng oCnn.TransactionStackCounter, 0, "stack empty after final commit"
    AssertEqStr oCnn.CheckIntegrity(), "ok", "integrity_check ok after savepoints"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ExecuteRaisesOnBadSql()
    Dim oCnn            As cConnection
    Dim bRaised         As Boolean

    If Not TestBegin("cConnection.ExecuteRaisesOnBadSql") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    On Error Resume Next
    oCnn.Execute "SELECT FROM WHERE"
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "Execute raises on invalid SQL"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ErrorInfo()
    Dim oCnn            As cConnection

    If Not TestBegin("cConnection.ErrorInfo") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    AssertTrue oCnn.CreateNewDB(":memory:"), "CreateNewDB(:memory:) succeeds"
    AssertTrue Len(oCnn.GetSqliteErrStr(1)) > 0, "GetSqliteErrStr(SQLITE_ERROR) not empty"
    On Error Resume Next
    oCnn.Execute "SELECT * FROM nosuchtable"
    On Error GoTo EH
    AssertTrue InStr(oCnn.LastDBError(), "nosuchtable") > 0, "LastDBError mentions the missing table"
    AssertEqLng oCnn.LastDBErrCode(), 1, "LastDBErrCode is SQLITE_ERROR"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_GetRs()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cConnection.GetRs") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER, name TEXT)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'a'), (2, 'b'), (3, 'c')"
    Set oRs = oCnn.GetRs("SELECT id, name FROM t WHERE id >= ? AND name <> ? ORDER BY id", 2, "c")
    AssertEqLng oRs.RecordCount, 1, "parameterised filter returns one row"
    AssertEqLng CLng(oRs.Fields("id").Value), 2, "bound integer parameter"
    AssertEqStr CStr(oRs.Fields("name").Value), "b", "bound text parameter"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ExecCmd()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset

    If Not TestBegin("cConnection.ExecCmd") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER, name TEXT, amount REAL, note TEXT)"
    oCnn.ExecCmd "INSERT INTO t(id, name, amount, note) VALUES(?, ?, ?, ?)", 5, "hello", 3.5, Null
    AssertEqLng oCnn.AffectedRows, 1, "ExecCmd inserts one row"
    Set oRs = oCnn.GetRs("SELECT id, name, amount, note FROM t WHERE id = ?", 5)
    AssertEqLng oRs.RecordCount, 1, "row retrievable by bound id"
    AssertEqStr CStr(oRs.Fields("name").Value), "hello", "bound text stored"
    AssertTrue oRs.Fields("amount").Value = 3.5, "bound double stored"
    AssertTrue IsNull(oRs.Fields("note").Value), "bound Null stored as NULL"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ExecCmdBlob()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim baIn(0 To 2)    As Byte
    Dim baOut()         As Byte

    If Not TestBegin("cConnection.ExecCmdBlob") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE b(data BLOB)"
    baIn(0) = 10
    baIn(1) = 20
    baIn(2) = 30
    oCnn.ExecCmd "INSERT INTO b VALUES(?)", baIn
    Set oRs = oCnn.GetRs("SELECT data FROM b")
    baOut = oRs.Fields("data").Value
    AssertEqLng UBound(baOut) - LBound(baOut) + 1, 3, "bound blob length round-trips"
    AssertEqLng CLng(baOut(0)), 10, "bound blob byte 0 round-trips"
    AssertEqLng CLng(baOut(2)), 30, "bound blob byte 2 round-trips"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_BindInt64AndDate()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim dSample         As Date

    If Not TestBegin("cConnection.BindInt64AndDate") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(big INTEGER, d TEXT, cur REAL)"
    dSample = DateSerial(2020, 1, 2) + TimeSerial(3, 4, 5)
    '--- 2^53+1 is not representable in a Double: proves no double round-trip
    oCnn.ExecCmd "INSERT INTO t VALUES(?, ?, ?)", CDec("9007199254740993"), dSample, CCur(12.34)
    Set oRs = oCnn.GetRs("SELECT big, d, cur, typeof(big) AS tb, typeof(d) AS td FROM t")
    AssertTrue oRs.Fields("big").Value = CDec("9007199254740993"), "integral Decimal binds via int64 (no precision loss)"
    AssertEqStr CStr(oRs.Fields("tb").Value), "integer", "Decimal stored with INTEGER affinity"
    AssertEqStr CStr(oRs.Fields("d").Value), "2020-01-02 03:04:05", "Date binds as ISO text"
    AssertEqStr CStr(oRs.Fields("td").Value), "text", "Date stored as TEXT"
    AssertTrue oRs.Fields("cur").Value = 12.34, "fractional Currency binds as double"
    '--- bound date must equality-match a GetDateString literal
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM t WHERE d = ?", dSample)
    AssertEqLng CLng(oRs.Fields("n").Value), 1, "WHERE d = ? matches the text-stored date"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_UniqueID64()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim vId1            As Variant
    Dim vId2            As Variant
    Dim dStamp          As Date
    Dim dblFrac         As Double

    If Not TestBegin("cConnection.UniqueID64") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    vId1 = oCnn.UniqueID64
    vId2 = oCnn.UniqueID64
    AssertEqLng VarType(vId1), 20, "returns a VT_I8 variant (matches RC6)"
    AssertTrue CDec(vId2) > CDec(vId1), "strictly increasing"
    '--- RC6 encoding: value / 10^14 = local VB date serial
    AssertTrue Abs(CDbl(CDec(vId1) / CDec("100000000000000")) - CDbl(Now)) < 0.01, "value encodes the current date serial"
    oCnn.CreateNewDB ":memory:"
    Set oRs = oCnn.OpenRecordset("SELECT 1")
    dblFrac = -1
    dStamp = oRs.UniqueID64ToVBDate(vId1, dblFrac)
    AssertTrue Abs(DateDiff("s", dStamp, Now)) < 120, "UniqueID64ToVBDate returns the encoded time"
    AssertTrue dblFrac >= 0 And dblFrac < 1, "ReturnFracSeconds is the sub-second remainder"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CopyDatabase()
    Dim oCnn            As cConnection
    Dim oDst            As cConnection

    If Not TestBegin("cConnection.CopyDatabase") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)"
    oCnn.Execute "INSERT INTO t(v) VALUES('a'), ('b'), ('c')"
    Set oDst = oCnn.CopyDatabase()
    AssertEqLng CLng(oDst.GetRs("SELECT COUNT(*) AS n FROM t").Fields("n").Value), 3, "copy holds all rows"
    AssertEqStr CStr(oDst.GetRs("SELECT v FROM t WHERE id = ?", 2).Fields("v").Value), "b", "copied values intact"
    '--- the copy is fully independent of the source
    oCnn.Execute "INSERT INTO t(v) VALUES('d')"
    AssertEqLng CLng(oDst.GetRs("SELECT COUNT(*) AS n FROM t").Fields("n").Value), 3, "copy unaffected by source changes"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CreateTable()
    Dim oCnn            As cConnection
    Dim oFds            As Collection
    Dim oRs             As cRecordset

    If Not TestBegin("cConnection.CreateTable") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    Set oFds = oCnn.NewFieldDefs()
    oFds.Add "F1 TEXT"
    oFds.Add "F2 INTEGER"
    oFds.Add "F3 DOUBLE"
    oCnn.CreateTable "TT", oFds
    Set oRs = oCnn.OpenRecordset("SELECT * FROM TT")
    AssertEqLng CLng(oRs.Fields.Count), 3, "created table has all columns"
    AssertEqStr oRs.Fields(0).Name, "F1", "first column name"
    AssertEqStr oRs.Fields(1).OriginalDataType, "INTEGER", "declared type applied"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_DateHelpers()
    Dim oCnn            As cConnection
    Dim dSample         As Date

    If Not TestBegin("cConnection.DateHelpers") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    dSample = DateSerial(2020, 1, 2) + TimeSerial(3, 4, 5)
    AssertEqStr oCnn.GetDateString(dSample), "2020-01-02 03:04:05", "GetDateString"
    AssertEqStr oCnn.GetShortDateString(dSample), "2020-01-02", "GetShortDateString"
    AssertEqStr oCnn.GetTimeString(dSample), "03:04:05", "GetTimeString"
    AssertEqStr oCnn.GetBooleanString(True), "1", "GetBooleanString(True)"
    AssertEqStr oCnn.GetBooleanString(False), "0", "GetBooleanString(False)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
