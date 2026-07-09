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
