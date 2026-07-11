Attribute VB_Name = "mdCommandTests"
'=========================================================================
' mdCommandTests - tests for cCommand/cSelectCommand/cCursor (run via
' mdTestRunner)
'=========================================================================
Option Explicit

Public Sub RunCommandTests()
    Test_CommandExecuteAndReuse
    Test_CommandNamedParams
    Test_CommandSetters
    Test_SelectCommand
    Test_Cursor
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

Private Sub Test_CommandExecuteAndReuse()
    Dim oCnn            As cConnection
    Dim oCmd            As cCommand
    Dim oRs             As cRecordset

    If Not TestBegin("cCommand.ExecuteAndReuse") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oCmd = oCnn.CreateCommand("INSERT INTO t(name, score) VALUES(?, ?)")
    AssertEqLng oCmd.ParameterCount(), 2, "ParameterCount"
    AssertTrue oCmd.StmtHdl <> 0, "StmtHdl non-zero after prepare"
    oCmd.SetText 1, "delta"
    oCmd.SetDouble 2, 4.5
    AssertTrue oCmd.Execute(), "first Execute succeeds"
    '--- statement is reset by Execute: re-bind and run again
    oCmd.SetText 1, "epsilon"
    oCmd.SetDouble 2, 5.5
    AssertTrue oCmd.Execute(), "second Execute (reuse) succeeds"
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM t")
    AssertEqLng CLng(oRs.Fields("n").Value), 5, "both prepared inserts landed"
    Set oRs = oCnn.GetRs("SELECT name FROM t WHERE score = ?", 5.5)
    AssertEqStr CStr(oRs.Fields("name").Value), "epsilon", "second bind used on reuse"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CommandNamedParams()
    Dim oCnn            As cConnection
    Dim oCmd            As cCommand
    Dim oRs             As cRecordset

    If Not TestBegin("cCommand.NamedParams") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oCmd = oCnn.CreateCommand("INSERT INTO t(name, score) VALUES(:nm, @sc)")
    AssertEqLng oCmd.NameToIdx(":nm"), 1, "NameToIdx with explicit prefix"
    AssertEqLng oCmd.NameToIdx("sc"), 2, "NameToIdx without prefix retries :/@/$"
    AssertEqStr oCmd.IdxToName(1), ":nm", "IdxToName returns the prefixed name"
    oCmd.SetText oCmd.NameToIdx("nm"), "zeta"
    oCmd.SetDouble oCmd.NameToIdx("sc"), 6.5
    AssertTrue oCmd.Execute(), "Execute with named params succeeds"
    Set oRs = oCnn.GetRs("SELECT name FROM t WHERE score = ?", 6.5)
    AssertEqStr CStr(oRs.Fields("name").Value), "zeta", "named-param insert landed"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CommandSetters()
    Dim oCnn            As cConnection
    Dim oCmd            As cCommand
    Dim oRs             As cRecordset
    Dim dSample         As Date
    Dim baBlob(0 To 1)  As Byte
    Dim baOut()         As Byte

    If Not TestBegin("cCommand.Setters") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE v(big INTEGER, d TEXT, sd TEXT, tm TEXT, b INTEGER, data BLOB, x TEXT)"
    dSample = DateSerial(2021, 3, 4) + TimeSerial(5, 6, 7)
    baBlob(0) = 7
    baBlob(1) = 9
    Set oCmd = oCnn.CreateCommand("INSERT INTO v VALUES(?, ?, ?, ?, ?, ?, ?)")
    oCmd.SetInt64 1, CDec("9007199254740993")
    oCmd.SetDate 2, dSample
    oCmd.SetShortDate 3, dSample
    oCmd.SetTime 4, dSample
    oCmd.SetBoolean 5, True
    oCmd.SetBlob 6, baBlob
    oCmd.SetNull 7
    AssertTrue oCmd.Execute(), "Execute with all setters succeeds"
    Set oRs = oCnn.GetRs("SELECT * FROM v")
    AssertTrue oRs.Fields("big").Value = CDec("9007199254740993"), "SetInt64 keeps precision"
    AssertEqStr CStr(oRs.Fields("d").Value), "2021-03-04 05:06:07", "SetDate as ISO text"
    AssertEqStr CStr(oRs.Fields("sd").Value), "2021-03-04", "SetShortDate as ISO date"
    AssertEqStr CStr(oRs.Fields("tm").Value), "05:06:07", "SetTime as ISO time"
    AssertEqLng CLng(oRs.Fields("b").Value), 1, "SetBoolean True as 1"
    baOut = oRs.Fields("data").Value
    AssertEqLng CLng(baOut(1)), 9, "SetBlob round-trips"
    AssertTrue IsNull(oRs.Fields("x").Value), "SetNull binds NULL"
    '--- SetAllParamsNull clears every binding
    oCmd.SetAllParamsNull
    AssertTrue oCmd.Execute(), "Execute after SetAllParamsNull succeeds"
    Set oRs = oCnn.GetRs("SELECT COUNT(*) AS n FROM v WHERE big IS NULL")
    AssertEqLng CLng(oRs.Fields("n").Value), 1, "cleared bindings insert NULLs"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_SelectCommand()
    Dim oCnn            As cConnection
    Dim oCmd            As cSelectCommand
    Dim oRs             As cRecordset

    If Not TestBegin("cSelectCommand.ExecuteAndReuse") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oCmd = oCnn.CreateSelectCommand("SELECT id, name FROM t WHERE id >= :min ORDER BY id")
    AssertEqLng oCmd.ParameterCount(), 1, "ParameterCount"
    oCmd.SetInt32 oCmd.NameToIdx("min"), 2
    Set oRs = oCmd.Execute()
    AssertEqLng oRs.RecordCount, 2, "first Execute returns two rows"
    AssertEqStr CStr(oRs.Fields("name").Value), "beta", "first row of first Execute"
    '--- re-bind and re-run the same prepared statement
    oCmd.SetInt32 1, 3
    Set oRs = oCmd.Execute()
    AssertEqLng oRs.RecordCount, 1, "second Execute (reuse) returns one row"
    AssertEqStr CStr(oRs.Fields("name").Value), "gamma", "first row of second Execute"
    '--- the returned recordset is disconnected: usable after another Execute
    oCmd.SetInt32 1, 1
    AssertEqLng oCmd.Execute().RecordCount, 3, "third Execute returns all rows"
    AssertEqLng oRs.RecordCount, 1, "earlier recordset unaffected (disconnected)"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Cursor()
    Dim oCnn            As cConnection
    Dim oCursor         As cCursor
    Dim lRows           As Long
    Dim sNames          As String

    If Not TestBegin("cCursor.StepAndReset") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oCursor = oCnn.CreateCursor("SELECT id, name FROM t WHERE id >= ? ORDER BY id")
    oCursor.SetInt32 1, 2
    AssertEqLng oCursor.ColCount, 2, "ColCount"
    Do While oCursor.Step()
        lRows = lRows + 1
        sNames = sNames & oCursor.ColVal(1) & ";"
    Loop
    AssertEqLng lRows, 2, "Step visits two rows"
    AssertEqStr sNames, "beta;gamma;", "ColVal reads each row"
    AssertEqLng oCursor.StepCounter, 2, "StepCounter counts steps"
    AssertEqStr oCursor.ColName(1), "name", "ColName"
    '--- Reset rewinds: same bindings, walk again
    oCursor.Reset
    AssertEqLng oCursor.StepCounter, 0, "Reset zeroes StepCounter"
    AssertTrue oCursor.Step(), "Step works again after Reset"
    AssertEqLng CLng(oCursor.ColVal(0)), 2, "first row again after Reset"
    AssertEqLng oCursor.ColType(0), mdSqliteApi.SQLITE_INTEGER, "ColType INTEGER"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
