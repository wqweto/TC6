Attribute VB_Name = "mdEventTests"
'=========================================================================
' mdEventTests - event timing tests for cRecordset/cConnection
'
' Every expected trace below was captured from the original RC6.dll 6.0.15
' with a VBScript event sink (scenarios driven identically against both
' engines), so these tests pin TC6's event timing to RC6's:
'   - Move/AddNew/Delete args are 0-based row indices, -1 = BOF, -2 = EOF
'   - Move-family events fire only when the position actually changes
'   - MoveNext at EOF / MovePrevious at BOF raise vbObjectError, no event
'   - MoveFirst/MoveLast on an empty recordset are silent no-ops
'   - AbsolutePosition is 1-based; Get: -1 = empty, -2 = BOF, -3 = EOF;
'     Let: out of range raises, same row is silent
'   - assigning a Field its current value is a no-op (no FieldChange)
'   - Sort Let rewinds to the first row and always raises Move(0)
'   - AddNew raises AddNew(newIdx) only (no Move despite the cursor move)
'   - ReQuery raises QueryFinished only (cursor rewinds without Move)
'   - Content Let raises nothing
' Documented divergences from RC6 (see NOTES.md): named-savepoint txn
' events (dead code in RC6), SortRefresh cursor drift, ResetChanges Move
' storm under an active sort.
'=========================================================================
Option Explicit

Public Sub RunEventTests()
    Test_MoveEvents
    Test_AbsolutePositionEvents
    Test_AddNewDeleteEvents
    Test_FieldChangeEvents
    Test_SortFindResetEvents
    Test_TxnEvents
End Sub

Private Function pvSeededDb() As cConnection
    Dim oCnn            As cConnection

    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, v TEXT)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'a'), (2, 'b'), (3, 'c')"
    Set pvSeededDb = oCnn
End Function

Private Sub Test_MoveEvents()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oSink           As cEventSink
    Dim lErr            As Long

    If Not TestBegin("cRecordset.MoveEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    Set oSink = New cEventSink
    oSink.Attach oRs
    '--- forward past the end (RC6: EOF surfaces as Move(-2))
    oRs.MoveNext
    oRs.MoveNext
    oRs.MoveNext
    AssertEqStr oSink.Trace, "Move(1);Move(2);Move(-2)", "MoveNext trail into EOF"
    '--- another MoveNext at EOF raises and stays silent
    lErr = 0
    On Error Resume Next
    oRs.MoveNext
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, vbObjectError, "MoveNext at EOF raises vbObjectError"
    AssertEqStr oSink.Trace, "Move(1);Move(2);Move(-2)", "no event from failed MoveNext"
    '--- back from EOF, then same-position moves are silent
    oSink.Clear
    oRs.MovePrevious
    oRs.MoveFirst
    oRs.MoveFirst
    oRs.MoveLast
    oRs.MoveLast
    AssertEqStr oSink.Trace, "Move(2);Move(0);Move(2)", "same-position moves are silent"
    '--- backward past the start (RC6: BOF surfaces as Move(-1))
    oSink.Clear
    oRs.MoveFirst
    oRs.MovePrevious
    AssertEqStr oSink.Trace, "Move(0);Move(-1)", "MovePrevious into BOF"
    lErr = 0
    On Error Resume Next
    oRs.MovePrevious
    lErr = Err.Number
    On Error GoTo EH
    AssertEqLng lErr, vbObjectError, "MovePrevious at BOF raises vbObjectError"
    '--- empty recordset: MoveFirst/MoveLast are silent no-ops
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t WHERE 0")
    oSink.Attach oRs
    oSink.Clear
    oRs.MoveFirst
    oRs.MoveLast
    AssertEqStr oSink.Trace, vbNullString, "MoveFirst/MoveLast silent on empty rs"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_AbsolutePositionEvents()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oSink           As cEventSink
    Dim lErr            As Long

    If Not TestBegin("cRecordset.AbsolutePositionEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    Set oSink = New cEventSink
    oSink.Attach oRs
    AssertEqLng oRs.AbsolutePosition, 1, "1-based initial position"
    '--- Let fires 0-based Move only on an actual change
    oRs.AbsolutePosition = 2
    oRs.AbsolutePosition = 2
    oRs.AbsolutePosition = 3
    AssertEqStr oSink.Trace, "Move(1);Move(2)", "AbsolutePosition moves (change-only)"
    AssertEqLng CLng(oRs.Fields("id").Value), 3, "positioned by 1-based value"
    '--- out-of-range assignments raise without moving
    lErr = 0
    On Error Resume Next
    oRs.AbsolutePosition = 0
    lErr = Err.Number
    oRs.AbsolutePosition = 4
    On Error GoTo EH
    AssertEqLng lErr, vbObjectError, "AbsolutePosition = 0 raises vbObjectError"
    AssertEqStr oSink.Trace, "Move(1);Move(2)", "no event from failed assignments"
    '--- Get codes at BOF/EOF/empty
    oRs.MoveNext
    AssertEqLng oRs.AbsolutePosition, -3, "EOF reports -3"
    oRs.MoveFirst
    oRs.MovePrevious
    AssertEqLng oRs.AbsolutePosition, -2, "BOF reports -2"
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t WHERE 0")
    AssertEqLng oRs.AbsolutePosition, -1, "empty recordset reports -1"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_AddNewDeleteEvents()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oSink           As cEventSink

    If Not TestBegin("cRecordset.AddNewDeleteEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    '--- AddNew raises AddNew(newIdx) only, no Move despite the cursor move
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    Set oSink = New cEventSink
    oSink.Attach oRs
    oRs.AddNew
    oRs.Fields("v").Value = "x"
    AssertEqStr oSink.Trace, "AddNew(3);FieldChange(3,1)", "AddNew + field set on new row"
    '--- delete mid row: cursor stays, the next row slides in
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    oSink.Attach oRs
    oSink.Clear
    oRs.MoveNext
    oRs.Delete
    AssertEqStr oSink.Trace, "Move(1);Delete(1)", "mid delete keeps position"
    AssertEqLng oRs.AbsolutePosition, 2, "still on 1-based position 2"
    AssertEqLng CLng(oRs.Fields("id").Value), 3, "next row slid into place"
    '--- delete the last row: cursor clamps to the new last row
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    oSink.Attach oRs
    oSink.Clear
    oRs.MoveLast
    oRs.Delete
    AssertEqStr oSink.Trace, "Move(2);Delete(1)", "last delete clamps to new last row"
    '--- delete the only row: -1 with BOF and EOF both set
    oCnn.Execute "CREATE TABLE t1(id INTEGER PRIMARY KEY)"
    oCnn.Execute "INSERT INTO t1 VALUES(1)"
    Set oRs = oCnn.OpenRecordset("SELECT id FROM t1")
    oSink.Attach oRs
    oSink.Clear
    oRs.Delete
    AssertEqStr oSink.Trace, "Delete(-1)", "only-row delete reports -1"
    AssertTrue oRs.BOF And oRs.EOF, "BOF and EOF after emptying"
    AssertEqLng oRs.AbsolutePosition, -1, "empty position code"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_FieldChangeEvents()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oSink           As cEventSink

    If Not TestBegin("cRecordset.FieldChangeEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    Set oSink = New cEventSink
    oSink.Attach oRs
    '--- assigning the current value is a complete no-op
    oRs.Fields("v").Value = "a"
    AssertEqStr oSink.Trace, vbNullString, "same-value assignment is silent"
    AssertTrue Not oRs.ContainsChanges, "same-value assignment leaves no changes"
    '--- a real change fires once; repeating it is silent again
    oRs.Fields("v").Value = "z"
    oRs.Fields("v").Value = "z"
    AssertEqStr oSink.Trace, "FieldChange(0,1)", "changed value fires once"
    AssertTrue oRs.ContainsChanges, "change pending"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_SortFindResetEvents()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oSink           As cEventSink

    If Not TestBegin("cRecordset.SortFindResetEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    '--- Sort rewinds to the first row and always raises Move(0)
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    Set oSink = New cEventSink
    oSink.Attach oRs
    oRs.MoveNext
    oRs.Sort = "v DESC"
    AssertEqStr oSink.Trace, "Move(1);Move(0)", "Sort rewinds with a single Move(0)"
    AssertEqLng oRs.AbsolutePosition, 1, "on first row after sort"
    AssertEqLng CLng(oRs.Fields("id").Value), 3, "DESC first row"
    oSink.Clear
    oRs.Sort = vbNullString
    AssertEqStr oSink.Trace, "Move(0)", "clearing Sort also raises Move(0)"
    '--- Find: Move on a hit, silence on a miss
    Set oRs = oCnn.OpenRecordset("SELECT id, v FROM t ORDER BY id")
    oSink.Attach oRs
    oSink.Clear
    AssertTrue oRs.FindFirst("v = 'b'"), "FindFirst hit"
    AssertTrue Not oRs.FindFirst("v = 'nope'"), "FindFirst miss"
    AssertEqStr oSink.Trace, "Move(1)", "Find fires Move on hit only"
    '--- ResetChanges reverts and raises Reset
    oSink.Clear
    oRs.Fields("v").Value = "q"
    oRs.ResetChanges
    AssertEqStr oSink.Trace, "FieldChange(1,1);Reset", "edit + ResetChanges"
    '--- ReQuery raises QueryFinished only, cursor rewinds without Move
    oSink.Clear
    oRs.MoveLast
    oSink.Clear
    oRs.ReQuery
    AssertEqStr oSink.Trace, "QueryFinished", "ReQuery raises QueryFinished only"
    AssertEqLng oRs.AbsolutePosition, 1, "rewound to first row after ReQuery"
    '--- Content Let raises nothing (RC6 verified)
    oSink.Clear
    oRs.Content = oRs.Content
    AssertEqStr oSink.Trace, vbNullString, "Content Let raises no events"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_TxnEvents()
    Dim oCnn            As cConnection
    Dim oSink           As cEventSink

    If Not TestBegin("cConnection.TxnEvents") Then Exit Sub
    On Error GoTo EH
    Set oCnn = pvSeededDb()
    Set oSink = New cEventSink
    oSink.AttachConnection oCnn
    '--- unnamed transactions match RC6 exactly
    oCnn.BeginTrans
    oCnn.CommitTrans
    oCnn.BeginTrans
    oCnn.RollbackTrans
    AssertEqStr oSink.Trace, "CommitTransComplete;RollbackTransComplete", "unnamed txn events"
    '--- named savepoints: TC6 implements the real SAVEPOINT semantics the
    '--- RC6 interface declares; RC6 6.0.15 ignores the names (counter-only
    '--- nesting, savepoint events never fire) - documented divergence
    oSink.Clear
    oCnn.BeginTrans
    oCnn.BeginTrans "sp"
    oCnn.RollbackTrans "sp"
    oCnn.CommitTrans
    AssertEqStr oSink.Trace, "RollbackToSavePoint(sp);CommitTransComplete", "savepoint rollback keeps outer txn"
    oSink.Clear
    oCnn.BeginTrans
    oCnn.BeginTrans "sp"
    oCnn.CommitTrans "sp"
    oCnn.RollbackTrans
    AssertEqStr oSink.Trace, "SavePointReleased(sp);CommitForSavePoint(sp);RollbackTransComplete", "savepoint release then outer rollback"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
