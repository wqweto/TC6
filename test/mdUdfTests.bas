Attribute VB_Name = "mdUdfTests"
'=========================================================================
' mdUdfTests - tests for user-defined functions/aggregates/collations
' (cUDFMethods + cConnection.AddUserDefined*), run via mdTestRunner
'=========================================================================
Option Explicit

Public Sub RunUdfTests()
    Test_ScalarUdf
    Test_AggregateUdf
    Test_Collation
End Sub

Private Sub Test_ScalarUdf()
    Dim oCnn            As cConnection
    Dim oFuncs          As cTestFuncs
    Dim oRs             As cRecordset
    Dim bRaised         As Boolean
    Dim sError          As String

    If Not TestBegin("cUDFMethods.ScalarUdf") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    Set oFuncs = New cTestFuncs
    AssertTrue oCnn.AddUserDefinedFunction(oFuncs), "AddUserDefinedFunction succeeds"
    Set oRs = oCnn.GetRs("SELECT REV('abc') AS r, TWICE(21) AS t")
    AssertEqStr CStr(oRs.Fields("r").Value), "cba", "REV reverses text (name index 0)"
    AssertEqLng CLng(oRs.Fields("t").Value), 42, "TWICE doubles int (name index 1)"
    '--- SetResultError surfaces as a statement error with the message
    On Error Resume Next
    oCnn.GetRs "SELECT ERRF()"
    bRaised = (Err.Number <> 0)
    sError = Err.Description
    On Error GoTo EH
    AssertTrue bRaised, "SetResultError fails the statement"
    AssertTrue InStr(1, sError, "boom", vbTextCompare) > 0, "error message carried through"
    '--- unregister: the function is gone
    AssertTrue oCnn.RemoveUserDefinedFunction(oFuncs), "RemoveUserDefinedFunction succeeds"
    On Error Resume Next
    oCnn.GetRs "SELECT REV('x')"
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "removed function no longer resolves"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_AggregateUdf()
    Dim oCnn            As cConnection
    Dim oAgg            As cTestAgg
    Dim oRs             As cRecordset

    If Not TestBegin("cUDFMethods.AggregateUdf") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(v REAL)"
    oCnn.Execute "INSERT INTO t VALUES(1), (2), (3)"
    Set oAgg = New cTestAgg
    AssertTrue oCnn.AddUserDefinedAggregateFunction(oAgg), "AddUserDefinedAggregateFunction succeeds"
    Set oRs = oCnn.GetRs("SELECT PROD(v) AS p FROM t")
    AssertTrue oRs.Fields("p").Value = 6, "PROD aggregates 1*2*3"
    Set oRs = oCnn.GetRs("SELECT PROD(v) AS p FROM t WHERE v > 1")
    AssertTrue oRs.Fields("p").Value = 6, "PROD over a filtered set (2*3)"
    AssertTrue oCnn.RemoveUserDefinedAggregateFunction(oAgg), "RemoveUserDefinedAggregateFunction succeeds"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_Collation()
    Dim oCnn            As cConnection
    Dim oColl           As cTestColl
    Dim oRs             As cRecordset
    Dim bRaised         As Boolean

    If Not TestBegin("cUDFMethods.Collation") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(k TEXT)"
    oCnn.Execute "INSERT INTO t VALUES('ab'), ('ba'), ('aa')"
    Set oColl = New cTestColl
    AssertTrue oCnn.AddUserDefinedCollation(oColl), "AddUserDefinedCollation succeeds"
    '--- REVCOL compares reversed strings: aa->aa, ba->ab, ab->ba
    Set oRs = oCnn.OpenRecordset("SELECT k FROM t ORDER BY k COLLATE REVCOL")
    AssertEqStr CStr(oRs.ValueMatrix(0, 0)), "aa", "first by reversed compare"
    AssertEqStr CStr(oRs.ValueMatrix(1, 0)), "ba", "second by reversed compare"
    AssertEqStr CStr(oRs.ValueMatrix(2, 0)), "ab", "third by reversed compare"
    AssertTrue oCnn.RemoveUserDefinedCollation(oColl), "RemoveUserDefinedCollation succeeds"
    On Error Resume Next
    oCnn.OpenRecordset "SELECT k FROM t ORDER BY k COLLATE REVCOL"
    bRaised = (Err.Number <> 0)
    On Error GoTo EH
    AssertTrue bRaised, "removed collation no longer resolves"
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
