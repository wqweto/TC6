Attribute VB_Name = "mdErrorTests"
'=========================================================================
' mdErrorTests - every user-facing Err.Raise is triggered identically on
' TC6 and the real RC6.dll and the error number + description compared
'=========================================================================
Option Explicit

Public Sub RunErrorTests()
    Test_ErrorRC6Compat
End Sub

Private Function pvTriggerError(oCnn As Object, ByVal lCase As Long) As String
    Dim oRs             As Object
    Dim oCmd            As Object
    Dim vDummy          As Variant

    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5), (2, 'beta', 2.5)"
    On Error Resume Next
    Select Case lCase
    Case 1      '--- OpenRecordset on invalid SQL
        Set oRs = oCnn.OpenRecordset("SELECT bogus syntax here")
    Case 2      '--- Execute on invalid SQL
        oCnn.Execute "CREATE BOGUS"
    Case 3      '--- ExecCmd on a missing table
        oCnn.ExecCmd "INSERT INTO nosuch VALUES(1)"
    Case 4      '--- AbsolutePosition out of range
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.AbsolutePosition = 0
    Case 5      '--- MoveNext at EOF
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.MoveNext
        oRs.MoveNext
        oRs.MoveNext
    Case 6      '--- MovePrevious at BOF
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.MovePrevious
        oRs.MovePrevious
    Case 7      '--- Delete at EOF
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.MoveNext
        oRs.MoveNext
        oRs.MoveNext
        Err.Clear
        oRs.Delete
    Case 8      '--- Field Value at EOF
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.MoveNext
        oRs.MoveNext
        oRs.MoveNext
        Err.Clear
        vDummy = oRs.Fields("id").Value
    Case 9      '--- Value Let on a non-updatable recordset
        Set oRs = oCnn.OpenRecordset("SELECT id + 1 AS e FROM t")
        oRs.Fields("e").Value = 5
    Case 10     '--- Value Let on an expression column of an updatable rs
        Set oRs = oCnn.OpenRecordset("SELECT id, name, id + 1 AS e FROM t ORDER BY id")
        oRs.Fields("e").Value = 5
    Case 11     '--- Sort on an unknown field
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.Sort = "nosuchfield"
    Case 12     '--- Sort with an unclosed bracket
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        oRs.Sort = "[unclosed"
    Case 13     '--- FindFirst with an unparsable criterion
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        vDummy = oRs.FindFirst("%%% garbage %%%")
    Case 14     '--- unknown field name
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        vDummy = oRs.Fields("nope").Value
    Case 15     '--- field index out of range
        Set oRs = oCnn.OpenRecordset("SELECT * FROM t ORDER BY id")
        vDummy = oRs.Fields(99).Value
    Case 16     '--- unknown table in the schema objects
        vDummy = oCnn.DataBases(0).Tables("nope").Name
    Case 17     '--- constraint violation on cCommand.Execute
        Set oCmd = oCnn.CreateCommand("INSERT INTO t(id) VALUES(?)")
        oCmd.SetInt32 1, 1
        oCmd.Execute
    Case 18     '--- constraint violation on UpdateBatch
        Set oRs = oCnn.OpenRecordset("SELECT id, name FROM t ORDER BY id")
        oRs.Fields("id").Value = 2
        oRs.UpdateBatch
    Case 19     '--- CreateCommand on invalid SQL
        Set oCmd = oCnn.CreateCommand("INSERT INTO (")
    Case 20     '--- CreateSelectCommand rebind to invalid SQL
        Set oCmd = oCnn.CreateSelectCommand("SELECT * FROM t WHERE id = ?")
        oCmd.SQL = "SELECT busted ("
    Case 21     '--- unknown column lookup
        vDummy = oCnn.DataBases(0).Tables("t").Columns("nope").Name
    Case 22     '--- unknown index lookup
        vDummy = oCnn.DataBases(0).Tables("t").Indexes("nope").Name
    Case 23     '--- unknown view lookup
        vDummy = oCnn.DataBases(0).Views("nope").Name
    Case 24     '--- unknown trigger lookup
        vDummy = oCnn.DataBases(0).Tables("t").Triggers("nope").Name
    Case 25     '--- unknown database lookup
        vDummy = oCnn.DataBases("nope").Name
    Case 26     '--- SQL Let with valid parameterless text on a select command
        Set oCmd = oCnn.CreateSelectCommand("SELECT * FROM t WHERE id = ?")
        oCmd.SQL = "SELECT * FROM t"
    Case 27     '--- cCommand rebind to invalid SQL
        Set oCmd = oCnn.CreateCommand("INSERT INTO t(name) VALUES(?)")
        oCmd.SQL = "bogus ("
    Case 28     '--- CreateSelectCommand on invalid SQL that has a parameter
        Set oCmd = oCnn.CreateSelectCommand("SELECT busted ( ?")
    Case 29     '--- cCommand rebind to valid parameterless SQL
        Set oCmd = oCnn.CreateCommand("INSERT INTO t(name) VALUES(?)")
        oCmd.SQL = "DELETE FROM t WHERE id = 1"
    Case 30     '--- cSelectCommand rebind to invalid SQL that has a parameter
        Set oCmd = oCnn.CreateSelectCommand("SELECT * FROM t WHERE id = ?")
        oCmd.SQL = "SELECT busted ( ?"
    Case 31     '--- constraint violation through Execute
        oCnn.Execute "INSERT INTO t(id) VALUES(1)"
    Case 32     '--- constraint violation through ExecCmd
        oCnn.ExecCmd "INSERT INTO t(id) VALUES(1)"
    Case 33     '--- CreateSelectCommand on valid parameterless SQL
        Set oCmd = oCnn.CreateSelectCommand("SELECT * FROM t")
    Case 34     '--- CreateSelectCommand on invalid parameterless SQL
        Set oCmd = oCnn.CreateSelectCommand("SELECT busted (")
    Case 35     '--- bind index out of range on cCommand
        Set oCmd = oCnn.CreateCommand("INSERT INTO t(name) VALUES(?)")
        oCmd.SetInt32 99, 1
    Case 36     '--- bind index out of range on cSelectCommand
        Set oCmd = oCnn.CreateSelectCommand("SELECT * FROM t WHERE id = ?")
        oCmd.SetInt32 99, 1
    Case 37     '--- runtime error while materialising a select
        Set oRs = oCnn.OpenRecordset("SELECT abs(-9223372036854775808)")
    End Select
    '--- RC6 never populates Err.Source (VB stamps the project name), so only
    '--- the number and description take part in the comparison
    pvTriggerError = Err.Number & "|" & Err.Description
End Function

Private Sub Test_ErrorRC6Compat()
    Const CASE_COUNT    As Long = 37
    Dim oRc6Probe       As Object
    Dim oCnn            As Object
    Dim lIdx            As Long
    Dim sTc6            As String
    Dim sRc6            As String

    If Not TestBegin("cConnection.ErrorRC6Compat") Then Exit Sub
    On Error Resume Next
    Set oRc6Probe = CreateObject("RC6.cConnection")
    If oRc6Probe Is Nothing Then
        TestSkipCurrent "RC6.dll not registered"
        Exit Sub
    End If
    On Error GoTo EH
    For lIdx = 1 To CASE_COUNT
        Set oCnn = New cConnection
        sTc6 = pvTriggerError(oCnn, lIdx)
        Set oCnn = CreateObject("RC6.cConnection")
        sRc6 = pvTriggerError(oCnn, lIdx)
        AssertEqStr sTc6, sRc6, "case " & lIdx & ": error number|description matches RC6"
    Next
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
