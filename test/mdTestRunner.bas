Attribute VB_Name = "mdTestRunner"
'=========================================================================
' mdTestRunner - test engine for the TC6 SQLite replacement
'
' Driven by the public cTestHost COM class (see cTestHost.cls). One
' standard module per class holds that class's tests (e.g.
' mdConnectionTests.bas) and exposes a Public Sub RunXxxTests that
' cTestHost calls. Each Test_ sub starts with
' `If Not TestBegin(name) Then Exit Sub` (honours the name filter) and
' traps its own errors, so one failing/skipped test never aborts the rest.
' Assertions (AssertTrue/AssertEqLng/AssertEqStr) record pass/fail and keep
' going. The full report is buffered in memory (TestReport) and optionally
' written to a file (TestWriteReport).
'=========================================================================
Option Explicit

Private m_sLog                        As String
Private m_sFilter                     As String
Private m_sCurrentTest                As String
Private m_lChecks                     As Long
Private m_lChecksFailed               As Long
Private m_lTestsRun                   As Long
Private m_lTestsFailed                As Long
Private m_lTestsSkipped               As Long
Private m_bCurrentFailed              As Boolean

Public Sub TestReset(sFilter As String)
    m_sLog = vbNullString
    m_sFilter = sFilter
    m_sCurrentTest = vbNullString
    m_lChecks = 0
    m_lChecksFailed = 0
    m_lTestsRun = 0
    m_lTestsFailed = 0
    m_lTestsSkipped = 0
    m_bCurrentFailed = False
    pvLogLine "=== TC6 SQLite replacement - test run " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    If Len(sFilter) > 0 Then
        pvLogLine "filter: " & sFilter
    End If
End Sub

Public Function TestBegin(sName As String) As Boolean
    If Not pvMatchesFilter(sName) Then
        m_lTestsSkipped = m_lTestsSkipped + 1
        Exit Function
    End If
    m_sCurrentTest = sName
    m_bCurrentFailed = False
    m_lTestsRun = m_lTestsRun + 1
    pvLogLine ""
    pvLogLine "--- RUN  " & sName
    TestBegin = True
End Function

Public Sub TestEnd()
    If m_bCurrentFailed Then
        m_lTestsFailed = m_lTestsFailed + 1
        pvLogLine "--- FAIL " & m_sCurrentTest
    Else
        pvLogLine "--- PASS " & m_sCurrentTest
    End If
End Sub

Public Sub TestErr()
    m_lChecksFailed = m_lChecksFailed + 1
    m_bCurrentFailed = True
    pvLogLine "    [ERROR] " & Err.Number & " - " & Err.Description
    Debug.Print "Critical error: " & Err.Description & " [" & m_sCurrentTest & "]"
    TestEnd
End Sub

Public Sub TestFinish()
    pvLogLine ""
    pvLogLine "-------------------------------------------------------"
    pvLogLine "Tests run: " & m_lTestsRun & "   failed: " & m_lTestsFailed & "   skipped: " & m_lTestsSkipped
    pvLogLine "Checks:    " & m_lChecks & "   failed: " & m_lChecksFailed
    pvLogLine "RESULT: " & IIf(m_lTestsFailed = 0, "PASS", "FAIL")
End Sub

Public Sub AssertTrue(ByVal bCond As Boolean, sMsg As String)
    m_lChecks = m_lChecks + 1
    If bCond Then
        pvLogLine "    [ok]   " & sMsg
    Else
        m_lChecksFailed = m_lChecksFailed + 1
        m_bCurrentFailed = True
        pvLogLine "    [FAIL] " & sMsg
    End If
End Sub

Public Sub AssertEqLng(ByVal lActual As Long, ByVal lExpected As Long, sMsg As String)
    AssertTrue lActual = lExpected, sMsg & " (expected=" & lExpected & " actual=" & lActual & ")"
End Sub

Public Sub AssertEqStr(sActual As String, sExpected As String, sMsg As String)
    AssertTrue sActual = sExpected, sMsg & " (expected=" & Chr$(34) & sExpected & Chr$(34) & " actual=" & Chr$(34) & sActual & Chr$(34) & ")"
End Sub

Public Function TestReport() As String
    TestReport = m_sLog
End Function

Public Function TestRunCount() As Long
    TestRunCount = m_lTestsRun
End Function

Public Function TestFailedCount() As Long
    TestFailedCount = m_lTestsFailed
End Function

Public Sub TestWriteReport(sFile As String)
    Dim iFile           As Integer

    If Len(sFile) = 0 Then
        Exit Sub
    End If
    iFile = FreeFile
    Open sFile For Output As #iFile
    Print #iFile, m_sLog
    Close #iFile
End Sub

Private Function pvMatchesFilter(sName As String) As Boolean
    Dim asTok()         As String
    Dim lIdx            As Long

    If Len(m_sFilter) = 0 Then
        pvMatchesFilter = True
        Exit Function
    End If
    asTok = Split(m_sFilter, ",")
    For lIdx = LBound(asTok) To UBound(asTok)
        If Len(Trim$(asTok(lIdx))) > 0 Then
            If InStr(1, sName, Trim$(asTok(lIdx)), vbTextCompare) > 0 Then
                pvMatchesFilter = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub pvLogLine(sText As String)
    If Len(m_sLog) > 0 Then
        m_sLog = m_sLog & vbCrLf
    End If
    m_sLog = m_sLog & sText
    Debug.Print sText
End Sub
