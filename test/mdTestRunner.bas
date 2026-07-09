Attribute VB_Name = "mdTestRunner"
'=========================================================================
' mdTestRunner - minimal test harness for the RC6 SQLite replacement
'
' One standard module per class holds that class's tests (e.g.
' mdConnectionTests.bas) and exposes a single Public Sub RunXxxTests that
' Main calls. Each Test_ sub traps its own errors (On Error GoTo EH) and
' brackets its body with TestBegin/TestEnd, so one failing test never
' aborts the rest. Assertions (AssertTrue/AssertEqLng/AssertEqStr) record
' pass/fail and keep going. Results go to test\testrun.log (headless).
'=========================================================================
Option Explicit

Private Const LOG_FILE                As String = "testrun.log"

Private m_sLog                        As String
Private m_sCurrentTest                As String
Private m_lChecks                     As Long
Private m_lChecksFailed               As Long
Private m_lTestsRun                   As Long
Private m_lTestsFailed                As Long
Private m_bCurrentFailed              As Boolean

Public Sub Main()
    pvLogLine "=== RC6 SQLite replacement - test run " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ==="
    '--- register per-class suites here
    RunConnectionTests
    pvLogLine ""
    pvLogLine "-------------------------------------------------------"
    pvLogLine "Tests run: " & m_lTestsRun & "   failed: " & m_lTestsFailed
    pvLogLine "Checks:    " & m_lChecks & "   failed: " & m_lChecksFailed
    pvLogLine "RESULT: " & IIf(m_lTestsFailed = 0, "PASS", "FAIL")
    pvWriteLog
End Sub

Public Sub TestBegin(sName As String)
    m_sCurrentTest = sName
    m_bCurrentFailed = False
    m_lTestsRun = m_lTestsRun + 1
    pvLogLine ""
    pvLogLine "--- RUN  " & sName
End Sub

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

Private Sub pvLogLine(sText As String)
    If Len(m_sLog) > 0 Then
        m_sLog = m_sLog & vbCrLf
    End If
    m_sLog = m_sLog & sText
    Debug.Print sText
End Sub

Private Sub pvWriteLog()
    Dim iFile           As Integer

    iFile = FreeFile
    Open App.Path & "\" & LOG_FILE For Output As #iFile
    Print #iFile, m_sLog
    Close #iFile
End Sub
