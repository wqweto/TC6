VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cTestHost - public COM entry point for the test suite
'
' Createable from PowerShell / VBScript (32-bit host, the DLL is x86):
'   Set o = CreateObject("TC6SQLiteTest.cTestHost")
'   rc = o.RunAll("C:\...\testrun.log")   ' rc = number of failed tests
'   WScript.Echo o.Report
' RunTests(Filter, OutputFile) runs only tests whose name contains any of
' the comma-separated Filter tokens (case-insensitive); empty = all.
'=========================================================================
Option Explicit

Public Function RunAll(Optional ByVal OutputFile As String) As Long
    RunAll = RunTests(vbNullString, OutputFile)
End Function

Public Function RunTests(Optional ByVal Filter As String, Optional ByVal OutputFile As String) As Long
    TestReset Filter
    '--- register per-class suites here
    RunConnectionTests
    RunRecordsetTests
    RunCommandTests
    RunSchemaTests
    TestFinish
    If Len(OutputFile) = 0 Then
        OutputFile = App.Path & "\testrun.log"
    End If
    TestWriteReport OutputFile
    RunTests = TestFailedCount()
End Function

Public Property Get Report() As String
    Report = TestReport()
End Property

Public Property Get TestsRun() As Long
    TestsRun = TestRunCount()
End Property

Public Property Get TestsFailed() As Long
    TestsFailed = TestFailedCount()
End Property
