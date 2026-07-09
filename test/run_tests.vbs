' Drives the RC6 SQLite replacement test suite via the cTestHost COM class.
' Run with the 32-bit script host (the DLL is x86):
'   C:\Windows\SysWOW64\cscript.exe //nologo run_tests.vbs [/out:log] [/filter:tokens]
' /out    - report file path (default <dll folder>\testrun.log)
' /filter - comma-separated, case-insensitive substrings; only matching
'           test names run (e.g. /filter:Savepoints,DateHelpers)
' Exit code = number of failed tests (0 = all pass).
Option Explicit

Dim oHost, oArgs, sOut, sFilter, lRc

Set oArgs = WScript.Arguments
sOut = ""
sFilter = ""
If oArgs.Named.Exists("out") Then
    sOut = oArgs.Named("out")
End If
If oArgs.Named.Exists("filter") Then
    sFilter = oArgs.Named("filter")
End If

Set oHost = CreateObject("RC6SQLiteTest.cTestHost")
lRc = oHost.RunTests(sFilter, sOut)
WScript.Echo oHost.Report
WScript.Echo "failed tests: " & lRc
WScript.Quit lRc
