' Drives the TC6 SQLite replacement test suite via the cTestHost COM class.
' Run with the 32-bit script host (the DLL is x86):
'   C:\Windows\SysWOW64\cscript.exe //nologo run_tests.vbs [/out:log] [/filter:tokens] [/bin]
' /out    - report file path (default <dll folder>\testrun.log)
' /filter - comma-separated, case-insensitive substrings; only matching
'           test names run (e.g. /filter:Savepoints,DateHelpers)
' /bin    - use TestRunnerBin.dll, which runs the same suite against the
'           registered compiled TC6SQLite.dll instead of the sources
' Exit code = number of failed tests (0 = all pass).
Option Explicit

Dim oHost, oArgs, sOut, sFilter, sProgId, lRc

Set oArgs = WScript.Arguments
sOut = ""
sFilter = ""
sProgId = "TC6SQLiteTest.cTestHost"
If oArgs.Named.Exists("out") Then
    sOut = oArgs.Named("out")
End If
If oArgs.Named.Exists("filter") Then
    sFilter = oArgs.Named("filter")
End If
If oArgs.Named.Exists("bin") Then
    sProgId = "TC6SQLiteTestBin.cTestHost"
End If

Set oHost = CreateObject(sProgId)
lRc = oHost.RunTests(sFilter, sOut)
WScript.Echo oHost.Report
WScript.Echo "failed tests: " & lRc
WScript.Quit lRc
