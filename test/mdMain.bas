Attribute VB_Name = "mdMain"
Option Explicit

Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Dst As LongPtr, ByVal Src As LongPtr, ByVal L As Long)

Private Function pvPtrToStrA(ByVal p As LongPtr) As String
    Dim n As Long
    Dim b() As Byte

    If p = 0 Then
        Exit Function
    End If
    n = lstrlenA(p)
    If n <= 0 Then
        Exit Function
    End If
    ReDim b(0 To n - 1)
    RtlMoveMemory VarPtr(b(0)), p, n
    pvPtrToStrA = StrConv(b, vbUnicode)
End Function

Private Function pvAscZ(ByVal s As String) As Byte()
    pvAscZ = StrConv(s & vbNullChar, vbFromUnicode)
End Function

Private Sub pvWriteResult(ByVal sText As String)
    Dim iFile As Integer

    iFile = FreeFile
    Open App.Path & "\apitest_out.txt" For Output As #iFile
    Print #iFile, sText
    Close #iFile
End Sub

Private Sub Main()
    Dim hDb As LongPtr
    Dim hStmt As LongPtr
    Dim rc As Long
    Dim v As Long
    Dim fn() As Byte
    Dim sql() As Byte
    Dim sOut As String

    sOut = "libversion=" & pvPtrToStrA(vbsqlite3_libversion())
    fn = pvAscZ(":memory:")
    rc = vbsqlite3_open(VarPtr(fn(0)), VarPtr(hDb))
    sOut = sOut & vbCrLf & "open rc=" & rc & " hDbNonZero=" & CBool(hDb <> 0)
    sql = pvAscZ("SELECT 40+2")
    rc = vbsqlite3_prepare_v2(hDb, VarPtr(sql(0)), -1, VarPtr(hStmt), 0)
    sOut = sOut & vbCrLf & "prepare rc=" & rc
    rc = vbsqlite3_step(hStmt)
    v = vbsqlite3_column_int(hStmt, 0)
    sOut = sOut & vbCrLf & "step rc=" & rc & " (100=SQLITE_ROW) col0=" & v
    rc = vbsqlite3_finalize(hStmt)
    rc = vbsqlite3_close(hDb)
    sOut = sOut & vbCrLf & "close rc=" & rc
    pvWriteResult sOut
End Sub