Attribute VB_Name = "mdBinShims"
'=========================================================================
' mdBinShims - minimal stand-ins for mdGlobals helpers used by the test
' modules, so TestRunnerBin.vbp can compile against the registered
' TC6SQLite.dll typelib instead of the sources
'=========================================================================
Option Explicit

#If Win64 = 0 And TWINBASIC = 0 Then
Public Enum LongPtr
    [_]
End Enum
#End If

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long

Public Function FromUtf8Array(baText() As Byte) As String
    Const CP_UTF8       As Long = 65001
    Dim lSize           As Long

    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function
