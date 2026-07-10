Attribute VB_Name = "mdGlobals"
'=========================================================================
' mdGlobals - shared helpers (UTF-8 <-> VB String marshaling)
'=========================================================================
Option Explicit

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As LongPtr) As Long

'--- live-instance counter (leak/cycle diagnostic, see cRecordset)
Public g_lLiveRecordsets            As Long

Public Function ToUtf8Array(sText As String) As Byte()
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function

Public Function FromUtf8Array(baText() As Byte) As String
    Dim lSize           As Long

    If UBound(baText) >= 0 Then
        FromUtf8Array = String$(2 * (UBound(baText) + 1), 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, baText(0), UBound(baText) + 1, StrPtr(FromUtf8Array), Len(FromUtf8Array))
        FromUtf8Array = Left$(FromUtf8Array, lSize)
    End If
End Function

Public Function FromUtf8Ptr(ByVal lpUtf8 As LongPtr) As String
    Dim lLen            As Long
    Dim lSize           As Long

    If lpUtf8 = 0 Then
        Exit Function
    End If
    lLen = lstrlenA(lpUtf8)
    If lLen > 0 Then
        FromUtf8Ptr = String$(lLen, 0)
        lSize = MultiByteToWideChar(CP_UTF8, 0, ByVal lpUtf8, lLen, StrPtr(FromUtf8Ptr), lLen)
        FromUtf8Ptr = Left$(FromUtf8Ptr, lSize)
    End If
End Function

Public Function ObjectFromPtr(ByVal pObj As LongPtr) As IUnknown
    Call vbaObjSetAddref(ObjectFromPtr, pObj)
End Function
