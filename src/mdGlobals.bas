Attribute VB_Name = "mdGlobals"
'=========================================================================
' mdGlobals - shared helpers (UTF-8 <-> VB String marshaling)
'=========================================================================
Option Explicit

'--- for WideCharToMultiByte
Private Const CP_UTF8                       As Long = 65001
'--- int64 VARIANT type (no VB6 vbLongLong constant)
Private Const VT_I8                         As Integer = 20

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare Function vbaObjSetAddref Lib "msvbvm60" Alias "__vbaObjSetAddref" (oDest As Any, ByVal lSrcPtr As LongPtr) As Long
Private Declare Sub GetSystemTimePreciseAsFileTime Lib "kernel32" (lpSystemTimeAsFileTime As Currency)
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As Currency, lpLocalFileTime As Currency) As Long

'--- live-instance counter (leak/cycle diagnostic, see cRecordset)
Public g_lLiveRecordsets            As Long
'--- monotonicity guard for CreateUniqueID64
Private m_decLastUniqueId           As Variant

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

Public Function BindVariant(ByVal hStmt As LongPtr, ByVal lIndex As Long, vValue As Variant) As Long
    Dim baBuf()         As Byte

    '--- bind a VB value to a 1-based statement parameter; text/blob use
    '--- SQLITE_TRANSIENT so SQLite copies the buffer before we free it
    Select Case VarType(vValue)
    Case vbNull, vbEmpty
        BindVariant = vbsqlite3_bind_null(hStmt, lIndex)
    Case vbBoolean
        BindVariant = vbsqlite3_bind_int(hStmt, lIndex, IIf(vValue, 1, 0))
    Case vbByte, vbInteger, vbLong
        BindVariant = vbsqlite3_bind_int(hStmt, lIndex, CLng(vValue))
    Case vbCurrency, vbDecimal
        '--- int64 carrier types: a double round-trip would lose precision
        '--- above 2^53, so integral values go through bind_int64
        If vValue = Int(vValue) Then
            BindVariant = BindInt64Value(hStmt, lIndex, vValue)
        Else
            BindVariant = vbsqlite3_bind_double(hStmt, lIndex, CDbl(vValue))
        End If
    Case VT_I8
        BindVariant = BindInt64Value(hStmt, lIndex, vValue)
    Case vbDate
        '--- dates are stored as ISO text (see cConnection.GetDateString)
        BindVariant = BindTextValue(hStmt, lIndex, Format$(vValue, "yyyy-mm-dd hh:nn:ss"))
    Case vbSingle, vbDouble
        BindVariant = vbsqlite3_bind_double(hStmt, lIndex, CDbl(vValue))
    Case vbByte + vbArray
        baBuf = vValue
        BindVariant = BindBlobValue(hStmt, lIndex, baBuf)
    Case Else
        BindVariant = BindTextValue(hStmt, lIndex, CStr(vValue))
    End Select
End Function

Public Function BindInt64Value(ByVal hStmt As LongPtr, ByVal lIndex As Long, vValue As Variant) As Long
#If Win64 Then
    BindInt64Value = vbsqlite3_bind_int64(hStmt, lIndex, CLngLng(vValue))
#Else
    '--- the x86 declare types the int64 as Currency (raw bits = value*10000),
    '--- so scale the integral value down by 10000 to place it into those bits
    BindInt64Value = vbsqlite3_bind_int64(hStmt, lIndex, CCur(CDec(vValue) / 10000))
#End If
End Function

Public Function BindTextValue(ByVal hStmt As LongPtr, ByVal lIndex As Long, sText As String) As Long
    Dim baBuf()         As Byte
    Dim lLen            As Long

    baBuf = ToUtf8Array(sText)
    lLen = pvArrayByteLen(baBuf)
    If lLen = 0 Then
        ReDim baBuf(0 To 0)
    End If
    BindTextValue = vbsqlite3_bind_text(hStmt, lIndex, VarPtr(baBuf(0)), lLen, SQLITE_TRANSIENT)
End Function

Public Function BindBlobValue(ByVal hStmt As LongPtr, ByVal lIndex As Long, baBuf() As Byte) As Long
    Dim lLen            As Long

    lLen = pvArrayByteLen(baBuf)
    If lLen > 0 Then
        BindBlobValue = vbsqlite3_bind_blob(hStmt, lIndex, VarPtr(baBuf(0)), lLen, SQLITE_TRANSIENT)
    Else
        '--- an empty (non-NULL) blob needs zeroblob; bind_blob with a NULL
        '--- pointer would bind SQL NULL instead
        BindBlobValue = vbsqlite3_bind_zeroblob(hStmt, lIndex, 0)
    End If
End Function

Public Function PrepareStatement(oCnn As cConnection, sSql As String) As LongPtr
    Dim baSql()         As Byte
    Dim hStmt           As LongPtr

    If oCnn Is Nothing Then
        Err.Raise vbObjectError, "PrepareStatement", "No active connection"
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    If vbsqlite3_prepare_v2(oCnn.frDbHandle, VarPtr(baSql(0)), -1, VarPtr(hStmt), 0) <> SQLITE_OK Then
        Err.Raise vbObjectError, "PrepareStatement", oCnn.LastDBError()
    End If
    PrepareStatement = hStmt
End Function

Public Function ReadColumnValue(ByVal hStmt As LongPtr, ByVal lCol As Long) As Variant
    '--- qualify the constants: cField.FieldType has case-identical members
    Select Case vbsqlite3_column_type(hStmt, lCol)
    Case mdSqliteApi.SQLITE_INTEGER
        ReadColumnValue = pvColumnInteger(hStmt, lCol)
    Case mdSqliteApi.SQLITE_FLOAT
        ReadColumnValue = vbsqlite3_column_double(hStmt, lCol)
    Case mdSqliteApi.SQLITE_TEXT
        ReadColumnValue = FromUtf8Ptr(vbsqlite3_column_text(hStmt, lCol))
    Case mdSqliteApi.SQLITE_BLOB
        ReadColumnValue = pvColumnBlob(hStmt, lCol)
    Case Else
        ReadColumnValue = Null
    End Select
End Function

Public Function StmtParamIndex(ByVal hStmt As LongPtr, sName As String) As Long
    '--- accept the parameter name with or without its :/@/$ prefix
    StmtParamIndex = pvParamIndex(hStmt, sName)
    If StmtParamIndex = 0 Then
        StmtParamIndex = pvParamIndex(hStmt, ":" & sName)
    End If
    If StmtParamIndex = 0 Then
        StmtParamIndex = pvParamIndex(hStmt, "@" & sName)
    End If
    If StmtParamIndex = 0 Then
        StmtParamIndex = pvParamIndex(hStmt, "$" & sName)
    End If
End Function

Public Function StmtParamName(ByVal hStmt As LongPtr, ByVal lIndex As Long) As String
    StmtParamName = FromUtf8Ptr(vbsqlite3_bind_parameter_name(hStmt, lIndex))
End Function

Public Function QuoteIdentifier(sName As String) As String
    QuoteIdentifier = """" & Replace(sName, """", """""") & """"
End Function

Public Function QuoteString(sText As String) As String
    QuoteString = "'" & Replace(sText, "'", "''") & "'"
End Function

Public Function Int64Variant(vValue As Variant) As Variant
    Dim vRet            As Variant
    Dim cyValue         As Currency
    Dim nVt             As Integer

    '--- build a true VT_I8 variant (matching RC6's int64 results): place
    '--- the raw int64 bits via the Currency carrier (value scaled by 10000)
    cyValue = CCur(CDec(vValue) / 10000)
    nVt = VT_I8
    Call CopyMemory(vRet, nVt, 2)
    Call CopyMemory(ByVal VarPtr(vRet) + 8, cyValue, 8)
    Int64Variant = vRet
End Function

Public Function CreateUniqueID64() As Variant
    Dim cyFileTime      As Currency
    Dim cyLocal         As Currency
    Dim decId           As Variant

    '--- RC6 format (verified against RC6.dll 3.42.0): local-time VB date
    '--- serial * 10^14, sub-ms precision from the system clock. FILETIME
    '--- read as Currency = milliseconds since 1601-01-01; 109205 days lie
    '--- between that epoch and the VB serial epoch (1899-12-30)
    Call GetSystemTimePreciseAsFileTime(cyFileTime)
    Call FileTimeToLocalFileTime(cyFileTime, cyLocal)
    decId = Int((CDec(cyLocal) / 86400000 - 109205) * CDec("100000000000000"))
    '--- strictly increasing even within clock granularity
    If Not IsEmpty(m_decLastUniqueId) Then
        If decId <= m_decLastUniqueId Then
            decId = m_decLastUniqueId + 1
        End If
    End If
    m_decLastUniqueId = decId
    CreateUniqueID64 = Int64Variant(decId)
End Function

Private Function pvParamIndex(ByVal hStmt As LongPtr, sName As String) As Long
    Dim baName()        As Byte

    baName = ToUtf8Array(sName & vbNullChar)
    pvParamIndex = vbsqlite3_bind_parameter_index(hStmt, VarPtr(baName(0)))
End Function

Private Function pvColumnInteger(ByVal hStmt As LongPtr, ByVal lCol As Long) As Variant
    Dim vDec            As Variant

    '--- recover the true int64: x86 returns raw bits as Currency (value*10000)
#If Win64 Then
    vDec = CDec(vbsqlite3_column_int64(hStmt, lCol))
#Else
    vDec = CDec(vbsqlite3_column_int64(hStmt, lCol)) * CDec(10000)
#End If
    If vDec >= -2147483648# And vDec <= 2147483647# Then
        pvColumnInteger = CLng(vDec)
    Else
        pvColumnInteger = vDec
    End If
End Function

Private Function pvColumnBlob(ByVal hStmt As LongPtr, ByVal lCol As Long) As Variant
    Dim lLen            As Long
    Dim lPtr            As LongPtr
    Dim baBuf()         As Byte

    lLen = vbsqlite3_column_bytes(hStmt, lCol)
    If lLen > 0 Then
        lPtr = vbsqlite3_column_blob(hStmt, lCol)
        ReDim baBuf(0 To lLen - 1)
        Call CopyMemory(baBuf(0), ByVal lPtr, lLen)
        pvColumnBlob = baBuf
    Else
        pvColumnBlob = baBuf
    End If
End Function

Private Function pvArrayByteLen(baBuf() As Byte) As Long
    On Error GoTo QH
    pvArrayByteLen = UBound(baBuf) - LBound(baBuf) + 1
QH:
End Function
