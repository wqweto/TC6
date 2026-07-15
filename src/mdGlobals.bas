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

Public Function FromUtf8PtrLen(ByVal lpUtf8 As LongPtr, ByVal lLen As Long) As String
    Dim lSize           As Long

    If lpUtf8 = 0 Or lLen <= 0 Then
        Exit Function
    End If
    FromUtf8PtrLen = String$(lLen, 0)
    lSize = MultiByteToWideChar(CP_UTF8, 0, ByVal lpUtf8, lLen, StrPtr(FromUtf8PtrLen), lLen)
    FromUtf8PtrLen = Left$(FromUtf8PtrLen, lSize)
End Function

Public Function BindVariant(ByVal hStmt As LongPtr, ByVal lIndex As Long, vValue As Variant) As Long
    Dim baBuf()         As Byte

    '--- bind a VB value to a 1-based statement parameter; text/blob use
    '--- SQLITE_TRANSIENT so SQLite copies the buffer before we free it
    Select Case VarType(vValue)
    Case vbNull, vbEmpty
        BindVariant = stub_sqlite3_bind_null(hStmt, lIndex)
    Case vbBoolean
        BindVariant = stub_sqlite3_bind_int(hStmt, lIndex, IIf(vValue, 1, 0))
    Case vbByte, vbInteger, vbLong
        BindVariant = stub_sqlite3_bind_int(hStmt, lIndex, CLng(vValue))
    Case vbCurrency, vbDecimal
        '--- int64 carrier types: a double round-trip would lose precision
        '--- above 2^53, so integral values go through bind_int64
        If vValue = Int(vValue) Then
            BindVariant = BindInt64Value(hStmt, lIndex, vValue)
        Else
            BindVariant = stub_sqlite3_bind_double(hStmt, lIndex, CDbl(vValue))
        End If
    Case VT_I8
        BindVariant = BindInt64Value(hStmt, lIndex, vValue)
    Case vbDate
        '--- dates are stored as ISO text (see cConnection.GetDateString)
        BindVariant = BindTextValue(hStmt, lIndex, Format$(vValue, "yyyy-mm-dd hh:nn:ss"))
    Case vbSingle, vbDouble
        BindVariant = stub_sqlite3_bind_double(hStmt, lIndex, CDbl(vValue))
    Case vbByte + vbArray
        baBuf = vValue
        BindVariant = BindBlobValue(hStmt, lIndex, baBuf)
    Case Else
        BindVariant = BindTextValue(hStmt, lIndex, CStr(vValue))
    End Select
End Function

Public Function BindInt64Value(ByVal hStmt As LongPtr, ByVal lIndex As Long, vValue As Variant) As Long
#If Win64 Then
    BindInt64Value = stub_sqlite3_bind_int64(hStmt, lIndex, CLngLng(vValue))
#Else
    '--- the x86 declare types the int64 as Currency (raw bits = value*10000),
    '--- so scale the integral value down by 10000 to place it into those bits
    BindInt64Value = stub_sqlite3_bind_int64(hStmt, lIndex, CCur(CDec(vValue) / 10000))
#End If
End Function

Public Function BindTextValue(ByVal hStmt As LongPtr, ByVal lIndex As Long, sText As String) As Long
    Dim baBuf()         As Byte
    Dim lLen            As Long

    baBuf = ToUtf8Array(sText)
    lLen = pvArrayByteLen(baBuf)
    If lLen = 0 Then
        ReDim baBuf(0 To 0) As Byte
    End If
    BindTextValue = stub_sqlite3_bind_text(hStmt, lIndex, VarPtr(baBuf(0)), lLen, SQLITE_TRANSIENT)
End Function

Public Function BindBlobValue(ByVal hStmt As LongPtr, ByVal lIndex As Long, baBuf() As Byte) As Long
    Dim lLen            As Long

    lLen = pvArrayByteLen(baBuf)
    If lLen > 0 Then
        BindBlobValue = stub_sqlite3_bind_blob(hStmt, lIndex, VarPtr(baBuf(0)), lLen, SQLITE_TRANSIENT)
    Else
        '--- an empty (non-NULL) blob needs zeroblob; bind_blob with a NULL
        '--- pointer would bind SQL NULL instead
        BindBlobValue = stub_sqlite3_bind_zeroblob(hStmt, lIndex, 0)
    End If
End Function

Public Function PrepareStatement(oCnn As cConnection, sSql As String) As LongPtr
    Dim baSql()         As Byte
    Dim hStmt           As LongPtr

    If oCnn Is Nothing Then
        Err.Raise vbObjectError, "PrepareStatement", "No active connection"
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    If stub_sqlite3_prepare_v2(oCnn.frDbHandle, VarPtr(baSql(0)), -1, VarPtr(hStmt), 0) <> SQLITE_OK Then
        Err.Raise vbObjectError, "PrepareStatement", oCnn.LastDBError()
    End If
    PrepareStatement = hStmt
End Function

Public Function ReadColumnValue(ByVal hStmt As LongPtr, ByVal lCol As Long) As Variant
    '--- qualify the constants: cField.FieldType has case-identical members
    Select Case stub_sqlite3_column_type(hStmt, lCol)
    Case sqlite3win32helper.SQLITE_INTEGER
        ReadColumnValue = pvColumnInteger(hStmt, lCol)
    Case sqlite3win32helper.SQLITE_FLOAT
        ReadColumnValue = stub_sqlite3_column_double(hStmt, lCol)
    Case sqlite3win32helper.SQLITE_TEXT
        ReadColumnValue = FromUtf8Ptr(stub_sqlite3_column_text(hStmt, lCol))
    Case sqlite3win32helper.SQLITE_BLOB
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
    StmtParamName = FromUtf8Ptr(stub_sqlite3_bind_parameter_name(hStmt, lIndex))
End Function

Public Function QuoteIdentifier(sName As String) As String
    QuoteIdentifier = """" & Replace(sName, """", """""") & """"
End Function

Public Function QuoteString(sText As String) As String
    QuoteString = "'" & Replace(sText, "'", "''") & "'"
End Function

Public Sub SaveCommandSql(oCnn As cConnection, sTable As String, sKey As String, sSql As String)
    Dim baSql()         As Byte

    '--- RC6 saved-commands table (exact DDL); the SQL is a UTF-16 blob
    If Not pvSavedTableExists(oCnn, sTable) Then
        oCnn.Execute "CREATE TABLE " & sTable & "(ID Integer Primary Key,CommandKey Text Collate NoCase Unique On Conflict Replace, SQL Blob)"
    End If
    baSql = sSql
    oCnn.ExecCmd "INSERT INTO " & sTable & "(CommandKey, SQL) VALUES(?, ?)", sKey, baSql
End Sub

Public Function LookupSavedSql(oCnn As cConnection, sTable As String, sKey As String) As String
    Dim oRs             As cRecordset
    Dim baSql()         As Byte

    If pvSavedTableExists(oCnn, sTable) Then
        Set oRs = oCnn.GetRs("SELECT SQL FROM " & sTable & " WHERE CommandKey = ?", sKey)
        If oRs.RecordCount > 0 Then
            baSql = oRs.Fields(0).Value
            LookupSavedSql = baSql
        End If
    End If
End Function

Private Function pvSavedTableExists(oCnn As cConnection, sTable As String) As Boolean
    Dim oRs             As cRecordset

    Set oRs = oCnn.GetRs("SELECT COUNT(*) FROM sqlite_master WHERE type = 'table' AND name = ?", sTable)
    pvSavedTableExists = (oRs.Fields(0).Value > 0)
End Function

Public Function AdoTypeToDeclType(ByVal lAdoType As Long, Optional ByVal lCharMaxLen As Long) As String
    Const adSmallInt                        As Long = 2
    Const adInteger                         As Long = 3
    Const adSingle                          As Long = 4
    Const adDouble                          As Long = 5
    Const adCurrency                        As Long = 6
    Const adDate                            As Long = 7
    Const adBoolean                         As Long = 11
    Const adDecimal                         As Long = 14
    Const adTinyInt                         As Long = 16
    Const adUnsignedTinyInt                 As Long = 17
    Const adUnsignedSmallInt                As Long = 18
    Const adUnsignedInt                     As Long = 19
    Const adBigInt                          As Long = 20
    Const adUnsignedBigInt                  As Long = 21
    Const adBinary                          As Long = 128
    Const adChar                            As Long = 129
    Const adWChar                           As Long = 130
    Const adNumeric                         As Long = 131
    Const adDBDate                          As Long = 133
    Const adDBTime                          As Long = 134
    Const adDBTimeStamp                     As Long = 135
    Const adVarNumeric                      As Long = 139
    Const adVarChar                         As Long = 200
    Const adVarWChar                        As Long = 202
    Const adVarBinary                       As Long = 204
    Const adLongVarBinary                   As Long = 205

    '--- RC6 mapping (verified against RC6.dll 6.0.15 via the converter
    '--- tests); column sizes appear only on the converter path, which
    '--- passes CHARACTER_MAXIMUM_LENGTH from the columns schema rowset
    Select Case lAdoType
    Case adSmallInt, adInteger, adTinyInt, adUnsignedTinyInt, adUnsignedSmallInt, adUnsignedInt, adBigInt, adUnsignedBigInt
        AdoTypeToDeclType = "INTEGER"
    Case adBoolean
        AdoTypeToDeclType = "BIT"
    Case adSingle, adDouble, adCurrency, adDecimal, adNumeric, adVarNumeric
        AdoTypeToDeclType = "REAL"
    Case adDate, adDBDate, adDBTime, adDBTimeStamp
        AdoTypeToDeclType = "DATE"
    Case adBinary, adVarBinary, adLongVarBinary
        AdoTypeToDeclType = "BLOB"
    Case adChar, adWChar, adVarChar, adVarWChar
        If lCharMaxLen > 0 Then
            AdoTypeToDeclType = "TEXT(" & lCharMaxLen & ")"
        Else
            AdoTypeToDeclType = "TEXT"
        End If
    Case Else
        AdoTypeToDeclType = "TEXT"
    End Select
End Function

Public Sub CopyAdoRsToTable(oCnn As cConnection, sTable As String, oAdoRs As Object, vDataTypes As Variant, ByVal bNoCase As Boolean, cPkNames As VBA.Collection, Optional oProgress As cConverter)
    Const adFldKeyColumn                    As Long = &H8000&
    Dim oFld            As Object
    Dim lCol            As Long
    Dim lCount          As Long
    Dim sDefs           As String
    Dim sVals           As String
    Dim sType           As String
    Dim sPk             As String
    Dim bIsPk           As Boolean
    Dim hStmt           As LongPtr
    Dim vValue          As Variant
    Dim lRow            As Long
    Dim lTotal          As Long
    Dim vName           As Variant

    lCount = oAdoRs.Fields.Count
    '--- pk columns: explicit list, else the adFldKeyColumn field attribute
    If cPkNames Is Nothing Then
        Set cPkNames = New VBA.Collection
        For lCol = 0 To lCount - 1
            If (oAdoRs.Fields(lCol).Attributes And adFldKeyColumn) <> 0 Then
                cPkNames.Add oAdoRs.Fields(lCol).Name
            End If
        Next
    End If
    For lCol = 0 To lCount - 1
        Set oFld = oAdoRs.Fields(lCol)
        If Len(sDefs) > 0 Then
            sDefs = sDefs & ", "
            sVals = sVals & ", "
        End If
        sType = vbNullString
        If IsArray(vDataTypes) Then
            sType = vDataTypes(lCol)
        End If
        If Len(sType) = 0 Then
            sType = AdoTypeToDeclType(oFld.Type)
        End If
        bIsPk = False
        For Each vName In cPkNames
            If StrComp(CStr(vName), oFld.Name, vbTextCompare) = 0 Then
                bIsPk = True
            End If
        Next
        '--- NOT NULL comes only through vDataTypes (the converter appends it
        '--- from IS_NULLABLE; plain CreateTableFromADORs never emits it)
        sDefs = sDefs & "[" & oFld.Name & "] " & sType
        If bIsPk And cPkNames.Count = 1 Then
            sDefs = sDefs & " PRIMARY KEY"
        End If
        If Left$(sType, 4) = "TEXT" And bNoCase Then
            sDefs = sDefs & " COLLATE NOCASE"
        End If
        sVals = sVals & "?"
    Next
    If cPkNames.Count > 1 Then
        For Each vName In cPkNames
            If Len(sPk) > 0 Then
                sPk = sPk & ", "
            End If
            sPk = sPk & "[" & vName & "]"
        Next
        sDefs = sDefs & ", PRIMARY KEY (" & sPk & ")"
    End If
    '--- RC6 quirk: a leading space after the paren when the table has no pk
    oCnn.Execute "CREATE TABLE [" & sTable & "] (" & IIf(cPkNames.Count = 0, " ", vbNullString) & sDefs & ")"
    '--- copy the rows with a single prepared INSERT
    If oAdoRs.EOF And oAdoRs.BOF Then
        Exit Sub
    End If
    On Error Resume Next
    lTotal = oAdoRs.RecordCount
    On Error GoTo 0
    hStmt = PrepareStatement(oCnn, "INSERT INTO [" & sTable & "] VALUES (" & sVals & ")")
    On Error GoTo EH
    oAdoRs.MoveFirst
    Do While Not oAdoRs.EOF
        For lCol = 0 To lCount - 1
            vValue = oAdoRs.Fields(lCol).Value
            If VarType(vValue) = vbDate Then
                vValue = Format$(vValue, "yyyy-mm-dd hh:nn:ss")
            End If
            Call BindVariant(hStmt, lCol + 1, vValue)
        Next
        If stub_sqlite3_step(hStmt) <> SQLITE_DONE Then
            Err.Raise vbObjectError, "CopyAdoRsToTable", oCnn.LastDBError()
        End If
        Call stub_sqlite3_reset(hStmt)
        lRow = lRow + 1
        If Not oProgress Is Nothing Then
            oProgress.frInsertProgress sTable, lTotal, lRow
        End If
        oAdoRs.MoveNext
    Loop
    Call stub_sqlite3_finalize(hStmt)
    Exit Sub
EH:
    Call stub_sqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub CreateTableFromRecordset(oCnn As cConnection, ByVal oSrc As cRecordset, sTable As String, ByVal bTempTable As Boolean, ByVal bWithPrimaryKeys As Boolean)
    Dim oField          As cField
    Dim sDefs           As String
    Dim sCols           As String
    Dim sVals           As String
    Dim sType           As String
    Dim lRow            As Long
    Dim lCol            As Long
    Dim aParams()       As Variant

    '--- column definitions from the source recordset's field metadata
    For Each oField In oSrc.Fields
        If Len(sDefs) > 0 Then
            sDefs = sDefs & ", "
            sCols = sCols & ", "
            sVals = sVals & ", "
        End If
        sType = oField.OriginalDataType
        If Len(sType) = 0 Then
            sType = "TEXT"
        End If
        sDefs = sDefs & QuoteIdentifier(oField.Name) & " " & sType
        If bWithPrimaryKeys Then
            If oField.PrimaryKey Then
                sDefs = sDefs & " PRIMARY KEY"
            End If
        End If
        sCols = sCols & QuoteIdentifier(oField.Name)
        sVals = sVals & "?"
    Next
    If Len(sDefs) = 0 Then
        Err.Raise 5, "CreateTableFromRecordset", "Source recordset has no fields"
    End If
    oCnn.Execute "CREATE " & IIf(bTempTable, "TEMP ", vbNullString) & "TABLE " & QuoteIdentifier(sTable) & " (" & sDefs & ")"
    If oSrc.RecordCount > 0 Then
        ReDim aParams(0 To oSrc.Fields.Count - 1) As Variant
        For lRow = 0 To oSrc.RecordCount - 1
            For lCol = 0 To oSrc.Fields.Count - 1
                aParams(lCol) = oSrc.ValueMatrix(lRow, lCol)
            Next
            oCnn.frExecCmd "INSERT INTO " & QuoteIdentifier(sTable) & " (" & sCols & ") VALUES (" & sVals & ")", aParams
        Next
    End If
End Sub

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

    '--- RC6 format (verified against RC6.dll 6.0.15): local-time VB date
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
    pvParamIndex = stub_sqlite3_bind_parameter_index(hStmt, VarPtr(baName(0)))
End Function

Private Function pvColumnInteger(ByVal hStmt As LongPtr, ByVal lCol As Long) As Variant
    Dim vDec            As Variant

    '--- recover the true int64: x86 returns raw bits as Currency (value*10000)
#If Win64 Then
    vDec = CDec(stub_sqlite3_column_int64(hStmt, lCol))
#Else
    vDec = CDec(stub_sqlite3_column_int64(hStmt, lCol)) * CDec(10000)
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

    lLen = stub_sqlite3_column_bytes(hStmt, lCol)
    If lLen > 0 Then
        lPtr = stub_sqlite3_column_blob(hStmt, lCol)
        ReDim baBuf(0 To lLen - 1) As Byte
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
