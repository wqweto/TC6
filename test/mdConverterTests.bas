Attribute VB_Name = "mdConverterTests"
'=========================================================================
' mdConverterTests - ADO interop tests (GetADORsFromContent, DataSource,
' CreateTableFromADORs, cConverter) incl. RC6 cross-checks
'=========================================================================
Option Explicit

Public Sub RunConverterTests()
    Test_GetADORsFromContent
    Test_CreateTableFromADORs
    Test_ConverterJetMdb
End Sub

Private Function pvSchemaDump(oCnn As Object) As String
    Dim oRs             As Object

    Set oRs = oCnn.OpenRecordset("SELECT type, name, sql FROM sqlite_master ORDER BY type, name")
    Do While Not oRs.EOF
        pvSchemaDump = pvSchemaDump & oRs.Fields(0).Value & "|" & oRs.Fields(1).Value & "|" & oRs.Fields(2).Value & vbLf
        If CLng(oRs.AbsolutePosition) = CLng(oRs.RecordCount) Then
            Exit Do
        End If
        oRs.MoveNext
    Loop
End Function

Private Function pvQuoteDump(sDbFile As String, sTable As String, sOrderBy As String) As String
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim sSql            As String
    Dim lCol            As Long

    '--- storage-level dump: TC6 opens the db file and quote()s every cell,
    '--- so the comparison is reader-independent
    Set oCnn = New cConnection
    oCnn.OpenDBReadOnly sDbFile
    Set oRs = oCnn.OpenRecordset("SELECT * FROM [" & sTable & "] WHERE 0")
    For lCol = 0 To oRs.Fields.Count - 1
        If Len(sSql) > 0 Then
            sSql = sSql & " || '|' || "
        End If
        sSql = sSql & "quote([" & oRs.Fields(lCol).Name & "])"
    Next
    Set oRs = oCnn.OpenRecordset("SELECT " & sSql & " FROM [" & sTable & "] ORDER BY " & sOrderBy)
    Do While Not oRs.EOF
        pvQuoteDump = pvQuoteDump & oRs.Fields(0).Value & vbLf
        If oRs.AbsolutePosition = oRs.RecordCount Then
            Exit Do
        End If
        oRs.MoveNext
    Loop
End Function

Private Function pvAdoRsDump(oArs As Object) As String
    Dim oFld            As Object
    Dim vValue          As Variant

    For Each oFld In oArs.Fields
        pvAdoRsDump = pvAdoRsDump & oFld.Name & ":" & oFld.Type & ":" & oFld.DefinedSize & ":" & Hex$(oFld.Attributes) & "|"
    Next
    pvAdoRsDump = pvAdoRsDump & vbLf
    If CLng(oArs.RecordCount) = 0 Then
        Exit Function
    End If
    oArs.MoveFirst
    Do While Not oArs.EOF
        For Each oFld In oArs.Fields
            vValue = oFld.Value
            If IsNull(vValue) Then
                pvAdoRsDump = pvAdoRsDump & "<null>|"
            ElseIf IsArray(vValue) Then
                pvAdoRsDump = pvAdoRsDump & "blob:" & (UBound(vValue) - LBound(vValue) + 1) & "|"
            Else
                pvAdoRsDump = pvAdoRsDump & TypeName(vValue) & ":" & vValue & "|"
            End If
        Next
        pvAdoRsDump = pvAdoRsDump & vbLf
        oArs.MoveNext
    Loop
End Function

Private Sub Test_GetADORsFromContent()
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim oArs            As Object
    Dim oRc6Cnn         As Object
    Dim oRc6Rs          As Object
    Dim sTc6            As String

    If Not TestBegin("cRecordset.GetADORsFromContent") Then Exit Sub
    On Error GoTo EH
    Set oCnn = New cConnection
    oCnn.CreateNewDB ":memory:"
    oCnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB, big INTEGER)"
    oCnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5, X'01FF', 9007199254740993), (2, NULL, NULL, NULL, NULL)"
    Set oRs = oCnn.OpenRecordset("SELECT id, name, score, data, big FROM t ORDER BY id")
    Set oArs = oRs.GetADORsFromContent()
    AssertEqStr TypeName(oArs), "Recordset", "disconnected ADO recordset"
    AssertEqLng CLng(oArs.RecordCount), 2, "ADO RecordCount"
    AssertEqLng CLng(oArs.Fields("id").Type), 3, "INTEGER pk maps to adInteger"
    AssertEqLng CLng(oArs.Fields("big").Type), 14, "INTEGER maps to adDecimal"
    AssertEqLng CLng(oArs.Fields("name").Type), 203, "TEXT maps to adLongVarWChar"
    AssertEqLng CLng(oArs.Fields("score").Type), 5, "REAL maps to adDouble"
    AssertEqLng CLng(oArs.Fields("data").Type), 205, "BLOB maps to adLongVarBinary"
    oArs.MoveFirst
    AssertEqLng CLng(oArs.Fields("id").Value), 1, "row0 id"
    AssertTrue oArs.Fields("big").Value = CDec("9007199254740993"), "int64 kept full precision"
    oArs.MoveNext
    AssertTrue IsNull(oArs.Fields("name").Value), "NULL cell surfaces as ADO Null"
    Set oArs = oRs.DataSource()
    AssertEqStr TypeName(oArs), "Recordset", "DataSource wraps the ADO recordset"
    '--- field-for-field + cell-for-cell against RC6
    On Error Resume Next
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    On Error GoTo EH
    If Not oRc6Cnn Is Nothing Then
        oRc6Cnn.CreateNewDB ":memory:"
        oRc6Cnn.Execute "CREATE TABLE t(id INTEGER PRIMARY KEY, name TEXT, score REAL, data BLOB, big INTEGER)"
        oRc6Cnn.Execute "INSERT INTO t VALUES(1, 'alpha', 1.5, X'01FF', 9007199254740993), (2, NULL, NULL, NULL, NULL)"
        Set oRc6Rs = oRc6Cnn.OpenRecordset("SELECT id, name, score, data, big FROM t ORDER BY id")
        sTc6 = pvAdoRsDump(oRs.GetADORsFromContent())
        AssertEqStr sTc6, pvAdoRsDump(oRc6Rs.GetADORsFromContent()), "ADO recordset matches RC6 field-for-field"
    End If
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_CreateTableFromADORs()
    Const adInteger     As Long = 3
    Const adDouble      As Long = 5
    Const adCurrency    As Long = 6
    Const adDate        As Long = 7
    Const adBoolean     As Long = 11
    Const adVarWChar    As Long = 202
    Const adLongVarBinary As Long = 205
    Const NULLABLE_ATTRIBS As Long = &H60           '--- adFldMayBeNull Or adFldIsNullable
    Const LONG_ATTRIBS  As Long = &HE0              '--- + adFldLong
    Dim oArs            As Object
    Dim oCnn            As cConnection
    Dim oRc6Cnn         As Object
    Dim oRc6Cnn2        As Object
    Dim sTc6            As String
    Dim sRc6            As String
    Dim baBlob(0 To 1)  As Byte
    Dim sDbTc6          As String
    Dim sDbRc6          As String

    If Not TestBegin("cConnection.CreateTableFromADORs") Then Exit Sub
    On Error GoTo EH
    '--- fabricated disconnected ADO recordset with assorted types
    Set oArs = CreateObject("ADODB.Recordset")
    oArs.Fields.Append "id", adInteger
    oArs.Fields.Append "nm", adVarWChar, 50, NULLABLE_ATTRIBS
    oArs.Fields.Append "amt", adCurrency, 0, NULLABLE_ATTRIBS
    oArs.Fields.Append "bd", adDate, 0, NULLABLE_ATTRIBS
    oArs.Fields.Append "act", adBoolean, 0, NULLABLE_ATTRIBS
    oArs.Fields.Append "frac", adDouble, 0, NULLABLE_ATTRIBS
    oArs.Fields.Append "pic", adLongVarBinary, -1, LONG_ATTRIBS
    oArs.Open
    baBlob(0) = 1
    baBlob(1) = 255
    oArs.AddNew
    oArs.Fields("id").Value = 1
    oArs.Fields("nm").Value = "Ann"
    oArs.Fields("amt").Value = CCur(1234.5)
    oArs.Fields("bd").Value = DateSerial(2020, 2, 1) + TimeSerial(10, 30, 0)
    oArs.Fields("act").Value = True
    oArs.Fields("frac").Value = 2.5
    oArs.Fields("pic").Value = baBlob
    oArs.Update
    oArs.AddNew
    oArs.Fields("id").Value = 2
    oArs.Update
    sDbTc6 = Environ$("TEMP") & "\tc6_adors_tc6.db"
    sDbRc6 = Environ$("TEMP") & "\tc6_adors_rc6.db"
    If Len(Dir$(sDbTc6)) > 0 Then
        Kill sDbTc6
    End If
    Set oCnn = New cConnection
    oCnn.CreateNewDB sDbTc6
    oCnn.CreateTableFromADORs oCnn, "conv", oArs
    sTc6 = pvSchemaDump(oCnn) & "==" & vbLf & pvQuoteDump(sDbTc6, "conv", "[id]")
    AssertTrue InStr(sTc6, "table|conv|") > 0, "table created from ADO recordset"
    AssertTrue InStr(sTc6, "'Ann'") > 0, "data copied"
    '--- same fabricated recordset through RC6's converter, then compare the
    '--- raw storage (both files quote-dumped by TC6)
    On Error Resume Next
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    On Error GoTo EH
    If Not oRc6Cnn Is Nothing Then
        If Len(Dir$(sDbRc6)) > 0 Then
            Kill sDbRc6
        End If
        oRc6Cnn.CreateNewDB sDbRc6
        Set oRc6Cnn2 = oRc6Cnn
        oRc6Cnn.CreateTableFromADORs oRc6Cnn2, "conv", oArs
        sRc6 = pvSchemaDump(oRc6Cnn) & "==" & vbLf & pvQuoteDump(sDbRc6, "conv", "[id]")
        AssertEqStr sTc6, sRc6, "schema + raw storage match RC6"
    End If
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub

Private Sub Test_ConverterJetMdb()
    Dim oCat            As Object
    Dim oAdoCnn         As Object
    Dim oConv           As cConverter
    Dim oSink           As cEventSink
    Dim oCnn            As cConnection
    Dim oRc6Cnn         As Object
    Dim oRc6Conv        As Object
    Dim sMdb            As String
    Dim sTc6            As String
    Dim sTrace          As String
    Dim sDbTc6          As String
    Dim sDbRc6          As String

    If Not TestBegin("cConverter.ConvertJetMdb") Then Exit Sub
    On Error Resume Next
    sMdb = Environ$("TEMP") & "\tc6_conv_test.mdb"
    If Len(Dir$(sMdb)) > 0 Then
        Kill sMdb
    End If
    Set oCat = CreateObject("ADOX.Catalog")
    oCat.Create "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sMdb
    If oCat Is Nothing Or Err.Number <> 0 Then
        TestSkipCurrent "Jet OLEDB provider not available"
        Exit Sub
    End If
    On Error GoTo EH
    Set oAdoCnn = oCat.ActiveConnection
    oAdoCnn.Execute "CREATE TABLE Emp (id LONG CONSTRAINT pk PRIMARY KEY, nm TEXT(50), sal CURRENCY, bd DATETIME, act YESNO, notes MEMO, qty LONG, frac DOUBLE)"
    oAdoCnn.Execute "INSERT INTO Emp VALUES (1, 'Ann', 1234.5, #2020-02-01 10:30:00#, true, 'long text here', 42, 2.5)"
    oAdoCnn.Execute "INSERT INTO Emp (id, act) VALUES (2, false)"
    oAdoCnn.Execute "CREATE INDEX ix_nm ON Emp (nm)"
    oAdoCnn.Execute "CREATE TABLE T2 (k TEXT(10) CONSTRAINT pk2 PRIMARY KEY, v LONG)"
    oAdoCnn.Execute "INSERT INTO T2 VALUES ('a', 1)"
    '--- TC6 conversion with event sink
    sDbTc6 = Environ$("TEMP") & "\tc6_conv_tc6.db"
    sDbRc6 = Environ$("TEMP") & "\tc6_conv_rc6.db"
    If Len(Dir$(sDbTc6)) > 0 Then
        Kill sDbTc6
    End If
    Set oCnn = New cConnection
    oCnn.CreateNewDB sDbTc6
    Set oConv = New cConverter
    Set oSink = New cEventSink
    oSink.AttachConverter oConv
    oConv.ConvertDatabase oAdoCnn, oCnn
    oConv.ConvertIndexes oAdoCnn, oCnn
    sTrace = oSink.Trace
    AssertTrue InStr(sTrace, "SchemaProgress(Emp,2,") > 0, "SchemaProgress fired per table"
    AssertTrue InStr(sTrace, "InsertProgress(Emp,2,2)") > 0, "InsertProgress fired per row"
    AssertTrue InStr(sTrace, "IndexProgress(Emp,ix_nm,") > 0, "IndexProgress fired"
    AssertEqLng CLng(oCnn.GetRs("SELECT COUNT(*) FROM Emp").Fields(0).Value), 2, "rows converted"
    AssertEqLng CLng(oCnn.GetRs("SELECT COUNT(*) FROM T2").Fields(0).Value), 1, "second table converted"
    AssertTrue InStr(pvSchemaDump(oCnn), "ix_nm") > 0, "index converted"
    '--- RC6 conversion of the same mdb must match schema + raw storage
    On Error Resume Next
    Set oRc6Cnn = CreateObject("RC6.cConnection")
    On Error GoTo EH
    If Not oRc6Cnn Is Nothing Then
        If Len(Dir$(sDbRc6)) > 0 Then
            Kill sDbRc6
        End If
        oRc6Cnn.CreateNewDB sDbRc6
        Set oRc6Conv = CreateObject("RC6.cConverter")
        oRc6Conv.ConvertDatabase oAdoCnn, oRc6Cnn
        oRc6Conv.ConvertIndexes oAdoCnn, oRc6Cnn
        sTc6 = pvSchemaDump(oCnn) & "==" & vbLf & pvQuoteDump(sDbTc6, "Emp", "[id]") & pvQuoteDump(sDbTc6, "T2", "[k]")
        AssertEqStr sTc6, pvSchemaDump(oRc6Cnn) & "==" & vbLf & pvQuoteDump(sDbRc6, "Emp", "[id]") & pvQuoteDump(sDbRc6, "T2", "[k]"), "converted schema + raw storage match RC6"
    End If
    oAdoCnn.Close
    TestEnd
    Exit Sub
EH:
    TestErr
End Sub
