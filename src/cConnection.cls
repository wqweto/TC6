VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cConnection - SQLite database connection wrapper (winsqlite3.dll)
'
' Implemented: open/create/close, Execute + parameterised GetRs/ExecCmd,
' transactions with savepoints, recordsets/commands/cursors, schema
' objects (DataBases), MemDB, CreateTable, CopyDatabase, UniqueID64,
' UDF/collation registration, error/version info, pragmas, date helpers.
' Still stubs: CreateTableFromRsContent (needs Content serialization) and
' the ADO migration surface (CreateTableFromADORs).
'=========================================================================
Option Explicit

Public Enum eConnectionSynchronousType
    SynchronousOff = 0
    SynchronousNormal = 1
    SynchronousFull = 2
    SynchronousExtra = 3
End Enum

Public Enum eCodecType
    CODEC_TYPE_AES128 = 1
    CODEC_TYPE_AES256 = 2
    CODEC_TYPE_CHACHA20 = 3
    CODEC_TYPE_SQLCIPHER = 4
    CODEC_TYPE_RC4 = 5
End Enum

Public Event CommitTransComplete()
Public Event RollbackTransComplete()
Public Event CommitForSavePointComplete(SavePointName As String)
Public Event SavePointReleased(SavePointName As String)
Public Event RollbackToSavePointComplete(SavePointName As String)

Private m_hDb                       As LongPtr
Private m_sFileName                 As String
Private m_bMapDbNullToEmpty         As Boolean
Private m_eCodecType                As eCodecType
Private m_dblBusyTimeout            As Double
Private m_bEnableNestedTransactions As Boolean
Private m_cTransactionStack         As VBA.Collection
Private m_oMemDB                    As cMemDB

Public Property Get MapDbNullToEmpty() As Boolean
    MapDbNullToEmpty = m_bMapDbNullToEmpty
End Property

Public Property Let MapDbNullToEmpty(ByVal MapDbNullToEmpty As Boolean)
    m_bMapDbNullToEmpty = MapDbNullToEmpty
End Property

Public Property Get DBHdl() As Long
    DBHdl = m_hDb
End Property

Public Property Get FileName() As String
    FileName = m_sFileName
End Property

Public Property Let FileName(FileName As String)
    m_sFileName = FileName
End Property

Public Property Get PageSize() As Long
    PageSize = pvGetScalarLong("PRAGMA page_size")
End Property

Public Property Get BusyTimeOutSeconds() As Double
    BusyTimeOutSeconds = m_dblBusyTimeout
End Property

Public Property Let BusyTimeOutSeconds(ByVal RHS As Double)
    m_dblBusyTimeout = RHS
    If m_hDb <> 0 Then
        Call stub_sqlite3_busy_timeout(m_hDb, CLng(RHS * 1000))
    End If
End Property

Public Property Get Synchronous() As eConnectionSynchronousType
    Synchronous = pvGetScalarLong("PRAGMA synchronous")
End Property

Public Property Let Synchronous(ByVal RHS As eConnectionSynchronousType)
    Execute "PRAGMA synchronous = " & CLng(RHS)
End Property

Public Property Get DataBases() As cDataBases
    Dim oDbs            As cDataBases

    Set oDbs = New cDataBases
    oDbs.frInit Me
    Set DataBases = oDbs
End Property

Public Property Get UniqueID64() As Variant
    UniqueID64 = CreateUniqueID64()
End Property

Public Property Get AffectedRows() As Long
    If m_hDb <> 0 Then
        AffectedRows = stub_sqlite3_changes(m_hDb)
    End If
End Property

Public Property Get TotalAffectedRowsInSession() As Long
    If m_hDb <> 0 Then
        TotalAffectedRowsInSession = stub_sqlite3_total_changes(m_hDb)
    End If
End Property

Public Property Get LastInsertAutoID() As Variant
    LastInsertAutoID = pvLastInsertId()
End Property

Public Property Get EnableNestedTransactions() As Boolean
    EnableNestedTransactions = m_bEnableNestedTransactions
End Property

Public Property Let EnableNestedTransactions(ByVal RHS As Boolean)
    m_bEnableNestedTransactions = RHS
End Property

Public Property Get TransactionStackCounter() As Long
    TransactionStackCounter = m_cTransactionStack.Count
End Property

Public Property Get TransactionStack() As VBA.Collection
    Set TransactionStack = m_cTransactionStack
End Property

Public Property Get MemDB() As cMemDB
    If m_oMemDB Is Nothing Then
        Set m_oMemDB = New cMemDB
    End If
    Set MemDB = m_oMemDB
End Property

Public Property Get CodecType() As eCodecType
    CodecType = m_eCodecType
End Property

Public Property Let CodecType(ByVal RHS As eCodecType)
    m_eCodecType = RHS
End Property

Public Function Version() As String
    Version = FromUtf8Ptr(stub_sqlite3_libversion())
End Function

Public Function CheckIntegrity() As String
    CheckIntegrity = pvGetScalarText("PRAGMA integrity_check")
End Function

Public Function CreateNewDB(Optional FileName As String = ":memory:", Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
    '--- winsqlite3.dll has no codec, so EncrKey is accepted but ignored
    CreateNewDB = pvOpenDatabase(FileName, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE)
End Function

Public Function OpenDB(Optional FileName As String, Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
    If Len(FileName) = 0 Then
        FileName = m_sFileName
    End If
    OpenDB = pvOpenDatabase(FileName, SQLITE_OPEN_READWRITE Or SQLITE_OPEN_CREATE)
End Function

Public Function OpenDBReadOnly(Optional FileName As String, Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
    If Len(FileName) = 0 Then
        FileName = m_sFileName
    End If
    OpenDBReadOnly = pvOpenDatabase(FileName, SQLITE_OPEN_READONLY)
End Function

Public Function OpenSchema() As cRecordset
    Set OpenSchema = OpenRecordset("SELECT type, name, tbl_name, rootpage, sql FROM sqlite_master ORDER BY type, name")
End Function

Public Sub ReKey(Optional NewEncrKey As String)
    '--- winsqlite3.dll ships without a codec; rekeying is unsupported
    Err.Raise vbObjectError, "cConnection.ReKey", "Encryption is not supported by winsqlite3.dll"
End Sub

Public Sub AttachDataBase(FileName As String, DataBaseName As String, Optional EncrKey As String)
    Execute "ATTACH DATABASE " & QuoteString(FileName) & " AS " & QuoteIdentifier(DataBaseName)
End Sub

Public Sub DetachDataBase(DataBaseName As String)
    Execute "DETACH DATABASE " & QuoteIdentifier(DataBaseName)
End Sub

Public Function OpenRecordset(SQL As String, Optional ByVal ReadOnly As Boolean) As cRecordset
    Dim oRs             As cRecordset

    Set oRs = New cRecordset
    oRs.frOpen Me, SQL, ReadOnly
    Set OpenRecordset = oRs
End Function

Public Sub BeginTrans(Optional SavePointName As String)
    Dim sName           As String

    '--- an outer BEGIN when the stack is empty and no explicit savepoint
    If m_cTransactionStack.Count = 0 And Len(SavePointName) = 0 Then
        Execute "BEGIN"
        m_cTransactionStack.Add ""
        Exit Sub
    End If
    '--- otherwise a (nested) savepoint - gated by EnableNestedTransactions
    If m_cTransactionStack.Count > 0 And Not m_bEnableNestedTransactions Then
        Err.Raise vbObjectError, "cConnection.BeginTrans", "Nested transactions are disabled"
    End If
    If Len(SavePointName) = 0 Then
        sName = "SP" & (m_cTransactionStack.Count + 1)
    Else
        sName = SavePointName
    End If
    Execute "SAVEPOINT " & QuoteIdentifier(sName)
    m_cTransactionStack.Add sName
End Sub

Public Sub CommitTrans(Optional SavePointName As String)
    Dim sName           As String

    If m_cTransactionStack.Count = 0 Then
        Exit Sub
    End If
    sName = m_cTransactionStack.Item(m_cTransactionStack.Count)
    m_cTransactionStack.Remove m_cTransactionStack.Count
    If Len(sName) = 0 Then
        Execute "COMMIT"
        RaiseEvent CommitTransComplete
    Else
        Execute "RELEASE SAVEPOINT " & QuoteIdentifier(sName)
        RaiseEvent SavePointReleased(sName)
        RaiseEvent CommitForSavePointComplete(sName)
    End If
End Sub

Public Sub RollbackTrans(Optional ToSavePointName As String, Optional ByVal ReUseSavePointForNewAttempts As Boolean)
    Dim sName           As String

    If m_cTransactionStack.Count = 0 Then
        Exit Sub
    End If
    '--- rollback to a named savepoint without unwinding the whole transaction
    If Len(ToSavePointName) > 0 Then
        Execute "ROLLBACK TO SAVEPOINT " & QuoteIdentifier(ToSavePointName)
        If Not ReUseSavePointForNewAttempts Then
            Execute "RELEASE SAVEPOINT " & QuoteIdentifier(ToSavePointName)
            pvPopTo ToSavePointName
        End If
        RaiseEvent RollbackToSavePointComplete(ToSavePointName)
        Exit Sub
    End If
    '--- unwind the innermost frame
    sName = m_cTransactionStack.Item(m_cTransactionStack.Count)
    m_cTransactionStack.Remove m_cTransactionStack.Count
    If Len(sName) = 0 Then
        Execute "ROLLBACK"
        RaiseEvent RollbackTransComplete
    Else
        Execute "ROLLBACK TO SAVEPOINT " & QuoteIdentifier(sName)
        Execute "RELEASE SAVEPOINT " & QuoteIdentifier(sName)
        RaiseEvent RollbackToSavePointComplete(sName)
    End If
End Sub

Public Function CopyDatabase(Optional ByVal DstDBFileName As String = ":memory:", Optional ByVal EncrKey As String, Optional ByVal CompactDstDB As Boolean) As cConnection
    Dim oDst            As cConnection
    Dim baMain()        As Byte
    Dim hBackup         As LongPtr
    Dim lRc             As Long

    If m_hDb = 0 Then
        Err.Raise vbObjectError, "cConnection.CopyDatabase", "Database is not open"
    End If
    Set oDst = New cConnection
    If Not oDst.CreateNewDB(DstDBFileName, EncrKey) Then
        Err.Raise vbObjectError, "cConnection.CopyDatabase", "Cannot create '" & DstDBFileName & "'"
    End If
    baMain = ToUtf8Array("main" & vbNullChar)
    hBackup = stub_sqlite3_backup_init(oDst.frDbHandle, VarPtr(baMain(0)), m_hDb, VarPtr(baMain(0)))
    If hBackup = 0 Then
        Err.Raise vbObjectError, "cConnection.CopyDatabase", oDst.LastDBError()
    End If
    lRc = stub_sqlite3_backup_step(hBackup, -1)
    Call stub_sqlite3_backup_finish(hBackup)
    If lRc <> SQLITE_DONE Then
        Err.Raise vbObjectError, "cConnection.CopyDatabase", oDst.LastDBError()
    End If
    If CompactDstDB Then
        oDst.CompactDataBase
    End If
    Set CopyDatabase = oDst
End Function

Public Sub CompactDataBase()
    Execute "VACUUM"
End Sub

Public Sub Cancel()
    If m_hDb <> 0 Then
        Call stub_sqlite3_interrupt(m_hDb)
    End If
End Sub

Public Function CreateCommand(SQL As String) As cCommand
    Dim oCmd            As cCommand
    Dim sSaved          As String

    '--- RC6 semantics: a saved CommandKey resolves to its stored SQL
    sSaved = LookupSavedSql(Me, "dhWriteCommands", SQL)
    Set oCmd = New cCommand
    oCmd.frInit Me, IIf(Len(sSaved) > 0, sSaved, SQL)
    Set CreateCommand = oCmd
End Function

Public Function CreateSelectCommand(SQL As String) As cSelectCommand
    Dim oCmd            As cSelectCommand
    Dim sSaved          As String

    sSaved = LookupSavedSql(Me, "dhSelectCommands", SQL)
    Set oCmd = New cSelectCommand
    oCmd.frInit Me, IIf(Len(sSaved) > 0, sSaved, SQL)
    Set CreateSelectCommand = oCmd
End Function

Public Function CreateCursor(SQL As String) As cCursor
    Dim oCursor         As cCursor

    Set oCursor = New cCursor
    oCursor.frInit Me, SQL
    Set CreateCursor = oCursor
End Function

Public Function Execute(SQL As String) As Boolean
    Dim lRc             As Long

    lRc = pvExec(SQL)
    If lRc <> SQLITE_OK Then
        Err.Raise vbObjectError, "cConnection.Execute", pvLastError()
    End If
    Execute = True
End Function

Public Function GetDateString(ByVal D As Date) As String
    GetDateString = Format$(D, "yyyy-mm-dd hh:nn:ss")
End Function

Public Function GetShortDateString(ByVal D As Date) As String
    GetShortDateString = Format$(D, "yyyy-mm-dd")
End Function

Public Function GetTimeString(ByVal D As Date) As String
    GetTimeString = Format$(D, "hh:nn:ss")
End Function

Public Function GetBooleanString(ByVal B As Boolean) As String
    GetBooleanString = IIf(B, "1", "0")
End Function

Public Function LastDBError() As String
    LastDBError = pvLastError()
End Function

Public Function GetSqliteErrStr(ByVal ResultCode As Long) As String
    GetSqliteErrStr = FromUtf8Ptr(stub_sqlite3_errstr(ResultCode))
End Function

Public Function LastDBErrCode(Optional ByVal GetExtendedCode As Boolean) As Long
    If m_hDb = 0 Then
        Exit Function
    End If
    If GetExtendedCode Then
        LastDBErrCode = stub_sqlite3_extended_errcode(m_hDb)
    Else
        LastDBErrCode = stub_sqlite3_errcode(m_hDb)
    End If
End Function

Public Function AddUserDefinedFunction(New_ObjectIFunc As IFunction) As Boolean
    AddUserDefinedFunction = UdfRegisterFunction(Me, New_ObjectIFunc)
End Function

Public Function RemoveUserDefinedFunction(New_ObjectIFunc As IFunction) As Boolean
    RemoveUserDefinedFunction = UdfUnregisterFunction(Me, New_ObjectIFunc)
End Function

Public Function AddUserDefinedAggregateFunction(New_ObjectIAggFunc As IAggregateFunction) As Boolean
    AddUserDefinedAggregateFunction = UdfRegisterAggregate(Me, New_ObjectIAggFunc)
End Function

Public Function RemoveUserDefinedAggregateFunction(New_ObjectIAggFunc As IAggregateFunction) As Boolean
    RemoveUserDefinedAggregateFunction = UdfUnregisterAggregate(Me, New_ObjectIAggFunc)
End Function

Public Function AddUserDefinedCollation(New_ObjectIColl As ICollation) As Boolean
    AddUserDefinedCollation = UdfRegisterCollation(Me, New_ObjectIColl)
End Function

Public Function RemoveUserDefinedCollation(New_ObjectIColl As ICollation) As Boolean
    RemoveUserDefinedCollation = UdfUnregisterCollation(Me, New_ObjectIColl)
End Function

Public Sub CreateTableFromRsContent(Content() As Byte, Optional NewTableName As String, Optional ByVal AsTempTable As Boolean, Optional ByVal WithPrimaryKeys As Boolean)
    Dim oRs             As cRecordset
    Dim sTable          As String
    Dim lCol            As Long
    Dim baTemp()        As Byte

    '--- VB6 bug: assigning a ByRef array PARAMETER straight to a Property
    '--- Let crashes at runtime - copy to a local first
    baTemp = Content
    Set oRs = New cRecordset
    oRs.Content = baTemp
    sTable = NewTableName
    If Len(sTable) = 0 Then
        '--- default to the origin table carried by the blob
        For lCol = 0 To oRs.frFieldCount - 1
            sTable = oRs.frColOriginTable(lCol)
            If Len(sTable) > 0 Then
                Exit For
            End If
        Next
    End If
    If Len(sTable) = 0 Then
        Err.Raise 5, "cConnection.CreateTableFromRsContent", "Table name required"
    End If
    CreateTableFromRecordset Me, oRs, sTable, AsTempTable, WithPrimaryKeys
End Sub

Public Sub CreateTableFromADORs(CnnDst As cConnection, ByVal TableName As String, AdoRs As Object, Optional arrDataTypes As Variant)
    Dim oCnn            As cConnection

    Set oCnn = CnnDst
    If oCnn Is Nothing Then
        Set oCnn = Me
    End If
    CopyAdoRsToTable oCnn, TableName, AdoRs, arrDataTypes, False, Nothing
End Sub

Public Sub CreateTable(TableName As String, Optional FieldDefs As Collection)
    Dim sDefs           As String
    Dim vElem           As Variant

    '--- each FieldDefs item is a verbatim column definition, e.g. "F1 Text"
    If FieldDefs Is Nothing Then
        Err.Raise 5, "cConnection.CreateTable", "FieldDefs required"
    End If
    For Each vElem In FieldDefs
        If Len(sDefs) > 0 Then
            sDefs = sDefs & ", "
        End If
        sDefs = sDefs & vElem
    Next
    If Len(sDefs) = 0 Then
        Err.Raise 5, "cConnection.CreateTable", "FieldDefs is empty"
    End If
    Execute "CREATE TABLE " & QuoteIdentifier(TableName) & " (" & sDefs & ")"
End Sub

Public Function NewFieldDefs() As Collection
    Set NewFieldDefs = New Collection
End Function

Public Function GetRs(ByVal SQL As String, ParamArray P() As Variant) As cRecordset
    Dim aParams()       As Variant

    aParams = P
    Set GetRs = frGetRs(SQL, aParams)
End Function

Public Sub ExecCmd(ByVal SQL As String, ParamArray P() As Variant)
    Dim aParams()       As Variant

    aParams = P
    frExecCmd SQL, aParams
End Sub

Friend Property Get frDbHandle() As LongPtr
    frDbHandle = m_hDb
End Property

Friend Function frGetRs(sSql As String, aParams() As Variant) As cRecordset
    Dim oRs             As cRecordset

    Set oRs = New cRecordset
    oRs.frOpenParams Me, sSql, aParams
    Set frGetRs = oRs
End Function

'--- private helpers ------------------------------------------------------

Private Function pvOpenDatabase(sFileName As String, ByVal lFlags As Long) As Boolean
    Dim baFile()        As Byte
    Dim lRc             As Long

    pvCloseDatabase
    baFile = ToUtf8Array(sFileName & vbNullChar)
    lRc = stub_sqlite3_open_v2(VarPtr(baFile(0)), VarPtr(m_hDb), lFlags, 0)
    If lRc <> SQLITE_OK Then
        pvCloseDatabase
        Exit Function
    End If
    m_sFileName = sFileName
    If m_dblBusyTimeout > 0 Then
        Call stub_sqlite3_busy_timeout(m_hDb, CLng(m_dblBusyTimeout * 1000))
    End If
    pvOpenDatabase = True
End Function

Private Sub pvCloseDatabase()
    If m_hDb <> 0 Then
        UdfReleaseDb m_hDb
        Call stub_sqlite3_close_v2(m_hDb)
        m_hDb = 0
    End If
    Do While m_cTransactionStack.Count > 0
        m_cTransactionStack.Remove m_cTransactionStack.Count
    Loop
End Sub

Private Function pvExec(sSql As String) As Long
    Dim baSql()         As Byte

    If m_hDb = 0 Then
        pvExec = SQLITE_MISUSE
        Exit Function
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    pvExec = stub_sqlite3_exec(m_hDb, VarPtr(baSql(0)), 0, 0, 0)
End Function

Friend Sub frExecCmd(sSql As String, aParams() As Variant)
    Dim baSql()         As Byte
    Dim hStmt           As LongPtr
    Dim lIdx            As Long
    Dim lRc             As Long

    If m_hDb = 0 Then
        Err.Raise vbObjectError, "cConnection.ExecCmd", "Database is not open"
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    If stub_sqlite3_prepare_v2(m_hDb, VarPtr(baSql(0)), -1, VarPtr(hStmt), 0) <> SQLITE_OK Then
        Err.Raise vbObjectError, "cConnection.ExecCmd", pvLastError()
    End If
    On Error GoTo EH
    For lIdx = LBound(aParams) To UBound(aParams)
        If BindVariant(hStmt, lIdx - LBound(aParams) + 1, aParams(lIdx)) <> SQLITE_OK Then
            Err.Raise vbObjectError, "cConnection.ExecCmd", pvLastError()
        End If
    Next
    '--- run to completion (tolerate a statement that yields rows)
    Do
        lRc = stub_sqlite3_step(hStmt)
    Loop While lRc = SQLITE_ROW
    If lRc <> SQLITE_DONE Then
        Err.Raise vbObjectError, "cConnection.ExecCmd", pvLastError()
    End If
    Call stub_sqlite3_finalize(hStmt)
    Exit Sub
EH:
    Call stub_sqlite3_finalize(hStmt)
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function pvGetScalarLong(sSql As String) As Long
    Dim baSql()         As Byte
    Dim hStmt           As LongPtr

    If m_hDb = 0 Then
        Exit Function
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    If stub_sqlite3_prepare_v2(m_hDb, VarPtr(baSql(0)), -1, VarPtr(hStmt), 0) = SQLITE_OK Then
        If stub_sqlite3_step(hStmt) = SQLITE_ROW Then
            pvGetScalarLong = stub_sqlite3_column_int(hStmt, 0)
        End If
        Call stub_sqlite3_finalize(hStmt)
    End If
End Function

Private Function pvGetScalarText(sSql As String) As String
    Dim baSql()         As Byte
    Dim hStmt           As LongPtr

    If m_hDb = 0 Then
        Exit Function
    End If
    baSql = ToUtf8Array(sSql & vbNullChar)
    If stub_sqlite3_prepare_v2(m_hDb, VarPtr(baSql(0)), -1, VarPtr(hStmt), 0) = SQLITE_OK Then
        If stub_sqlite3_step(hStmt) = SQLITE_ROW Then
            pvGetScalarText = FromUtf8Ptr(stub_sqlite3_column_text(hStmt, 0))
        End If
        Call stub_sqlite3_finalize(hStmt)
    End If
End Function

Private Function pvLastInsertId() As Variant
    If m_hDb = 0 Then
        pvLastInsertId = CVar(0)
        Exit Function
    End If
    '--- on x86 the API returns the raw int64 bits reinterpreted as Currency
    '--- (value * 10000), so scale back up through Decimal to recover it
#If Win64 Then
    pvLastInsertId = stub_sqlite3_last_insert_rowid(m_hDb)
#Else
    pvLastInsertId = CDec(stub_sqlite3_last_insert_rowid(m_hDb)) * CDec(10000)
#End If
End Function

Private Function pvLastError() As String
    If m_hDb <> 0 Then
        pvLastError = FromUtf8Ptr(stub_sqlite3_errmsg(m_hDb))
    End If
End Function

Private Sub pvPopTo(sName As String)
    '--- drop stack frames down to and including the named savepoint
    Do While m_cTransactionStack.Count > 0
        If m_cTransactionStack.Item(m_cTransactionStack.Count) = sName Then
            m_cTransactionStack.Remove m_cTransactionStack.Count
            Exit Do
        End If
        m_cTransactionStack.Remove m_cTransactionStack.Count
    Loop
End Sub

Private Sub Class_Initialize()
    Set m_cTransactionStack = New VBA.Collection
    m_bMapDbNullToEmpty = True
    m_bEnableNestedTransactions = True
End Sub

Private Sub Class_Terminate()
    pvCloseDatabase
End Sub
