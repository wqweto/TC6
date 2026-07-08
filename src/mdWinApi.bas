Attribute VB_Name = "mdWinApi"
'=========================================================================
' mdWinApi - API declares for the built-in Windows SQLite (winsqlite3.dll)
'
' - All declared functions are StdCall, so the module compiles unchanged
'   in VB6 (x86) and twinBASIC (x86/x64). The 8 CDecl/variadic printf-style
'   exports are intentionally NOT declared (see footer).
' - Handles/pointers are LongPtr; sqlite3_int64 is Currency (8-byte, kept
'   VB6-compatible); int is Long
' - SQLite is UTF-8: pass/receive LongPtr to byte buffers, not VB String
' - Only core constants are declared here; for the full set of result
'   codes, flags and options refer to doc\sqlite3.h
'=========================================================================
Option Explicit

#If Win64 = 0 And TWINBASIC = 0 Then
Public Enum LongPtr
    [_]
End Enum
#End If

'--- Fundamental result codes (see doc\sqlite3.h for the full set)
Public Const SQLITE_OK                      As Long = 0
Public Const SQLITE_ERROR                   As Long = 1
Public Const SQLITE_BUSY                    As Long = 5
Public Const SQLITE_LOCKED                  As Long = 6
Public Const SQLITE_NOMEM                   As Long = 7
Public Const SQLITE_CONSTRAINT              As Long = 19
Public Const SQLITE_MISUSE                  As Long = 21
Public Const SQLITE_RANGE                   As Long = 25
Public Const SQLITE_ROW                     As Long = 100
Public Const SQLITE_DONE                    As Long = 101

'--- Fundamental datatypes (column/value type)
Public Const SQLITE_INTEGER                 As Long = 1
Public Const SQLITE_FLOAT                   As Long = 2
Public Const SQLITE_TEXT                    As Long = 3
Public Const SQLITE_BLOB                    As Long = 4
Public Const SQLITE_NULL                    As Long = 5

'--- Text encodings / function flags (create_function eTextRep)
Public Const SQLITE_UTF8                    As Long = 1
Public Const SQLITE_UTF16LE                 As Long = 2
Public Const SQLITE_UTF16BE                 As Long = 3
Public Const SQLITE_UTF16                   As Long = 4
Public Const SQLITE_DETERMINISTIC           As Long = &H800&
Public Const SQLITE_DIRECTONLY              As Long = &H80000

'--- Open flags (sqlite3_open_v2)
Public Const SQLITE_OPEN_READONLY           As Long = &H1
Public Const SQLITE_OPEN_READWRITE          As Long = &H2
Public Const SQLITE_OPEN_CREATE             As Long = &H4
Public Const SQLITE_OPEN_URI                As Long = &H40&
Public Const SQLITE_OPEN_MEMORY             As Long = &H80&
Public Const SQLITE_OPEN_NOMUTEX            As Long = &H8000&
Public Const SQLITE_OPEN_FULLMUTEX          As Long = &H10000
Public Const SQLITE_OPEN_SHAREDCACHE        As Long = &H20000
Public Const SQLITE_OPEN_PRIVATECACHE       As Long = &H40000
Public Const SQLITE_OPEN_WAL                As Long = &H80000

'--- Prepare flags (sqlite3_prepare_v3)
Public Const SQLITE_PREPARE_PERSISTENT      As Long = &H1
Public Const SQLITE_PREPARE_NO_VTAB         As Long = &H4

'--- Destructor sentinels (last arg of bind_*/result_* text/blob)
#If Win64 Then
Public Const SQLITE_STATIC                  As LongPtr = 0
Public Const SQLITE_TRANSIENT               As LongPtr = -1
#Else
Public Const SQLITE_STATIC                  As Long = 0
Public Const SQLITE_TRANSIENT               As Long = -1
#End If

'--- winsqlite3.dll exports (275 declares)
Public Declare Function sqlite3_aggregate_context Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal nBytes As Long) As LongPtr
Public Declare Function sqlite3_aggregate_count Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_auto_extension Lib "winsqlite3" (ByVal xEntryPoint As LongPtr) As Long
Public Declare Function sqlite3_autovacuum_pages Lib "winsqlite3" (ByVal db As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare Function sqlite3_backup_finish Lib "winsqlite3" (ByVal p As LongPtr) As Long
Public Declare Function sqlite3_backup_init Lib "winsqlite3" (ByVal pDest As LongPtr, ByVal zDestName As LongPtr, ByVal pSource As LongPtr, ByVal zSourceName As LongPtr) As LongPtr
Public Declare Function sqlite3_backup_pagecount Lib "winsqlite3" (ByVal p As LongPtr) As Long
Public Declare Function sqlite3_backup_remaining Lib "winsqlite3" (ByVal p As LongPtr) As Long
Public Declare Function sqlite3_backup_step Lib "winsqlite3" (ByVal p As LongPtr, ByVal nPage As Long) As Long
Public Declare Function sqlite3_bind_blob Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal n As Long, ByVal cb5 As LongPtr) As Long
Public Declare Function sqlite3_bind_blob64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Currency, ByVal cb5 As LongPtr) As Long
Public Declare Function sqlite3_bind_double Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Double) As Long
Public Declare Function sqlite3_bind_int Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Long) As Long
Public Declare Function sqlite3_bind_int64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Currency) As Long
Public Declare Function sqlite3_bind_null Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare Function sqlite3_bind_parameter_count Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_bind_parameter_index Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zName As LongPtr) As Long
Public Declare Function sqlite3_bind_parameter_name Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_bind_pointer Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal cb5 As LongPtr) As Long
Public Declare Function sqlite3_bind_text Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
Public Declare Function sqlite3_bind_text16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
Public Declare Function sqlite3_bind_text64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Currency, ByVal cb5 As LongPtr, ByVal encoding As Long) As Long
Public Declare Function sqlite3_bind_value Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr) As Long
Public Declare Function sqlite3_bind_zeroblob Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal n As Long) As Long
Public Declare Function sqlite3_bind_zeroblob64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Currency) As Long
Public Declare Function sqlite3_blob_bytes Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_blob_close Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_blob_open Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zDb As LongPtr, ByVal zTable As LongPtr, ByVal zColumn As LongPtr, ByVal iRow As Currency, ByVal flags As Long, ByVal ppBlob As LongPtr) As Long
Public Declare Function sqlite3_blob_read Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal Z As LongPtr, ByVal N As Long, ByVal iOffset As Long) As Long
Public Declare Function sqlite3_blob_reopen Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Currency) As Long
Public Declare Function sqlite3_blob_write Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As Long, ByVal iOffset As Long) As Long
Public Declare Function sqlite3_busy_handler Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As Long
Public Declare Function sqlite3_busy_timeout Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal ms As Long) As Long
Public Declare Function sqlite3_cancel_auto_extension Lib "winsqlite3" (ByVal xEntryPoint As LongPtr) As Long
Public Declare Function sqlite3_changes Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_changes64 Lib "winsqlite3" (ByVal p1 As LongPtr) As Currency
Public Declare Function sqlite3_clear_bindings Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_close Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_close_v2 Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_collation_needed Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare Function sqlite3_collation_needed16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare Function sqlite3_column_blob Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function sqlite3_column_bytes Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_bytes16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_count Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Function sqlite3_column_database_name Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_database_name16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_decltype Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_decltype16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_double Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Double
Public Declare Function sqlite3_column_int Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_int64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Currency
Public Declare Function sqlite3_column_name Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function sqlite3_column_name16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function sqlite3_column_origin_name Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_origin_name16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_table_name Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_table_name16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_column_text Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function sqlite3_column_text16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function sqlite3_column_type Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function sqlite3_column_value Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function sqlite3_commit_hook Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function sqlite3_compileoption_get Lib "winsqlite3" (ByVal N As Long) As LongPtr
Public Declare Function sqlite3_compileoption_used Lib "winsqlite3" (ByVal zOptName As LongPtr) As Long
Public Declare Function sqlite3_complete Lib "winsqlite3" (ByVal sql As LongPtr) As Long
Public Declare Function sqlite3_complete16 Lib "winsqlite3" (ByVal sql As LongPtr) As Long
Public Declare Function sqlite3_context_db_handle Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_create_collation Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare Function sqlite3_create_collation_v2 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function sqlite3_create_collation16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare Function sqlite3_create_filename Lib "winsqlite3" (ByVal zDatabase As LongPtr, ByVal zJournal As LongPtr, ByVal zWal As LongPtr, ByVal nParam As Long, ByVal azParam As LongPtr) As LongPtr
Public Declare Function sqlite3_create_function Lib "winsqlite3" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare Function sqlite3_create_function_v2 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function sqlite3_create_function16 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare Function sqlite3_create_module Lib "winsqlite3" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr) As Long
Public Declare Function sqlite3_create_module_v2 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function sqlite3_create_window_function Lib "winsqlite3" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xValue As LongPtr, ByVal xInverse As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function sqlite3_data_count Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Function sqlite3_database_file_object Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_db_cacheflush Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_db_filename Lib "winsqlite3" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As LongPtr
Public Declare Function sqlite3_db_handle Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_db_mutex Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_db_name Lib "winsqlite3" (ByVal db As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function sqlite3_db_readonly Lib "winsqlite3" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As Long
Public Declare Function sqlite3_db_release_memory Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_db_status Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal op As Long, ByVal pCur As LongPtr, ByVal pHiwtr As LongPtr, ByVal resetFlg As Long) As Long
Public Declare Function sqlite3_db_status64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal p5 As Long) As Long
Public Declare Function sqlite3_declare_vtab Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zSQL As LongPtr) As Long
Public Declare Function sqlite3_deserialize Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal pData As LongPtr, ByVal szDb As Currency, ByVal szBuf As Currency, ByVal mFlags As Long) As Long
Public Declare Function sqlite3_drop_modules Lib "winsqlite3" (ByVal db As LongPtr, ByVal azKeep As LongPtr) As Long
Public Declare Function sqlite3_enable_load_extension Lib "winsqlite3" (ByVal db As LongPtr, ByVal onoff As Long) As Long
Public Declare Function sqlite3_enable_shared_cache Lib "winsqlite3" (ByVal p1 As Long) As Long
Public Declare Function sqlite3_errcode Lib "winsqlite3" (ByVal db As LongPtr) As Long
Public Declare Function sqlite3_errmsg Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_errmsg16 Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_error_offset Lib "winsqlite3" (ByVal db As LongPtr) As Long
Public Declare Function sqlite3_errstr Lib "winsqlite3" (ByVal p1 As Long) As LongPtr
Public Declare Function sqlite3_exec Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal sql As LongPtr, ByVal callback As LongPtr, ByVal p4 As LongPtr, ByVal errmsg As LongPtr) As Long
Public Declare Function sqlite3_expanded_sql Lib "winsqlite3" (ByVal pStmt As LongPtr) As LongPtr
Public Declare Function sqlite3_expired Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_extended_errcode Lib "winsqlite3" (ByVal db As LongPtr) As Long
Public Declare Function sqlite3_extended_result_codes Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal onoff As Long) As Long
Public Declare Function sqlite3_file_control Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zDbName As LongPtr, ByVal op As Long, ByVal p4 As LongPtr) As Long
Public Declare Function sqlite3_filename_database Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_filename_journal Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_filename_wal Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_finalize Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Sub sqlite3_free Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_free_filename Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_free_table Lib "winsqlite3" (ByVal result As LongPtr)
Public Declare Function sqlite3_get_autocommit Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_get_auxdata Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function sqlite3_get_clientdata Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare Function sqlite3_get_table Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal pazResult As LongPtr, ByVal pnRow As LongPtr, ByVal pnColumn As LongPtr, ByVal pzErrmsg As LongPtr) As Long
Public Declare Function sqlite3_global_recover Lib "winsqlite3" () As Long
Public Declare Function sqlite3_hard_heap_limit64 Lib "winsqlite3" (ByVal N As Currency) As Currency
Public Declare Function sqlite3_initialize Lib "winsqlite3" () As Long
Public Declare Sub sqlite3_interrupt Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Function sqlite3_is_interrupted Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_keyword_check Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare Function sqlite3_keyword_count Lib "winsqlite3" () As Long
Public Declare Function sqlite3_keyword_name Lib "winsqlite3" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr) As Long
Public Declare Function sqlite3_last_insert_rowid Lib "winsqlite3" (ByVal p1 As LongPtr) As Currency
Public Declare Function sqlite3_libversion Lib "winsqlite3" () As LongPtr
Public Declare Function sqlite3_libversion_number Lib "winsqlite3" () As Long
Public Declare Function sqlite3_limit Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal id As Long, ByVal newVal As Long) As Long
Public Declare Function sqlite3_load_extension Lib "winsqlite3" (ByVal db As LongPtr, ByVal zFile As LongPtr, ByVal zProc As LongPtr, ByVal pzErrMsg As LongPtr) As Long
Public Declare Function sqlite3_malloc Lib "winsqlite3" (ByVal p1 As Long) As LongPtr
Public Declare Function sqlite3_malloc64 Lib "winsqlite3" (ByVal p1 As Currency) As LongPtr
Public Declare Function sqlite3_memory_alarm Lib "winsqlite3" (ByVal cb1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency) As Long
Public Declare Function sqlite3_memory_highwater Lib "winsqlite3" (ByVal resetFlag As Long) As Currency
Public Declare Function sqlite3_memory_used Lib "winsqlite3" () As Currency
Public Declare Function sqlite3_msize Lib "winsqlite3" (ByVal p1 As LongPtr) As Currency
Public Declare Function sqlite3_mutex_alloc Lib "winsqlite3" (ByVal p1 As Long) As LongPtr
Public Declare Sub sqlite3_mutex_enter Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_mutex_free Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_mutex_leave Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Function sqlite3_mutex_try Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_next_stmt Lib "winsqlite3" (ByVal pDb As LongPtr, ByVal pStmt As LongPtr) As LongPtr
Public Declare Function sqlite3_open Lib "winsqlite3" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare Function sqlite3_open_v2 Lib "winsqlite3" (ByVal filename As LongPtr, ByVal ppDb As LongPtr, ByVal flags As Long, ByVal zVfs As LongPtr) As Long
Public Declare Function sqlite3_open16 Lib "winsqlite3" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare Function sqlite3_os_end Lib "winsqlite3" () As Long
Public Declare Function sqlite3_os_init Lib "winsqlite3" () As Long
Public Declare Function sqlite3_overload_function Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zFuncName As LongPtr, ByVal nArg As Long) As Long
Public Declare Function sqlite3_prepare Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_prepare_v2 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_prepare_v3 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_prepare16 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_prepare16_v2 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_prepare16_v3 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function sqlite3_profile Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal xProfile As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Sub sqlite3_progress_handler Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal cb3 As LongPtr, ByVal p4 As LongPtr)
Public Declare Sub sqlite3_randomness Lib "winsqlite3" (ByVal N As Long, ByVal P As LongPtr)
Public Declare Function sqlite3_realloc Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_realloc64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Currency) As LongPtr
Public Declare Function sqlite3_release_memory Lib "winsqlite3" (ByVal p1 As Long) As Long
Public Declare Function sqlite3_reset Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Sub sqlite3_reset_auto_extension Lib "winsqlite3" ()
Public Declare Sub sqlite3_result_blob Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_blob64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_double Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Double)
Public Declare Sub sqlite3_result_error Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare Sub sqlite3_result_error_code Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare Sub sqlite3_result_error_nomem Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_result_error_toobig Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_result_error16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare Sub sqlite3_result_int Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare Sub sqlite3_result_int64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Currency)
Public Declare Sub sqlite3_result_null Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Sub sqlite3_result_pointer Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_subtype Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare Sub sqlite3_result_text Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_text16 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_text16be Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_text16le Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub sqlite3_result_text64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As Currency, ByVal cb4 As LongPtr, ByVal encoding As Long)
Public Declare Sub sqlite3_result_value Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr)
Public Declare Sub sqlite3_result_zeroblob Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal n As Long)
Public Declare Function sqlite3_result_zeroblob64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal n As Currency) As Long
Public Declare Function sqlite3_rollback_hook Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function sqlite3_rtree_geometry_callback Lib "winsqlite3" (ByVal db As LongPtr, ByVal zGeom As LongPtr, ByVal xGeom As LongPtr, ByVal pContext As LongPtr) As Long
Public Declare Function sqlite3_rtree_query_callback Lib "winsqlite3" (ByVal db As LongPtr, ByVal zQueryFunc As LongPtr, ByVal xQueryFunc As LongPtr, ByVal pContext As LongPtr, ByVal xDestructor As LongPtr) As Long
Public Declare Function sqlite3_serialize Lib "winsqlite3" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal piSize As LongPtr, ByVal mFlags As Long) As LongPtr
Public Declare Function sqlite3_set_authorizer Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal xAuth As LongPtr, ByVal pUserData As LongPtr) As Long
Public Declare Sub sqlite3_set_auxdata Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal N As Long, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare Function sqlite3_set_clientdata Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare Function sqlite3_set_errmsg Lib "winsqlite3" (ByVal db As LongPtr, ByVal errcode As Long, ByVal zErrMsg As LongPtr) As Long
Public Declare Sub sqlite3_set_last_insert_rowid Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Currency)
Public Declare Function sqlite3_setlk_timeout Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal ms As Long, ByVal flags As Long) As Long
Public Declare Function sqlite3_shutdown Lib "winsqlite3" () As Long
Public Declare Function sqlite3_sleep Lib "winsqlite3" (ByVal p1 As Long) As Long
Public Declare Sub sqlite3_soft_heap_limit Lib "winsqlite3" (ByVal N As Long)
Public Declare Function sqlite3_soft_heap_limit64 Lib "winsqlite3" (ByVal N As Currency) As Currency
Public Declare Function sqlite3_sourceid Lib "winsqlite3" () As LongPtr
Public Declare Function sqlite3_sql Lib "winsqlite3" (ByVal pStmt As LongPtr) As LongPtr
Public Declare Function sqlite3_status Lib "winsqlite3" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare Function sqlite3_status64 Lib "winsqlite3" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare Function sqlite3_step Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_stmt_busy Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_stmt_explain Lib "winsqlite3" (ByVal pStmt As LongPtr, ByVal eMode As Long) As Long
Public Declare Function sqlite3_stmt_isexplain Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Function sqlite3_stmt_readonly Lib "winsqlite3" (ByVal pStmt As LongPtr) As Long
Public Declare Function sqlite3_stmt_status Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal op As Long, ByVal resetFlg As Long) As Long
Public Declare Sub sqlite3_str_append Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zIn As LongPtr, ByVal N As Long)
Public Declare Sub sqlite3_str_appendall Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zIn As LongPtr)
Public Declare Sub sqlite3_str_appendchar Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal N As Long, ByVal C As Long)
Public Declare Function sqlite3_str_errcode Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_str_finish Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_str_length Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_str_new Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Sub sqlite3_str_reset Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Function sqlite3_str_value Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Sub sqlite3_str_vappendf Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zFormat As LongPtr, ByVal va_list As LongPtr)
Public Declare Function sqlite3_strglob Lib "winsqlite3" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr) As Long
Public Declare Function sqlite3_stricmp Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare Function sqlite3_strlike Lib "winsqlite3" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr, ByVal cEsc As Long) As Long
Public Declare Function sqlite3_strnicmp Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long) As Long
Public Declare Function sqlite3_system_errno Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_table_column_metadata Lib "winsqlite3" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal zTableName As LongPtr, ByVal zColumnName As LongPtr, ByVal pzDataType As LongPtr, ByVal pzCollSeq As LongPtr, ByVal pNotNull As LongPtr, ByVal pPrimaryKey As LongPtr, ByVal pAutoinc As LongPtr) As Long
Public Declare Sub sqlite3_thread_cleanup Lib "winsqlite3" ()
Public Declare Function sqlite3_threadsafe Lib "winsqlite3" () As Long
Public Declare Function sqlite3_total_changes Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_total_changes64 Lib "winsqlite3" (ByVal p1 As LongPtr) As Currency
Public Declare Function sqlite3_trace Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal xTrace As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function sqlite3_trace_v2 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal uMask As Long, ByVal xCallback As LongPtr, ByVal pCtx As LongPtr) As Long
Public Declare Function sqlite3_transfer_bindings Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare Function sqlite3_txn_state Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal zSchema As LongPtr) As Long
Public Declare Function sqlite3_update_hook Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function sqlite3_uri_boolean Lib "winsqlite3" (ByVal z As LongPtr, ByVal zParam As LongPtr, ByVal bDefault As Long) As Long
Public Declare Function sqlite3_uri_int64 Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency) As Currency
Public Declare Function sqlite3_uri_key Lib "winsqlite3" (ByVal z As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function sqlite3_uri_parameter Lib "winsqlite3" (ByVal z As LongPtr, ByVal zParam As LongPtr) As LongPtr
Public Declare Function sqlite3_user_data Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_blob Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_bytes Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_bytes16 Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_double Lib "winsqlite3" (ByVal p1 As LongPtr) As Double
Public Declare Function sqlite3_value_dup Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_encoding Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Sub sqlite3_value_free Lib "winsqlite3" (ByVal p1 As LongPtr)
Public Declare Function sqlite3_value_frombind Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_int Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_int64 Lib "winsqlite3" (ByVal p1 As LongPtr) As Currency
Public Declare Function sqlite3_value_nochange Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_numeric_type Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_pointer Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_subtype Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_value_text Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_text16 Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_text16be Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_text16le Lib "winsqlite3" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function sqlite3_value_type Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_vfs_find Lib "winsqlite3" (ByVal zVfsName As LongPtr) As LongPtr
Public Declare Function sqlite3_vfs_register Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal makeDflt As Long) As Long
Public Declare Function sqlite3_vfs_unregister Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_vmprintf Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare Function sqlite3_vsnprintf Lib "winsqlite3" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare Function sqlite3_vtab_collation Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function sqlite3_vtab_distinct Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_vtab_in Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal iCons As Long, ByVal bHandle As Long) As Long
Public Declare Function sqlite3_vtab_in_first Lib "winsqlite3" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare Function sqlite3_vtab_in_next Lib "winsqlite3" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare Function sqlite3_vtab_nochange Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_vtab_on_conflict Lib "winsqlite3" (ByVal p1 As LongPtr) As Long
Public Declare Function sqlite3_vtab_rhs_value Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal ppVal As LongPtr) As Long
Public Declare Function sqlite3_wal_autocheckpoint Lib "winsqlite3" (ByVal db As LongPtr, ByVal N As Long) As Long
Public Declare Function sqlite3_wal_checkpoint Lib "winsqlite3" (ByVal db As LongPtr, ByVal zDb As LongPtr) As Long
Public Declare Function sqlite3_wal_checkpoint_v2 Lib "winsqlite3" (ByVal db As LongPtr, ByVal zDb As LongPtr, ByVal eMode As Long, ByVal pnLog As LongPtr, ByVal pnCkpt As LongPtr) As Long
Public Declare Function sqlite3_wal_hook Lib "winsqlite3" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function sqlite3_win32_set_directory Lib "winsqlite3" (ByVal vType As Long, ByVal zValue As LongPtr) As LongPtr
Public Declare Function sqlite3_win32_set_directory16 Lib "winsqlite3" (ByVal vType As Long, ByVal zValue As LongPtr) As Long
Public Declare Function sqlite3_win32_set_directory8 Lib "winsqlite3" (ByVal vType As Long, ByVal zValue As LongPtr) As Long

'--- CDecl / C-variadic exports (8) - intentionally NOT declared so the
'--- module stays pure StdCall for VB6 (x86) and tB (x86/x64). Do string
'--- formatting VB-side; these are not needed on the data-access path:
'---   sqlite3_config, sqlite3_db_config, sqlite3_log, sqlite3_mprintf, sqlite3_snprintf, sqlite3_str_appendf, sqlite3_test_control, sqlite3_vtab_config

'--- The following winsqlite3.dll exports are data symbols or internal
'--- helpers with no public prototype in sqlite3.h. Declare when needed:
'---   data:     sqlite3_version (char[]), sqlite3_temp_directory,
'---             sqlite3_data_directory, sqlite3_fts3_may_be_corrupt,
'---             sqlite3_unsupported_selecttrace
'---   internal: sqlite3_win32_is_nt, sqlite3_win32_sleep,
'---             sqlite3_win32_write_debug, sqlite3_win32_mbcs_to_utf8[_v2],
'---             sqlite3_win32_utf8_to_mbcs[_v2], sqlite3_win32_unicode_to_utf8,
'---             sqlite3_win32_utf8_to_unicode
