Attribute VB_Name = "sqlite3win32stubs"
'=========================================================================
' sqlite3win32stubs - stub_sqlite3_* declares for the built-in Windows
' SQLite (winsqlite3.dll)
'
' - All declared functions are StdCall, so the module compiles unchanged
'   in VB6 (x86) and twinBASIC (x86/x64). The 8 CDecl/variadic printf-style
'   exports are intentionally NOT declared (see footer).
' - Handles/pointers are LongPtr; int is Long; sqlite3_int64 is LongLong on
'   x64 and Currency on x86/VB6 (both 8-byte) - see #If Win64 blocks below
' - SQLite is UTF-8: pass/receive LongPtr to byte buffers, not VB String
' - The core constants live in sqlite3win32helper.bas; for the full set of
'   result codes, flags and options refer to doc\sqlite3.h
'=========================================================================
Option Explicit

#If Win64 = 0 And TWINBASIC = 0 Then
Public Enum LongPtr
    [_]
End Enum
#End If

'--- winsqlite3.dll exports (275 declares)
Public Declare Function stub_sqlite3_aggregate_context Lib "winsqlite3" Alias "sqlite3_aggregate_context" (ByVal p1 As LongPtr, ByVal nBytes As Long) As LongPtr
Public Declare Function stub_sqlite3_aggregate_count Lib "winsqlite3" Alias "sqlite3_aggregate_count" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_auto_extension Lib "winsqlite3" Alias "sqlite3_auto_extension" (ByVal xEntryPoint As LongPtr) As Long
Public Declare Function stub_sqlite3_autovacuum_pages Lib "winsqlite3" Alias "sqlite3_autovacuum_pages" (ByVal db As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare Function stub_sqlite3_backup_finish Lib "winsqlite3" Alias "sqlite3_backup_finish" (ByVal p As LongPtr) As Long
Public Declare Function stub_sqlite3_backup_init Lib "winsqlite3" Alias "sqlite3_backup_init" (ByVal pDest As LongPtr, ByVal zDestName As LongPtr, ByVal pSource As LongPtr, ByVal zSourceName As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_backup_pagecount Lib "winsqlite3" Alias "sqlite3_backup_pagecount" (ByVal p As LongPtr) As Long
Public Declare Function stub_sqlite3_backup_remaining Lib "winsqlite3" Alias "sqlite3_backup_remaining" (ByVal p As LongPtr) As Long
Public Declare Function stub_sqlite3_backup_step Lib "winsqlite3" Alias "sqlite3_backup_step" (ByVal p As LongPtr, ByVal nPage As Long) As Long
Public Declare Function stub_sqlite3_bind_blob Lib "winsqlite3" Alias "sqlite3_bind_blob" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal n As Long, ByVal cb5 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_bind_blob64 Lib "winsqlite3" Alias "sqlite3_bind_blob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongLong, ByVal cb5 As LongPtr) As Long
#Else
Public Declare Function stub_sqlite3_bind_blob64 Lib "winsqlite3" Alias "sqlite3_bind_blob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Currency, ByVal cb5 As LongPtr) As Long
#End If
Public Declare Function stub_sqlite3_bind_double Lib "winsqlite3" Alias "sqlite3_bind_double" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Double) As Long
Public Declare Function stub_sqlite3_bind_int Lib "winsqlite3" Alias "sqlite3_bind_int" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Long) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_bind_int64 Lib "winsqlite3" Alias "sqlite3_bind_int64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongLong) As Long
#Else
Public Declare Function stub_sqlite3_bind_int64 Lib "winsqlite3" Alias "sqlite3_bind_int64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Currency) As Long
#End If
Public Declare Function stub_sqlite3_bind_null Lib "winsqlite3" Alias "sqlite3_bind_null" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare Function stub_sqlite3_bind_parameter_count Lib "winsqlite3" Alias "sqlite3_bind_parameter_count" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_bind_parameter_index Lib "winsqlite3" Alias "sqlite3_bind_parameter_index" (ByVal p1 As LongPtr, ByVal zName As LongPtr) As Long
Public Declare Function stub_sqlite3_bind_parameter_name Lib "winsqlite3" Alias "sqlite3_bind_parameter_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_bind_pointer Lib "winsqlite3" Alias "sqlite3_bind_pointer" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal cb5 As LongPtr) As Long
Public Declare Function stub_sqlite3_bind_text Lib "winsqlite3" Alias "sqlite3_bind_text" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
Public Declare Function stub_sqlite3_bind_text16 Lib "winsqlite3" Alias "sqlite3_bind_text16" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_bind_text64 Lib "winsqlite3" Alias "sqlite3_bind_text64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongLong, ByVal cb5 As LongPtr, ByVal encoding As Long) As Long
#Else
Public Declare Function stub_sqlite3_bind_text64 Lib "winsqlite3" Alias "sqlite3_bind_text64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Currency, ByVal cb5 As LongPtr, ByVal encoding As Long) As Long
#End If
Public Declare Function stub_sqlite3_bind_value Lib "winsqlite3" Alias "sqlite3_bind_value" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr) As Long
Public Declare Function stub_sqlite3_bind_zeroblob Lib "winsqlite3" Alias "sqlite3_bind_zeroblob" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal n As Long) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_bind_zeroblob64 Lib "winsqlite3" Alias "sqlite3_bind_zeroblob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongLong) As Long
#Else
Public Declare Function stub_sqlite3_bind_zeroblob64 Lib "winsqlite3" Alias "sqlite3_bind_zeroblob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Currency) As Long
#End If
Public Declare Function stub_sqlite3_blob_bytes Lib "winsqlite3" Alias "sqlite3_blob_bytes" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_blob_close Lib "winsqlite3" Alias "sqlite3_blob_close" (ByVal p1 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_blob_open Lib "winsqlite3" Alias "sqlite3_blob_open" (ByVal p1 As LongPtr, ByVal zDb As LongPtr, ByVal zTable As LongPtr, ByVal zColumn As LongPtr, ByVal iRow As LongLong, ByVal flags As Long, ByVal ppBlob As LongPtr) As Long
#Else
Public Declare Function stub_sqlite3_blob_open Lib "winsqlite3" Alias "sqlite3_blob_open" (ByVal p1 As LongPtr, ByVal zDb As LongPtr, ByVal zTable As LongPtr, ByVal zColumn As LongPtr, ByVal iRow As Currency, ByVal flags As Long, ByVal ppBlob As LongPtr) As Long
#End If
Public Declare Function stub_sqlite3_blob_read Lib "winsqlite3" Alias "sqlite3_blob_read" (ByVal p1 As LongPtr, ByVal Z As LongPtr, ByVal N As Long, ByVal iOffset As Long) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_blob_reopen Lib "winsqlite3" Alias "sqlite3_blob_reopen" (ByVal p1 As LongPtr, ByVal p2 As LongLong) As Long
#Else
Public Declare Function stub_sqlite3_blob_reopen Lib "winsqlite3" Alias "sqlite3_blob_reopen" (ByVal p1 As LongPtr, ByVal p2 As Currency) As Long
#End If
Public Declare Function stub_sqlite3_blob_write Lib "winsqlite3" Alias "sqlite3_blob_write" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As Long, ByVal iOffset As Long) As Long
Public Declare Function stub_sqlite3_busy_handler Lib "winsqlite3" Alias "sqlite3_busy_handler" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As Long
Public Declare Function stub_sqlite3_busy_timeout Lib "winsqlite3" Alias "sqlite3_busy_timeout" (ByVal p1 As LongPtr, ByVal ms As Long) As Long
Public Declare Function stub_sqlite3_cancel_auto_extension Lib "winsqlite3" Alias "sqlite3_cancel_auto_extension" (ByVal xEntryPoint As LongPtr) As Long
Public Declare Function stub_sqlite3_changes Lib "winsqlite3" Alias "sqlite3_changes" (ByVal p1 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_changes64 Lib "winsqlite3" Alias "sqlite3_changes64" (ByVal p1 As LongPtr) As LongLong
#Else
Public Declare Function stub_sqlite3_changes64 Lib "winsqlite3" Alias "sqlite3_changes64" (ByVal p1 As LongPtr) As Currency
#End If
Public Declare Function stub_sqlite3_clear_bindings Lib "winsqlite3" Alias "sqlite3_clear_bindings" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_close Lib "winsqlite3" Alias "sqlite3_close" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_close_v2 Lib "winsqlite3" Alias "sqlite3_close_v2" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_collation_needed Lib "winsqlite3" Alias "sqlite3_collation_needed" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare Function stub_sqlite3_collation_needed16 Lib "winsqlite3" Alias "sqlite3_collation_needed16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare Function stub_sqlite3_column_blob Lib "winsqlite3" Alias "sqlite3_column_blob" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function stub_sqlite3_column_bytes Lib "winsqlite3" Alias "sqlite3_column_bytes" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function stub_sqlite3_column_bytes16 Lib "winsqlite3" Alias "sqlite3_column_bytes16" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function stub_sqlite3_column_count Lib "winsqlite3" Alias "sqlite3_column_count" (ByVal pStmt As LongPtr) As Long
Public Declare Function stub_sqlite3_column_database_name Lib "winsqlite3" Alias "sqlite3_column_database_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_database_name16 Lib "winsqlite3" Alias "sqlite3_column_database_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_decltype Lib "winsqlite3" Alias "sqlite3_column_decltype" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_decltype16 Lib "winsqlite3" Alias "sqlite3_column_decltype16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_double Lib "winsqlite3" Alias "sqlite3_column_double" (ByVal p1 As LongPtr, ByVal iCol As Long) As Double
Public Declare Function stub_sqlite3_column_int Lib "winsqlite3" Alias "sqlite3_column_int" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_column_int64 Lib "winsqlite3" Alias "sqlite3_column_int64" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongLong
#Else
Public Declare Function stub_sqlite3_column_int64 Lib "winsqlite3" Alias "sqlite3_column_int64" (ByVal p1 As LongPtr, ByVal iCol As Long) As Currency
#End If
Public Declare Function stub_sqlite3_column_name Lib "winsqlite3" Alias "sqlite3_column_name" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_column_name16 Lib "winsqlite3" Alias "sqlite3_column_name16" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_column_origin_name Lib "winsqlite3" Alias "sqlite3_column_origin_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_origin_name16 Lib "winsqlite3" Alias "sqlite3_column_origin_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_table_name Lib "winsqlite3" Alias "sqlite3_column_table_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_table_name16 Lib "winsqlite3" Alias "sqlite3_column_table_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_column_text Lib "winsqlite3" Alias "sqlite3_column_text" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function stub_sqlite3_column_text16 Lib "winsqlite3" Alias "sqlite3_column_text16" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function stub_sqlite3_column_type Lib "winsqlite3" Alias "sqlite3_column_type" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare Function stub_sqlite3_column_value Lib "winsqlite3" Alias "sqlite3_column_value" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare Function stub_sqlite3_commit_hook Lib "winsqlite3" Alias "sqlite3_commit_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_compileoption_get Lib "winsqlite3" Alias "sqlite3_compileoption_get" (ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_compileoption_used Lib "winsqlite3" Alias "sqlite3_compileoption_used" (ByVal zOptName As LongPtr) As Long
Public Declare Function stub_sqlite3_complete Lib "winsqlite3" Alias "sqlite3_complete" (ByVal sql As LongPtr) As Long
Public Declare Function stub_sqlite3_complete16 Lib "winsqlite3" Alias "sqlite3_complete16" (ByVal sql As LongPtr) As Long
Public Declare Function stub_sqlite3_context_db_handle Lib "winsqlite3" Alias "sqlite3_context_db_handle" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_create_collation Lib "winsqlite3" Alias "sqlite3_create_collation" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare Function stub_sqlite3_create_collation_v2 Lib "winsqlite3" Alias "sqlite3_create_collation_v2" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function stub_sqlite3_create_collation16 Lib "winsqlite3" Alias "sqlite3_create_collation16" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare Function stub_sqlite3_create_filename Lib "winsqlite3" Alias "sqlite3_create_filename" (ByVal zDatabase As LongPtr, ByVal zJournal As LongPtr, ByVal zWal As LongPtr, ByVal nParam As Long, ByVal azParam As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_create_function Lib "winsqlite3" Alias "sqlite3_create_function" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare Function stub_sqlite3_create_function_v2 Lib "winsqlite3" Alias "sqlite3_create_function_v2" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function stub_sqlite3_create_function16 Lib "winsqlite3" Alias "sqlite3_create_function16" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare Function stub_sqlite3_create_module Lib "winsqlite3" Alias "sqlite3_create_module" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr) As Long
Public Declare Function stub_sqlite3_create_module_v2 Lib "winsqlite3" Alias "sqlite3_create_module_v2" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function stub_sqlite3_create_window_function Lib "winsqlite3" Alias "sqlite3_create_window_function" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xValue As LongPtr, ByVal xInverse As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare Function stub_sqlite3_data_count Lib "winsqlite3" Alias "sqlite3_data_count" (ByVal pStmt As LongPtr) As Long
Public Declare Function stub_sqlite3_database_file_object Lib "winsqlite3" Alias "sqlite3_database_file_object" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_db_cacheflush Lib "winsqlite3" Alias "sqlite3_db_cacheflush" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_db_filename Lib "winsqlite3" Alias "sqlite3_db_filename" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_db_handle Lib "winsqlite3" Alias "sqlite3_db_handle" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_db_mutex Lib "winsqlite3" Alias "sqlite3_db_mutex" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_db_name Lib "winsqlite3" Alias "sqlite3_db_name" (ByVal db As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_db_readonly Lib "winsqlite3" Alias "sqlite3_db_readonly" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As Long
Public Declare Function stub_sqlite3_db_release_memory Lib "winsqlite3" Alias "sqlite3_db_release_memory" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_db_status Lib "winsqlite3" Alias "sqlite3_db_status" (ByVal p1 As LongPtr, ByVal op As Long, ByVal pCur As LongPtr, ByVal pHiwtr As LongPtr, ByVal resetFlg As Long) As Long
Public Declare Function stub_sqlite3_db_status64 Lib "winsqlite3" Alias "sqlite3_db_status64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal p5 As Long) As Long
Public Declare Function stub_sqlite3_declare_vtab Lib "winsqlite3" Alias "sqlite3_declare_vtab" (ByVal p1 As LongPtr, ByVal zSQL As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_deserialize Lib "winsqlite3" Alias "sqlite3_deserialize" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal pData As LongPtr, ByVal szDb As LongLong, ByVal szBuf As LongLong, ByVal mFlags As Long) As Long
#Else
Public Declare Function stub_sqlite3_deserialize Lib "winsqlite3" Alias "sqlite3_deserialize" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal pData As LongPtr, ByVal szDb As Currency, ByVal szBuf As Currency, ByVal mFlags As Long) As Long
#End If
Public Declare Function stub_sqlite3_drop_modules Lib "winsqlite3" Alias "sqlite3_drop_modules" (ByVal db As LongPtr, ByVal azKeep As LongPtr) As Long
Public Declare Function stub_sqlite3_enable_load_extension Lib "winsqlite3" Alias "sqlite3_enable_load_extension" (ByVal db As LongPtr, ByVal onoff As Long) As Long
Public Declare Function stub_sqlite3_enable_shared_cache Lib "winsqlite3" Alias "sqlite3_enable_shared_cache" (ByVal p1 As Long) As Long
Public Declare Function stub_sqlite3_errcode Lib "winsqlite3" Alias "sqlite3_errcode" (ByVal db As LongPtr) As Long
Public Declare Function stub_sqlite3_errmsg Lib "winsqlite3" Alias "sqlite3_errmsg" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_errmsg16 Lib "winsqlite3" Alias "sqlite3_errmsg16" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_error_offset Lib "winsqlite3" Alias "sqlite3_error_offset" (ByVal db As LongPtr) As Long
Public Declare Function stub_sqlite3_errstr Lib "winsqlite3" Alias "sqlite3_errstr" (ByVal p1 As Long) As LongPtr
Public Declare Function stub_sqlite3_exec Lib "winsqlite3" Alias "sqlite3_exec" (ByVal p1 As LongPtr, ByVal sql As LongPtr, ByVal callback As LongPtr, ByVal p4 As LongPtr, ByVal errmsg As LongPtr) As Long
Public Declare Function stub_sqlite3_expanded_sql Lib "winsqlite3" Alias "sqlite3_expanded_sql" (ByVal pStmt As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_expired Lib "winsqlite3" Alias "sqlite3_expired" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_extended_errcode Lib "winsqlite3" Alias "sqlite3_extended_errcode" (ByVal db As LongPtr) As Long
Public Declare Function stub_sqlite3_extended_result_codes Lib "winsqlite3" Alias "sqlite3_extended_result_codes" (ByVal p1 As LongPtr, ByVal onoff As Long) As Long
Public Declare Function stub_sqlite3_file_control Lib "winsqlite3" Alias "sqlite3_file_control" (ByVal p1 As LongPtr, ByVal zDbName As LongPtr, ByVal op As Long, ByVal p4 As LongPtr) As Long
Public Declare Function stub_sqlite3_filename_database Lib "winsqlite3" Alias "sqlite3_filename_database" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_filename_journal Lib "winsqlite3" Alias "sqlite3_filename_journal" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_filename_wal Lib "winsqlite3" Alias "sqlite3_filename_wal" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_finalize Lib "winsqlite3" Alias "sqlite3_finalize" (ByVal pStmt As LongPtr) As Long
Public Declare Sub stub_sqlite3_free Lib "winsqlite3" Alias "sqlite3_free" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_free_filename Lib "winsqlite3" Alias "sqlite3_free_filename" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_free_table Lib "winsqlite3" Alias "sqlite3_free_table" (ByVal result As LongPtr)
Public Declare Function stub_sqlite3_get_autocommit Lib "winsqlite3" Alias "sqlite3_get_autocommit" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_get_auxdata Lib "winsqlite3" Alias "sqlite3_get_auxdata" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_get_clientdata Lib "winsqlite3" Alias "sqlite3_get_clientdata" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_get_table Lib "winsqlite3" Alias "sqlite3_get_table" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal pazResult As LongPtr, ByVal pnRow As LongPtr, ByVal pnColumn As LongPtr, ByVal pzErrmsg As LongPtr) As Long
Public Declare Function stub_sqlite3_global_recover Lib "winsqlite3" Alias "sqlite3_global_recover" () As Long
#If Win64 Then
Public Declare Function stub_sqlite3_hard_heap_limit64 Lib "winsqlite3" Alias "sqlite3_hard_heap_limit64" (ByVal N As LongLong) As LongLong
#Else
Public Declare Function stub_sqlite3_hard_heap_limit64 Lib "winsqlite3" Alias "sqlite3_hard_heap_limit64" (ByVal N As Currency) As Currency
#End If
Public Declare Function stub_sqlite3_initialize Lib "winsqlite3" Alias "sqlite3_initialize" () As Long
Public Declare Sub stub_sqlite3_interrupt Lib "winsqlite3" Alias "sqlite3_interrupt" (ByVal p1 As LongPtr)
Public Declare Function stub_sqlite3_is_interrupted Lib "winsqlite3" Alias "sqlite3_is_interrupted" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_keyword_check Lib "winsqlite3" Alias "sqlite3_keyword_check" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare Function stub_sqlite3_keyword_count Lib "winsqlite3" Alias "sqlite3_keyword_count" () As Long
Public Declare Function stub_sqlite3_keyword_name Lib "winsqlite3" Alias "sqlite3_keyword_name" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_last_insert_rowid Lib "winsqlite3" Alias "sqlite3_last_insert_rowid" (ByVal p1 As LongPtr) As LongLong
#Else
Public Declare Function stub_sqlite3_last_insert_rowid Lib "winsqlite3" Alias "sqlite3_last_insert_rowid" (ByVal p1 As LongPtr) As Currency
#End If
Public Declare Function stub_sqlite3_libversion Lib "winsqlite3" Alias "sqlite3_libversion" () As LongPtr
Public Declare Function stub_sqlite3_libversion_number Lib "winsqlite3" Alias "sqlite3_libversion_number" () As Long
Public Declare Function stub_sqlite3_limit Lib "winsqlite3" Alias "sqlite3_limit" (ByVal p1 As LongPtr, ByVal id As Long, ByVal newVal As Long) As Long
Public Declare Function stub_sqlite3_load_extension Lib "winsqlite3" Alias "sqlite3_load_extension" (ByVal db As LongPtr, ByVal zFile As LongPtr, ByVal zProc As LongPtr, ByVal pzErrMsg As LongPtr) As Long
Public Declare Function stub_sqlite3_malloc Lib "winsqlite3" Alias "sqlite3_malloc" (ByVal p1 As Long) As LongPtr
#If Win64 Then
Public Declare Function stub_sqlite3_malloc64 Lib "winsqlite3" Alias "sqlite3_malloc64" (ByVal p1 As LongLong) As LongPtr
#Else
Public Declare Function stub_sqlite3_malloc64 Lib "winsqlite3" Alias "sqlite3_malloc64" (ByVal p1 As Currency) As LongPtr
#End If
#If Win64 Then
Public Declare Function stub_sqlite3_memory_alarm Lib "winsqlite3" Alias "sqlite3_memory_alarm" (ByVal cb1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong) As Long
#Else
Public Declare Function stub_sqlite3_memory_alarm Lib "winsqlite3" Alias "sqlite3_memory_alarm" (ByVal cb1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency) As Long
#End If
#If Win64 Then
Public Declare Function stub_sqlite3_memory_highwater Lib "winsqlite3" Alias "sqlite3_memory_highwater" (ByVal resetFlag As Long) As LongLong
#Else
Public Declare Function stub_sqlite3_memory_highwater Lib "winsqlite3" Alias "sqlite3_memory_highwater" (ByVal resetFlag As Long) As Currency
#End If
#If Win64 Then
Public Declare Function stub_sqlite3_memory_used Lib "winsqlite3" Alias "sqlite3_memory_used" () As LongLong
#Else
Public Declare Function stub_sqlite3_memory_used Lib "winsqlite3" Alias "sqlite3_memory_used" () As Currency
#End If
#If Win64 Then
Public Declare Function stub_sqlite3_msize Lib "winsqlite3" Alias "sqlite3_msize" (ByVal p1 As LongPtr) As LongLong
#Else
Public Declare Function stub_sqlite3_msize Lib "winsqlite3" Alias "sqlite3_msize" (ByVal p1 As LongPtr) As Currency
#End If
Public Declare Function stub_sqlite3_mutex_alloc Lib "winsqlite3" Alias "sqlite3_mutex_alloc" (ByVal p1 As Long) As LongPtr
Public Declare Sub stub_sqlite3_mutex_enter Lib "winsqlite3" Alias "sqlite3_mutex_enter" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_mutex_free Lib "winsqlite3" Alias "sqlite3_mutex_free" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_mutex_leave Lib "winsqlite3" Alias "sqlite3_mutex_leave" (ByVal p1 As LongPtr)
Public Declare Function stub_sqlite3_mutex_try Lib "winsqlite3" Alias "sqlite3_mutex_try" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_next_stmt Lib "winsqlite3" Alias "sqlite3_next_stmt" (ByVal pDb As LongPtr, ByVal pStmt As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_open Lib "winsqlite3" Alias "sqlite3_open" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare Function stub_sqlite3_open_v2 Lib "winsqlite3" Alias "sqlite3_open_v2" (ByVal filename As LongPtr, ByVal ppDb As LongPtr, ByVal flags As Long, ByVal zVfs As LongPtr) As Long
Public Declare Function stub_sqlite3_open16 Lib "winsqlite3" Alias "sqlite3_open16" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare Function stub_sqlite3_os_end Lib "winsqlite3" Alias "sqlite3_os_end" () As Long
Public Declare Function stub_sqlite3_os_init Lib "winsqlite3" Alias "sqlite3_os_init" () As Long
Public Declare Function stub_sqlite3_overload_function Lib "winsqlite3" Alias "sqlite3_overload_function" (ByVal p1 As LongPtr, ByVal zFuncName As LongPtr, ByVal nArg As Long) As Long
Public Declare Function stub_sqlite3_prepare Lib "winsqlite3" Alias "sqlite3_prepare" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_prepare_v2 Lib "winsqlite3" Alias "sqlite3_prepare_v2" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_prepare_v3 Lib "winsqlite3" Alias "sqlite3_prepare_v3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_prepare16 Lib "winsqlite3" Alias "sqlite3_prepare16" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_prepare16_v2 Lib "winsqlite3" Alias "sqlite3_prepare16_v2" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_prepare16_v3 Lib "winsqlite3" Alias "sqlite3_prepare16_v3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare Function stub_sqlite3_profile Lib "winsqlite3" Alias "sqlite3_profile" (ByVal p1 As LongPtr, ByVal xProfile As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Sub stub_sqlite3_progress_handler Lib "winsqlite3" Alias "sqlite3_progress_handler" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal cb3 As LongPtr, ByVal p4 As LongPtr)
Public Declare Sub stub_sqlite3_randomness Lib "winsqlite3" Alias "sqlite3_randomness" (ByVal N As Long, ByVal P As LongPtr)
Public Declare Function stub_sqlite3_realloc Lib "winsqlite3" Alias "sqlite3_realloc" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
#If Win64 Then
Public Declare Function stub_sqlite3_realloc64 Lib "winsqlite3" Alias "sqlite3_realloc64" (ByVal p1 As LongPtr, ByVal p2 As LongLong) As LongPtr
#Else
Public Declare Function stub_sqlite3_realloc64 Lib "winsqlite3" Alias "sqlite3_realloc64" (ByVal p1 As LongPtr, ByVal p2 As Currency) As LongPtr
#End If
Public Declare Function stub_sqlite3_release_memory Lib "winsqlite3" Alias "sqlite3_release_memory" (ByVal p1 As Long) As Long
Public Declare Function stub_sqlite3_reset Lib "winsqlite3" Alias "sqlite3_reset" (ByVal pStmt As LongPtr) As Long
Public Declare Sub stub_sqlite3_reset_auto_extension Lib "winsqlite3" Alias "sqlite3_reset_auto_extension" ()
Public Declare Sub stub_sqlite3_result_blob Lib "winsqlite3" Alias "sqlite3_result_blob" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
#If Win64 Then
Public Declare Sub stub_sqlite3_result_blob64 Lib "winsqlite3" Alias "sqlite3_result_blob64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong, ByVal cb4 As LongPtr)
#Else
Public Declare Sub stub_sqlite3_result_blob64 Lib "winsqlite3" Alias "sqlite3_result_blob64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency, ByVal cb4 As LongPtr)
#End If
Public Declare Sub stub_sqlite3_result_double Lib "winsqlite3" Alias "sqlite3_result_double" (ByVal p1 As LongPtr, ByVal p2 As Double)
Public Declare Sub stub_sqlite3_result_error Lib "winsqlite3" Alias "sqlite3_result_error" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare Sub stub_sqlite3_result_error_code Lib "winsqlite3" Alias "sqlite3_result_error_code" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare Sub stub_sqlite3_result_error_nomem Lib "winsqlite3" Alias "sqlite3_result_error_nomem" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_result_error_toobig Lib "winsqlite3" Alias "sqlite3_result_error_toobig" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_result_error16 Lib "winsqlite3" Alias "sqlite3_result_error16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare Sub stub_sqlite3_result_int Lib "winsqlite3" Alias "sqlite3_result_int" (ByVal p1 As LongPtr, ByVal p2 As Long)
#If Win64 Then
Public Declare Sub stub_sqlite3_result_int64 Lib "winsqlite3" Alias "sqlite3_result_int64" (ByVal p1 As LongPtr, ByVal p2 As LongLong)
#Else
Public Declare Sub stub_sqlite3_result_int64 Lib "winsqlite3" Alias "sqlite3_result_int64" (ByVal p1 As LongPtr, ByVal p2 As Currency)
#End If
Public Declare Sub stub_sqlite3_result_null Lib "winsqlite3" Alias "sqlite3_result_null" (ByVal p1 As LongPtr)
Public Declare Sub stub_sqlite3_result_pointer Lib "winsqlite3" Alias "sqlite3_result_pointer" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare Sub stub_sqlite3_result_subtype Lib "winsqlite3" Alias "sqlite3_result_subtype" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare Sub stub_sqlite3_result_text Lib "winsqlite3" Alias "sqlite3_result_text" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub stub_sqlite3_result_text16 Lib "winsqlite3" Alias "sqlite3_result_text16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub stub_sqlite3_result_text16be Lib "winsqlite3" Alias "sqlite3_result_text16be" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare Sub stub_sqlite3_result_text16le Lib "winsqlite3" Alias "sqlite3_result_text16le" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
#If Win64 Then
Public Declare Sub stub_sqlite3_result_text64 Lib "winsqlite3" Alias "sqlite3_result_text64" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As LongLong, ByVal cb4 As LongPtr, ByVal encoding As Long)
#Else
Public Declare Sub stub_sqlite3_result_text64 Lib "winsqlite3" Alias "sqlite3_result_text64" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As Currency, ByVal cb4 As LongPtr, ByVal encoding As Long)
#End If
Public Declare Sub stub_sqlite3_result_value Lib "winsqlite3" Alias "sqlite3_result_value" (ByVal p1 As LongPtr, ByVal p2 As LongPtr)
Public Declare Sub stub_sqlite3_result_zeroblob Lib "winsqlite3" Alias "sqlite3_result_zeroblob" (ByVal p1 As LongPtr, ByVal n As Long)
#If Win64 Then
Public Declare Function stub_sqlite3_result_zeroblob64 Lib "winsqlite3" Alias "sqlite3_result_zeroblob64" (ByVal p1 As LongPtr, ByVal n As LongLong) As Long
#Else
Public Declare Function stub_sqlite3_result_zeroblob64 Lib "winsqlite3" Alias "sqlite3_result_zeroblob64" (ByVal p1 As LongPtr, ByVal n As Currency) As Long
#End If
Public Declare Function stub_sqlite3_rollback_hook Lib "winsqlite3" Alias "sqlite3_rollback_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_rtree_geometry_callback Lib "winsqlite3" Alias "sqlite3_rtree_geometry_callback" (ByVal db As LongPtr, ByVal zGeom As LongPtr, ByVal xGeom As LongPtr, ByVal pContext As LongPtr) As Long
Public Declare Function stub_sqlite3_rtree_query_callback Lib "winsqlite3" Alias "sqlite3_rtree_query_callback" (ByVal db As LongPtr, ByVal zQueryFunc As LongPtr, ByVal xQueryFunc As LongPtr, ByVal pContext As LongPtr, ByVal xDestructor As LongPtr) As Long
Public Declare Function stub_sqlite3_serialize Lib "winsqlite3" Alias "sqlite3_serialize" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal piSize As LongPtr, ByVal mFlags As Long) As LongPtr
Public Declare Function stub_sqlite3_set_authorizer Lib "winsqlite3" Alias "sqlite3_set_authorizer" (ByVal p1 As LongPtr, ByVal xAuth As LongPtr, ByVal pUserData As LongPtr) As Long
Public Declare Sub stub_sqlite3_set_auxdata Lib "winsqlite3" Alias "sqlite3_set_auxdata" (ByVal p1 As LongPtr, ByVal N As Long, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare Function stub_sqlite3_set_clientdata Lib "winsqlite3" Alias "sqlite3_set_clientdata" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare Function stub_sqlite3_set_errmsg Lib "winsqlite3" Alias "sqlite3_set_errmsg" (ByVal db As LongPtr, ByVal errcode As Long, ByVal zErrMsg As LongPtr) As Long
#If Win64 Then
Public Declare Sub stub_sqlite3_set_last_insert_rowid Lib "winsqlite3" Alias "sqlite3_set_last_insert_rowid" (ByVal p1 As LongPtr, ByVal p2 As LongLong)
#Else
Public Declare Sub stub_sqlite3_set_last_insert_rowid Lib "winsqlite3" Alias "sqlite3_set_last_insert_rowid" (ByVal p1 As LongPtr, ByVal p2 As Currency)
#End If
Public Declare Function stub_sqlite3_setlk_timeout Lib "winsqlite3" Alias "sqlite3_setlk_timeout" (ByVal p1 As LongPtr, ByVal ms As Long, ByVal flags As Long) As Long
Public Declare Function stub_sqlite3_shutdown Lib "winsqlite3" Alias "sqlite3_shutdown" () As Long
Public Declare Function stub_sqlite3_sleep Lib "winsqlite3" Alias "sqlite3_sleep" (ByVal p1 As Long) As Long
Public Declare Sub stub_sqlite3_soft_heap_limit Lib "winsqlite3" Alias "sqlite3_soft_heap_limit" (ByVal N As Long)
#If Win64 Then
Public Declare Function stub_sqlite3_soft_heap_limit64 Lib "winsqlite3" Alias "sqlite3_soft_heap_limit64" (ByVal N As LongLong) As LongLong
#Else
Public Declare Function stub_sqlite3_soft_heap_limit64 Lib "winsqlite3" Alias "sqlite3_soft_heap_limit64" (ByVal N As Currency) As Currency
#End If
Public Declare Function stub_sqlite3_sourceid Lib "winsqlite3" Alias "sqlite3_sourceid" () As LongPtr
Public Declare Function stub_sqlite3_sql Lib "winsqlite3" Alias "sqlite3_sql" (ByVal pStmt As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_status Lib "winsqlite3" Alias "sqlite3_status" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare Function stub_sqlite3_status64 Lib "winsqlite3" Alias "sqlite3_status64" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare Function stub_sqlite3_step Lib "winsqlite3" Alias "sqlite3_step" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_stmt_busy Lib "winsqlite3" Alias "sqlite3_stmt_busy" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_stmt_explain Lib "winsqlite3" Alias "sqlite3_stmt_explain" (ByVal pStmt As LongPtr, ByVal eMode As Long) As Long
Public Declare Function stub_sqlite3_stmt_isexplain Lib "winsqlite3" Alias "sqlite3_stmt_isexplain" (ByVal pStmt As LongPtr) As Long
Public Declare Function stub_sqlite3_stmt_readonly Lib "winsqlite3" Alias "sqlite3_stmt_readonly" (ByVal pStmt As LongPtr) As Long
Public Declare Function stub_sqlite3_stmt_status Lib "winsqlite3" Alias "sqlite3_stmt_status" (ByVal p1 As LongPtr, ByVal op As Long, ByVal resetFlg As Long) As Long
Public Declare Sub stub_sqlite3_str_append Lib "winsqlite3" Alias "sqlite3_str_append" (ByVal p1 As LongPtr, ByVal zIn As LongPtr, ByVal N As Long)
Public Declare Sub stub_sqlite3_str_appendall Lib "winsqlite3" Alias "sqlite3_str_appendall" (ByVal p1 As LongPtr, ByVal zIn As LongPtr)
Public Declare Sub stub_sqlite3_str_appendchar Lib "winsqlite3" Alias "sqlite3_str_appendchar" (ByVal p1 As LongPtr, ByVal N As Long, ByVal C As Long)
Public Declare Function stub_sqlite3_str_errcode Lib "winsqlite3" Alias "sqlite3_str_errcode" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_str_finish Lib "winsqlite3" Alias "sqlite3_str_finish" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_str_length Lib "winsqlite3" Alias "sqlite3_str_length" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_str_new Lib "winsqlite3" Alias "sqlite3_str_new" (ByVal p1 As LongPtr) As LongPtr
Public Declare Sub stub_sqlite3_str_reset Lib "winsqlite3" Alias "sqlite3_str_reset" (ByVal p1 As LongPtr)
Public Declare Function stub_sqlite3_str_value Lib "winsqlite3" Alias "sqlite3_str_value" (ByVal p1 As LongPtr) As LongPtr
Public Declare Sub stub_sqlite3_str_vappendf Lib "winsqlite3" Alias "sqlite3_str_vappendf" (ByVal p1 As LongPtr, ByVal zFormat As LongPtr, ByVal va_list As LongPtr)
Public Declare Function stub_sqlite3_strglob Lib "winsqlite3" Alias "sqlite3_strglob" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr) As Long
Public Declare Function stub_sqlite3_stricmp Lib "winsqlite3" Alias "sqlite3_stricmp" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare Function stub_sqlite3_strlike Lib "winsqlite3" Alias "sqlite3_strlike" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr, ByVal cEsc As Long) As Long
Public Declare Function stub_sqlite3_strnicmp Lib "winsqlite3" Alias "sqlite3_strnicmp" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long) As Long
Public Declare Function stub_sqlite3_system_errno Lib "winsqlite3" Alias "sqlite3_system_errno" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_table_column_metadata Lib "winsqlite3" Alias "sqlite3_table_column_metadata" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal zTableName As LongPtr, ByVal zColumnName As LongPtr, ByVal pzDataType As LongPtr, ByVal pzCollSeq As LongPtr, ByVal pNotNull As LongPtr, ByVal pPrimaryKey As LongPtr, ByVal pAutoinc As LongPtr) As Long
Public Declare Sub stub_sqlite3_thread_cleanup Lib "winsqlite3" Alias "sqlite3_thread_cleanup" ()
Public Declare Function stub_sqlite3_threadsafe Lib "winsqlite3" Alias "sqlite3_threadsafe" () As Long
Public Declare Function stub_sqlite3_total_changes Lib "winsqlite3" Alias "sqlite3_total_changes" (ByVal p1 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_total_changes64 Lib "winsqlite3" Alias "sqlite3_total_changes64" (ByVal p1 As LongPtr) As LongLong
#Else
Public Declare Function stub_sqlite3_total_changes64 Lib "winsqlite3" Alias "sqlite3_total_changes64" (ByVal p1 As LongPtr) As Currency
#End If
Public Declare Function stub_sqlite3_trace Lib "winsqlite3" Alias "sqlite3_trace" (ByVal p1 As LongPtr, ByVal xTrace As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_trace_v2 Lib "winsqlite3" Alias "sqlite3_trace_v2" (ByVal p1 As LongPtr, ByVal uMask As Long, ByVal xCallback As LongPtr, ByVal pCtx As LongPtr) As Long
Public Declare Function stub_sqlite3_transfer_bindings Lib "winsqlite3" Alias "sqlite3_transfer_bindings" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare Function stub_sqlite3_txn_state Lib "winsqlite3" Alias "sqlite3_txn_state" (ByVal p1 As LongPtr, ByVal zSchema As LongPtr) As Long
Public Declare Function stub_sqlite3_update_hook Lib "winsqlite3" Alias "sqlite3_update_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_uri_boolean Lib "winsqlite3" Alias "sqlite3_uri_boolean" (ByVal z As LongPtr, ByVal zParam As LongPtr, ByVal bDefault As Long) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_uri_int64 Lib "winsqlite3" Alias "sqlite3_uri_int64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong) As LongLong
#Else
Public Declare Function stub_sqlite3_uri_int64 Lib "winsqlite3" Alias "sqlite3_uri_int64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Currency) As Currency
#End If
Public Declare Function stub_sqlite3_uri_key Lib "winsqlite3" Alias "sqlite3_uri_key" (ByVal z As LongPtr, ByVal N As Long) As LongPtr
Public Declare Function stub_sqlite3_uri_parameter Lib "winsqlite3" Alias "sqlite3_uri_parameter" (ByVal z As LongPtr, ByVal zParam As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_user_data Lib "winsqlite3" Alias "sqlite3_user_data" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_blob Lib "winsqlite3" Alias "sqlite3_value_blob" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_bytes Lib "winsqlite3" Alias "sqlite3_value_bytes" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_bytes16 Lib "winsqlite3" Alias "sqlite3_value_bytes16" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_double Lib "winsqlite3" Alias "sqlite3_value_double" (ByVal p1 As LongPtr) As Double
Public Declare Function stub_sqlite3_value_dup Lib "winsqlite3" Alias "sqlite3_value_dup" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_encoding Lib "winsqlite3" Alias "sqlite3_value_encoding" (ByVal p1 As LongPtr) As Long
Public Declare Sub stub_sqlite3_value_free Lib "winsqlite3" Alias "sqlite3_value_free" (ByVal p1 As LongPtr)
Public Declare Function stub_sqlite3_value_frombind Lib "winsqlite3" Alias "sqlite3_value_frombind" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_int Lib "winsqlite3" Alias "sqlite3_value_int" (ByVal p1 As LongPtr) As Long
#If Win64 Then
Public Declare Function stub_sqlite3_value_int64 Lib "winsqlite3" Alias "sqlite3_value_int64" (ByVal p1 As LongPtr) As LongLong
#Else
Public Declare Function stub_sqlite3_value_int64 Lib "winsqlite3" Alias "sqlite3_value_int64" (ByVal p1 As LongPtr) As Currency
#End If
Public Declare Function stub_sqlite3_value_nochange Lib "winsqlite3" Alias "sqlite3_value_nochange" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_numeric_type Lib "winsqlite3" Alias "sqlite3_value_numeric_type" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_pointer Lib "winsqlite3" Alias "sqlite3_value_pointer" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_subtype Lib "winsqlite3" Alias "sqlite3_value_subtype" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_value_text Lib "winsqlite3" Alias "sqlite3_value_text" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_text16 Lib "winsqlite3" Alias "sqlite3_value_text16" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_text16be Lib "winsqlite3" Alias "sqlite3_value_text16be" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_text16le Lib "winsqlite3" Alias "sqlite3_value_text16le" (ByVal p1 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_value_type Lib "winsqlite3" Alias "sqlite3_value_type" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_vfs_find Lib "winsqlite3" Alias "sqlite3_vfs_find" (ByVal zVfsName As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_vfs_register Lib "winsqlite3" Alias "sqlite3_vfs_register" (ByVal p1 As LongPtr, ByVal makeDflt As Long) As Long
Public Declare Function stub_sqlite3_vfs_unregister Lib "winsqlite3" Alias "sqlite3_vfs_unregister" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_vmprintf Lib "winsqlite3" Alias "sqlite3_vmprintf" (ByVal p1 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_vsnprintf Lib "winsqlite3" Alias "sqlite3_vsnprintf" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_vtab_collation Lib "winsqlite3" Alias "sqlite3_vtab_collation" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare Function stub_sqlite3_vtab_distinct Lib "winsqlite3" Alias "sqlite3_vtab_distinct" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_vtab_in Lib "winsqlite3" Alias "sqlite3_vtab_in" (ByVal p1 As LongPtr, ByVal iCons As Long, ByVal bHandle As Long) As Long
Public Declare Function stub_sqlite3_vtab_in_first Lib "winsqlite3" Alias "sqlite3_vtab_in_first" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare Function stub_sqlite3_vtab_in_next Lib "winsqlite3" Alias "sqlite3_vtab_in_next" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare Function stub_sqlite3_vtab_nochange Lib "winsqlite3" Alias "sqlite3_vtab_nochange" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_vtab_on_conflict Lib "winsqlite3" Alias "sqlite3_vtab_on_conflict" (ByVal p1 As LongPtr) As Long
Public Declare Function stub_sqlite3_vtab_rhs_value Lib "winsqlite3" Alias "sqlite3_vtab_rhs_value" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal ppVal As LongPtr) As Long
Public Declare Function stub_sqlite3_wal_autocheckpoint Lib "winsqlite3" Alias "sqlite3_wal_autocheckpoint" (ByVal db As LongPtr, ByVal N As Long) As Long
Public Declare Function stub_sqlite3_wal_checkpoint Lib "winsqlite3" Alias "sqlite3_wal_checkpoint" (ByVal db As LongPtr, ByVal zDb As LongPtr) As Long
Public Declare Function stub_sqlite3_wal_checkpoint_v2 Lib "winsqlite3" Alias "sqlite3_wal_checkpoint_v2" (ByVal db As LongPtr, ByVal zDb As LongPtr, ByVal eMode As Long, ByVal pnLog As LongPtr, ByVal pnCkpt As LongPtr) As Long
Public Declare Function stub_sqlite3_wal_hook Lib "winsqlite3" Alias "sqlite3_wal_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_win32_set_directory Lib "winsqlite3" Alias "sqlite3_win32_set_directory" (ByVal vType As Long, ByVal zValue As LongPtr) As LongPtr
Public Declare Function stub_sqlite3_win32_set_directory16 Lib "winsqlite3" Alias "sqlite3_win32_set_directory16" (ByVal vType As Long, ByVal zValue As LongPtr) As Long
Public Declare Function stub_sqlite3_win32_set_directory8 Lib "winsqlite3" Alias "sqlite3_win32_set_directory8" (ByVal vType As Long, ByVal zValue As LongPtr) As Long

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
