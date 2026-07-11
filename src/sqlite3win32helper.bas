Attribute VB_Name = "sqlite3win32helper"
'=========================================================================
' sqlite3win32helper - core SQLite constants for the sqlite3win32stubs
' declares. Only a curated set is included; for the full set of result
' codes, flags and options refer to doc\sqlite3.h
'=========================================================================
Option Explicit

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
