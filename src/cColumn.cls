VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cColumn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cColumn - one column's definition, filled by cColumns from PRAGMA
' table_info + sqlite3_table_column_metadata + the unique-index scan. The
' members that exist only in the CREATE TABLE text (constraint names,
' check expressions, conflict algorithms, PK sort order) need DDL parsing
' and are not implemented yet - they return defaults.
'=========================================================================
Option Explicit

Public Enum ConflictAlgorithms
    OnConflict_UseTableDefault = 0
    OnConflict_ABORT = 1
    OnConflict_IGNORE = 2
    OnConflict_FAIL = 3
    OnConflict_ROLLBACK = 4
    OnConflict_REPLACE = 5
End Enum

Private m_sName                     As String
Private m_sType                     As String
Private m_sDefault                  As String
Private m_bNotNull                  As Boolean
Private m_lPkOrdinal                As Long
Private m_bAutoInc                  As Boolean
Private m_sCollate                  As String
Private m_bUnique                   As Boolean

Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Get NameInBrackets() As String
    NameInBrackets = "[" & m_sName & "]"
End Property

Public Property Get ColumnType() As String
    ColumnType = m_sType
End Property

Public Property Get DefaultValue() As String
    DefaultValue = m_sDefault
End Property

Public Property Get OriginalConstraint() As String
End Property

Public Property Get ConstraintName() As String
End Property

Public Property Get CheckExpression() As String
End Property

Public Property Get Collate() As String
    Collate = m_sCollate
End Property

Public Property Get PrimaryKey() As Boolean
    PrimaryKey = (m_lPkOrdinal > 0)
End Property

Public Property Get PrimarySortOrder() As String
End Property

Public Property Get PrimaryConflictAlgorithm() As ConflictAlgorithms
    PrimaryConflictAlgorithm = OnConflict_UseTableDefault
End Property

Public Property Get PrimaryAutoIncrement() As Boolean
    PrimaryAutoIncrement = m_bAutoInc
End Property

Public Property Get NotNullConstraint() As Boolean
    NotNullConstraint = m_bNotNull
End Property

Public Property Get NotNullConflictAlgorithm() As ConflictAlgorithms
    NotNullConflictAlgorithm = OnConflict_UseTableDefault
End Property

Public Property Get UniqueConstraint() As Boolean
    UniqueConstraint = m_bUnique
End Property

Public Property Get UniqueConflictAlgorithm() As ConflictAlgorithms
    UniqueConflictAlgorithm = OnConflict_UseTableDefault
End Property

Friend Sub frInit(sName As String, sType As String, sDefault As String, ByVal bNotNull As Boolean, ByVal lPkOrdinal As Long, ByVal bAutoInc As Boolean, sCollate As String)
    m_sName = sName
    m_sType = sType
    m_sDefault = sDefault
    m_bNotNull = bNotNull
    m_lPkOrdinal = lPkOrdinal
    m_bAutoInc = bAutoInc
    m_sCollate = sCollate
End Sub

Friend Sub frMarkUnique()
    m_bUnique = True
End Sub
