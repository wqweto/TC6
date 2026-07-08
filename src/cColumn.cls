VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
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
Option Explicit

Public Enum ConflictAlgorithms
    OnConflict_UseTableDefault = 0
    OnConflict_ABORT = 1
    OnConflict_IGNORE = 2
    OnConflict_FAIL = 3
    OnConflict_ROLLBACK = 4
    OnConflict_REPLACE = 5
End Enum

Public Property Get Name() As String
End Property

Public Property Get NameInBrackets() As String
End Property

Public Property Get ColumnType() As String
End Property

Public Property Get DefaultValue() As String
End Property

Public Property Get OriginalConstraint() As String
End Property

Public Property Get ConstraintName() As String
End Property

Public Property Get CheckExpression() As String
End Property

Public Property Get Collate() As String
End Property

Public Property Get PrimaryKey() As Boolean
End Property

Public Property Get PrimarySortOrder() As String
End Property

Public Property Get PrimaryConflictAlgorithm() As ConflictAlgorithms
End Property

Public Property Get PrimaryAutoIncrement() As Boolean
End Property

Public Property Get NotNullConstraint() As Boolean
End Property

Public Property Get NotNullConflictAlgorithm() As ConflictAlgorithms
End Property

Public Property Get UniqueConstraint() As Boolean
End Property

Public Property Get UniqueConflictAlgorithm() As ConflictAlgorithms
End Property
