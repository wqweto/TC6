VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum FieldType
    SQLite_INTEGER = 1
    SQLite_DOUBLE = 2
    SQLite_TEXT = 3
    SQLite_BLOB = 4
    SQLite_NULL = 5
    VB_DATE_AutoConverted = 6
    VB_Boolean_AutoConverted = 7
    VB_ShortDate_AutoConverted = 8
    VB_Time_AutoConverted = 9
End Enum

Public Property Get IndexInFieldList() As Long
End Property

Public Property Get Name() As String
End Property

Public Property Get DefaultValue() As String
End Property

Public Property Get Updateable() As Boolean
End Property

Public Property Get OriginalColumnName() As String
End Property

Public Property Get OriginalTableName() As String
End Property

Public Property Get OriginalDataBaseName() As String
End Property

Public Property Get OriginalDataType() As String
End Property

Public Property Get CollationSequence() As String
End Property

Public Property Get NotNullConstraint() As Boolean
End Property

Public Property Get PrimaryKey() As Boolean
End Property

Public Property Get AutoIncrement() As Boolean
End Property

Public Property Get UniqueConstraint() As Boolean
End Property

Public Property Get ColumnType() As FieldType
End Property

Public Property Get DefinedSize() As Long
End Property

Public Property Get ActualSize() As Long
End Property

Public Property Get Changed() As Boolean
End Property

Public Property Get UnderlyingValue() As Variant
End Property

Public Property Get OriginalValue() As Variant
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
End Property

Public Property Let Value(RHS As Variant)
End Property
