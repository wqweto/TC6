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
'=========================================================================
' cField - live accessor for one column of the parent recordset's current
' row. Created and wired by cRecordset (frInit); Value reads the cell at
' the recordset's current row. Updates are not implemented yet.
'=========================================================================
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

Private m_oRs                       As cRecordset
Private m_lCol                      As Long

Public Property Get IndexInFieldList() As Long
    IndexInFieldList = m_lCol
End Property

Public Property Get Name() As String
    Name = m_oRs.frFieldName(m_lCol)
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
    ColumnType = pvValueType(m_oRs.frCellValue(m_lCol))
End Property

Public Property Get DefinedSize() As Long
End Property

Public Property Get ActualSize() As Long
    ActualSize = pvValueSize(m_oRs.frCellValue(m_lCol))
End Property

Public Property Get Changed() As Boolean
End Property

Public Property Get UnderlyingValue() As Variant
    UnderlyingValue = m_oRs.frCellValue(m_lCol)
End Property

Public Property Get OriginalValue() As Variant
    OriginalValue = m_oRs.frCellValue(m_lCol)
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_oRs.frCellValue(m_lCol)
End Property

Public Property Let Value(RHS As Variant)
    Err.Raise vbObjectError, "cField.Value", "Recordset updates are not implemented yet"
End Property

Friend Sub frInit(oRs As cRecordset, ByVal lCol As Long)
    Set m_oRs = oRs
    m_lCol = lCol
End Sub

Private Function pvValueType(vValue As Variant) As FieldType
    '--- infer the SQLite storage class from the materialised VB value
    Select Case VarType(vValue)
    Case vbNull, vbEmpty
        pvValueType = SQLite_NULL
    Case vbString
        pvValueType = SQLite_TEXT
    Case vbSingle, vbDouble
        pvValueType = SQLite_DOUBLE
    Case vbByte + vbArray
        pvValueType = SQLite_BLOB
    Case Else
        pvValueType = SQLite_INTEGER
    End Select
End Function

Private Function pvValueSize(vValue As Variant) As Long
    Select Case VarType(vValue)
    Case vbString
        pvValueSize = Len(vValue)
    Case vbByte + vbArray
        If pvArrayLen(vValue) > 0 Then
            pvValueSize = pvArrayLen(vValue)
        End If
    End Select
End Function

Private Function pvArrayLen(vArr As Variant) As Long
    On Error GoTo QH
    pvArrayLen = UBound(vArr) - LBound(vArr) + 1
QH:
End Function
