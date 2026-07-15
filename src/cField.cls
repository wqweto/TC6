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
' row. Created and wired by cRecordset (frInit); Value reads/writes the
' cell at the recordset's current row (writes are batched - see
' cRecordset.UpdateBatch).
'
' NOTE: the weak m_pRs pointer is intentional and final - it matches the
' original RC6, where every recordset-dependent member of an orphaned
' field raises error 91 while IndexInFieldList keeps working (verified
' against RC6.dll 6.0.15); a strong-ref cRowset split was considered and
' rejected - not going to implement.
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

#If Win64 Then
Private Const PTR_SIZE                      As Long = 8
#Else
Private Const PTR_SIZE                      As Long = 4
#End If

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private m_pRs                       As cRecordset
Private m_lCol                      As Long

Public Property Get IndexInFieldList() As Long
    IndexInFieldList = m_lCol
End Property

Public Property Get Name() As String
    Name = m_pRs.frFieldName(m_lCol)
End Property

Public Property Get DefaultValue() As String
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        DefaultValue = oColumn.DefaultValue
    End If
End Property

Public Property Get Updateable() As Boolean
    Updateable = m_pRs.frColumnUpdatable(m_lCol)
End Property

Public Property Get OriginalColumnName() As String
    OriginalColumnName = m_pRs.frColOriginName(m_lCol)
End Property

Public Property Get OriginalTableName() As String
    OriginalTableName = m_pRs.frColOriginTable(m_lCol)
End Property

Public Property Get OriginalDataBaseName() As String
    OriginalDataBaseName = m_pRs.frColOriginDb(m_lCol)
End Property

Public Property Get OriginalDataType() As String
    OriginalDataType = m_pRs.frColDeclType(m_lCol)
End Property

Public Property Get CollationSequence() As String
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        CollationSequence = oColumn.Collate
    End If
End Property

Public Property Get NotNullConstraint() As Boolean
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        NotNullConstraint = oColumn.NotNullConstraint
    End If
End Property

Public Property Get PrimaryKey() As Boolean
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        PrimaryKey = oColumn.PrimaryKey
    End If
End Property

Public Property Get AutoIncrement() As Boolean
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        AutoIncrement = oColumn.PrimaryAutoIncrement
    End If
End Property

Public Property Get UniqueConstraint() As Boolean
    Dim oColumn         As cColumn

    Set oColumn = m_pRs.frColSchema(m_lCol)
    If Not oColumn Is Nothing Then
        UniqueConstraint = oColumn.UniqueConstraint
    End If
End Property

Public Property Get ColumnType() As FieldType
    ColumnType = pvValueType(m_pRs.frCellValue(m_lCol))
End Property

Public Property Get DefinedSize() As Long
    Dim sType           As String
    Dim lPos            As Long
    Dim lEnd            As Long

    '--- the size in parentheses of the declared type, e.g. VARCHAR(50)
    sType = m_pRs.frColDeclType(m_lCol)
    lPos = InStr(sType, "(")
    If lPos > 0 Then
        lEnd = InStr(lPos, sType, ")")
        If lEnd > lPos Then
            DefinedSize = Val(Mid$(sType, lPos + 1, lEnd - lPos - 1))
        End If
    End If
End Property

Public Property Get ActualSize() As Long
    ActualSize = pvValueSize(m_pRs.frCellValue(m_lCol))
End Property

Public Property Get Changed() As Boolean
    Changed = m_pRs.frCellDirty(m_lCol)
End Property

Public Property Get UnderlyingValue() As Variant
    UnderlyingValue = m_pRs.frCellValue(m_lCol)
End Property

Public Property Get OriginalValue() As Variant
    OriginalValue = m_pRs.frCellOriginal(m_lCol)
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_pRs.frCellValue(m_lCol)
End Property

Public Property Let Value(RHS As Variant)
    m_pRs.frSetCellValue m_lCol, RHS
End Property

Friend Sub frInit(oRs As cRecordset, ByVal lCol As Long)
    Call CopyMemory(m_pRs, ObjPtr(oRs), PTR_SIZE)
    m_lCol = lCol
End Sub

Friend Sub frTerminate()
    Call CopyMemory(m_pRs, ObjPtr(Nothing), PTR_SIZE)
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

Private Sub Class_Terminate()
    frTerminate
End Sub
