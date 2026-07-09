VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum ResetEnum
    ResetAll = 0
    ResetDeletesOnly = 1
    ResetInsertsOnly = 2
    ResetUpdatesOnly = 3
End Enum

Public Event FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
Public Event Move(ByVal NewRowIdxZeroBased As Long)
Public Event AddNew(ByVal NewRowIdxZeroBased As Long)
Public Event Delete(ByVal NewRowIdxZeroBased As Long)
Public Event QueryFinished()
Public Event Reset()

Public Property Get ContainsChanges() As Boolean
End Property

Public Property Get AutoCreateUniqueID64() As Boolean
End Property

Public Property Let AutoCreateUniqueID64(ByVal RHS As Boolean)
End Property

Public Property Get RecordCount() As Long
End Property

Public Property Get Updatable() As Boolean
End Property

Public Property Get UniqueID64() As Variant
End Property

Public Property Get BOF() As Boolean
End Property

Public Property Get EOF() As Boolean
End Property

Public Property Get AbsolutePosition() As Long
End Property

Public Property Let AbsolutePosition(ByVal RHS As Long)
End Property

Public Property Get Bookmark() As Long
End Property

Public Property Let Bookmark(ByVal RHS As Long)
End Property

Public Property Get Sort() As String
End Property

Public Property Let Sort(ByVal RHS As String)
End Property

Public Property Get Fields() As cFields
Attribute Fields.VB_UserMemId = 0
End Property

Public Property Get SQL() As String
End Property

Public Property Let SQL(SQL As String)
End Property

Public Property Get ValueMatrix(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long) As Variant
End Property

Public Property Let ValueMatrix(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long, RHS As Variant)
End Property

Public Property Get ActiveConnection() As cConnection
End Property

Public Property Set ActiveConnection(RHS As cConnection)
End Property

Public Property Get ContentChangesOnly() As Byte()
End Property

Public Property Get Content() As Byte()
End Property

Public Property Let Content(RHS() As Byte)
End Property

Public Function UniqueID64ToVBDate(UniqueID64 As Variant, Optional ReturnFracSeconds As Double) As Date
End Function

Public Sub AddNew()
End Sub

Public Sub Delete()
End Sub

Public Sub UpdateBatch()
End Sub

Public Sub MoveFirst()
End Sub

Public Sub MoveLast()
End Sub

Public Sub MoveNext()
End Sub

Public Sub MovePrevious()
End Sub

Public Sub SortRefresh()
End Sub

Public Function FindFirst(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
End Function

Public Function FindNext(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
End Function

Public Function FindPrevious(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
End Function

Public Function FindLast(Criterion As String, Optional ByVal DistinctNullValues As Boolean = True) As Boolean
End Function

Public Sub OpenRecordset(Optional SQL As String, Optional Cnn As Variant, Optional ByVal ReadOnly As Boolean)
End Sub

Public Function GetRows(Optional ByVal RowCount As Long = -1, Optional ByVal StartIdxZeroBased As Long, Optional FieldListCommaSeparated As String, Optional ByVal TransPosed As Boolean) As Variant()
End Function

Public Function GetRowsWithHeaders(Optional ByVal RowCount As Long = -1, Optional ByVal StartIdxZeroBased As Long, Optional FieldListCommaSeparated As String, Optional ByVal TransPosed As Boolean, Optional ByVal HeaderAtIdxMinus1 As Boolean) As Variant()
End Function

Public Sub ReQuery()
End Sub

Public Sub ResetChanges(Optional ByVal WhichChanges As ResetEnum = ResetAll)
End Sub

Public Function DataSource() As Object
End Function

Public Function GetADORsFromContent() As Object
End Function

Public Function ToJSONUTF8(Optional ByVal ValueArrAsColsRows As Boolean, Optional ByVal Indent As Byte, Optional ByVal UniEscaping As Boolean, Optional ByVal LfChar As Byte) As Byte()
End Function
