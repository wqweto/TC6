VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMemDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get Cnn() As cConnection
End Property

Public Property Set Cnn(RHS As cConnection)
End Property

Public Function GetRs(ByVal SQL As String, ParamArray P() As Variant) As cRecordset
End Function

Public Sub CreateTable(TableName As String, Optional FieldDefs As Collection)
End Sub

Public Function NewFieldDefs() As Collection
End Function

Public Function GetTable(TableName As String, Optional ByVal WhereClause As String, Optional ByVal OrderByClause As String) As cRecordset
End Function

Public Function GetSingleVal(SQL As String) As Variant
End Function

Public Function GetSum(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
End Function

Public Function GetAvg(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
End Function

Public Function GetMin(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
End Function

Public Function GetMax(TableName As String, Field As String, Optional ByVal WhereClause As String) As Variant
End Function

Public Function GetCount(TableName As String, Optional ByVal Field As String, Optional ByVal WhereClause As String) As Variant
End Function

Public Sub Exec(SQL As String)
End Sub

Public Sub ExecCmd(ByVal SQL As String, ParamArray P() As Variant)
End Sub

Public Function CreateCommand(SQL As String) As cCommand
End Function

Public Function CreateSelectCommand(SQL As String) As cSelectCommand
End Function

Public Sub BeginTrans(Optional SavePointName As String)
End Sub

Public Sub CommitTrans(Optional SavePointName As String)
End Sub

Public Sub RollbackTrans(Optional ToSavePointName As String, Optional ByVal ReUseSavePointForNewAttempts As Boolean)
End Sub

Public Sub CreateTableFromRs(AdoOrSQLiteRsOrByteContent As Variant, ByVal DstTableName As String, Optional ByVal WithPrimaryKeys As Boolean)
End Sub

Public Function CreateCursor(SQL As String) As cCursor
End Function
