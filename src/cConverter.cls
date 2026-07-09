VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event SchemaProgress(ByVal TableName As String, ByVal TablesCount As Long, ByVal TablesCreated As Long)
Public Event InsertProgress(ByVal TableName As String, ByVal RecordCount As Long, ByVal RecordsInserted As Long)
Public Event IndexProgress(ByVal TableName As String, ByVal IndexName As String, ByVal IndexesCount As Long, ByVal IndexesCreated As Long)

Public Property Get CaseInSensitive_TextColumns() As Boolean
End Property

Public Property Let CaseInSensitive_TextColumns(ByVal RHS As Boolean)
End Property

Public Sub ConvertDatabase(Cnn As Object, sCnn As cConnection)
End Sub

Public Sub ConvertIndexes(Cnn As Object, sCnn As cConnection)
End Sub
