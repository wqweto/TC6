VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum eConnectionSynchronousType
    SynchronousOff = 0
    SynchronousNormal = 1
    SynchronousFull = 2
    SynchronousExtra = 3
End Enum

Public Enum eCodecType
    CODEC_TYPE_AES128 = 1
    CODEC_TYPE_AES256 = 2
    CODEC_TYPE_CHACHA20 = 3
    CODEC_TYPE_SQLCIPHER = 4
    CODEC_TYPE_RC4 = 5
End Enum

Public Event CommitTransComplete()
Public Event RollbackTransComplete()
Public Event CommitForSavePointComplete(SavePointName As String)
Public Event SavePointReleased(SavePointName As String)
Public Event RollbackToSavePointComplete(SavePointName As String)

Public Property Get MapDbNullToEmpty() As Boolean
End Property

Public Property Let MapDbNullToEmpty(ByVal MapDbNullToEmpty As Boolean)
End Property

Public Property Get DBHdl() As Long
End Property

Public Property Get FileName() As String
End Property

Public Property Let FileName(FileName As String)
End Property

Public Property Get PageSize() As Long
End Property

Public Property Get BusyTimeOutSeconds() As Double
End Property

Public Property Let BusyTimeOutSeconds(ByVal RHS As Double)
End Property

Public Property Get Synchronous() As eConnectionSynchronousType
End Property

Public Property Let Synchronous(ByVal RHS As eConnectionSynchronousType)
End Property

Public Property Get DataBases() As cDataBases
End Property

Public Property Get UniqueID64() As Variant
End Property

Public Property Get AffectedRows() As Long
End Property

Public Property Get TotalAffectedRowsInSession() As Long
End Property

Public Property Get LastInsertAutoID() As Variant
End Property

Public Property Get EnableNestedTransactions() As Boolean
End Property

Public Property Let EnableNestedTransactions(ByVal RHS As Boolean)
End Property

Public Property Get TransactionStackCounter() As Long
End Property

Public Property Get TransactionStack() As cCollection
End Property

Public Property Get MemDB() As cMemDB
End Property

Public Property Get CodecType() As eCodecType
End Property

Public Property Let CodecType(ByVal RHS As eCodecType)
End Property

Public Function Version() As String
End Function

Public Function CheckIntegrity() As String
End Function

Public Function CreateNewDB(Optional FileName As String = ":memory:", Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
End Function

Public Function OpenDB(Optional FileName As String, Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
End Function

Public Function OpenDBReadOnly(Optional FileName As String, Optional EncrKey As String, Optional ByVal EnableVBFunctions As Boolean = True) As Boolean
End Function

Public Function OpenSchema() As cRecordset
End Function

Public Sub ReKey(Optional NewEncrKey As String)
End Sub

Public Sub AttachDataBase(FileName As String, DataBaseName As String, Optional EncrKey As String)
End Sub

Public Sub DetachDataBase(DataBaseName As String)
End Sub

Public Function OpenRecordset(SQL As String, Optional ByVal ReadOnly As Boolean) As cRecordset
End Function

Public Sub BeginTrans(Optional SavePointName As String)
End Sub

Public Sub CommitTrans(Optional SavePointName As String)
End Sub

Public Sub RollbackTrans(Optional ToSavePointName As String, Optional ByVal ReUseSavePointForNewAttempts As Boolean)
End Sub

Public Function CopyDatabase(Optional ByVal DstDBFileName As String = ":memory:", Optional ByVal EncrKey As String, Optional ByVal CompactDstDB As Boolean) As cConnection
End Function

Public Sub CompactDataBase()
End Sub

Public Sub Cancel()
End Sub

Public Function CreateCommand(SQL As String) As cCommand
End Function

Public Function CreateSelectCommand(SQL As String) As cSelectCommand
End Function

Public Function CreateCursor(SQL As String) As cCursor
End Function

Public Function Execute(SQL As String) As Boolean
End Function

Public Function GetDateString(ByVal D As Date) As String
End Function

Public Function GetShortDateString(ByVal D As Date) As String
End Function

Public Function GetTimeString(ByVal D As Date) As String
End Function

Public Function GetBooleanString(ByVal B As Boolean) As String
End Function

Public Function LastDBError() As String
End Function

Public Function GetSqliteErrStr(ByVal ResultCode As Long) As String
End Function

Public Function LastDBErrCode(Optional ByVal GetExtendedCode As Boolean) As Long
End Function

Public Function AddUserDefinedFunction(New_ObjectIFunc As IFunction) As Boolean
End Function

Public Function RemoveUserDefinedFunction(New_ObjectIFunc As IFunction) As Boolean
End Function

Public Function AddUserDefinedAggregateFunction(New_ObjectIAggFunc As IAggregateFunction) As Boolean
End Function

Public Function RemoveUserDefinedAggregateFunction(New_ObjectIAggFunc As IAggregateFunction) As Boolean
End Function

Public Function AddUserDefinedCollation(New_ObjectIColl As ICollation) As Boolean
End Function

Public Function RemoveUserDefinedCollation(New_ObjectIColl As ICollation) As Boolean
End Function

Public Sub CreateTableFromRsContent(Content() As Byte, Optional NewTableName As String, Optional ByVal AsTempTable As Boolean, Optional ByVal WithPrimaryKeys As Boolean)
End Sub

Public Sub CreateTableFromADORs(CnnDst As cConnection, ByVal TableName As String, AdoRs As Object, Optional arrDataTypes As Variant)
End Sub

Public Sub CreateTable(TableName As String, Optional FieldDefs As Collection)
End Sub

Public Function NewFieldDefs() As Collection
End Function

Public Function GetRs(ByVal SQL As String, ParamArray P() As Variant) As cRecordset
End Function

Public Sub ExecCmd(ByVal SQL As String, ParamArray P() As Variant)
End Sub
