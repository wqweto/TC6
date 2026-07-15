VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEventSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' cEventSink - WithEvents trace recorder for event-timing tests
'
' Records every cRecordset/cConnection event as "Name(args)" tokens joined
' with ";" so tests can assert exact event sequences against the traces
' observed from the original RC6.dll.
'=========================================================================
Option Explicit

Private WithEvents m_oRs             As cRecordset
Attribute m_oRs.VB_VarHelpID = -1
Private WithEvents m_oCnn            As cConnection
Attribute m_oCnn.VB_VarHelpID = -1
Private WithEvents m_oConv           As cConverter
Attribute m_oConv.VB_VarHelpID = -1
Private m_sTrace                     As String

Public Property Get Trace() As String
    Trace = m_sTrace
End Property

Public Sub Attach(oRs As cRecordset)
    Set m_oRs = oRs
End Sub

Public Sub AttachConnection(oCnn As cConnection)
    Set m_oCnn = oCnn
End Sub

Public Sub AttachConverter(oConv As cConverter)
    Set m_oConv = oConv
End Sub

Public Sub Clear()
    m_sTrace = vbNullString
End Sub

Private Sub pvAppend(sText As String)
    If Len(m_sTrace) > 0 Then
        m_sTrace = m_sTrace & ";"
    End If
    m_sTrace = m_sTrace & sText
End Sub

Private Sub m_oRs_FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
    pvAppend "FieldChange(" & RowIdxZeroBased & "," & ColIdxZeroBased & ")"
End Sub

Private Sub m_oRs_Move(ByVal NewRowIdxZeroBased As Long)
    pvAppend "Move(" & NewRowIdxZeroBased & ")"
End Sub

Private Sub m_oRs_AddNew(ByVal NewRowIdxZeroBased As Long)
    pvAppend "AddNew(" & NewRowIdxZeroBased & ")"
End Sub

Private Sub m_oRs_Delete(ByVal NewRowIdxZeroBased As Long)
    pvAppend "Delete(" & NewRowIdxZeroBased & ")"
End Sub

Private Sub m_oRs_QueryFinished()
    pvAppend "QueryFinished"
End Sub

Private Sub m_oRs_Reset()
    pvAppend "Reset"
End Sub

Private Sub m_oCnn_CommitTransComplete()
    pvAppend "CommitTransComplete"
End Sub

Private Sub m_oCnn_RollbackTransComplete()
    pvAppend "RollbackTransComplete"
End Sub

Private Sub m_oCnn_CommitForSavePointComplete(SavePointName As String)
    pvAppend "CommitForSavePoint(" & SavePointName & ")"
End Sub

Private Sub m_oCnn_SavePointReleased(SavePointName As String)
    pvAppend "SavePointReleased(" & SavePointName & ")"
End Sub

Private Sub m_oCnn_RollbackToSavePointComplete(SavePointName As String)
    pvAppend "RollbackToSavePoint(" & SavePointName & ")"
End Sub

Private Sub m_oConv_SchemaProgress(ByVal TableName As String, ByVal TablesCount As Long, ByVal TablesCreated As Long)
    pvAppend "SchemaProgress(" & TableName & "," & TablesCount & "," & TablesCreated & ")"
End Sub

Private Sub m_oConv_InsertProgress(ByVal TableName As String, ByVal RecordCount As Long, ByVal RecordsInserted As Long)
    pvAppend "InsertProgress(" & TableName & "," & RecordCount & "," & RecordsInserted & ")"
End Sub

Private Sub m_oConv_IndexProgress(ByVal TableName As String, ByVal IndexName As String, ByVal IndexesCount As Long, ByVal IndexesCreated As Long)
    pvAppend "IndexProgress(" & TableName & "," & IndexName & "," & IndexesCount & "," & IndexesCreated & ")"
End Sub
