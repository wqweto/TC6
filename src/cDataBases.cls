VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataBases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cDataBases - the databases attached to a connection (PRAGMA
' database_list), snapshotted at creation and after Attach/Detach. Item
' accepts a 0-based index or a database name (keyed, O(1)).
'=========================================================================
Option Explicit

Private m_oConnection               As cConnection
Private m_cItems                    As VBA.Collection

Public Property Get Count() As Long
    Count = m_cItems.Count
End Property

Public Property Get Item(Idx_Or_DatabaseName As Variant) As cDataBase
Attribute Item.VB_UserMemId = 0
    Dim lIdx            As Long

    If IsNumeric(Idx_Or_DatabaseName) Then
        lIdx = CLng(Idx_Or_DatabaseName)
        If lIdx < 0 Or lIdx >= m_cItems.Count Then
            Err.Raise 9
        End If
        Set Item = m_cItems.Item(lIdx + 1)
    Else
        Set Item = pvItemByName(CStr(Idx_Or_DatabaseName))
    End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Public Sub AttachDataBase(FileName As String, DataBaseName As String, Optional EncrKey As String)
    m_oConnection.AttachDataBase FileName, DataBaseName, EncrKey
    pvReload
End Sub

Public Sub DetachDataBase(DataBaseName As String)
    m_oConnection.DetachDataBase DataBaseName
    pvReload
End Sub

Friend Sub frInit(oCnn As cConnection)
    Set m_oConnection = oCnn
    pvReload
End Sub

Private Sub pvReload()
    Dim oRs             As cRecordset
    Dim oDb             As cDataBase
    Dim lRow            As Long

    Set m_cItems = New VBA.Collection
    '--- database_list columns: seq, name, file
    Set oRs = m_oConnection.GetRs("PRAGMA database_list")
    For lRow = 0 To oRs.RecordCount - 1
        Set oDb = New cDataBase
        oDb.frInit m_oConnection, CStr(oRs.ValueMatrix(lRow, 1))
        m_cItems.Add oDb, CStr(oRs.ValueMatrix(lRow, 1))
    Next
End Sub

Private Function pvItemByName(sName As String) As cDataBase
    On Error GoTo EH
    Set pvItemByName = m_cItems.Item(sName)
    Exit Function
EH:
    '--- clear the pending collection error so 5 gets its default description
    Err.Clear
    Err.Raise 5
End Function
