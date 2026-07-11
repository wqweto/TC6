VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIndexes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cIndexes - the indexes of one table (sqlite_master, including implicit
' sqlite_autoindex_* entries whose SQL is empty), snapshotted at creation.
' Item accepts a 0-based index or an index name (keyed, O(1)).
'=========================================================================
Option Explicit

Private m_cItems                    As VBA.Collection

Public Property Get Count() As Long
    Count = m_cItems.Count
End Property

Public Property Get Item(Idx_Or_IndexName As Variant) As cIndex
Attribute Item.VB_UserMemId = 0
    Dim lIdx            As Long

    If IsNumeric(Idx_Or_IndexName) Then
        lIdx = CLng(Idx_Or_IndexName)
        If lIdx < 0 Or lIdx >= m_cItems.Count Then
            Err.Raise 9
        End If
        Set Item = m_cItems.Item(lIdx + 1)
    Else
        Set Item = pvItemByName(CStr(Idx_Or_IndexName))
    End If
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Friend Sub frInit(oCnn As cConnection, sDb As String, sTable As String)
    Dim oRs             As cRecordset
    Dim oIndex          As cIndex
    Dim lRow            As Long
    Dim sSql            As String

    Set m_cItems = New VBA.Collection
    Set oRs = oCnn.GetRs("SELECT name, sql FROM " & QuoteIdentifier(sDb) & ".sqlite_master" & _
        " WHERE type='index' AND tbl_name = ? ORDER BY name", sTable)
    For lRow = 0 To oRs.RecordCount - 1
        sSql = vbNullString
        If Not IsNull(oRs.ValueMatrix(lRow, 1)) Then
            sSql = CStr(oRs.ValueMatrix(lRow, 1))
        End If
        Set oIndex = New cIndex
        oIndex.frInit CStr(oRs.ValueMatrix(lRow, 0)), sSql
        m_cItems.Add oIndex, CStr(oRs.ValueMatrix(lRow, 0))
    Next
End Sub

Private Function pvItemByName(sName As String) As cIndex
    On Error GoTo EH
    Set pvItemByName = m_cItems.Item(sName)
    Exit Function
EH:
    Err.Raise 9, "cIndexes", "Unknown index '" & sName & "'"
End Function
