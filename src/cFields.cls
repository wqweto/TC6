VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFields"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cFields - the columns of a cRecordset. One cField per column, created
' and wired by cRecordset (frInit). Item accepts a 0-based index or a
' field name; For Each iterates the fields in order.
'
' The backing collection is keyed on field name, so Item(name)/Exists are
' O(1) lookups (no per-access scan of the parent recordset). SQLite allows
' duplicate result-column names, so only the first occurrence of a name is
' keyed - later duplicates are still reachable by index (matching the
' first-match semantics of a name lookup).
'=========================================================================
Option Explicit

Private m_cItems                    As VBA.Collection

Public Property Get Count() As Variant
    Count = m_cItems.Count
End Property

Public Property Get Item(ZeroBasedIdx_Or_FieldName As Variant) As cField
Attribute Item.VB_UserMemId = 0
    Dim lIdx            As Long

    If IsNumeric(ZeroBasedIdx_Or_FieldName) Then
        lIdx = CLng(ZeroBasedIdx_Or_FieldName)
        If lIdx < 0 Or lIdx >= m_cItems.Count Then
            Err.Raise vbObjectError, "cFields.Item", "Field-Index out of range"
        End If
        Set Item = m_cItems.Item(lIdx + 1)
    Else
        Set Item = pvItemByName(CStr(ZeroBasedIdx_Or_FieldName))
    End If
End Property

Public Function Exists(FieldName As String) As Boolean
    Dim oField          As cField

    On Error GoTo QH
    Set oField = m_cItems.Item(FieldName)
    Exists = True
QH:
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Friend Sub frInit(oRs As cRecordset)
    Dim lCol            As Long
    Dim oField          As cField

    Set m_cItems = New VBA.Collection
    For lCol = 0 To oRs.frFieldCount - 1
        Set oField = New cField
        oField.frInit oRs, lCol
        pvAdd oField, oRs.frFieldName(lCol)
    Next
End Sub

Friend Sub frTerminate()
    Dim oField          As cField

    If m_cItems Is Nothing Then
        Exit Sub
    End If
    For Each oField In m_cItems
        oField.frTerminate
    Next
End Sub

Private Sub pvAdd(oField As cField, sName As String)
    '--- key on the field name; a duplicate name (or an empty one) falls back
    '--- to a positional-only Add so the collection stays 1:1 with the columns
    On Error GoTo QH
    If Len(sName) > 0 Then
        m_cItems.Add oField, sName
        Exit Sub
    End If
QH:
    m_cItems.Add oField
End Sub

Private Function pvItemByName(sName As String) As cField
    On Error GoTo EH
    Set pvItemByName = m_cItems.Item(sName)
    Exit Function
EH:
    Err.Raise vbObjectError, "cFields.Item", "No such Field-Def: " & sName
End Function
