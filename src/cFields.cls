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
'=========================================================================
Option Explicit

Private m_oRs                       As cRecordset
Private m_cItems                    As VBA.Collection

Public Property Get Count() As Variant
    Count = m_cItems.Count
End Property

Public Property Get Item(ZeroBasedIdx_Or_FieldName As Variant) As cField
Attribute Item.VB_UserMemId = 0
    Dim lIdx            As Long

    If IsNumeric(ZeroBasedIdx_Or_FieldName) Then
        lIdx = CLng(ZeroBasedIdx_Or_FieldName)
    Else
        lIdx = m_oRs.frFieldIndex(CStr(ZeroBasedIdx_Or_FieldName))
    End If
    If lIdx < 0 Or lIdx >= m_cItems.Count Then
        Err.Raise 9
    End If
    Set Item = m_cItems.Item(lIdx + 1)
End Property

Public Function Exists(FieldName As String) As Boolean
    Exists = (m_oRs.frFieldIndex(FieldName) >= 0)
End Function

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = m_cItems.[_NewEnum]
End Property

Friend Sub frInit(oRs As cRecordset)
    Dim lCol            As Long
    Dim oField          As cField

    Set m_oRs = oRs
    Set m_cItems = New VBA.Collection
    For lCol = 0 To oRs.frFieldCount - 1
        Set oField = New cField
        oField.frInit oRs, lCol
        m_cItems.Add oField
    Next
End Sub
