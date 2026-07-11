VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cView - one view: name + CREATE VIEW DDL (SQLForCreate); SQL is the
' best-effort SELECT body (the text after the first " AS ")
'=========================================================================
Option Explicit

Private m_sName                     As String
Private m_sSql                      As String

Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Get NameInBrackets() As String
    NameInBrackets = "[" & m_sName & "]"
End Property

Public Property Get SQLForCreate() As String
    SQLForCreate = m_sSql
End Property

Public Property Get SQL() As String
    Dim lPos            As Long

    lPos = InStr(1, m_sSql, " AS ", vbTextCompare)
    If lPos > 0 Then
        SQL = Trim$(Mid$(m_sSql, lPos + 4))
    End If
End Property

Friend Sub frInit(sName As String, sSql As String)
    m_sName = sName
    m_sSql = sSql
End Sub
