VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cIndex - one index: name + CREATE INDEX DDL (empty for the implicit
' sqlite_autoindex_* indexes)
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

Public Property Get SQL() As String
    SQL = m_sSql
End Property

Friend Sub frInit(sName As String, sSql As String)
    m_sName = sName
    m_sSql = sSql
End Sub
