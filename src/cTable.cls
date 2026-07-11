VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cTable - one table's schema. Columns/Indexes/Triggers build a fresh
' snapshot on every access. Constraint (the table-level constraint clause)
' needs DDL parsing and is not implemented yet.
'=========================================================================
Option Explicit

Private m_oConnection               As cConnection
Private m_sDb                       As String
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

Public Property Get Columns() As cColumns
    Dim oColumns        As cColumns

    Set oColumns = New cColumns
    oColumns.frInit m_oConnection, m_sDb, m_sName
    Set Columns = oColumns
End Property

Public Property Get Constraint() As String
End Property

Public Property Get Indexes() As cIndexes
    Dim oIndexes        As cIndexes

    Set oIndexes = New cIndexes
    oIndexes.frInit m_oConnection, m_sDb, m_sName
    Set Indexes = oIndexes
End Property

Public Property Get Triggers() As cTriggers
    Dim oTriggers       As cTriggers

    Set oTriggers = New cTriggers
    oTriggers.frInit m_oConnection, m_sDb, m_sName
    Set Triggers = oTriggers
End Property

Friend Sub frInit(oCnn As cConnection, sDb As String, sName As String, sSql As String)
    Set m_oConnection = oCnn
    m_sDb = sDb
    m_sName = sName
    m_sSql = sSql
End Sub
