VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDataBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'=========================================================================
' cDataBase - one attached database. Tables/Views build a fresh snapshot
' from sqlite_master on every access, so ReScanSchemaInfo has nothing to
' invalidate.
'=========================================================================
Option Explicit

Private m_oConnection               As cConnection
Private m_sName                     As String

Public Property Get Name() As String
    Name = m_sName
End Property

Public Property Get NameInBrackets() As String
    NameInBrackets = "[" & m_sName & "]"
End Property

Public Property Get Tables() As cTables
    Dim oTables         As cTables

    Set oTables = New cTables
    oTables.frInit m_oConnection, m_sName
    Set Tables = oTables
End Property

Public Property Get Views() As cViews
    Dim oViews          As cViews

    Set oViews = New cViews
    oViews.frInit m_oConnection, m_sName
    Set Views = oViews
End Property

Public Sub ReScanSchemaInfo()
    '--- schema collections are re-queried on every access - nothing cached
End Sub

Friend Sub frInit(oCnn As cConnection, sName As String)
    Set m_oConnection = oCnn
    m_sName = sName
End Sub
