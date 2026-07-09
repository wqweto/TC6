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
Attribute VB_Exposed = False
Option Explicit

Public Property Get Name() As String
End Property

Public Property Get NameInBrackets() As String
End Property

Public Property Get SQLForCreate() As String
End Property

Public Property Get Columns() As cColumns
End Property

Public Property Get Constraint() As String
End Property

Public Property Get Indexes() As cIndexes
End Property

Public Property Get Triggers() As cTriggers
End Property
