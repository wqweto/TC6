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
Attribute VB_Exposed = False
Option Explicit

Public Property Get Count() As Long
End Property

Public Property Get Item(Idx_Or_DatabaseName As Variant) As cDataBase
Attribute Item.VB_UserMemId = 0
End Property

Public Sub AttachDataBase(FileName As String, DataBaseName As String, Optional EncrKey As String)
End Sub

Public Sub DetachDataBase(DataBaseName As String)
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Property
