VERSION 1.0 CLASS
BEGIN
  MultiUse = 0  'False
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--- test aggregate UDF: PROD multiplies the values of a group
Option Explicit

Implements IAggregateFunction

Private m_dblProd                   As Double
Private m_bHasRows                  As Boolean

Private Property Get IAggregateFunction_DefinedNames() As String
    IAggregateFunction_DefinedNames = "PROD"
End Property

Private Sub IAggregateFunction_CallbackStep(ByVal ZeroBasedNameIndex As Long, ByVal ParamCount As Long, UDF As cUDFMethods)
    If Not m_bHasRows Then
        m_dblProd = 1
        m_bHasRows = True
    End If
    m_dblProd = m_dblProd * UDF.GetDouble(1)
End Sub

Private Sub IAggregateFunction_CallbackFinal(ByVal ZeroBasedNameIndex As Long, UDF As cUDFMethods)
    If m_bHasRows Then
        UDF.SetResultDouble m_dblProd
    Else
        UDF.SetResultNull
    End If
    m_bHasRows = False
End Sub
