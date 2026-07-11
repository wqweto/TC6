Attribute VB_Name = "mdUdf"
'=========================================================================
' mdUdf - registry and AddressOf trampolines for user-defined functions,
' aggregates and collations (see cConnection.AddUserDefined*).
'
' winsqlite3.dll is built with SQLITE_CALLBACK = __stdcall (the whole ABI
' is StdCall), so plain VB6 module procedures work as callbacks directly.
' Each registered name gets a slot holding the implementing object and its
' ZeroBasedNameIndex; the 1-based slot number travels through SQLite as
' the user-data pointer and comes back via sqlite3_user_data.
'=========================================================================
Option Explicit

Private Type UDFREG
    hDb                 As LongPtr
    oFunc               As IFunction
    oAggFunc            As IAggregateFunction
    oColl               As ICollation
    lNameIdx            As Long
End Type

Private m_aRegs()                   As UDFREG
Private m_lRegCount                 As Long

Public Function UdfRegisterFunction(oCnn As cConnection, oFunc As IFunction) As Boolean
    Dim aNames()        As String
    Dim lIdx            As Long
    Dim baName()        As Byte
    Dim lSlot           As Long

    aNames = Split(oFunc.DefinedNames, ",")
    For lIdx = 0 To UBound(aNames)
        If Len(Trim$(aNames(lIdx))) > 0 Then
            lSlot = pvAllocSlot(oCnn.frDbHandle)
            Set m_aRegs(lSlot).oFunc = oFunc
            m_aRegs(lSlot).lNameIdx = lIdx
            baName = ToUtf8Array(Trim$(aNames(lIdx)) & vbNullChar)
            If stub_sqlite3_create_function_v2(oCnn.frDbHandle, VarPtr(baName(0)), -1, SQLITE_UTF8, lSlot, _
                    AddressOf UdfFuncCallback, 0, 0, 0) <> SQLITE_OK Then
                pvFreeSlot lSlot
                Exit Function
            End If
        End If
    Next
    UdfRegisterFunction = True
End Function

Public Function UdfUnregisterFunction(oCnn As cConnection, oFunc As IFunction) As Boolean
    pvUnregisterNames oCnn, oFunc.DefinedNames, False
    pvFreeByObject oCnn.frDbHandle, oFunc, Nothing, Nothing
    UdfUnregisterFunction = True
End Function

Public Function UdfRegisterAggregate(oCnn As cConnection, oAggFunc As IAggregateFunction) As Boolean
    Dim aNames()        As String
    Dim lIdx            As Long
    Dim baName()        As Byte
    Dim lSlot           As Long

    aNames = Split(oAggFunc.DefinedNames, ",")
    For lIdx = 0 To UBound(aNames)
        If Len(Trim$(aNames(lIdx))) > 0 Then
            lSlot = pvAllocSlot(oCnn.frDbHandle)
            Set m_aRegs(lSlot).oAggFunc = oAggFunc
            m_aRegs(lSlot).lNameIdx = lIdx
            baName = ToUtf8Array(Trim$(aNames(lIdx)) & vbNullChar)
            If stub_sqlite3_create_function_v2(oCnn.frDbHandle, VarPtr(baName(0)), -1, SQLITE_UTF8, lSlot, _
                    0, AddressOf UdfStepCallback, AddressOf UdfFinalCallback, 0) <> SQLITE_OK Then
                pvFreeSlot lSlot
                Exit Function
            End If
        End If
    Next
    UdfRegisterAggregate = True
End Function

Public Function UdfUnregisterAggregate(oCnn As cConnection, oAggFunc As IAggregateFunction) As Boolean
    pvUnregisterNames oCnn, oAggFunc.DefinedNames, False
    pvFreeByObject oCnn.frDbHandle, Nothing, oAggFunc, Nothing
    UdfUnregisterAggregate = True
End Function

Public Function UdfRegisterCollation(oCnn As cConnection, oColl As ICollation) As Boolean
    Dim aNames()        As String
    Dim lIdx            As Long
    Dim baName()        As Byte
    Dim lSlot           As Long

    aNames = Split(oColl.DefinedNames, ",")
    For lIdx = 0 To UBound(aNames)
        If Len(Trim$(aNames(lIdx))) > 0 Then
            lSlot = pvAllocSlot(oCnn.frDbHandle)
            Set m_aRegs(lSlot).oColl = oColl
            m_aRegs(lSlot).lNameIdx = lIdx
            baName = ToUtf8Array(Trim$(aNames(lIdx)) & vbNullChar)
            If stub_sqlite3_create_collation(oCnn.frDbHandle, VarPtr(baName(0)), SQLITE_UTF8, lSlot, _
                    AddressOf UdfCollateCallback) <> SQLITE_OK Then
                pvFreeSlot lSlot
                Exit Function
            End If
        End If
    Next
    UdfRegisterCollation = True
End Function

Public Function UdfUnregisterCollation(oCnn As cConnection, oColl As ICollation) As Boolean
    pvUnregisterNames oCnn, oColl.DefinedNames, True
    pvFreeByObject oCnn.frDbHandle, Nothing, Nothing, oColl
    UdfUnregisterCollation = True
End Function

Public Sub UdfReleaseDb(ByVal hDb As LongPtr)
    Dim lIdx            As Long

    '--- called when a connection closes: SQLite drops its registrations,
    '--- here we drop the object references the slots keep alive
    For lIdx = 1 To m_lRegCount
        If m_aRegs(lIdx).hDb = hDb Then
            pvFreeSlot lIdx
        End If
    Next
End Sub

Public Sub UdfFuncCallback(ByVal hCtx As LongPtr, ByVal lArgc As Long, ByVal lpArgv As LongPtr)
    Dim lSlot           As Long
    Dim oUdf            As cUDFMethods

    On Error GoTo EH
    lSlot = stub_sqlite3_user_data(hCtx)
    Set oUdf = New cUDFMethods
    oUdf.frInit hCtx, lArgc, lpArgv
    m_aRegs(lSlot).oFunc.Callback m_aRegs(lSlot).lNameIdx, lArgc, oUdf
    Exit Sub
EH:
    '--- a VB error must never propagate into SQLite
    pvResultError hCtx, Err.Description
End Sub

Public Sub UdfStepCallback(ByVal hCtx As LongPtr, ByVal lArgc As Long, ByVal lpArgv As LongPtr)
    Dim lSlot           As Long
    Dim oUdf            As cUDFMethods

    On Error GoTo EH
    lSlot = stub_sqlite3_user_data(hCtx)
    Set oUdf = New cUDFMethods
    oUdf.frInit hCtx, lArgc, lpArgv
    m_aRegs(lSlot).oAggFunc.CallbackStep m_aRegs(lSlot).lNameIdx, lArgc, oUdf
    Exit Sub
EH:
    pvResultError hCtx, Err.Description
End Sub

Public Sub UdfFinalCallback(ByVal hCtx As LongPtr)
    Dim lSlot           As Long
    Dim oUdf            As cUDFMethods

    On Error GoTo EH
    lSlot = stub_sqlite3_user_data(hCtx)
    Set oUdf = New cUDFMethods
    oUdf.frInit hCtx, 0, 0
    m_aRegs(lSlot).oAggFunc.CallbackFinal m_aRegs(lSlot).lNameIdx, oUdf
    Exit Sub
EH:
    pvResultError hCtx, Err.Description
End Sub

Public Function UdfCollateCallback(ByVal lpArg As LongPtr, ByVal lLen1 As Long, ByVal lpStr1 As LongPtr, ByVal lLen2 As Long, ByVal lpStr2 As LongPtr) As Long
    Dim sStr1           As String
    Dim sStr2           As String

    On Error GoTo EH
    sStr1 = FromUtf8PtrLen(lpStr1, lLen1)
    sStr2 = FromUtf8PtrLen(lpStr2, lLen2)
    UdfCollateCallback = m_aRegs(lpArg).oColl.CallbackCollate(m_aRegs(lpArg).lNameIdx, sStr1, sStr2)
    Exit Function
EH:
    '--- collations cannot report errors: treat a failure as "equal"
    UdfCollateCallback = 0
End Function

Private Sub pvUnregisterNames(oCnn As cConnection, sDefinedNames As String, ByVal bCollation As Boolean)
    Dim aNames()        As String
    Dim lIdx            As Long
    Dim baName()        As Byte

    aNames = Split(sDefinedNames, ",")
    For lIdx = 0 To UBound(aNames)
        If Len(Trim$(aNames(lIdx))) > 0 Then
            baName = ToUtf8Array(Trim$(aNames(lIdx)) & vbNullChar)
            If bCollation Then
                Call stub_sqlite3_create_collation(oCnn.frDbHandle, VarPtr(baName(0)), SQLITE_UTF8, 0, 0)
            Else
                Call stub_sqlite3_create_function_v2(oCnn.frDbHandle, VarPtr(baName(0)), -1, SQLITE_UTF8, 0, 0, 0, 0, 0)
            End If
        End If
    Next
End Sub

Private Sub pvFreeByObject(ByVal hDb As LongPtr, oFunc As IFunction, oAggFunc As IAggregateFunction, oColl As ICollation)
    Dim lIdx            As Long

    For lIdx = 1 To m_lRegCount
        If m_aRegs(lIdx).hDb = hDb Then
            If m_aRegs(lIdx).oFunc Is oFunc And m_aRegs(lIdx).oAggFunc Is oAggFunc And m_aRegs(lIdx).oColl Is oColl Then
                If Not (oFunc Is Nothing And oAggFunc Is Nothing And oColl Is Nothing) Then
                    pvFreeSlot lIdx
                End If
            End If
        End If
    Next
End Sub

Private Function pvAllocSlot(ByVal hDb As LongPtr) As Long
    Dim lIdx            As Long

    For lIdx = 1 To m_lRegCount
        If m_aRegs(lIdx).hDb = 0 Then
            pvAllocSlot = lIdx
            m_aRegs(lIdx).hDb = hDb
            Exit Function
        End If
    Next
    If m_lRegCount = 0 Then
        ReDim m_aRegs(1 To 16)
    ElseIf m_lRegCount >= UBound(m_aRegs) Then
        ReDim Preserve m_aRegs(1 To 2 * m_lRegCount)
    End If
    m_lRegCount = m_lRegCount + 1
    m_aRegs(m_lRegCount).hDb = hDb
    pvAllocSlot = m_lRegCount
End Function

Private Sub pvFreeSlot(ByVal lSlot As Long)
    m_aRegs(lSlot).hDb = 0
    m_aRegs(lSlot).lNameIdx = 0
    Set m_aRegs(lSlot).oFunc = Nothing
    Set m_aRegs(lSlot).oAggFunc = Nothing
    Set m_aRegs(lSlot).oColl = Nothing
End Sub

Private Sub pvResultError(ByVal hCtx As LongPtr, sMsg As String)
    If Len(sMsg) = 0 Then
        sMsg = "Unhandled error in user-defined function"
    End If
    Call stub_sqlite3_result_error16(hCtx, StrPtr(sMsg), -1)
End Sub
