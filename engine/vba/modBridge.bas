Option Explicit
' Read-only bridge between Python (Application.Run) and Oracle Smart View (HsAddin.dll).
' Every function returns a value and traps its own errors: an unhandled VBA error opens the
' debugger, and the debugger blocks COM. Status strings are "OK" or "ERR|<rc>|<message>|<last error>".
' Deliberately absent: submit, calculate, business rules, and anything else that writes to EPM.

' Smart View acts on sheets of the ACTIVE workbook; Python activates the right workbook before each call
' (the bridge's own grids, or a copied user workbook being refreshed).
Private Function Ws(ByVal sheetName As String) As Worksheet
    Set Ws = ActiveWorkbook.Worksheets(sheetName)
End Function

Private Function St(ByVal rc As Long) As String
    Dim msg As Variant, code As Variant, emsg As Variant, edesc As Variant
    If rc = 0 Then St = "OK": Exit Function
    On Error Resume Next
    HypGetErrorInfo rc, msg
    HypGetLastError code, emsg, edesc
    St = "ERR|" & rc & "|" & CStr(msg) & "|" & CStr(emsg) & " " & CStr(edesc)
End Function

Private Function VbaErr(ByVal where As String) As String
    VbaErr = "ERR|VBA" & Err.Number & "|" & where & ": " & Err.Description & "|"
End Function

Public Function B_Ping() As String
    B_Ping = "pong"
End Function

' Test hook for the watchdog: blocks Excel's thread for `seconds` without yielding.
Public Function B_TestBusy(ByVal seconds As Double) As String
    Dim t As Double
    t = Timer + seconds
    Do While Timer < t
    Loop
    B_TestBusy = "done"
End Function

' ---------- options ----------
Public Function B_SetOption(ByVal item As Long, ByVal value As Variant, ByVal sheetName As String) As String
    On Error GoTo fail
    B_SetOption = St(HypSetOption(item, value, sheetName))
    Exit Function
fail:
    B_SetOption = VbaErr("SetOption")
End Function

Public Function B_GetOption(ByVal item As Long, ByVal sheetName As String) As Variant
    On Error GoTo fail
    Dim v As Variant, rc As Long
    rc = HypGetOption(item, v, sheetName)
    If rc = 0 Then B_GetOption = Array("OK", v) Else B_GetOption = Array(St(rc))
    Exit Function
fail:
    B_GetOption = Array(VbaErr("GetOption"))
End Function

Public Function B_SetAliasTable(ByVal sheetName As String, ByVal tableName As String) As String
    On Error GoTo fail
    B_SetAliasTable = St(HypSetAliasTable(sheetName, tableName))
    Exit Function
fail:
    B_SetAliasTable = VbaErr("SetAliasTable")
End Function

Public Function B_SetConnAliasTable(ByVal conn As String, ByVal tableName As String) As String
    On Error GoTo fail
    B_SetConnAliasTable = St(HypSetConnAliasTable(conn, tableName))
    Exit Function
fail:
    B_SetConnAliasTable = VbaErr("SetConnAliasTable")
End Function

' ---------- connection ----------
Public Function B_Connect(ByVal sheetName As String, ByVal conn As String) As String
    On Error GoTo fail
    B_Connect = St(HypUIConnect(sheetName, Empty, Empty, conn))
    Exit Function
fail:
    B_Connect = VbaErr("Connect")
End Function

Public Function B_Connected(ByVal sheetName As String) As String
    On Error GoTo fail
    B_Connected = CStr(HypConnected(sheetName))
    Exit Function
fail:
    B_Connected = VbaErr("Connected")
End Function

Public Function B_ConnectionExists(ByVal conn As String) As String
    On Error GoTo fail
    B_ConnectionExists = CStr(HypConnectionExists(conn))
    Exit Function
fail:
    B_ConnectionExists = VbaErr("ConnectionExists")
End Function

Public Function B_Disconnect(ByVal sheetName As String) As String
    On Error GoTo fail
    B_Disconnect = St(HypDisconnect(sheetName, False))
    Exit Function
fail:
    B_Disconnect = VbaErr("Disconnect")
End Function

Public Function B_DisconnectAll() As String
    On Error GoTo fail
    B_DisconnectAll = St(HypDisconnectAll())
    Exit Function
fail:
    B_DisconnectAll = VbaErr("DisconnectAll")
End Function

' ---------- ad hoc ----------
Public Function B_Retrieve(ByVal sheetName As String) As String
    On Error GoTo fail
    Ws(sheetName).Activate
    B_Retrieve = St(HypRetrieve(sheetName))
    Exit Function
fail:
    B_Retrieve = VbaErr("Retrieve")
End Function

Public Function B_RetrieveRange(ByVal sheetName As String, ByVal addr As String, ByVal conn As String) As String
    On Error GoTo fail
    B_RetrieveRange = St(HypRetrieveRange(sheetName, Ws(sheetName).Range(addr), conn))
    Exit Function
fail:
    B_RetrieveRange = VbaErr("RetrieveRange")
End Function

Public Function B_ZoomIn(ByVal sheetName As String, ByVal addr As String, ByVal level As Long, ByVal across As Boolean) As String
    On Error GoTo fail
    B_ZoomIn = St(HypZoomIn(sheetName, Ws(sheetName).Range(addr), level, across))
    Exit Function
fail:
    B_ZoomIn = VbaErr("ZoomIn")
End Function

Public Function B_ZoomOut(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo fail
    B_ZoomOut = St(HypZoomOut(sheetName, Ws(sheetName).Range(addr)))
    Exit Function
fail:
    B_ZoomOut = VbaErr("ZoomOut")
End Function

Public Function B_PivotToGrid(ByVal sheetName As String, ByVal dimName As String, ByVal addr As String) As String
    On Error GoTo fail
    B_PivotToGrid = St(HypPivotToGrid(sheetName, dimName, Ws(sheetName).Range(addr)))
    Exit Function
fail:
    B_PivotToGrid = VbaErr("PivotToGrid")
End Function

Public Function B_PivotToPOV(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo fail
    B_PivotToPOV = St(HypPivotToPOV(sheetName, Ws(sheetName).Range(addr)))
    Exit Function
fail:
    B_PivotToPOV = VbaErr("PivotToPOV")
End Function

Public Function B_Pivot(ByVal sheetName As String, ByVal startAddr As String, ByVal endAddr As String) As String
    On Error GoTo fail
    B_Pivot = St(HypPivot(sheetName, Ws(sheetName).Range(startAddr), Ws(sheetName).Range(endAddr)))
    Exit Function
fail:
    B_Pivot = VbaErr("Pivot")
End Function

Public Function B_KeepOnly(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo fail
    B_KeepOnly = St(HypKeepOnly(sheetName, Ws(sheetName).Range(addr)))
    Exit Function
fail:
    B_KeepOnly = VbaErr("KeepOnly")
End Function

Public Function B_RemoveOnly(ByVal sheetName As String, ByVal addr As String) As String
    On Error GoTo fail
    B_RemoveOnly = St(HypRemoveOnly(sheetName, Ws(sheetName).Range(addr)))
    Exit Function
fail:
    B_RemoveOnly = VbaErr("RemoveOnly")
End Function

' (Removed on purpose: HypExecuteQuery - an MDX "insert" could write - and HypMenuVRefreshAll - it would refresh
'  every workbook open in the user's Excel. The MCP refreshes only the sheets it is asked to.)

' ---------- metadata ----------
Public Function B_GetDimensions(ByVal sheetName As String) As Variant
    On Error GoTo fail
    Dim names As Variant, types As Variant, rc As Long
    rc = HypGetDimensions(sheetName, names, types)
    If rc = 0 Then B_GetDimensions = Array("OK", names, types) Else B_GetDimensions = Array(St(rc))
    Exit Function
fail:
    B_GetDimensions = Array(VbaErr("GetDimensions"))
End Function

Public Function B_GetPOV(ByVal sheetName As String) As Variant
    On Error GoTo fail
    Dim dims As Variant, mbrs As Variant, typ As Variant, rc As Long
    rc = HypGetPOV(sheetName, dims, mbrs, typ)
    If rc = 0 Then B_GetPOV = Array("OK", dims, mbrs, typ) Else B_GetPOV = Array(St(rc))
    Exit Function
fail:
    B_GetPOV = Array(VbaErr("GetPOV"))
End Function

Public Function B_GetMembers(ByVal sheetName As String, ByVal dimName As String) As Variant
    On Error GoTo fail
    Dim names As Variant, descs As Variant, rc As Long
    rc = HypGetMembers(sheetName, dimName, names, descs)
    If rc = 0 Then B_GetMembers = Array("OK", names, descs) Else B_GetMembers = Array(St(rc))
    Exit Function
fail:
    B_GetMembers = Array(VbaErr("GetMembers"))
End Function

Public Function B_FindMember(ByVal sheetName As String, ByVal mbr As String, ByVal aliasTable As String) As Variant
    On Error GoTo fail
    Dim dn As Variant, al As Variant, gn As Variant, lv As Variant, rc As Long
    rc = HypFindMember(sheetName, mbr, aliasTable, dn, al, gn, lv)
    If rc = 0 Then B_FindMember = Array("OK", dn, al, gn, lv) Else B_FindMember = Array(St(rc))
    Exit Function
fail:
    B_FindMember = Array(VbaErr("FindMember"))
End Function

' predicate: 1 children, 2 descendants, 3 bottom level, 4 siblings, 5 same level, 6 same generation,
' 7 parent, 8 dimension, 9 named generation, 10 named level, 11 search, 12 wild search
Public Function B_QueryMembers(ByVal sheetName As String, ByVal mbr As String, ByVal predicate As Long, _
                               ByVal opt As Variant, ByVal dimName As String, ByVal in1 As Variant, ByVal in2 As Variant) As Variant
    On Error GoTo fail
    Dim arr As Variant, rc As Long
    rc = HypQueryMembers(sheetName, mbr, predicate, opt, dimName, in1, in2, arr)
    If rc = 0 Then B_QueryMembers = Array("OK", arr) Else B_QueryMembers = Array(St(rc))
    Exit Function
fail:
    B_QueryMembers = Array(VbaErr("QueryMembers"))
End Function

Public Function B_OtlMemberInfo(ByVal sheetName As String, ByVal dimName As String, ByVal mbr As String, ByVal predicate As Long) As Variant
    On Error GoTo fail
    Dim arr As Variant, rc As Long
    rc = HypOtlGetMemberInfo(sheetName, dimName, mbr, predicate, arr)
    If rc = 0 Then B_OtlMemberInfo = Array("OK", arr) Else B_OtlMemberInfo = Array(St(rc))
    Exit Function
fail:
    B_OtlMemberInfo = Array(VbaErr("OtlGetMemberInfo"))
End Function

Public Function B_MemberInfoEx(ByVal sheetName As String, ByVal mbr As String) As Variant
    On Error GoTo fail
    Dim pn As Variant, pv As Variant, ps As Variant, rc As Long
    rc = HypGetMemberInformationEx(sheetName, mbr, pn, pv, ps)
    If rc = 0 Then B_MemberInfoEx = Array("OK", pn, pv, ps) Else B_MemberInfoEx = Array(St(rc))
    Exit Function
fail:
    B_MemberInfoEx = Array(VbaErr("GetMemberInformationEx"))
End Function

Public Function B_GetParent(ByVal sheetName As String, ByVal mbr As String) As Variant
    On Error GoTo fail
    Dim p As Variant, rc As Long
    rc = HypGetParent(sheetName, mbr, p)
    If rc = 0 Then B_GetParent = Array("OK", p) Else B_GetParent = Array(St(rc))
    Exit Function
fail:
    B_GetParent = Array(VbaErr("GetParent"))
End Function

Public Function B_GetChildren(ByVal sheetName As String, ByVal mbr As String, ByVal count As Integer) As Variant
    On Error GoTo fail
    Dim kids As Variant, rc As Long
    rc = HypGetChildren(sheetName, mbr, count, kids)
    If rc = 0 Then B_GetChildren = Array("OK", kids) Else B_GetChildren = Array(St(rc))
    Exit Function
fail:
    B_GetChildren = Array(VbaErr("GetChildren"))
End Function

Public Function B_ListApplications(ByVal url As String, ByVal server As String) As Variant
    On Error GoTo fail
    Dim apps As Variant, descs As Variant, rc As Long
    rc = HypListApplications(url, server, Empty, Empty, apps, descs)
    If rc = 0 Then B_ListApplications = Array("OK", apps, descs) Else B_ListApplications = Array(St(rc))
    Exit Function
fail:
    B_ListApplications = Array(VbaErr("ListApplications"))
End Function

Public Function B_ListDatabases(ByVal url As String, ByVal server As String, ByVal app As String) As Variant
    On Error GoTo fail
    Dim dbs As Variant, rc As Long
    rc = HypListDatabases(url, server, Empty, Empty, app, dbs)
    If rc = 0 Then B_ListDatabases = Array("OK", dbs) Else B_ListDatabases = Array(St(rc))
    Exit Function
fail:
    B_ListDatabases = Array(VbaErr("ListDatabases"))
End Function

Public Function B_SubVar(ByVal sheetName As String, ByVal app As String, ByVal db As String, ByVal varName As String) As Variant
    On Error GoTo fail
    Dim names As Variant, vals As Variant, rc As Long
    rc = HypGetSubstitutionVariable(sheetName, app, db, varName, names, vals)
    If rc = 0 Then B_SubVar = Array("OK", names, vals) Else B_SubVar = Array(St(rc))
    Exit Function
fail:
    B_SubVar = Array(VbaErr("GetSubstitutionVariable"))
End Function

Public Function B_SheetInfo(ByVal sheetName As String) As Variant
    On Error GoTo fail
    Dim n As Variant, v As Variant, rc As Long
    rc = HypGetSheetInfo(sheetName, n, v)
    If rc = 0 Then B_SheetInfo = Array("OK", n, v) Else B_SheetInfo = Array(St(rc))
    Exit Function
fail:
    B_SheetInfo = Array(VbaErr("GetSheetInfo"))
End Function

Public Function B_IsFreeForm(ByVal sheetName As String) As String
    On Error GoTo fail
    B_IsFreeForm = CStr(HypIsFreeForm(sheetName))
    Exit Function
fail:
    B_IsFreeForm = VbaErr("IsFreeForm")
End Function

Public Function B_LastError() As String
    On Error GoTo fail
    Dim code As Variant, emsg As Variant, edesc As Variant
    HypGetLastError code, emsg, edesc
    B_LastError = CStr(code) & "|" & CStr(emsg) & "|" & CStr(edesc)
    Exit Function
fail:
    B_LastError = VbaErr("LastError")
End Function

' ---------- drill-through (read-only: lists the reports on a cell; NEVER executes one - that opens a browser) ----------
Public Function B_DrillReports(ByVal sheetName As String, ByVal addr As String) As Variant
    Dim ids As Variant, nms As Variant, urls As Variant, tmpls As Variant, typs As Variant, rc As Long
    On Error GoTo fail
    rc = HypGetDrillThroughReports(sheetName, Ws(sheetName).Range(addr), ids, nms, urls, tmpls, typs)
    B_DrillReports = Array(St(rc), nms, urls, typs)
    Exit Function
fail:
    B_DrillReports = Array(VbaErr("B_DrillReports"), Empty, Empty, Empty)
End Function
