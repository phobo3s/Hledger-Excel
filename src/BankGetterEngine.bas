Attribute VB_Name = "BankGetterEngine"
Option Explicit

Private Type StepRow
    BankID As String
    Seq As Integer
    StepType As String
    Predicate As String
    Param1 As String
    Param2 As String
    Param3 As String
    AccountName As String
    DateCol As Integer
    DescCol As Integer
    AmountCol As Integer
    RawCol As Integer
    SkipRows As Integer
    AmountSign As Integer
    LoopLabel As String
    HookName As String
End Type

Private Type EngineState
    chrome As stdChrome
    loopVar As String
    writeRow As Long
    originCell As Range
End Type

' Column index constants for BANKS sheet (1-based)
Private Const COL_BANKID As Integer = 1
Private Const COL_SEQ As Integer = 2
Private Const COL_STEPTYPE As Integer = 3
Private Const COL_PREDICATE As Integer = 4
Private Const COL_PARAM1 As Integer = 5
Private Const COL_PARAM2 As Integer = 6
Private Const COL_PARAM3 As Integer = 7
Private Const COL_ACCOUNTNAME As Integer = 8
Private Const COL_DATECOL As Integer = 9
Private Const COL_DESCCOL As Integer = 10
Private Const COL_AMOUNTCOL As Integer = 11
Private Const COL_RAWCOL As Integer = 12
Private Const COL_SKIPROWS As Integer = 13
Private Const COL_AMOUNTSIGN As Integer = 14
Private Const COL_LOOPLABEL As Integer = 15
Private Const COL_HOOKNAME As Integer = 16

Public Sub RunBank(bankID As String)
    LogManager.LogInfo "=== BankGetterEngine: Starting bank '" & bankID & "' ==="
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Set ws = Application.ActiveWorkbook.Worksheets("Bank_Info")
    ws.Activate
    ws.Cells.Delete
    ws.Range("B2").Select

    Dim steps() As StepRow
    If Not LoadSteps(bankID, steps) Then
        MsgBox "No steps found for bank '" & bankID & "' in BANKS sheet.", vbExclamation, "BankGetterEngine"
        Exit Sub
    End If

    Dim state As EngineState
    state.writeRow = 0
    Set state.originCell = ws.Range("B2")

    ExecuteSteps steps, state, 0, UBound(steps)

    FormatBankInfo ws, state.originCell, state.writeRow

    ws.Range("B1").Select
    LogManager.LogInfo "=== BankGetterEngine: '" & bankID & "' completed. " & state.writeRow & " rows. ==="
    MsgBox "bitti (" & state.writeRow & " işlem)", vbInformation, "BankGetterEngine"
    Exit Sub

ErrorHandler:
    LogManager.LogError "BankGetterEngine.RunBank failed for '" & bankID & "': " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "BankGetterEngine Error"
End Sub

Private Function LoadSteps(bankID As String, ByRef steps() As StepRow) As Boolean
    Dim banksWs As Worksheet
    On Error Resume Next
    Set banksWs = Application.ActiveWorkbook.Worksheets("BANKS")
    On Error GoTo 0
    If banksWs Is Nothing Then
        LogManager.LogError "BANKS sheet not found. Run BankGetterSetup.CreateBANKSSheet first."
        LoadSteps = False
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = banksWs.Cells(banksWs.Rows.Count, COL_BANKID).End(xlUp).Row

    Dim count As Integer
    count = 0
    Dim r As Long
    For r = 2 To lastRow
        If UCase(Trim(banksWs.Cells(r, COL_BANKID).value)) = UCase(Trim(bankID)) Then
            count = count + 1
        End If
    Next r

    If count = 0 Then
        LoadSteps = False
        Exit Function
    End If

    ReDim steps(0 To count - 1)
    Dim idx As Integer
    idx = 0
    For r = 2 To lastRow
        If UCase(Trim(banksWs.Cells(r, COL_BANKID).value)) = UCase(Trim(bankID)) Then
            With steps(idx)
                .BankID = Trim(banksWs.Cells(r, COL_BANKID).value)
                .Seq = CInt(banksWs.Cells(r, COL_SEQ).value)
                .StepType = UCase(Trim(banksWs.Cells(r, COL_STEPTYPE).value))
                .Predicate = Trim(banksWs.Cells(r, COL_PREDICATE).value)
                .Param1 = Trim(banksWs.Cells(r, COL_PARAM1).value)
                .Param2 = Trim(banksWs.Cells(r, COL_PARAM2).value)
                .Param3 = Trim(banksWs.Cells(r, COL_PARAM3).value)
                .AccountName = Trim(banksWs.Cells(r, COL_ACCOUNTNAME).value)
                .DateCol = SafeInt(banksWs.Cells(r, COL_DATECOL).value)
                .DescCol = SafeInt(banksWs.Cells(r, COL_DESCCOL).value)
                .AmountCol = SafeInt(banksWs.Cells(r, COL_AMOUNTCOL).value)
                .RawCol = SafeInt(banksWs.Cells(r, COL_RAWCOL).value)
                .SkipRows = SafeInt(banksWs.Cells(r, COL_SKIPROWS).value)
                .AmountSign = SafeIntDefault(banksWs.Cells(r, COL_AMOUNTSIGN).value, 1)
                .LoopLabel = Trim(banksWs.Cells(r, COL_LOOPLABEL).value)
                .HookName = Trim(banksWs.Cells(r, COL_HOOKNAME).value)
            End With
            idx = idx + 1
        End If
    Next r

    LoadSteps = True
End Function

Private Sub ExecuteSteps(steps() As StepRow, state As EngineState, fromIdx As Long, toIdx As Long)
    Dim i As Long
    i = fromIdx
    Do While i <= toIdx
        Dim st As StepRow
        st = steps(i)
        LogManager.LogDebug "Step " & st.Seq & ": " & st.StepType
        Select Case st.StepType
            Case "ATTACH_WINDOW": ExecAttachWindow st, state
            Case "NAVIGATE":      ExecNavigate st, state
            Case "CLICK":         ExecClick st, state
            Case "CLICK_IF_EXISTS": ExecClickIfExists st, state
            Case "WAIT":          ExecWait st, state
            Case "SET_VALUE":     ExecSetValue st, state
            Case "EXTRACT_TABLE": ExecExtractTable st, state
            Case "CALL_HOOK":     ExecCallHook st, state
            Case "RESET_CURSOR":  ExecResetCursor st, state
            Case "LOOP_FOR_EACH": i = ExecLoopForEach(steps, state, i, toIdx)
            Case "LOOP_WHILE":    i = ExecLoopWhile(steps, state, i, toIdx)
            Case "LOOP_END":      ' consumed by loop handlers
        End Select
        i = i + 1
    Loop
End Sub

Private Sub ExecAttachWindow(st As StepRow, state As EngineState)
    Dim hwnd As LongPtr
    Call BringWindowToFront.GetHandleFromPartialCaption(hwnd, st.Param1)
    Dim extWin As stdWindow
    Set extWin = stdWindow.CreateFromHwnd(hwnd)
    Set state.chrome = stdChrome.CreateFromExisting(extWin)
    LogManager.LogInfo "Attached to window: " & st.Param1
End Sub

Private Sub ExecNavigate(st As StepRow, state As EngineState)
    state.chrome.Address = st.Param1
End Sub

Private Sub ExecClick(st As StepRow, state As EngineState)
    Dim pred As String
    pred = SubstituteVars(st.Predicate, state)
    Call state.chrome.AwaitForAccElement(stdLambda.Create(pred)).DoDefaultAction
End Sub

Private Sub ExecClickIfExists(st As StepRow, state As EngineState)
    Dim pred As String
    pred = SubstituteVars(st.Predicate, state)
    Dim timeout As Integer
    timeout = 3
    If Len(st.Param1) > 0 Then timeout = CInt(st.Param1)
    Dim el As stdAcc
    Set el = state.chrome.AwaitForAccElement(stdLambda.Create(pred), , timeout)
    If Not el Is Nothing Then el.DoDefaultAction
End Sub

Private Sub ExecWait(st As StepRow, state As EngineState)
    Dim pred As String
    pred = SubstituteVars(st.Predicate, state)
    Dim timeout As Integer
    timeout = -1
    If Len(st.Param1) > 0 Then timeout = CInt(st.Param1)
    Call state.chrome.AwaitForAccElement(stdLambda.Create(pred), , timeout)
End Sub

Private Sub ExecSetValue(st As StepRow, state As EngineState)
    Dim pred As String
    pred = SubstituteVars(st.Predicate, state)
    Dim val As String
    val = SubstituteVars(st.Param1, state)
    state.chrome.AwaitForAccElement(stdLambda.Create(pred)).value = val
End Sub

Private Sub ExecExtractTable(st As StepRow, state As EngineState)
    Dim pred As String
    pred = SubstituteVars(st.Predicate, state)
    Dim accountName As String
    accountName = SubstituteVars(st.AccountName, state)

    Call state.chrome.AwaitForAccElement(stdLambda.Create(pred))
    Dim tbl As stdAcc
    Set tbl = state.chrome.accMain.FindFirst(stdLambda.Create(pred))
    If tbl Is Nothing Then
        LogManager.LogWarning "EXTRACT_TABLE: table not found for predicate: " & pred
        Exit Sub
    End If

    Dim amountSign As Integer
    amountSign = st.AmountSign
    If amountSign = 0 Then amountSign = 1

    Dim skipLeft As Integer
    skipLeft = st.SkipRows
    Dim childi As Variant, itm As Variant
    Dim i As Long, j As Integer

    For Each childi In tbl.children
        If skipLeft > 0 Then
            skipLeft = skipLeft - 1
        Else
            Dim rowDate As Variant, rowDesc As String, rowAmount As Double, rowRaw As String
            rowDate = Empty: rowDesc = "": rowAmount = 0: rowRaw = ""
            j = 0
            For Each itm In childi.children
                j = j + 1
                Dim cellText As String
                On Error Resume Next
                cellText = itm.children.item(1).name
                If Err.Number <> 0 Then cellText = itm.name: Err.Clear
                On Error GoTo 0

                If j = st.DateCol Then
                    On Error Resume Next
                    rowDate = CDate(Replace(Replace(cellText, "/", "."), "(*)", ""))
                    If Err.Number <> 0 Then rowDate = Empty: Err.Clear
                    On Error GoTo 0
                ElseIf j = st.DescCol Then
                    rowDesc = cellText
                ElseIf j = st.AmountCol Then
                    On Error Resume Next
                    rowAmount = CDbl(cellText) * amountSign
                    If Err.Number <> 0 Then rowAmount = 0: Err.Clear
                    On Error GoTo 0
                ElseIf st.RawCol > 0 And j = st.RawCol Then
                    rowRaw = cellText
                End If
            Next itm

            If Not IsEmpty(rowDate) Then
                Dim orig As Range
                Set orig = state.originCell
                orig.offset(state.writeRow, 0).value = accountName
                orig.offset(state.writeRow, 1).value = rowDate
                orig.offset(state.writeRow, 2).value = rowDesc
                orig.offset(state.writeRow, 3).value = rowAmount
                If Len(rowRaw) > 0 Then
                    orig.offset(state.writeRow, 4).value = "'" & rowRaw
                End If
                state.writeRow = state.writeRow + 1
            End If
        End If
    Next childi
End Sub

Private Sub ExecCallHook(st As StepRow, state As EngineState)
    If Len(st.HookName) = 0 Then
        LogManager.LogWarning "CALL_HOOK: HookName is empty"
        Exit Sub
    End If
    ' Position ActiveCell so hooks that write via ActiveCell land on the correct row
    state.originCell.offset(state.writeRow, 0).Select
    On Error GoTo HookError
    CallByName BankGetterHooks, st.HookName, VbMethod, state.chrome, st.Param1, st.Param2, st.Param3
    SyncWriteRow state  ' Hook may have written rows — re-scan to find new position
    Exit Sub
HookError:
    LogManager.LogError "CALL_HOOK '" & st.HookName & "' failed: " & Err.Description
End Sub

' Scans down from originCell to find how many rows have been written (date col is offset 1)
Private Sub SyncWriteRow(state As EngineState)
    Dim r As Long
    Do While Len(Trim(CStr(state.originCell.offset(r, 1).value))) > 0
        r = r + 1
    Loop
    state.writeRow = r
End Sub

' Adds headers, sorts by date descending, and auto-fits Bank_Info
Private Sub FormatBankInfo(ws As Worksheet, originCell As Range, dataRows As Long)
    If dataRows = 0 Then Exit Sub

    Dim hRow As Long
    hRow = originCell.Row - 1
    Dim hCol As Long
    hCol = originCell.Column

    With ws
        .Cells(hRow, hCol).value = "Hesap"
        .Cells(hRow, hCol + 1).value = "Tarih"
        .Cells(hRow, hCol + 2).value = "A" & ChrW(231) & ChrW(305) & "klama"
        .Cells(hRow, hCol + 3).value = "Tutar"
        .Cells(hRow, hCol + 4).value = "Ham Veri"
        .Rows(hRow).Font.Bold = True
        .Rows(hRow).Interior.Color = RGB(68, 114, 196)
        .Rows(hRow).Font.Color = RGB(255, 255, 255)
    End With

    ' Sort data range by date (offset 1) descending — newest first
    Dim dataRange As Range
    Set dataRange = originCell.Resize(dataRows, 5)
    dataRange.Sort Key1:=originCell.offset(0, 1), Order1:=xlDescending, Header:=xlNo

    ws.Columns(hCol).ColumnWidth = 22
    ws.Columns(hCol + 1).NumberFormat = "dd.mm.yyyy"
    ws.Columns(hCol + 1).ColumnWidth = 12
    ws.Columns(hCol + 2).ColumnWidth = 42
    ws.Columns(hCol + 3).ColumnWidth = 14
    ws.Columns(hCol + 4).ColumnWidth = 30
End Sub

Private Sub ExecResetCursor(st As StepRow, state As EngineState)
    Dim rOff As Long, cOff As Long
    If Len(st.Param1) > 0 Then rOff = CLng(st.Param1)
    If Len(st.Param2) > 0 Then cOff = CLng(st.Param2)
    Set state.originCell = state.originCell.offset(rOff, cOff)
    state.writeRow = 0
End Sub

Private Function ExecLoopForEach(steps() As StepRow, state As EngineState, startIdx As Long, toIdx As Long) As Long
    Dim endIdx As Long
    endIdx = FindLoopEnd(steps, steps(startIdx).LoopLabel, startIdx + 1, toIdx)
    If endIdx < 0 Then
        LogManager.LogError "LOOP_FOR_EACH: no matching LOOP_END for label '" & steps(startIdx).LoopLabel & "'"
        ExecLoopForEach = startIdx
        Exit Function
    End If

    Dim items() As String
    items = Split(steps(startIdx).Param1, ",")
    Dim item As Variant
    For Each item In items
        state.loopVar = Trim(CStr(item))
        ExecuteSteps steps, state, startIdx + 1, endIdx - 1
    Next item

    ExecLoopForEach = endIdx
End Function

Private Function ExecLoopWhile(steps() As StepRow, state As EngineState, startIdx As Long, toIdx As Long) As Long
    Dim endIdx As Long
    endIdx = FindLoopEnd(steps, steps(startIdx).LoopLabel, startIdx + 1, toIdx)
    If endIdx < 0 Then
        LogManager.LogError "LOOP_WHILE: no matching LOOP_END for label '" & steps(startIdx).LoopLabel & "'"
        ExecLoopWhile = startIdx
        Exit Function
    End If

    Dim pred As String
    pred = SubstituteVars(steps(startIdx).Predicate, state)
    Dim timeout As Integer
    timeout = 3
    If Len(steps(startIdx).Param1) > 0 Then timeout = CInt(steps(startIdx).Param1)

    Do
        Dim el As stdAcc
        Set el = state.chrome.AwaitForAccElement(stdLambda.Create(pred), , timeout)
        If el Is Nothing Then Exit Do
        ExecuteSteps steps, state, startIdx + 1, endIdx - 1
        el.DoDefaultAction
    Loop

    ExecLoopWhile = endIdx
End Function

Private Function FindLoopEnd(steps() As StepRow, label As String, fromIdx As Long, toIdx As Long) As Long
    Dim depth As Integer
    depth = 0
    Dim i As Long
    For i = fromIdx To toIdx
        Select Case steps(i).StepType
            Case "LOOP_FOR_EACH", "LOOP_WHILE"
                If steps(i).LoopLabel = label Then depth = depth + 1
            Case "LOOP_END"
                If steps(i).LoopLabel = label Then
                    If depth = 0 Then
                        FindLoopEnd = i
                        Exit Function
                    End If
                    depth = depth - 1
                End If
        End Select
    Next i
    FindLoopEnd = -1
End Function

Private Function SubstituteVars(s As String, state As EngineState) As String
    SubstituteVars = Replace(s, "{LOOP_VAR}", state.loopVar)
End Function

Private Function SafeInt(v As Variant) As Integer
    If IsNumeric(v) Then SafeInt = CInt(v) Else SafeInt = 0
End Function

Private Function SafeIntDefault(v As Variant, defaultVal As Integer) As Integer
    If IsNumeric(v) And CDbl(v) <> 0 Then
        SafeIntDefault = CInt(v)
    Else
        SafeIntDefault = defaultVal
    End If
End Function
