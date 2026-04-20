Attribute VB_Name = "BankGetterRecorder"
Option Explicit

' Real-time click recorder for the BANKS sheet.
' Records Name, Role, Description and a ready-made predicate for every
' element you click in the target app — no manual tree inspection needed.
'
' Usage:
'   1. Open the bank app (CEPTETEB, Garanti browser tab, etc.)
'   2. Run:  BankGetterRecorder.StartRecording "CEPTETEB"
'      (use any partial text from the window title bar)
'   3. Switch to the bank app and click through your normal workflow.
'   4. Press ESC (in any window) to stop recording.
'   5. Open the RECORDING sheet — every click is captured with:
'        Step# | Role | Name | Description | Predicate | Suggested StepType
'   6. Run:  BankGetterRecorder.ConvertToBANKS "MyBank"
'      to append the recorded steps to the BANKS sheet.
'   7. Review BANKS sheet, adjust seq numbers, fill EXTRACT_TABLE columns.

Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As RECORDER_POINT) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)

Private Type RECORDER_POINT
    x As Long
    y As Long
End Type

' VK codes
Private Const VK_LBUTTON As Long = &H1
Private Const VK_ESCAPE As Long = &H1B

' Module-level recording state
Private gIsRecording As Boolean
Private gWasButtonDown As Boolean
Private gRecordRow As Long
Private gRecordWs As Worksheet

Public Sub StartRecording(Optional windowTitle As String = "")
    If Len(windowTitle) = 0 Then
        windowTitle = InputBox("Pencere başlığının bir kısmını gir (örn: CEPTETEB, Garanti):", _
                               "BankGetterRecorder")
        If Len(windowTitle) = 0 Then Exit Sub
    End If

    ' Prepare RECORDING sheet
    Set gRecordWs = GetOrCreateSheet("RECORDING")
    gRecordWs.Cells.Delete
    WriteHeader gRecordWs
    gRecordRow = 2

    gIsRecording = True
    gWasButtonDown = False

    MsgBox "Kayıt başladı!" & vbNewLine & vbNewLine & _
           "Banka uygulamasına geç ve normal şekilde tıkla." & vbNewLine & _
           "Bitirmek için herhangi bir yerde ESC'ye bas." & vbNewLine & vbNewLine & _
           "Bu mesajı kapat, ardından banka uygulamasına geç.", _
           vbInformation, "BankGetterRecorder — Kayıt"

    ' Recording loop: poll at 50ms intervals
    Do While gIsRecording
        DoEvents
        sleep 50

        ' ESC anywhere stops recording
        If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Then
            gIsRecording = False
            Exit Do
        End If

        ' Detect left mouse button release (click completed)
        Dim isDown As Boolean
        isDown = (GetAsyncKeyState(VK_LBUTTON) And &H8000) <> 0

        If Not isDown And gWasButtonDown Then
            ' Read element under cursor at moment of release
            Dim pT As RECORDER_POINT
            GetCursorPos pT
            RecordAtPoint pT.x, pT.y
        End If

        gWasButtonDown = isDown
    Loop

    gRecordWs.Columns("A:F").AutoFit
    gRecordWs.activate
    MsgBox "Kayıt tamamlandı! " & (gRecordRow - 2) & " adım kaydedildi." & vbNewLine & _
           "RECORDING sheet'ini incele, sonra ConvertToBANKS çalıştır.", _
           vbInformation, "BankGetterRecorder"
End Sub

Public Sub StopRecording()
    gIsRecording = False
End Sub

' Appends recorded steps from RECORDING sheet to BANKS sheet for the given bankID.
' Run after reviewing and cleaning up the RECORDING sheet.
Public Sub ConvertToBANKS(Optional bankID As String = "")
    If Len(bankID) = 0 Then
        bankID = InputBox("Bu kayıt için BankID gir (örn: Garanti, TEB, Akbank):", _
                          "ConvertToBANKS")
        If Len(bankID) = 0 Then Exit Sub
    End If

    Dim recWs As Worksheet
    On Error Resume Next
    Set recWs = Application.ActiveWorkbook.Worksheets("RECORDING")
    On Error GoTo 0
    If recWs Is Nothing Then
        MsgBox "RECORDING sheet bulunamadı. Önce StartRecording çalıştır.", vbExclamation
        Exit Sub
    End If

    Dim banksWs As Worksheet
    Set banksWs = GetOrCreateSheet("BANKS")

    ' Find next available row in BANKS and highest existing seq for this bankID
    Dim banksLastRow As Long
    banksLastRow = banksWs.Cells(banksWs.Rows.count, 1).End(xlUp).Row
    If banksLastRow = 1 Then
        ' Empty BANKS sheet — write headers
        BankGetterSetup.CreateBANKSHeaders banksWs
        banksLastRow = 1
    End If

    ' Find max seq already used for this bankID
    Dim maxSeq As Long
    maxSeq = 0
    Dim br As Long
    For br = 2 To banksLastRow
        If UCase(Trim(banksWs.Cells(br, 1).value)) = UCase(Trim(bankID)) Then
            Dim s As Long
            s = CLng(0 & banksWs.Cells(br, 2).value)
            If s > maxSeq Then maxSeq = s
        End If
    Next br

    ' Copy from RECORDING to BANKS
    Dim recLastRow As Long
    recLastRow = recWs.Cells(recWs.Rows.count, 1).End(xlUp).Row

    Dim addedCount As Long
    addedCount = 0
    Dim rr As Long
    For rr = 2 To recLastRow
        Dim stepType As String
        Dim predicate As String
        Dim nm As String
        Dim rl As String
        stepType = Trim(recWs.Cells(rr, 6).value)  ' Suggested StepType
        predicate = Trim(recWs.Cells(rr, 5).value)  ' Predicate
        nm = Trim(recWs.Cells(rr, 3).value)
        rl = Trim(recWs.Cells(rr, 2).value)
        If Len(stepType) = 0 Then stepType = "CLICK"

        maxSeq = maxSeq + 10
        banksLastRow = banksLastRow + 1

        banksWs.Cells(banksLastRow, 1).value = bankID
        banksWs.Cells(banksLastRow, 2).value = maxSeq
        banksWs.Cells(banksLastRow, 3).value = stepType
        banksWs.Cells(banksLastRow, 4).value = predicate
        ' For EXTRACT_TABLE steps leave other columns empty (user fills DateCol etc.)
        banksWs.Cells(banksLastRow, 17).value = "Recorded from: " & nm & " (" & rl & ")"
        addedCount = addedCount + 1
    Next rr

    banksWs.Columns("A:Q").AutoFit
    banksWs.activate
    MsgBox addedCount & " adım BANKS sheet'ine eklendi (BankID=" & bankID & ")." & vbNewLine & _
           "EXTRACT_TABLE satırları için DateCol, DescCol, AmountCol, SkipRows sütunlarını doldurmayı unutma.", _
           vbInformation, "ConvertToBANKS"
End Sub

Private Sub RecordAtPoint(x As Long, y As Long)
    Dim acc As stdAcc
    On Error Resume Next
    Set acc = stdAcc.CreateFromPoint(x, y)
    On Error GoTo 0
    If acc Is Nothing Then Exit Sub

    Dim nm As String, rl As String, desc As String, da As String
    On Error Resume Next
    nm = acc.name
    rl = acc.Role
    desc = acc.Description
    da = acc.DefaultAction
    On Error GoTo 0

    ' Skip container/background elements that aren't actionable
    Select Case rl
        Case "ROLE_CLIENT", "ROLE_WINDOW", "ROLE_PANE", "ROLE_DOCUMENT", _
             "ROLE_SCROLLBAR", "ROLE_GRIP", "ROLE_BORDER"
            Exit Sub
    End Select

    ' Skip if name is empty (usually non-interactive)
    If Len(Trim(nm)) = 0 And Len(Trim(desc)) = 0 Then Exit Sub

    ' Build predicate
    Dim pred As String
    If Len(nm) > 0 Then
        pred = "$1.Name = """ & nm & """ and $1.Role = """ & rl & """"
    ElseIf Len(desc) > 0 Then
        pred = "$1.Description like """ & desc & """ and $1.Role = """ & rl & """"
    Else
        pred = "$1.Role = """ & rl & """"
    End If

    ' Suggest step type
    Dim stepType As String
    Select Case rl
        Case "ROLE_LINK", "ROLE_PUSHBUTTON", "ROLE_MENUITEM", "ROLE_OUTLINEITEM"
            stepType = "CLICK"
        Case "ROLE_TABLE"
            stepType = "EXTRACT_TABLE"
        Case "ROLE_TEXT", "ROLE_COMBOBOX", "ROLE_DROPLIST"
            stepType = "SET_VALUE"
        Case "ROLE_LISTITEM"
            stepType = "CLICK"
        Case "ROLE_CHECKBUTTON", "ROLE_RADIOBUTTON"
            stepType = "CALL_HOOK"
        Case Else
            stepType = "CLICK"
    End Select

    ' Write to RECORDING sheet
    With gRecordWs
        .Cells(gRecordRow, 1).value = gRecordRow - 1
        .Cells(gRecordRow, 2).value = rl
        .Cells(gRecordRow, 3).value = nm
        .Cells(gRecordRow, 4).value = desc
        .Cells(gRecordRow, 5).value = pred
        .Cells(gRecordRow, 6).value = stepType
    End With
    gRecordRow = gRecordRow + 1

    LogManager.LogDebug "Recorded: [" & rl & "] " & nm
End Sub

Private Sub WriteHeader(ws As Worksheet)
    ws.Cells(1, 1).value = "Step#"
    ws.Cells(1, 2).value = "Role"
    ws.Cells(1, 3).value = "Name"
    ws.Cells(1, 4).value = "Description"
    ws.Cells(1, 5).value = "Predicate (BANKS'e yapıştır)"
    ws.Cells(1, 6).value = "Önerilen StepType"
    ws.Rows(1).Font.Bold = True
    ws.Range("A1:F1").Interior.Color = RGB(68, 114, 196)
    ws.Range("A1:F1").Font.Color = RGB(255, 255, 255)
End Sub

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Application.ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Application.ActiveWorkbook.Worksheets.Add( _
            After:=Application.ActiveWorkbook.Worksheets(Application.ActiveWorkbook.Worksheets.count))
        ws.name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function


