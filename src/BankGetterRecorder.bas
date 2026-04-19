Attribute VB_Name = "BankGetterRecorder"
Option Explicit

' Attaches to any running app window and dumps its accessibility tree
' to the ACC_TREE sheet so you can discover element Names and Roles
' needed to write BANKS sheet rows.
'
' Usage:
'   1. Open the bank app and navigate to the page you want to automate.
'   2. Run:  BankGetterRecorder.DumpTree "CEPTETEB"
'      (use any partial text from the window title bar)
'   3. Open the ACC_TREE sheet. Each row is one element.
'      Columns: Level | Path | Name | Role | Description | DefaultAction
'   4. Find the button/link you want to click. Read its Name and Role.
'   5. Write a BANKS row:
'        StepType = CLICK
'        Predicate = $1.Name = "that name" and $1.Role = "that role"
'
' Tip: Filter ACC_TREE col D (Role) to "ROLE_LINK" or "ROLE_PUSHBUTTON"
'      to quickly see all clickable elements on the current page.
Public Sub DumpTree(Optional windowTitle As String = "")
    If Len(windowTitle) = 0 Then
        windowTitle = InputBox("Enter partial window title to attach to:", "BankGetterRecorder")
        If Len(windowTitle) = 0 Then Exit Sub
    End If

    Dim hwnd As LongPtr
    Call BringWindowToFront.GetHandleFromPartialCaption(hwnd, windowTitle)
    If hwnd = 0 Then
        MsgBox "Window not found: " & windowTitle, vbExclamation, "BankGetterRecorder"
        Exit Sub
    End If

    Dim extWin As stdWindow
    Set extWin = stdWindow.CreateFromHwnd(hwnd)
    Dim chrome As stdChrome
    Set chrome = stdChrome.CreateFromExisting(extWin)

    Dim ws As Worksheet
    ws = GetOrCreateSheet("ACC_TREE")
    ws.Cells.Delete

    ' Headers
    ws.Cells(1, 1).value = "Level"
    ws.Cells(1, 2).value = "Path"
    ws.Cells(1, 3).value = "Name"
    ws.Cells(1, 4).value = "Role"
    ws.Cells(1, 5).value = "Description"
    ws.Cells(1, 6).value = "DefaultAction"
    ws.Cells(1, 7).value = "Value"
    ws.Rows(1).Font.Bold = True

    ' Freeze pane and auto-filter
    ws.Activate
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1:G1").AutoFilter

    Dim nextRow As Long
    nextRow = 2
    WalkTree chrome.accMain, 0, "root", ws, nextRow

    ws.Columns("A:G").AutoFit
    ws.Range("A1").Select
    MsgBox "Done. " & (nextRow - 2) & " elements found in ACC_TREE sheet." & vbNewLine & _
           "Tip: Filter column D (Role) to ROLE_LINK or ROLE_PUSHBUTTON to see clickable items.", _
           vbInformation, "BankGetterRecorder"
End Sub

' Dumps only ROLE_LINK and ROLE_PUSHBUTTON elements — faster for finding clickable items.
Public Sub DumpClickable(Optional windowTitle As String = "")
    If Len(windowTitle) = 0 Then
        windowTitle = InputBox("Enter partial window title:", "BankGetterRecorder")
        If Len(windowTitle) = 0 Then Exit Sub
    End If

    Dim hwnd As LongPtr
    Call BringWindowToFront.GetHandleFromPartialCaption(hwnd, windowTitle)
    If hwnd = 0 Then
        MsgBox "Window not found: " & windowTitle, vbExclamation
        Exit Sub
    End If

    Dim extWin As stdWindow
    Set extWin = stdWindow.CreateFromHwnd(hwnd)
    Dim chrome As stdChrome
    Set chrome = stdChrome.CreateFromExisting(extWin)

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet("ACC_TREE")
    ws.Cells.Delete

    ws.Cells(1, 1).value = "Level"
    ws.Cells(1, 2).value = "Path"
    ws.Cells(1, 3).value = "Name"
    ws.Cells(1, 4).value = "Role"
    ws.Cells(1, 5).value = "BANKS Predicate (copy-paste ready)"
    ws.Rows(1).Font.Bold = True

    Dim nextRow As Long
    nextRow = 2
    WalkClickable chrome.accMain, 0, "root", ws, nextRow

    ws.Columns("A:E").AutoFit
    ws.Range("E2").Select
    MsgBox "Done. " & (nextRow - 2) & " clickable elements found.", vbInformation, "BankGetterRecorder"
End Sub

Private Sub WalkTree(acc As stdAcc, level As Long, path As String, ws As Worksheet, ByRef nextRow As Long)
    If acc Is Nothing Then Exit Sub
    Dim child As stdAcc
    Dim i As Long
    i = 0
    For Each child In acc.children
        i = i + 1
        Dim childPath As String
        childPath = path & "." & i
        Dim nm As String, rl As String, desc As String, da As String, val As String
        On Error Resume Next
        nm = child.name
        rl = child.Role
        desc = child.Description
        da = child.DefaultAction
        val = child.value
        On Error GoTo 0
        ws.Cells(nextRow, 1).value = level + 1
        ws.Cells(nextRow, 2).value = childPath
        ws.Cells(nextRow, 3).value = nm
        ws.Cells(nextRow, 4).value = rl
        ws.Cells(nextRow, 5).value = desc
        ws.Cells(nextRow, 6).value = da
        ws.Cells(nextRow, 7).value = val
        nextRow = nextRow + 1
        WalkTree child, level + 1, childPath, ws, nextRow
    Next child
End Sub

Private Sub WalkClickable(acc As stdAcc, level As Long, path As String, ws As Worksheet, ByRef nextRow As Long)
    If acc Is Nothing Then Exit Sub
    Dim child As stdAcc
    Dim i As Long
    i = 0
    For Each child In acc.children
        i = i + 1
        Dim childPath As String
        childPath = path & "." & i
        Dim nm As String, rl As String
        On Error Resume Next
        nm = child.name
        rl = child.Role
        On Error GoTo 0
        If rl = "ROLE_LINK" Or rl = "ROLE_PUSHBUTTON" Or rl = "ROLE_MENUITEM" Then
            Dim pred As String
            pred = "$1.Name = """ & nm & """ and $1.Role = """ & rl & """"
            ws.Cells(nextRow, 1).value = level + 1
            ws.Cells(nextRow, 2).value = childPath
            ws.Cells(nextRow, 3).value = nm
            ws.Cells(nextRow, 4).value = rl
            ws.Cells(nextRow, 5).value = pred
            nextRow = nextRow + 1
        End If
        WalkClickable child, level + 1, childPath, ws, nextRow
    Next child
End Sub

Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Application.ActiveWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = Application.ActiveWorkbook.Worksheets.Add( _
            After:=Application.ActiveWorkbook.Worksheets(Application.ActiveWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
