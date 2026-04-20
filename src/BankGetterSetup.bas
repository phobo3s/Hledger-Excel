Attribute VB_Name = "BankGetterSetup"
Option Explicit

' Creates the BANKS worksheet and populates it with TEB script rows.
' Run this once after importing the new modules into the workbook.
Public Sub CreateBANKSSheet()
    Dim wb As Workbook
    Set wb = Application.ActiveWorkbook

    ' Remove existing BANKS sheet if present
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets("BANKS")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "BANKS"
    CreateBANKSHeaders ws

    ' Populate TEB rows
    PopulateTEB ws

    ' Auto-fit columns
    ws.Columns("A:Q").AutoFit
    MsgBox "BANKS sheet created and populated with TEB script.", vbInformation, "BankGetterSetup"
End Sub

Public Sub CreateBANKSHeaders(ws As Worksheet)
    Dim headers As Variant
    headers = Array("BankID", "Seq", "StepType", "Predicate", "Param1", "Param2", "Param3", _
                    "AccountName", "DateCol", "DescCol", "AmountCol", "RawCol", _
                    "SkipRows", "AmountSign", "LoopLabel", "HookName", "Notes")
    Dim c As Integer
    For c = 0 To UBound(headers)
        ws.Cells(1, c + 1).value = headers(c)
        ws.Cells(1, c + 1).Font.Bold = True
    Next c
End Sub

Private Sub PopulateTEB(ws As Worksheet)
    Dim r As Long
    r = 2

    ' Helper: write a step row
    ' WriteStep ws, r, bankID, seq, stepType, predicate, p1, p2, p3, acct, dc, descc, ac, rc, sk, sign, lbl, hook, notes

    ' === ATTACH WINDOW ===
    WriteStep ws, r, "TEB", 10, "ATTACH_WINDOW", "", "CEPTETEB", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Find CEPTETEB desktop app window"
    r = r + 1

    ' === FETCH ACCOUNTS (positional, handled by hook) ===
    WriteStep ws, r, "TEB", 20, "CLICK", "$1.Name = ""Hesaplar"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Navigate to Accounts section"
    r = r + 1
    WriteStep ws, r, "TEB", 30, "CALL_HOOK", "", "4", "", "", "", 0, 0, 0, 0, 0, 0, "", "Hook_TEB_FetchAccounts", "Fetch 4 deposit accounts (positional link iteration)"
    r = r + 1

    ' === FETCH INVESTMENTS ===
    WriteStep ws, r, "TEB", 100, "CLICK", "$1.Name = ""Yat" & ChrW(305) & "r" & ChrW(305) & "mlar"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Navigate to Investments"
    r = r + 1
    WriteStep ws, r, "TEB", 110, "CLICK", "$1.Name = ""Hisse " & ChrW(304) & ChrW(351) & "lemleri"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Stock transactions"
    r = r + 1
    WriteStep ws, r, "TEB", 120, "CLICK", "$1.Name = """ & ChrW(304) & ChrW(351) & "lemlerim"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "My transactions"
    r = r + 1
    WriteStep ws, r, "TEB", 130, "CLICK", "$1.Name = ""Hesap Ekstresi"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Account statement"
    r = r + 1
    WriteStep ws, r, "TEB", 140, "WAIT", "$1.Name = ""Tarih Aral" & ChrW(305) & ChrW(287) & "" & ChrW(305) & """ and $1.Role = ""ROLE_STATICTEXT""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Wait for date range label"
    r = r + 1
    WriteStep ws, r, "TEB", 150, "CALL_HOOK", "", "31", "", "", "", 0, 0, 0, 0, 0, 0, "", "Hook_TEB_DatePicker", "Select start date 31 days ago"
    r = r + 1
    WriteStep ws, r, "TEB", 160, "CALL_HOOK", "", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "Hook_TEB_AccountFilter", "Select Cari Hesap checkbox"
    r = r + 1
    WriteStep ws, r, "TEB", 170, "CLICK", "$1.Name = ""Devam"" and $1.Role = ""ROLE_PUSHBUTTON""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Confirm and fetch"
    r = r + 1
    ' First page extract (always)
    WriteStep ws, r, "TEB", 180, "EXTRACT_TABLE", "$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""", "", "", "", "TEB Yat" & ChrW(305) & "r" & ChrW(305) & "m Hesab" & ChrW(305), 1, 4, 2, 0, 1, 1, "", "", "Extract investments page 1"
    r = r + 1
    ' Pagination loop
    WriteStep ws, r, "TEB", 190, "LOOP_WHILE", "$1.Name = ""Sonraki Sayfa"" and $1.Role = ""ROLE_PUSHBUTTON""", "3", "", "", "", 0, 0, 0, 0, 0, 0, "INV_PAGE", "", "While next page exists"
    r = r + 1
    WriteStep ws, r, "TEB", 200, "EXTRACT_TABLE", "$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""", "", "", "", "TEB Yat" & ChrW(305) & "r" & ChrW(305) & "m Hesab" & ChrW(305), 1, 4, 2, 0, 1, 1, "INV_PAGE", "", "Extract subsequent investment pages"
    r = r + 1
    WriteStep ws, r, "TEB", 210, "LOOP_END", "", "", "", "", "", 0, 0, 0, 0, 0, 0, "INV_PAGE", "", "Click next page and re-check"
    r = r + 1
    WriteStep ws, r, "TEB", 220, "CLICK", "$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Return home"
    r = r + 1

    ' === FETCH CREDIT CARDS — her kart ayrı sütun bloğunda (G, L, Q...) ===
    WriteStep ws, r, "TEB", 300, "RESET_CURSOR", "", "0", "5", "", "", 0, 0, 0, 0, 0, 0, "", "", "Kart bölümüne geç: originCell B2'den G2'ye"
    r = r + 1
    WriteStep ws, r, "TEB", 310, "CLICK", "$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Navigate to Cards section"
    r = r + 1
    WriteStep ws, r, "TEB", 320, "LOOP_FOR_EACH", "", "TEB BONUS CARD,TEB SHE CARD", "", "", "", 0, 0, 0, 0, 0, 0, "CARDS", "", "Iterate over card names"
    r = r + 1
    WriteStep ws, r, "TEB", 330, "CLICK", "$1.Name = ""{LOOP_VAR}"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Click card link"
    r = r + 1
    WriteStep ws, r, "TEB", 340, "EXTRACT_TABLE", "$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""", "", "", "", "Kart-{LOOP_VAR}", 1, 2, 4, 5, 2, -1, "", "", "Extract card transactions"
    r = r + 1
    WriteStep ws, r, "TEB", 350, "CLICK", "$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Back to cards list"
    r = r + 1
    WriteStep ws, r, "TEB", 360, "RESET_CURSOR", "", "0", "5", "", "", 0, 0, 0, 0, 0, 0, "CARDS", "", "Sonraki kart için 5 sütun sağa"
    r = r + 1
    WriteStep ws, r, "TEB", 370, "LOOP_END", "", "", "", "", "", 0, 0, 0, 0, 0, 0, "CARDS", "", "Next card iteration"
    r = r + 1
    WriteStep ws, r, "TEB", 380, "CLICK", "$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""", "", "", "", "", 0, 0, 0, 0, 0, 0, "", "", "Return home"
End Sub

Private Sub WriteStep(ws As Worksheet, r As Long, _
    bankID As String, seq As Integer, stepType As String, predicate As String, _
    p1 As String, p2 As String, p3 As String, accountName As String, _
    dateCol As Integer, descCol As Integer, amountCol As Integer, rawCol As Integer, _
    skipRows As Integer, amountSign As Integer, _
    loopLabel As String, hookName As String, notes As String)

    ws.Cells(r, 1).value = bankID
    ws.Cells(r, 2).value = seq
    ws.Cells(r, 3).value = stepType
    ws.Cells(r, 4).value = predicate
    ws.Cells(r, 5).value = p1
    ws.Cells(r, 6).value = p2
    ws.Cells(r, 7).value = p3
    ws.Cells(r, 8).value = accountName
    If dateCol > 0 Then ws.Cells(r, 9).value = dateCol
    If descCol > 0 Then ws.Cells(r, 10).value = descCol
    If amountCol > 0 Then ws.Cells(r, 11).value = amountCol
    If rawCol > 0 Then ws.Cells(r, 12).value = rawCol
    If skipRows > 0 Then ws.Cells(r, 13).value = skipRows
    If amountSign <> 0 Then ws.Cells(r, 14).value = amountSign
    ws.Cells(r, 15).value = loopLabel
    ws.Cells(r, 16).value = hookName
    ws.Cells(r, 17).value = notes
End Sub
