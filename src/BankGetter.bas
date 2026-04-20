Attribute VB_Name = "BankGetter"
Option Explicit

' Module-level state shared across the three fetch helpers
Private gBankWs As Worksheet
Private gBankRow As Long       ' accounts + investments bölümü (B sütunu)
Private gCardCol As Long       ' aktif kartın başlangıç sütunu (G, L, Q...)
Private gCardRow As Long       ' aktif kartın aktif satırı

Private Const BANK_COL_ACCOUNT As Integer = 2  ' B
Private Const BANK_COL_DATE    As Integer = 3  ' C
Private Const BANK_COL_DESC    As Integer = 4  ' D
Private Const BANK_COL_AMOUNT  As Integer = 5  ' E
Private Const BANK_COL_RAW     As Integer = 6  ' F
Private Const CARD_COL_OFFSET  As Integer = 5  ' Her kart 5 sütun sağa

' Entry point for the data-driven engine. bankID must match a BankID in the BANKS sheet.
' Run BankGetterSetup.CreateBANKSSheet once to initialize the BANKS sheet before using this.
Public Sub BankGetterRun(Optional bankID As String = "TEB")
    BankGetterEngine.RunBank bankID
End Sub

Public Sub BankGetterTEB()
    LogManager.LogInfo "=== BankGetter: TEB Data Fetch Started ==="
    On Error GoTo ErrorHandler

    Set gBankWs = Application.ActiveWorkbook.Worksheets("Bank_Info")
    gBankWs.Cells.Delete
    gBankRow = 2
    gCardCol = BANK_COL_ACCOUNT + CARD_COL_OFFSET  ' G (7) — ilk kart
    gCardRow = 2

    Dim chrome As stdChrome
    Dim hwnd As LongPtr
    Call BringWindowToFront.GetHandleFromPartialCaption(hwnd, "CEPTETEB")
    Dim extWin As stdWindow
    Set extWin = stdWindow.CreateFromHwnd(hwnd)
    Set chrome = stdChrome.CreateFromExisting(extWin)

    LogManager.LogInfo "Fetching Account Transactions..."
    Call BankGetter_FetchAccounts(chrome)

    LogManager.LogInfo "Fetching Investment Transactions..."
    Call BankGetter_FetchInvestments(chrome)

    LogManager.LogInfo "Fetching Credit Card Transactions..."
    Call BankGetter_FetchCards(chrome)

    SortAndFormatBankInfo

    gBankWs.Range("B1").Select
    LogManager.LogInfo "BankGetter: TEB Data Fetch Completed. " & (gBankRow - 2) & " rows."
    MsgBox "bitti (" & (gBankRow - 2) & " işlem)", vbInformation, "BankGetter"
    Exit Sub

ErrorHandler:
    LogManager.LogError "BankGetter failed: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical, "BankGetter Error"
End Sub

Private Sub BankGetter_FetchAccounts(chrome As stdChrome)
    On Error GoTo ErrorHandler

    Dim detailsLinks As Collection
    Dim detailItem As Variant
    Dim detailNum As Integer
    Dim childi As Variant, itm As Variant
    Dim q As Integer, j As Integer
    Dim skipLineCount As Integer
    Dim dateVal As Date, descVal As String, amountVal As Double
    Dim hasDate As Boolean

    With chrome
        .accMain.PrintDescTexts
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hesaplar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        For q = 1 To 4
            Set detailsLinks = Nothing
            Call .AwaitForAccElement(stdLambda.Create("$1.Name like ""Detay"" and $1.Role = ""ROLE_LINK"""))
            Set detailsLinks = .accMain.FindAll(stdLambda.Create("$1.Name = ""Detay"" and $1.Role = ""ROLE_LINK"""))
            detailNum = detailNum + 1
            Set detailItem = detailsLinks.item(q)
            Call detailItem.DoDefaultAction
            Call .AwaitForAccElement(stdLambda.Create("$1.Name like ""Hesap " & ChrW(304) & ChrW(351) & "lemleri"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
            Call .AwaitForAccElement(stdLambda.Create("$1.Name like ""Hesap Hareketleri"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
            Call .AwaitForAccElement(stdLambda.Create("$1.Name like ""1 Ay"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
            Call .AwaitForAccElement(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""))

            skipLineCount = 1
            For Each childi In .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""")).children
                If skipLineCount > 0 Then
                    skipLineCount = skipLineCount - 1
                Else
                    hasDate = False: j = 0: descVal = "": amountVal = 0
                    For Each itm In childi.children
                        j = j + 1
                        Dim cellText1 As String
                        cellText1 = SafeChildText(itm)
                        Select Case j
                            Case 1
                                On Error Resume Next
                                dateVal = CDate(Replace(Replace(cellText1, "/", "."), "(*)", ""))
                                hasDate = (Err.Number = 0): Err.Clear
                                On Error GoTo 0
                            Case 4: descVal = cellText1
                            Case 5
                                On Error Resume Next
                                amountVal = CDbl(cellText1)
                                If Err.Number <> 0 Then amountVal = 0: Err.Clear
                                On Error GoTo 0
                        End Select
                    Next itm
                    If hasDate Then
                        gBankWs.Cells(gBankRow, BANK_COL_ACCOUNT).value = "Hesap-" & detailNum
                        gBankWs.Cells(gBankRow, BANK_COL_DATE).value = dateVal
                        gBankWs.Cells(gBankRow, BANK_COL_DESC).value = descVal
                        gBankWs.Cells(gBankRow, BANK_COL_AMOUNT).value = amountVal
                        gBankRow = gBankRow + 1
                    End If
                End If
            Next childi
            Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hesaplar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Next q
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With

    LogManager.LogInfo "Account transactions fetched"
    Exit Sub
ErrorHandler:
    LogManager.LogError "BankGetter_FetchAccounts failed: " & Err.Description
End Sub

Private Sub BankGetter_FetchInvestments(chrome As stdChrome)
    On Error GoTo ErrorHandler

    Dim dateStr2 As String
    Dim childi As Variant, itm As Variant
    Dim skipLineCount As Integer, j As Integer
    Dim investmentTable As stdAcc
    Dim exitLoop As Boolean
    Dim dateVal As Date, descVal As String, amountVal As Double
    Dim hasDate As Boolean

    With chrome
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Yat" & ChrW(305) & "r" & ChrW(305) & "mlar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hisse " & ChrW(304) & ChrW(351) & "lemleri"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = """ & ChrW(304) & ChrW(351) & "lemlerim"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hesap Ekstresi"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Tarih Aral" & ChrW(305) & ChrW(287) & "" & ChrW(305) & """ and $1.Role = ""ROLE_STATICTEXT"""))

        dateStr2 = Replace(CStr(Date - 31), ".", "/")
        If Len(dateStr2) = 9 Then dateStr2 = "0" & dateStr2

        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""..."" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select year"" and $1.Role = ""ROLE_COMBOBOX""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select year"" and $1.Role = ""ROLE_COMBOBOX""")). _
            AwaitForElement(stdLambda.Create("$1.Name = """ & Mid(dateStr2, 7) & """ and $1.Role = ""ROLE_LISTITEM""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select month"" and $1.Role = ""ROLE_COMBOBOX""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select month"" and $1.Role = ""ROLE_COMBOBOX""")). _
            AwaitForElement(stdLambda.Create("$1.Name = """ & CastMonthName(CInt(Mid(dateStr2, 4, 2))) & """ and $1.Role = ""ROLE_LISTITEM""")).DoDefaultAction
        Call .accMain.AwaitForElement(stdLambda.Create("$1.Name = """ & CStr(CInt(Mid(dateStr2, 1, 2))) & """ and $1.Role = ""ROLE_LINK""")).DoDefaultAction

        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Cari Hesap"" and $1.Role = ""ROLE_CELL"""))
        chrome.AwaitForAccElement(stdLambda.Create("$1.Name = ""Cari Hesap"" and $1.Role = ""ROLE_CELL""")).AwaitForElement(stdLambda.Create("$1.Name = """" $1.Role = ""ROLE_CHECKBUTTON""")).parent.children.item(2).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Devam"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction

        exitLoop = False
        Do Until exitLoop
            If Not .accMain.AwaitForElement(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""), , 10) Is Nothing Then
                Set investmentTable = .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""))
                skipLineCount = 1
                For Each childi In investmentTable.children
                    If skipLineCount > 0 Then
                        skipLineCount = skipLineCount - 1
                    Else
                        hasDate = False: j = 0: descVal = "": amountVal = 0
                        For Each itm In childi.children
                            j = j + 1
                            Dim cellText2 As String
                            cellText2 = SafeChildText(itm)
                            Select Case j
                                Case 1
                                    On Error Resume Next
                                    dateVal = CDate(Replace(Replace(cellText2, "/", "."), "(*)", ""))
                                    hasDate = (Err.Number = 0): Err.Clear
                                    On Error GoTo 0
                                Case 2
                                    On Error Resume Next
                                    amountVal = CDbl(cellText2)
                                    If Err.Number <> 0 Then amountVal = 0: Err.Clear
                                    On Error GoTo 0
                                Case 4: descVal = cellText2
                            End Select
                        Next itm
                        If hasDate Then
                            gBankWs.Cells(gBankRow, BANK_COL_ACCOUNT).value = "TEB Yat" & ChrW(305) & "r" & ChrW(305) & "m Hesab" & ChrW(305)
                            gBankWs.Cells(gBankRow, BANK_COL_DATE).value = dateVal
                            gBankWs.Cells(gBankRow, BANK_COL_DESC).value = descVal
                            gBankWs.Cells(gBankRow, BANK_COL_AMOUNT).value = amountVal
                            gBankRow = gBankRow + 1
                        End If
                    End If
                Next childi
            End If
            If Not .AwaitForAccElement(stdLambda.Create("$1.Name = ""Sonraki Sayfa"" and $1.Role = ""ROLE_PUSHBUTTON"""), , 3) Is Nothing Then
                Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Sonraki Sayfa"" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
            Else
                exitLoop = True
            End If
        Loop

        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With

    LogManager.LogInfo "Investment transactions fetched"
    Exit Sub
ErrorHandler:
    LogManager.LogError "BankGetter_FetchInvestments failed: " & Err.Description
End Sub

Private Sub BankGetter_FetchCards(chrome As stdChrome)
    On Error GoTo ErrorHandler

    Dim cardsArr As Variant
    Dim cardName As Variant
    Dim cardTable As stdAcc
    Dim childi As Variant, itm As Variant
    Dim skipLineCount As Integer, detailNum As Integer, j As Integer
    Dim dateVal As Date, descVal As String, amountVal As Double, rawVal As String
    Dim hasDate As Boolean

    cardsArr = Array("TEB BONUS CARD", "TEB SHE CARD")

    With chrome
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction

        For Each cardName In cardsArr
            detailNum = detailNum + 1
            gCardRow = 2  ' her kart kendi bloğunda row 2'den başlar
            Call .accMain.AwaitForElement(stdLambda.Create("$1.Name = """ & cardName & """ and $1.Role = ""ROLE_LINK""")).DoDefaultAction

            If Not .accMain.AwaitForElement(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""), , 10) Is Nothing Then
                Set cardTable = .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""))
                skipLineCount = 2
                For Each childi In cardTable.children
                    If skipLineCount > 0 Then
                        skipLineCount = skipLineCount - 1
                    Else
                        hasDate = False: j = 0: descVal = "": amountVal = 0: rawVal = ""
                        For Each itm In childi.children
                            j = j + 1
                            Dim cellText3 As String
                            cellText3 = SafeChildText(itm)
                            Select Case j
                                Case 1
                                    On Error Resume Next
                                    dateVal = CDate(Replace(Replace(cellText3, "/", "."), "(*)", ""))
                                    hasDate = (Err.Number = 0): Err.Clear
                                    On Error GoTo 0
                                Case 2: descVal = cellText3
                                Case 4
                                    On Error Resume Next
                                    amountVal = -1 * CDbl(cellText3)
                                    If Err.Number <> 0 Then amountVal = 0: Err.Clear
                                    On Error GoTo 0
                                Case 5: rawVal = cellText3
                            End Select
                        Next itm
                        If hasDate Then
                            gBankWs.Cells(gCardRow, gCardCol).value = "Kart-" & cardName
                            gBankWs.Cells(gCardRow, gCardCol + 1).value = dateVal
                            gBankWs.Cells(gCardRow, gCardCol + 2).value = descVal
                            gBankWs.Cells(gCardRow, gCardCol + 3).value = amountVal
                            If Len(rawVal) > 0 Then gBankWs.Cells(gCardRow, gCardCol + 4).value = "'" & rawVal
                            gCardRow = gCardRow + 1
                        End If
                    End If
                Next childi
            End If

            Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
            gCardCol = gCardCol + CARD_COL_OFFSET  ' sonraki kart için 5 sütun sağa
        Next cardName

        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With

    LogManager.LogInfo "Credit card transactions fetched: " & detailNum & " cards"
    Exit Sub
ErrorHandler:
    LogManager.LogError "BankGetter_FetchCards failed: " & Err.Description
End Sub

Private Sub SortAndFormatBankInfo()
    ' --- Hesaplar + Yatırımlar bölümü (B-F) ---
    Dim accLastRow As Long
    accLastRow = gBankRow - 1
    If accLastRow >= 2 Then
        FormatSection BANK_COL_ACCOUNT, accLastRow, "Hesap", "Tarih", _
                      "A" & ChrW(231) & ChrW(305) & "klama", "Tutar", ""
    End If

    ' --- Kart bölümleri (G-K, L-P, ...) ---
    Dim col As Long
    col = BANK_COL_ACCOUNT + CARD_COL_OFFSET  ' G (7) — ilk kart
    Do While col < gCardCol  ' gCardCol bir sonraki boş karta işaret eder
        Dim cardLastRow As Long
        cardLastRow = gBankWs.Cells(gBankWs.Rows.count, col + 1).End(xlUp).Row
        If cardLastRow >= 2 Then
            Dim cardLabel As String
            cardLabel = gBankWs.Cells(2, col).value  ' "Kart-TEB BONUS CARD" vb.
            FormatSection col, cardLastRow, cardLabel, "Tarih", _
                          "A" & ChrW(231) & ChrW(305) & "klama", "Tutar", "Ham Veri"
        End If
        col = col + CARD_COL_OFFSET
    Loop
End Sub

Private Sub FormatSection(baseCol As Long, lastRow As Long, _
                           hAccount As String, hDate As String, _
                           hDesc As String, hAmount As String, hRaw As String)
    ' Başlık satırı
    With gBankWs
        .Cells(1, baseCol).value = hAccount
        .Cells(1, baseCol + 1).value = hDate
        .Cells(1, baseCol + 2).value = hDesc
        .Cells(1, baseCol + 3).value = hAmount
        If Len(hRaw) > 0 Then .Cells(1, baseCol + 4).value = hRaw
        With .Range(.Cells(1, baseCol), .Cells(1, baseCol + 4))
            .Font.Bold = True
            .Interior.Color = RGB(68, 114, 196)
            .Font.Color = RGB(255, 255, 255)
        End With
    End With

    ' Tarihe göre azalan sıralama (en yeni üstte)
    Dim dataRange As Range
    Set dataRange = gBankWs.Range(gBankWs.Cells(2, baseCol), _
                                   gBankWs.Cells(lastRow, baseCol + 4))
    dataRange.Sort Key1:=gBankWs.Cells(2, baseCol + 1), _
                   Order1:=xlDescending, Header:=xlNo

    ' Tarih formatı ve sütun genişlikleri
    gBankWs.Columns(baseCol + 1).NumberFormat = "dd.mm.yyyy"
    gBankWs.Columns(baseCol).ColumnWidth = 22
    gBankWs.Columns(baseCol + 1).ColumnWidth = 12
    gBankWs.Columns(baseCol + 2).ColumnWidth = 42
    gBankWs.Columns(baseCol + 3).ColumnWidth = 14
    If Len(hRaw) > 0 Then gBankWs.Columns(baseCol + 4).ColumnWidth = 30
End Sub

' Safely reads the display text of an accessibility element's first child
Private Function SafeChildText(itm As Variant) As String
    On Error Resume Next
    SafeChildText = itm.children.item(1).name
    If Err.Number <> 0 Then SafeChildText = itm.name: Err.Clear
    On Error GoTo 0
End Function

Private Function CastMonthName(monthNum As Integer) As String
    CastMonthName = Array("Ocak", ChrW(350) & "ubat", "Mart", "Nisan", "May" & ChrW(305) & "s", _
                          "Haziran", "Temmuz", "A" & ChrW(287) & "ustos", _
                          "Eyl" & ChrW(252) & "l", "Ekim", _
                          "Kas" & ChrW(305) & "m", "Aral" & ChrW(305) & "k")(monthNum - 1)
End Function


