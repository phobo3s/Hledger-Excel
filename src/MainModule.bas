Attribute VB_Name = "MainModule"
Option Explicit
Dim commDict As Object

'possible Ledger column headers
Dim dateCol As Integer
Dim transCodeCol As Integer
Dim descriptionCol As Integer
Dim notesCol As Integer
Dim currencyCol As Integer
Dim operationCol As Integer
Dim tagCol As Integer
Dim accNameCol As Integer
Dim amountCol As Integer
Dim rateCol As Integer
Dim reconCol As Integer

Private Type CommodityObject
    Type As String
    Amount As Double
    Ticker As String
    Price As Double
    Currency As String
    OldBuyPrice As Double
End Type

' Paths and constants moved to Config.bas
    
Sub CreateAllFilesAKATornado()
    LogManager.LogInfo "=== CreateAllFilesAKATornado Started ==="

    'aggregate all sub account pages
    Call AggregateAccounts(PopulateAccountsSheets)
    'main hledger file creator module
    Call ExportHledgerFile
    Dim sh As Object
    Set sh = CreateObject("Wscript.Shell")
    Call sh.Run("cmd.exe /u /c chcp 65001 && cd /d """ & ThisWorkbook.path & """ && hledger-ui -f ./Main.hledger -w -3 -X TRY --infer-market-prices -E --theme=terminal", 1, False)
    
End Sub

Private Sub PopulateColumnHeaderIndexes(pageName As String)
    Dim sh As Worksheet
    Dim shName As String
    
    dateCol = 0
    transCodeCol = 0
    descriptionCol = 0
    notesCol = 0
    currencyCol = 0
    operationCol = 0
    tagCol = 0
    accNameCol = 0
    amountCol = 0
    rateCol = 0
    reconCol = 0
        
    For Each sh In ThisWorkbook.Worksheets
        If sh.codeName = pageName Then shName = sh.name: Exit For
    Next sh
    On Error Resume Next
    Dim c As Variant
    With ThisWorkbook.Worksheets(shName)
        'check
        Set c = .Rows(1).find("Date", LookAt:=xlWhole)
        If c Is Nothing Then Err.Raise 1001, , "Date column missing"
        dateCol = .Cells(1, 1).EntireRow.find("Date").Column
        transCodeCol = .Cells(1, 1).EntireRow.find("Transaction Code").Column
        descriptionCol = .Cells(1, 1).EntireRow.find("Payee|Note").Column
        notesCol = .Cells(1, 1).EntireRow.find("Notes").Column
        currencyCol = .Cells(1, 1).EntireRow.find("Commodity/Currency").Column
        operationCol = .Cells(1, 1).EntireRow.find("Operation").Column
        tagCol = .Cells(1, 1).EntireRow.find("Tag/Note").Column
        accNameCol = .Cells(1, 1).EntireRow.find("Full Account Name").Column
        amountCol = .Cells(1, 1).EntireRow.find("Amount").Column
        rateCol = .Cells(1, 1).EntireRow.find("Rate/Price").Column
        reconCol = .Cells(1, 1).EntireRow.find("Reconciliation").Column
    End With
    On Error GoTo 0
End Sub
Private Function PopulateAccountsSheets() As Variant
    
    Dim sh As Worksheet
    Dim shName As String
    Dim result() As Variant
    Dim i As Long
    For Each sh In ThisWorkbook.Worksheets
        If sh.Tab.Color = Config.COLOR_LIGHT_GREEN Then 'ligth green
            ReDim Preserve result(0 To i)
            result(UBound(result)) = sh.name
            i = i + 1
        Else
        End If
    Next sh
    PopulateAccountsSheets = result
    
End Function
Private Sub AggregateAccounts(accountsArr As Variant)

    Dim lastRowNum As Long
    lastRowNum = MAIN_LEDGER.Cells(MAIN_LEDGER.Rows.Count, 8).End(xlUp).Row
    If lastRowNum > 1 Then MAIN_LEDGER.Range("A2").Resize(lastRowNum - 1, 13).value = ""
    If lastRowNum > 1 Then MAIN_LEDGER.Range("A2").Resize(lastRowNum - 1, 13).Interior.ColorIndex = -4142
    Dim sh As Worksheet
    Dim i As Long
    For i = LBound(accountsArr) To UBound(accountsArr)
        Set sh = Application.ActiveWorkbook.Worksheets(accountsArr(i))
        lastRowNum = sh.Cells(sh.Cells.Rows.Count, sh.Range("H2").Column).End(xlUp).Row
        If lastRowNum <> 1 Then
            sh.Cells(2, 1).Resize(lastRowNum - 1, 13).Copy _
            MAIN_LEDGER.Cells(MAIN_LEDGER.Cells(MAIN_LEDGER.Cells.Rows.Count, MAIN_LEDGER.Range("H2").Column).End(xlUp).Row + 1, 1)
        Else
        End If
    Next i
    
End Sub

Private Sub ExportHledgerFile()
    LogManager.LogInfo "=== Hledger File Generation Started ==="

    Dim rowNum As Long
    Dim splitRowNum As Long
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim hledgerFile As Object
    Set hledgerFile = fso.OpenTextFile(Config.TEMP_FILE_ADDR, 2, True, -2)

    hledgerFile.WriteLine "; created by me at " & Now()
    hledgerFile.WriteLine
    hledgerFile.WriteLine "; Options"
    hledgerFile.WriteLine "decimal-mark ,"
    hledgerFile.WriteLine
    hledgerFile.WriteLine "; Special Accounts"
    hledgerFile.WriteLine "account Varlıklar                    ; type: A"
    hledgerFile.WriteLine "account Borçlar                      ; type: L"
    hledgerFile.WriteLine "account Özkaynaklar                  ; type: E"
    hledgerFile.WriteLine "account Gelir                        ; type: R"
    'hledgerFile.WriteLine "account Gelir:Yatırım                ; type: R" 'for special purposes
    hledgerFile.WriteLine "account Gider                        ; type: X"
    hledgerFile.WriteLine "account Varlıklar:Dönen Varlıklar    ; type: C"
    hledgerFile.WriteLine "account Gelir:Yatırım"
    hledgerFile.WriteLine
    
    ' > Accounts Listing
    Dim Accountes As Object
    Dim anAccount As Variant
    Dim longestAccountNameLen As Long
    Set Accountes = GetTransactionAccountNames
    ACCOUNTS.Cells.Clear
    For Each anAccount In Accountes
        ACCOUNTS.Cells(ACCOUNTS.Rows.Count, 1).End(xlUp).offset(1, 0).value = anAccount
    Next anAccount
    With ACCOUNTS.Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    hledgerFile.WriteLine "; Accounts"
    For Each anAccount In Accountes
        hledgerFile.WriteLine "account " & anAccount
        If Len(anAccount) > longestAccountNameLen Then longestAccountNameLen = Len(anAccount)
    Next anAccount
    ' < Accounts Listing
    hledgerFile.WriteLine
    ' > Commodities Listing
    Dim commoditiees As Object
    Dim aCommodity As Variant
    Set commoditiees = GetTransactionCommodityNames
    COMMODITIES.Cells.Clear
    For Each aCommodity In commoditiees
        COMMODITIES.Cells(COMMODITIES.Rows.Count, 1).End(xlUp).offset(1, 0).value = aCommodity
    Next aCommodity
    With COMMODITIES.Sort
        .SetRange Range("A:A")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    hledgerFile.WriteLine "; Commodities"
    hledgerFile.WriteLine "commodity 1.000,00 TRY"
    For Each aCommodity In commoditiees
        hledgerFile.WriteLine "commodity 1.000,00 " & aCommodity
    Next aCommodity
    ' < Commodities Listing
    hledgerFile.WriteLine
    ' > Includes
    hledgerFile.WriteLine "include Commodity-Prices.hledger"
    hledgerFile.WriteLine "include Budget.hledger"
    ' < Includes
    Dim entities As Object
    hledgerFile.WriteLine
    Dim ddate As Date
    Dim ddateStr As String
    Dim transCode As String
    Dim Description As String
    Dim account As String
    Dim Amount As String
    Dim mainCurrncy As String
    Dim splitCurrncy As String
    Dim commentTransaction As String
    Dim commentSplit1 As String
    Dim commentSplitN As String
    Dim blanks As String
    Dim assertion As String
    Dim profit As String
    Dim currSymbol As String
    'Portfolio csv data
    Dim portfolioData As String
    Dim commodityCommandsDict As scripting.Dictionary
    Set commodityCommandsDict = New scripting.Dictionary
    
    ' > Address column headers
    Call PopulateColumnHeaderIndexes("MAIN_LEDGER")
    'böylece headerleri enum gibi kullanabiliyorum.
    ' < Address column headers
    
    Dim dataArr As Variant
    Dim lastDataRow As Long
    
    rowNum = 2
    With MAIN_LEDGER
        ' > get all data to array
        lastDataRow = .Cells(.Rows.Count, accNameCol).End(xlUp).Row
        dataArr = .UsedRange.value
       ' < get all data to array
        Do While rowNum <= UBound(dataArr, 1)
            'If rowNum = 11037 Then Stop
            ' Get transaction values
            ddate = Replace(dataArr(rowNum, dateCol), ".", "-")
            ddateStr = Year(ddate) & "-" & IIf(Month(ddate) < 10, "0", "") & Month(ddate) & "-" & IIf(Day(ddate) < 10, "0", "") & Day(ddate)
            transCode = "(" & dataArr(rowNum, transCodeCol) & ")"
            Description = dataArr(rowNum, descriptionCol)
            account = dataArr(rowNum, accNameCol)
            Amount = dataArr(rowNum, amountCol)
            mainCurrncy = GetCurrency(rowNum, currencyCol, dataArr)
            commentTransaction = dataArr(rowNum, notesCol)
            commentSplit1 = dataArr(rowNum, tagCol)
            
            '> Parse transaction data
            'first line
            hledgerFile.WriteLine ddateStr & _
                IIf(transCode = "(!)", "  !  ", "  *  ") & transCode & "  " & Description '& "  " & IIf(commentTransaction <> "", ";" & commentTransaction, "")
            'first Posting
            blanks = String(longestAccountNameLen - Len(account) + IIf(left(Amount, 1) = "-", 0, 1), " ")
            assertion = IIf(dataArr(rowNum, reconCol) = "", "", "=" & dataArr(rowNum, reconCol) & " " & mainCurrncy & "  ")
            hledgerFile.WriteLine "  " & account & blanks & "  " & Amount & " " & mainCurrncy & "  " & assertion & IIf(commentSplit1 <> "", ";" & commentSplit1, "")
            
            splitRowNum = rowNum + 1
                Do While dataArr(splitRowNum, dateCol) = "" And dataArr(splitRowNum, accNameCol) <> ""
                    account = dataArr(splitRowNum, accNameCol)
                    Amount = dataArr(splitRowNum, amountCol)
                    splitCurrncy = GetCurrency(splitRowNum, currencyCol, dataArr, mainCurrncy)
                    commentSplitN = dataArr(splitRowNum, tagCol)
                    'Nth posting
                    blanks = String(longestAccountNameLen - Len(account) + IIf(left(Amount, 1) = "-", 0, 1), " ")
                    assertion = IIf(dataArr(splitRowNum, reconCol) = "", "", "=" & dataArr(splitRowNum, reconCol) & " " & splitCurrncy & "  ")
                    hledgerFile.WriteLine "  " & account & blanks & "  " & Amount & " " & splitCurrncy & "  " & assertion & IIf(commentSplitN <> "", ";" & commentSplitN, "")
                    'lots special part
                    If dataArr(splitRowNum, operationCol) = "Buy" Then
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & Amount & "::" & splitCurrncy & "::BUY", 1  'stock command
                        
                    ElseIf dataArr(splitRowNum, operationCol) = "Sell" Then
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & Amount & "::" & splitCurrncy & "::SELL", 1 'unstock command
                    
                    ElseIf left(dataArr(splitRowNum, operationCol), 5) = "Split" Then
                        ddate = CDate(Mid(dataArr(rowNum, descriptionCol), InStr(1, dataArr(rowNum, descriptionCol), "Tarih:") + 6))
                        ddateStr = Year(ddate) & "-" & IIf(Month(ddate) < 10, "0", "") & Month(ddate) & "-" & IIf(Day(ddate) < 10, "0", "") & Day(ddate)
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::" & Mid(dataArr(splitRowNum, operationCol), 7) & "::" & splitCurrncy & "::SPLIT", 1 'split command
                    
                    ElseIf left(dataArr(splitRowNum, operationCol), 8) = "Dividend" Then
                        Dim commName As String: commName = Split(splitCurrncy, " @ ")(0)
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::1::" & commName & _
                            " @ " & -1 * CDec(Amount) & " " & splitCurrncy & "::DIVIDEND", 1   'dividend command
                    ElseIf left(dataArr(splitRowNum, operationCol), 8) = "Interest" Then
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::1::" & "TEB" & _
                            " @ " & -1 * CDec(Amount) & " " & splitCurrncy & "::INTEREST", 1   'interest command
                    ElseIf dataArr(splitRowNum, operationCol) = "Withdrawal" Then
                        commodityCommandsDict.Add (hledgerFile.line) & "::" & ddateStr & "::1::" & "TEB" & _
                            " @ " & -1 * CDec(Amount) & " " & splitCurrncy & "::WITHDRAWAL", 1   'withdrawal command
                    Else
                        'If dataArr(splitRowNum, operationCol) <> "" Then Debug.Print dataArr(splitRowNum, operationCol)
                    End If
                    splitRowNum = splitRowNum + 1
                    If splitRowNum > UBound(dataArr, 1) Then Exit Do
                Loop
            '< Parse transaction data
            rowNum = splitRowNum
            hledgerFile.WriteLine
        Loop
    End With
    
    'Parse Lots
    Dim writeCommands As scripting.Dictionary
    Set writeCommands = ParseCommodities(commodityCommandsDict)
    'write lots
    Dim aCommand As Variant
    Dim i As Long
    Dim streamLine As String
    hledgerFile.Close
    Dim tempReadFile As Object
    Set tempReadFile = fso.OpenTextFile(Config.TEMP_FILE_ADDR, 1, False, -2)
    Set hledgerFile = fso.OpenTextFile(Config.HLEDGER_FILE_ADDR, 2, True, -2) 'to go to the begining of the file
    
    'get to the position and change writing according to buy sell stuff.
    Dim leftCountTemp As Long
    Dim commandPartArray() As String
    Dim commandPartNum As Long
    Dim tempString As String
    Do While Not tempReadFile.AtEndOfStream
        streamLine = tempReadFile.ReadLine
        If Not writeCommands Is Nothing Then
            If writeCommands.Exists(CStr(tempReadFile.line)) Then
                Dim commandType As String: commandType = Split(writeCommands(CStr(tempReadFile.line)), "|")(1)
                Dim commandText As String: commandText = Split(writeCommands(CStr(tempReadFile.line)), "|")(0)
                If commandType Like "DIVIDEND*" Then
                    'nothing
                    writeCommands(CStr(tempReadFile.line)) = commandText & vbCrLf
                    leftCountTemp = Len(streamLine)
                    Do While Mid(streamLine, leftCountTemp - 2, 3) <> "   "
                        leftCountTemp = leftCountTemp - 1
                    Loop
                    hledgerFile.WriteLine left(streamLine, leftCountTemp) & CDbl(Split(Mid(streamLine, leftCountTemp), " ")(1)) & " " & Split(Mid(streamLine, leftCountTemp), " ")(5)
                ElseIf commandType Like "INTEREST*" Then
                    writeCommands(CStr(tempReadFile.line)) = commandText & vbCrLf
                    hledgerFile.WriteLine streamLine
                ElseIf commandType Like "BUY*" Then
                    writeCommands(CStr(tempReadFile.line)) = commandText & vbCrLf
                    leftCountTemp = Len(streamLine)
                    Do While Mid(streamLine, leftCountTemp - 2, 3) <> "   "
                        leftCountTemp = leftCountTemp - 1
                    Loop
                    hledgerFile.WriteLine left(streamLine, leftCountTemp) & writeCommands(CStr(tempReadFile.line))
                ElseIf commandType Like "SELL*" Then 'sell command
                    For i = LBound(Split(writeCommands(CStr(tempReadFile.line)), "|SELL")) To UBound(Split(writeCommands(CStr(tempReadFile.line)), "|SELL")) - 1
                        tempString = Split(writeCommands(CStr(tempReadFile.line)), "|SELL" & vbCrLf)(i) & vbCrLf
                        leftCountTemp = Len(streamLine)
                        Do While Mid(streamLine, leftCountTemp - 2, 3) <> "   "
                            leftCountTemp = leftCountTemp - 1
                        Loop
                        commandPartArray = Split(tempString, vbCrLf)
                        hledgerFile.WriteLine left(streamLine, leftCountTemp) & commandPartArray((commandPartNum - 0) * 4)
                    Next i
                    hledgerFile.WriteLine "  Gelir:Yatırım"
                ElseIf commandType Like "WITHDRAWAL*" Then 'withdrawal command
                    writeCommands(CStr(tempReadFile.line)) = commandText & vbCrLf
                    hledgerFile.WriteLine streamLine
                Else
                    MsgBox "err-25"
                End If
            Else
                hledgerFile.WriteLine streamLine
            End If
        Else
            hledgerFile.WriteLine streamLine
        End If
    Loop
    
    'permission denied
    'fso.DeleteFile Config.TEMP_FILE_ADDR, True

    rowNum = 2
    hledgerFile.WriteLine
    'Create Config.PORTFOLIO_CSV_PATH for portfolio-performance
    Dim line As String
    Dim cashMovementsLine As String
    Dim tempLine As Variant
    Dim fileNo As Variant
    Dim tempStockObj As Variant
    Dim StockObj As CommodityObject
    '*********************
    Dim eurInvestingAcc As String: eurInvestingAcc = "Yatırım_EUR"
    Dim tryInvestingAcc As String: tryInvestingAcc = "Yatırım_TRY"
    Dim usdInvestingAcc As String: usdInvestingAcc = "Yatırım_USD"
    Dim investingAcc As String
    Dim eurOffsetInvestingAcc As String: eurOffsetInvestingAcc = "TEB_Vadesiz_EUR"
    Dim tryOffsetInvestingAcc As String: tryOffsetInvestingAcc = "TEB_Vadesiz_TRY"
    Dim usdOffsetInvestingAcc As String: usdOffsetInvestingAcc = "TEB_Vadesiz_USD"
    Dim offsetInvestingAcc As String
    '********************
    Dim commodityCommandsDictDates As Variant
    Set commodityCommandsDictDates = New scripting.Dictionary
    For i = 0 To commodityCommandsDict.Count - 1
        line = commodityCommandsDict.keys(i)
        commodityCommandsDictDates.item(Split(line, "::")(0)) = Split(line, "::")(1)
    Next i
    line = ""
    cashMovementsLine = ""
    fileNo = FreeFile
    Open Config.PORTFOLIO_CSV_PATH For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    If Not writeCommands Is Nothing Then
        For i = 0 To writeCommands.Count - 1
            For Each tempLine In Split(writeCommands.Items(i), vbCrLf)
                Do While left(tempLine, 1) = " "
                    tempLine = Mid(tempLine, 2)
                Loop
                'Debug.Print tempLine
                If left(tempLine, 1) <> "[" And tempLine <> "" Then
                    ' > StockObj Conversion
                    tempStockObj = Split(tempLine, " ")
                    StockObj.Type = IIf(IsNumeric(tempStockObj(0)) = False, tempStockObj(0), "BUY/SELL")
                    StockObj.Amount = IIf(IsNumeric(tempStockObj(0)) = True, tempStockObj(0), 0)
                    StockObj.Ticker = tempStockObj(1)
                    StockObj.Price = tempStockObj(3)
                    StockObj.Currency = tempStockObj(4)
                    StockObj.Currency = Split(StockObj.Currency, ";")(0)
                    If UBound(tempStockObj) = 6 Then
                         StockObj.OldBuyPrice = Split(tempStockObj(6), "|")(0)
                    Else
                         StockObj.OldBuyPrice = 0
                    End If
                    ' < StockObj Conversion
                    ' > investing account selection
                    investingAcc = tryInvestingAcc
                    offsetInvestingAcc = tryOffsetInvestingAcc
                    Select Case StockObj.Currency
                    Case "EUR"
                        investingAcc = eurInvestingAcc
                        offsetInvestingAcc = eurOffsetInvestingAcc
                    Case "USD"
                        investingAcc = usdInvestingAcc
                        offsetInvestingAcc = usdOffsetInvestingAcc
                    End Select
                    ' < investing account selection
                    If StockObj.Type = "DIVIDEND" Then
                        cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                        cashMovementsLine = cashMovementsLine & ";" & "Dividend" & ";" & 1 * StockObj.Price
                        cashMovementsLine = cashMovementsLine & ";" & StockObj.Ticker & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                        cashMovementsLine = cashMovementsLine & vbCrLf
                        cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                        cashMovementsLine = cashMovementsLine & ";" & "Removal" & ";" & -1 * StockObj.Price & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                        cashMovementsLine = cashMovementsLine & vbCrLf
                    ElseIf StockObj.Type = "INTEREST" Then
                        cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                        cashMovementsLine = cashMovementsLine & ";" & "Interest" & ";" & 1 * StockObj.Price & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                        cashMovementsLine = cashMovementsLine & ";" & ""
                        cashMovementsLine = cashMovementsLine & vbCrLf
                        cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                        cashMovementsLine = cashMovementsLine & ";" & "Removal" & ";" & -1 * StockObj.Price & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                        cashMovementsLine = cashMovementsLine & vbCrLf
                    ElseIf StockObj.Type = "BUY/SELL" Then  'buy or sell action stockObj.Type = "BUY/SELL"
                        line = line & commodityCommandsDictDates(writeCommands.keys(i))
                        If CDbl(StockObj.Amount) > 0 Then 'buying transaction
                            If StockObj.Currency = "TRY" Then 'if the currency is TRY(main currency) then deposit from thin air. else no!
                                line = line & ";" & "Buy" & ";" & StockObj.Amount
                                line = line & ";" & StockObj.Ticker & ";" & Abs(CDbl(StockObj.Price)) & ";" & StockObj.Amount * Abs(CDbl(StockObj.Price))
                                line = line & ";" & investingAcc
                                line = line & ";" & offsetInvestingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                                cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                                cashMovementsLine = cashMovementsLine & ";" & IIf(CDbl(StockObj.Amount) < 0, "Removal", "Deposit") & ";" & _
                                StockObj.Amount * StockObj.Price & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                            Else
                                line = line & ";" & "Buy" & ";" & StockObj.Amount
                                line = line & ";" & StockObj.Ticker & ";" & Abs(CDbl(StockObj.Price)) & ";" & StockObj.Amount * Abs(CDbl(StockObj.Price))
                                line = line & ";" & investingAcc
                                line = line & ";" & offsetInvestingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                                line = line & commodityCommandsDictDates(writeCommands.keys(i))
                                line = line & ";" & "Sell" & ";" & StockObj.Amount * Abs(CDbl(StockObj.Price))
                                line = line & ";" & StockObj.Currency & ";" & 1 & ";" & StockObj.Amount * Abs(CDbl(StockObj.Price))
                                line = line & ";" & tryInvestingAcc
                                line = line & ";" & investingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                            End If
                        Else ' selling transaction
                            If StockObj.Currency = "TRY" Then 'if the currency is TRY(main currency) then deposit from thin air. else no!
                                line = line & ";" & "Sell" & ";" & StockObj.Amount
                                line = line & ";" & StockObj.Ticker & ";" & Abs(StockObj.OldBuyPrice) & ";" & StockObj.Amount * Abs(StockObj.OldBuyPrice)
                                line = line & ";" & investingAcc
                                line = line & ";" & offsetInvestingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                                cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                                cashMovementsLine = cashMovementsLine & ";" & IIf(CDbl(StockObj.Amount) < 0, "Removal", "Deposit") & ";" & _
                                StockObj.Amount * StockObj.OldBuyPrice & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                            Else
                                line = line & ";" & "Sell" & ";" & StockObj.Amount
                                line = line & ";" & StockObj.Ticker & ";" & Abs(StockObj.OldBuyPrice) & ";" & StockObj.Amount * Abs(StockObj.OldBuyPrice)
                                line = line & ";" & investingAcc
                                line = line & ";" & offsetInvestingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                                line = line & commodityCommandsDictDates(writeCommands.keys(i))
                                line = line & ";" & "Buy" & ";" & Abs(StockObj.Amount * Abs(StockObj.OldBuyPrice))
                                line = line & ";" & StockObj.Currency & ";" & 1 & ";" & Abs(StockObj.Amount * Abs(StockObj.OldBuyPrice))
                                line = line & ";" & investingAcc
                                line = line & ";" & tryInvestingAcc
                                line = line & ";" & StockObj.Currency
                                line = line & vbCrLf
                            End If
                        End If
                        cashMovementsLine = cashMovementsLine & vbCrLf
                    End If
                ElseIf StockObj.Type = "WITHDRAWAL" Then ' ben burada yapılan withdrawal'ın döviz cinsinde bir commodity olduğunu kabul ettim.
                    'Stop
                    'buralar full yaratıcılık ve saçmalık buraların bir elden geçemesi iyi olacaktır. Döviz ile ödemeler yaptığımda düşen döviz stoğumun PP yazılımına gitmesi
                    'için yapılan csv manipülayonları
                    line = line & commodityCommandsDictDates(writeCommands.keys(i))
                    line = line & ";" & "Sell" & ";" & Abs(StockObj.Price)
                    line = line & ";" & StockObj.Currency & ";" & Abs(StockObj.Price) & ";" & Abs(StockObj.Price)
                    line = line & ";" & investingAcc
                    line = line & ";" & offsetInvestingAcc
                    line = line & ";" & StockObj.Currency
                    line = line & vbCrLf
    '                line = line & commodityCommandsDictDates(writeCommands.keys(i))
    '                line = line & ";" & "Buy" & ";" & Abs(stockObj.Amount * Abs(stockObj.OldBuyPrice))
    '                line = line & ";" & stockObj.Currency & ";" & 1 & ";" & Abs(stockObj.Amount * Abs(stockObj.OldBuyPrice))
    '                line = line & ";" & investingAcc
    '                line = line & ";" & tryInvestingAcc
    '                line = line & ";" & stockObj.Currency
    '                line = line & vbCrLf
                    cashMovementsLine = cashMovementsLine & commodityCommandsDictDates(writeCommands.keys(i))
                    cashMovementsLine = cashMovementsLine & ";" & "Removal" & ";" & -1 * StockObj.Price & ";" & ";" & offsetInvestingAcc & ";" & StockObj.Currency
                    cashMovementsLine = cashMovementsLine & vbCrLf
                Else
                    'virtual account
                End If
            Next tempLine
        Next i
    Else
    End If
    line = Replace(line, """", "")
    
    cashMovementsLine = Replace(cashMovementsLine, """", "")
    Print #fileNo, line 'portfolioData Writing
    Close #fileNo
       
    fileNo = FreeFile
    Open Config.PORTFOLIO_CASH_CSV_PATH For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    Print #fileNo, cashMovementsLine
    Close #fileNo

    fileNo = FreeFile
    Open Config.PORTFOLIO_CSV_PATH_INVESTING For Output As #fileNo 'Open file for overwriting! Replace Output with Append to append
    Dim stockObjSub As Variant
    Dim stockObjSubSub As Variant
    If commDict.Count > 0 Then
        For Each tempStockObj In commDict.keys
            For Each stockObjSub In commDict(tempStockObj).keys
                For Each stockObjSubSub In commDict(tempStockObj)(stockObjSub).keys
                    line = line & tempStockObj & ";" & Mid(stockObjSub, 7, 2) & "." & Mid(stockObjSub, 5, 2) & "." & Mid(stockObjSub, 1, 4) & _
                        ";" & Split(commDict(tempStockObj)(stockObjSub)(stockObjSubSub), "@")(0) & _
                        ";" & Format(Split(commDict(tempStockObj)(stockObjSub)(stockObjSubSub), "@")(1), "0.0#######") & vbCrLf
                Next stockObjSubSub
            Next stockObjSub
        Next tempStockObj
    End If
    line = Replace(line, """", "")
    Print #fileNo, line
    Close #fileNo
    
    hledgerFile.Close
    Call ConvertTxttoUTF(Config.HLEDGER_FILE_ADDR, Config.HLEDGER_FILE_ADDR)
    Set tempReadFile = Nothing
    fso.DeleteFile Config.TEMP_FILE_ADDR, True

End Sub

Private Function ParseCommodities(ByRef commodityCommandsDict As Object) As Object

Dim command As Variant
Dim i As Long
Dim commandLineNum() As String
If commodityCommandsDict.Count = 0 Then Exit Function
ReDim commandLineNum(1 To commodityCommandsDict.Count)

Dim commands() As Variant
commands = commodityCommandsDict.keys()

Dim commName As String
Dim commPrice As Variant
Dim commDateStamp As String
Dim commQuantity As Variant
Dim tempString() As String
Dim isBuy As Boolean
Dim commLineNum As String
Dim minDateStamp As String
Set commDict = New scripting.Dictionary
Dim aStock As scripting.Dictionary
Dim aDate As scripting.Dictionary
Dim buyedPrice As Variant
Dim buyedQuant As Variant
Dim minSameDayNumber As Long
Dim keyy As Variant
Dim mainCurrency As String

Dim writeCommands As scripting.Dictionary
Set writeCommands = New scripting.Dictionary
Dim splitCommands As scripting.Dictionary
Set splitCommands = New scripting.Dictionary

'get split commands
For i = UBound(commands) To LBound(commands) Step -1
    tempString = Split(commands(i), "::")
    If tempString(4) = "SPLIT" Then
        commName = left(tempString(3), InStr(tempString(3), " ") - 1)
        commDateStamp = Replace(tempString(1), "-", "")
        commQuantity = CDec(tempString(2))
        splitCommands.item(commDateStamp & "|" & commName) = commQuantity
'    ElseIf tempString(4) = "DIVIDEND" Then
'        'Stop
    Else
    End If
Next i

'parse commodity splits
Dim aSplit As Variant
Dim splitPerc As Double
Dim tempStringPart As Variant
Dim tempVal As String
For Each aSplit In splitCommands
    commName = Split(aSplit, "|")(1)
    commDateStamp = Split(aSplit, "|")(0)
    splitPerc = (CDbl(splitCommands(aSplit)) / 100) + 1
    For i = UBound(commands) To LBound(commands) Step -1
        tempString = Split(commands(i), "::")
        If tempString(4) <> "SPLIT" And tempString(4) <> "DIVIDEND" And tempString(4) <> "INTEREST" Then
            If left(tempString(3), InStr(1, tempString(3), " ") - 1) = commName _
                And CLng(Replace(tempString(1), "-", "")) <= CLng(commDateStamp) Then
                tempString(0) = CStr(CLng(tempString(0)) - 0)
                tempString(2) = CStr(CDbl(tempString(2)) * splitPerc)
                tempStringPart = tempString(3)
                tempStringPart = Mid(tempStringPart, InStr(1, tempStringPart, "@") + 2)
                tempStringPart = left(tempStringPart, InStr(1, tempStringPart, " ") - 1)
                tempString(3) = Replace(tempString(3), tempStringPart, CStr(CDbl(tempStringPart) / splitPerc))
                commands(i) = ""
                For Each tempStringPart In tempString
                    commands(i) = commands(i) & "::" & tempStringPart
                Next tempStringPart
                commands(i) = Mid(commands(i), 3)
            Else
            End If
        Else
        End If
    Next i
Next aSplit

For i = UBound(commands) To LBound(commands) Step -1
    tempString = Split(commands(i), "::")
    commLineNum = tempString(0)
    'If commLineNum = "17486" Then Stop
    isBuy = IIf(tempString(4) = "BUY", True, False)
    commName = left(tempString(3), InStr(tempString(3), " ") - 1)
    commQuantity = CDec(tempString(2))
    commDateStamp = Replace(tempString(1), "-", "")
    mainCurrency = Split(tempString(3), " ")(UBound(Split(tempString(3), " ")))
    tempString(3) = Mid(tempString(3), InStr(tempString(3), "@") + 2)
    tempString(3) = left(tempString(3), InStr(tempString(3), " ") - 1)
    commPrice = Format(CDec(tempString(3)), "0.0#########################################")
    'If commName = "TKM" Then Stop
    'If commLineNum = "2673" Then Stop
    If commDict.Exists(commName) Then
        If IIf(tempString(4) = "BUY", True, False) Then
            Set aStock = commDict(commName)
            If aStock.Exists(commDateStamp) Then
                Set aDate = aStock(commDateStamp)
                minSameDayNumber = 0
                For Each keyy In aDate.keys
                    If keyy > minSameDayNumber Then minSameDayNumber = keyy 'actually it is the largest sameday num @TODO
                Next keyy
                aDate.Add minSameDayNumber + 1, commQuantity & "@" & commPrice
            Else
                Set aDate = New scripting.Dictionary
                aDate.Add 1, commQuantity & "@" & commPrice
                Set aStock.item(commDateStamp) = aDate
            End If
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                "   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & " " & mainCurrency & "|BUY" & vbCrLf
        ElseIf IIf(tempString(4) = "SELL", True, False) Then
            'If commName = "EBEBK" Then Stop
            Set aStock = commDict(commName)
            Do
                minDateStamp = "99999999"
                For Each keyy In aStock.keys
                    If keyy < minDateStamp Then minDateStamp = keyy
                Next keyy
                minSameDayNumber = 99999
                For Each keyy In aStock(minDateStamp).keys
                    If keyy < minSameDayNumber Then minSameDayNumber = keyy
                Next keyy
                tempString = Split(aStock(minDateStamp).item(minSameDayNumber), "@")
                buyedPrice = CDec(tempString(1))
                buyedQuant = CDec(tempString(0))
                'SELLING OPTIONS
                'If commName = "TKM" Then Stop
                If Abs(commQuantity) > buyedQuant Then
                    aStock(minDateStamp).Remove (minSameDayNumber)
                    If aStock(minDateStamp).Count = 0 Then aStock.Remove (minDateStamp)
                    If aStock.Count = 0 Then commDict.Remove (commName)
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(buyedQuant) & " " & commName & " @ " & _
                                                      Format(buyedPrice, "0.0###################################") & " " & _
                                                      mainCurrency & ";Sold @ " & Format(commPrice, "0.0###################################") & "|SELL" & vbCrLf
                    commQuantity = commQuantity + buyedQuant
                ElseIf Abs(commQuantity) < buyedQuant Then
                    aStock(minDateStamp).item(minSameDayNumber) = buyedQuant - Abs(commQuantity) & "@" & buyedPrice
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(commQuantity) & " " & commName & " @ " & _
                                                      Format(buyedPrice, "0.0###################################") & " " & _
                                                      mainCurrency & ";Sold @ " & Format(commPrice, "0.0###################################") & "|SELL" & vbCrLf
                    commQuantity = 0
                Else
                    aStock(minDateStamp).Remove (minSameDayNumber)
                    If aStock(minDateStamp).Count = 0 Then aStock.Remove (minDateStamp)
                    If aStock.Count = 0 Then commDict.Remove (commName)
                    writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                                      "  -" & Abs(commQuantity) & " " & commName & " @ " & _
                                                      Format(buyedPrice, "0.0###################################") & " " & _
                                                      mainCurrency & ";Sold @ " & Format(commPrice, "0.0###################################") & "|SELL" & vbCrLf
                    commQuantity = 0
                End If
                'END OF SELLING OPTIONS
            Loop While commQuantity <> 0
        ElseIf IIf(tempString(4) = "SPLIT", True, False) Then
            '
        ElseIf IIf(tempString(4) = "DIVIDEND", True, False) Then
            'Stop
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                              "  " & "DIVIDEND" & " " & commName & " @ " & _
                                              Format(commPrice, "0.0###################################") & " " & mainCurrency & "|DIVIDEND" & vbCrLf
        Else
            
        End If
    Else
        If isBuy Then
            Set aStock = New scripting.Dictionary
            Set aDate = New scripting.Dictionary
            aDate.Add 1, commQuantity & "@" & commPrice
            aStock.Add commDateStamp, aDate
            commDict.Add commName, aStock
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                "   " & Abs(commQuantity) & " " & commName & " @ " & Format(commPrice, "0.0###################################") & _
                " " & mainCurrency & "|BUY" & vbCrLf
        ElseIf IIf(tempString(4) = "INTEREST", True, False) Then
            'Stop
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                              "  " & "INTEREST" & " " & commName & " @ " & _
                                              Format(commPrice, "0.0###################################") & " " & mainCurrency & "|INTEREST" & vbCrLf
        ElseIf IIf(tempString(4) = "WITHDRAWAL", True, False) Then
            writeCommands.item(commLineNum) = IIf(writeCommands.item(commLineNum) = "", "", writeCommands.item(commLineNum)) & _
                                              "  " & "WITHDRAWAL" & " " & commName & " @ " & _
                                              Format(commPrice, "0.0###################################") & " " & mainCurrency & "|WITHDRAWAL" & vbCrLf
        Else
            MsgBox ("err-1")
        End If
    End If
Next i

Set ParseCommodities = writeCommands

End Function

Private Function GetCurrency(ByVal rowNum As Long, ByVal colnum As Long, ByRef dataArr As Variant, Optional ByVal mainCurrency As String) As String

    If dataArr(rowNum, colnum) Like "CURRENCY::*" Then
        GetCurrency = Replace(dataArr(rowNum, colnum), "CURRENCY::", "")
    ElseIf dataArr(rowNum, colnum + 1) = "Buy" Or _
           dataArr(rowNum, colnum + 1) = "Sell" Or _
           left(dataArr(rowNum, colnum + 1), 5) = "Split" Or _
           dataArr(rowNum, colnum + 1) = "Dividend" Then
        GetCurrency = dataArr(rowNum, colnum + 3)
        Do Until InStr(GetCurrency, ":") = 0
            GetCurrency = Mid(GetCurrency, InStr(GetCurrency, ":") + 1)
        Loop
        'Currency isminde karakterin sayısal - number olmasına karşı önlem "TI3" vakası.
        With CreateObject("VBScript.RegExp")
            .Pattern = "(\d)"
            'If GetCurrency = "A1CAP" Then Stop
            If .test(GetCurrency) Then GetCurrency = """" & GetCurrency & """"
        End With
        '
        GetCurrency = GetCurrency & " @ " & Format(dataArr(rowNum, colnum + 5), "0.0#######################################") & " " & mainCurrency 'No scientific for you!!
    Else
        Do While GetCurrency = ""
            GetCurrency = GetCurrency(rowNum - 1, colnum, dataArr)
            'Set rng = rng.offset(-1, 0)
        Loop
    End If

End Function

'Private Sub ExportAsCSV()
'
'    Dim MyFileName As String
'    Dim CurrentWB As Workbook, TempWB As Workbook
'    Set CurrentWB = ActiveWorkbook
'    ActiveWorkbook.ActiveSheet.UsedRange.Copy
'    Set TempWB = Application.Workbooks.Add(1)
'    With TempWB.Sheets(1).Range("A1")
'        .PasteSpecial xlPasteValues
'        .PasteSpecial xlPasteFormats
'    End With
'    'MyFileName = CurrentWB.Path & "\" & Left(CurrentWB.Name, InStrRev(CurrentWB.Name, ".") - 1) & ".csv"
'    'Optionally, comment previous line and uncomment next one to save as the current sheet name
'    MyFileName = CurrentWB.path & "\" & CurrentWB.ActiveSheet.name & ".csv"
'    Application.DisplayAlerts = False
'    TempWB.SaveAs Filename:=MyFileName, FileFormat:=xlCSV, CreateBackup:=False, local:=True
'    TempWB.Close SaveChanges:=False
'    Application.DisplayAlerts = True
'    Call ConvertTxttoUTF(MyFileName, MyFileName)
'
'End Sub

Private Sub ConvertTxttoUTF(sInFilePath As String, sOutFilePath As String)
    Dim objFS  As Object
    Dim objFSNoBOM  As Object
    Dim iFile       As Variant
    Dim sFileData   As String
    
    'Init
    iFile = FreeFile
    Open sInFilePath For Input As #iFile
        sFileData = Input$(LOF(iFile), iFile)
        sFileData = sFileData & vbCrLf
    Close iFile
    'Open & Write
    Set objFS = CreateObject("ADODB.Stream")
    With objFS
        .Charset = "UTF-8"
        .Open
        .WriteText sFileData
        .Position = 0
        .Type = 2
        .Position = 3
    End With
    Set objFSNoBOM = CreateObject("ADODB.Stream")
    With objFSNoBOM
        .Type = 1
        .Open
        objFS.CopyTo objFSNoBOM
    End With
        
    'Save & Close
    objFSNoBOM.SaveToFile sOutFilePath, 2   '2: Create Or Update
    objFSNoBOM.Close
    objFS.Close
    
    'Completed
    Application.StatusBar = "Completed"
End Sub

Public Function GetTransactionMinMaxEntityDates() As Object
      
    Dim rowNum As Long
    Dim splitRowNum As Long
    rowNum = 2 'starting row
    Dim entityDict As Object
    Set entityDict = CreateObject("scripting.Dictionary")
    Dim entityName As String
    Dim ddate As Long 'transaction date
    Dim dates(2) As Variant 'dates and count data
    Dim entityCount As Long 'entity transaction count
    With MAIN_LEDGER
        Do While .Cells(rowNum, 1) <> ""
            ddate = .Cells(rowNum, 1)
            splitRowNum = rowNum + 1
            Do While .Cells(splitRowNum, 1) = "" And .Cells(splitRowNum, 8) <> ""
                If .Cells(splitRowNum, 6) = "Buy" Or .Cells(splitRowNum, 6) = "Sell" Then
                    entityCount = .Cells(splitRowNum, 9)
                    entityName = .Cells(splitRowNum, 8)
                    Do While InStr(entityName, ":") <> 0
                        entityName = Mid(entityName, InStr(entityName, ":") + 1)
                    Loop
                    'If left(entityName, 2) = "W_" Then Stop
                    'If ddate = 44796 Then Stop
                    If entityDict.Exists(entityName) Then
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityDict(entityName)(0) + entityCount
                        dates(1) = entityDict(entityName)(1)
                        dates(2) = entityDict(entityName)(2)
                    Else
                        'If entityName = "TKM" Then Debug.Print entityCount
                        dates(0) = entityCount
                        dates(1) = CDate(Date)
                        dates(2) = CDate(25569) '1.1.1970 in long format
                    End If
                    If ddate < dates(1) Then dates(1) = CDate(ddate)
                    If ddate > dates(2) Then dates(2) = CDate(ddate)
                    entityDict.item(entityName) = dates
                Else
                
                End If
                splitRowNum = splitRowNum + 1
            Loop
            rowNum = splitRowNum
        Loop
        Debug.Print rowNum
    End With
    Dim entityKey As Variant
    Dim i As Long
    Dim entArray As Variant
    For Each entityKey In entityDict.keys
        'If entityKey = "GARFA" Then Stop
        i = i + 1
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 0).value = entityKey
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 1).value = entityDict(entityKey)(0)
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 2).value = entityDict(entityKey)(1)
'        ActiveSheet.Cells(1, 1).End(xlUp).offset(i, 3).value = CDate(IIf(entityDict(entityKey)(0) = 0, entityDict(entityKey)(2), CLng(Date)))
        If entityDict(entityKey)(0) <> 0 Then
            entArray = entityDict(entityKey)
            entArray(2) = Date
            entityDict(entityKey) = entArray
        Else
        End If
    Next entityKey
    
    Set GetTransactionMinMaxEntityDates = entityDict
    
End Function

Private Function GetTransactionCommodityNames() As Object

    Dim rowNum As Long
    Dim splitRowNum As Long
    rowNum = 2 'starting row
    Dim commodityDict As Object
    Set commodityDict = CreateObject("scripting.Dictionary")
    Dim commodityName As String

    With MAIN_LEDGER
        Do While .Cells(rowNum, 1) <> ""
            splitRowNum = rowNum + 1
            If UCase(.Cells(splitRowNum, 6)) = "BUY" Then
                commodityName = GetCurrency(1, 5, .Cells(splitRowNum, 1).EntireRow.value)
                'cleanup for price directive
                commodityName = Strings.left(commodityName, InStr(1, commodityName, " @") - 1)
                commodityDict.item(commodityName) = 1
            Else
            End If
            Do While .Cells(splitRowNum, 1) = "" And .Cells(splitRowNum, 8) <> ""
                splitRowNum = splitRowNum + 1
            Loop
            rowNum = splitRowNum
        Loop
    End With

    Set GetTransactionCommodityNames = commodityDict

End Function
    
    
Private Function GetTransactionAccountNames() As Object
      
    Dim rowNum As Long
    Dim splitRowNum As Long
    rowNum = 2 'starting row
    Dim accountDict As Object
    Set accountDict = CreateObject("scripting.Dictionary")
    Dim accountName As String
    
    With MAIN_LEDGER
        Do While .Cells(rowNum, 1) <> ""
            splitRowNum = rowNum + 1
            accountName = .Cells(rowNum, 8)
            accountDict.item(accountName) = 1
            Do While .Cells(splitRowNum, 1) = "" And .Cells(splitRowNum, 8) <> ""
                accountName = .Cells(splitRowNum, 8)
                accountDict.item(accountName) = 1
                splitRowNum = splitRowNum + 1
            Loop
            rowNum = splitRowNum
        Loop
    End With
    
    Set GetTransactionAccountNames = accountDict

End Function










