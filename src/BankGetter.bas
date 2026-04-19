Attribute VB_Name = "BankGetter"
Option Explicit

' Entry point for the data-driven engine. bankID must match a BankID in the BANKS sheet.
' Run BankGetterSetup.CreateBANKSSheet once to initialize the BANKS sheet before using this.
Public Sub BankGetterRun(Optional bankID As String = "TEB")
    BankGetterEngine.RunBank bankID
End Sub

Public Sub BankGetterTEB()
    LogManager.LogInfo "=== BankGetter: TEB Data Fetch Started ==="
    On Error GoTo ErrorHandler

    Application.ActiveWorkbook.Worksheets("Bank_Info").activate
    ActiveSheet.Cells.Delete
    ActiveSheet.Range("b2").Select

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

    ActiveSheet.Range("b2").Select
    LogManager.LogInfo "BankGetter: TEB Data Fetch Completed"
    MsgBox "bitti", vbInformation, "BankGetter"
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
    Dim childi As Variant
    Dim skipLineCount As Integer
    Dim q As Integer, i As Integer, j As Integer, k As Integer, itm As Variant

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
            i = -1
            j = 0
            For Each childi In .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""")).children
                i = i + 1
                If skipLineCount > 0 Then
                    i = i - 1
                    skipLineCount = skipLineCount - 1
                Else
                    For Each itm In childi.children
                        j = j + 1
                        If j = 1 Then
                            k = 1
                            On Error Resume Next
                            ActiveCell.offset(i, k).value = CDate(Replace(Replace(itm.children.item(1).name, "/", "."), "(*)", ""))
                            If Err.Number <> 0 Then
                                j = j - 1: Err.Clear
                            Else
                                ActiveCell.offset(i, k - 1).value = "Hesap-" & detailNum
                            End If
                            On Error GoTo 0
                        ElseIf j = 4 Then
                            k = 2
                            ActiveCell.offset(i, k).value = itm.children.item(1).name
                        ElseIf j = 5 Then
                            k = 3
                            ActiveCell.offset(i, k).value = CDbl(itm.children.item(1).name)
                        End If
                    Next itm
                    j = 0
                End If
            Next childi
            Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hesaplar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Next q
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        ActiveCell.Cells(ActiveSheet.Rows.Count - 1, ActiveCell.Column - 1).End(xlUp).offset(1, 0).Select
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
    Dim i As Integer, j As Integer, k As Integer, skipLineCount As Integer
    Dim investmentTable As stdAcc
    Dim exitLoop As Boolean

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

        i = -1: j = 0: exitLoop = False
        Do Until exitLoop
            If Not .accMain.AwaitForElement(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""), , 10) Is Nothing Then
                Set investmentTable = .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""))
                skipLineCount = 1
                For Each childi In investmentTable.children
                    i = i + 1
                    If skipLineCount > 0 Then
                        i = i - 1: skipLineCount = skipLineCount - 1
                    Else
                        For Each itm In childi.children
                            j = j + 1
                            If j = 1 Then
                                k = 1
                                On Error Resume Next
                                ActiveCell.offset(i, k).value = CDate(Replace(Replace(itm.children.item(1).name, "/", "."), "(*)", ""))
                                If Err.Number <> 0 Then
                                    j = j - 1: Err.Clear
                                Else
                                    ActiveCell.offset(i, k - 1).value = "TEB Yat" & ChrW(305) & "r" & ChrW(305) & "m Hesab" & ChrW(305)
                                End If
                                On Error GoTo 0
                            ElseIf j = 2 Then
                                k = 3: ActiveCell.offset(i, k).value = CDbl(itm.children.item(1).name)
                            ElseIf j = 4 Then
                                k = 2: ActiveCell.offset(i, k).value = itm.children.item(1).name
                            End If
                        Next itm
                        j = 0
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
        ActiveCell.End(xlUp).Select
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
    Dim i As Integer, j As Integer, k As Integer
    Dim skipLineCount As Integer, detailNum As Integer

    cardsArr = Array("TEB BONUS CARD", "TEB SHE CARD")

    With chrome
        ActiveCell.offset(0, 5).Select
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction

        i = -1: j = 0: detailNum = 0
        For Each cardName In cardsArr
            detailNum = detailNum + 1
            Call .accMain.AwaitForElement(stdLambda.Create("$1.Name = """ & cardName & """ and $1.Role = ""ROLE_LINK""")).DoDefaultAction

            If Not .accMain.AwaitForElement(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""), , 10) Is Nothing Then
                Set cardTable = .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE"""))
                skipLineCount = 2
                For Each childi In cardTable.children
                    i = i + 1
                    If skipLineCount > 0 Then
                        i = i - 1: skipLineCount = skipLineCount - 1
                    Else
                        For Each itm In childi.children
                            j = j + 1
                            If j = 1 Then
                                k = 1
                                ActiveCell.offset(i, k).value = Replace(Replace(itm.children.item(1).name, "/", "."), "(*)", "")
                                On Error Resume Next
                                ActiveCell.offset(i, k).value = CDate(ActiveCell.offset(i, k).value)
                                If Err.Number <> 0 And i <> 1 Then
                                    Err.Clear
                                    ActiveCell.offset(i, k).value = ""
                                    ActiveCell.offset(0, 5).Select
                                    ActiveCell.offset(0, k).value = Replace(Replace(itm.children.item(1).name, "/", "."), "(*)", "")
                                    i = -1
                                ElseIf Err.Number <> 0 Then
                                    Err.Clear
                                End If
                                On Error GoTo 0
                                ActiveCell.offset(i, k - 1).value = "Kart-" & cardName
                            ElseIf j = 2 Then
                                k = 2: ActiveCell.offset(i, k).value = itm.children.item(1).name
                            ElseIf j = 4 Then
                                k = 3: ActiveCell.offset(i, k).value = -1 * CDbl(itm.children.item(1).name)
                            ElseIf j = 5 Then
                                k = 4: ActiveCell.offset(i, k).value = "'" & itm.name
                            End If
                        Next itm
                        j = 0
                    End If
                Next childi
            End If

            Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Kartlar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
            i = -1
            ActiveCell.offset(0, 5).Select
        Next cardName

        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With

    LogManager.LogInfo "Credit card transactions fetched: " & detailNum & " cards"
    Exit Sub
ErrorHandler:
    LogManager.LogError "BankGetter_FetchCards failed: " & Err.Description
End Sub

Private Function CastMonthName(monthNum As Integer) As String
    CastMonthName = Array("Ocak", ChrW(350) & "ubat", "Mart", "Nisan", "May" & ChrW(305) & "s", _
                          "Haziran", "Temmuz", "A" & ChrW(287) & "ustos", _
                          "Eyl" & ChrW(252) & "l", "Ekim", _
                          "Kas" & ChrW(305) & "m", "Aral" & ChrW(305) & "k")(monthNum - 1)
End Function










