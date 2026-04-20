Attribute VB_Name = "Importer"
Option Explicit

Public Sub ImporterBegin()
    LogManager.LogInfo "=== CSV Import Process Started ==="

    Dim ws As Worksheet
    Set ws = IMPORT_SH
    ws.activate
    
    Dim targetWs As Worksheet
    If ws.Cells(2, 1).value = "" Then Exit Sub
    On Error GoTo WRONGPAGENAME
    Set targetWs = ActiveWorkbook.Worksheets(ws.Cells(2, 1).value)
    On Error GoTo 0
    
    Dim datesRange As Range
    Dim notesRange As Range
    Dim amountRange As Range
    Dim expenseCategoryRange As Range
    Dim specialCategoryRange As Range
    Dim bankDescRange As Range
    Set datesRange = Application.InputBox("Select date columns first value", "Date Column", , , , , , 8)
    If datesRange.offset(-1, 0).value = "Date" And datesRange.offset(-1, 1).value = "Desc" And datesRange.offset(-1, 2).value = "Amount" Then
        Set notesRange = datesRange.offset(0, 1)
        Set amountRange = datesRange.offset(0, 2)
    Else
        Set notesRange = Application.InputBox("Select note columns first value", "Note Column", , , , , , 8)
        Set amountRange = Application.InputBox("Select amount columns first value", "Amount Column", , , , , , 8)
    End If
    ' Check selected Ranges
    If datesRange.Cells.count <> 1 Then GoTo WRONGDATACOUNT
    If notesRange.Cells.count <> 1 Then GoTo WRONGDATACOUNT
    If amountRange.Cells.count <> 1 Then GoTo WRONGDATACOUNT
    ' Resize selection to get data
    Set expenseCategoryRange = ws.Cells(amountRange.Row, Application.WorksheetFunction.max(datesRange.Column, notesRange.Column, amountRange.Column) + 1)
    Set specialCategoryRange = ws.Cells(amountRange.Row, Application.WorksheetFunction.max(datesRange.Column, notesRange.Column, amountRange.Column) + 2)
    Set bankDescRange = ws.Cells(amountRange.Row, Application.WorksheetFunction.max(datesRange.Column, notesRange.Column, amountRange.Column) + 3)
    Set datesRange = datesRange.Resize(ws.Cells(ws.Rows.count, datesRange.Column).End(xlUp).Row - datesRange.Row + 1, 1)
    Set notesRange = notesRange.Resize(ws.Cells(ws.Rows.count, notesRange.Column).End(xlUp).Row - notesRange.Row + 1, 1)
    Set amountRange = amountRange.Resize(ws.Cells(ws.Rows.count, amountRange.Column).End(xlUp).Row - amountRange.Row + 1, 1)
    Set expenseCategoryRange = expenseCategoryRange.Resize(amountRange.Rows.count, 1)
    Set specialCategoryRange = specialCategoryRange.Resize(amountRange.Rows.count, 1)
    Set bankDescRange = bankDescRange.Resize(amountRange.Rows.count, 1)
    ' Check data
    If Not (datesRange.Cells.count = notesRange.Cells.count And datesRange.Cells.count = amountRange.Cells.count) Then GoTo WRONGDATACOUNT
    '*********************
    'check descriptions
    '*********************
    Dim foundDescRange As Range
    Dim answer As Variant
    Dim i As Integer
    Dim ruleCheckResult As scripting.Dictionary
    Load frmDescription
    
    'Put bank description to the side
    For i = datesRange.Rows.count To 1 Step -1
        If bankDescRange.Cells(i, 1).value = "" Then bankDescRange.Cells(i, 1).value = notesRange.Cells(i, 1).value
    Next i
    
    ' >>> RULES PART
    For i = datesRange.Rows.count To 1 Step -1
        If expenseCategoryRange.Cells(i, 1).value = "" Then
            Set ruleCheckResult = New scripting.Dictionary
            Set ruleCheckResult = Rules.CheckRules(notesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs.name)
            If ruleCheckResult.count <> 0 Then
                expenseCategoryRange.Cells(i, 1).value = ruleCheckResult("toAccount")
                specialCategoryRange.Cells(i, 1).value = ruleCheckResult("special")
                If ruleCheckResult("newDescription") <> "" Then notesRange.Cells(i, 1).value = ruleCheckResult("newDescription")
            Else
                'no rules applied
            End If
        Else
            'full row. no need to check for rules.
        End If
    Next i
    ' <<< RULES PART

    If MsgBox("Kural ataması tamamlandı. Devam etmek istiyor musunuz?", _
          vbYesNo + vbQuestion, "Importer") = vbNo Then Exit Sub
    
    ' >>> SAME SEARCHING PART
    For i = datesRange.Rows.count To 1 Step -1
        'Check duplicates and fully entered records
            If CheckDuplicate(datesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs, CStr(bankDescRange.Cells(i, 1).value)) = 0 And _
                                                                        (expenseCategoryRange.Cells(i, 1).value = "" Or _
                                                                        expenseCategoryRange.Cells(i, 1).value = "Gider:Bilinmeyen") Then
            'SAME SEARCHING
            If expenseCategoryRange.Cells(i, 1).value = "" Or expenseCategoryRange.Cells(i, 1).value = "Gider:Bilinmeyen" Then
                Set foundDescRange = Nothing
                'List search
                Dim foundRanges As Range
                On Error Resume Next
                Set foundRanges = modFindAll64.FindAll(targetWs.Cells(1, 13).EntireColumn, notesRange.Cells(i, 1).value, xlValues, xlPart)
                On Error GoTo 0
                If Not foundRanges Is Nothing Then
                    frmSimilarSelector.lblBankDesc.Caption = notesRange.Cells(i, 1).value
                    frmSimilarSelector.lblAmount.Caption = amountRange.Cells(i, 1).value
                    frmSimilarSelector.lblDate.Caption = datesRange.Cells(i, 1).value
                    Dim tempArr() As Variant
                    ReDim tempArr(foundRanges.count - 1, 2)
                    Dim rng As Variant
                    Dim j As Integer: j = 0
                    For Each rng In foundRanges
                        tempArr(j, 0) = rng.offset(0, -10).value
                        tempArr(j, 1) = rng.offset(1, -5).value
                        tempArr(j, 2) = rng.offset(1, -7).value
                        j = j + 1
                    Next rng
                    frmSimilarSelector.lbxSimilars.List() = tempArr
                    frmSimilarSelector.show
                    If frmSimilarSelector.FrmAnswer = "Cancel" Then
                        Unload frmSimilarSelector
                    ElseIf frmSimilarSelector.FrmAnswer = "Update" Then
                        notesRange.Cells(i, 1).value = frmSimilarSelector.tbxDesc.Text
                        expenseCategoryRange.Cells(i, 1).value = frmSimilarSelector.tbxToAcct.Text
                        specialCategoryRange.Cells(i, 1).value = frmSimilarSelector.cbxSpecial.Text
                        Set foundRanges = Nothing
                        Unload frmSimilarSelector
                    Else
                        MsgBox "error selection."
                        Unload frmSimilarSelector
                    End If
                Else
                    'no similar can be found
                End If
                'SAME SEARCHING
            End If
        Else
            'Record fully entered no need to check for same or similar entry.
        End If
    Next i
    
    If MsgBox("Kategorisiz kayıtlar gözden geçirildi. Yinelenen değerler kontrol edilsin mi?", _
          vbYesNo + vbQuestion, "Importer") = vbNo Then Exit Sub
    '<<< SAME SEARCHING PART
       
    '>>> DUPLICATE FIND PART
    For i = datesRange.Rows.count To 1 Step -1
        If CheckDuplicate(datesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs, CStr(bankDescRange.Cells(i, 1).value)) = 0 Then
            'unique record.
        Else
            'duplicate entry found
            Debug.Print "Duplicate found"
            datesRange.Cells(i, 1).Interior.ColorIndex = 3
            amountRange.Cells(i, 1).Interior.ColorIndex = 3
        End If
    Next i
     
    If MsgBox("Kategorisiz kayıtlar gözden geçirildi. Verileri tabloya yazmaya devam edilsin mi?", _
          vbYesNo + vbQuestion, "Importer") = vbNo Then Exit Sub
    '<<< DUPLICATE FIND PART
    
    '********************************
    'Make way for and add new data
    '********************************
    Dim startRow As Long
    Dim reconcileNoteRow As Long
    For i = datesRange.Rows.count To 1 Step -1
        'REAL Check duplicates!
        If CheckDuplicate(datesRange.Cells(i, 1).value, CDbl(amountRange.Cells(i, 1).value), targetWs, CStr(bankDescRange.Cells(i, 1).value)) = 0 Then
        '            '> find start row for yourself
            startRow = 2 'by default it is 2
            Do Until datesRange.Cells(i, 1).value >= targetWs.Cells(startRow, 1).value And targetWs.Cells(startRow, 1).value <> ""
                startRow = startRow + 1
            Loop
            '< find start row for yourself
            targetWs.Cells(startRow, 1).EntireRow.Insert
            targetWs.Cells(startRow, 1).EntireRow.Insert
            targetWs.Cells(startRow, 1).value = datesRange.Cells(i, 1).value
            targetWs.Cells(startRow, 2).value = "!" 'Random id generator??
            targetWs.Cells(startRow, 3).value = notesRange.Cells(i, 1).value
            targetWs.Cells(startRow, 5).value = "CURRENCY::TRY"
            targetWs.Cells(startRow, 8).value = targetWs.Cells(startRow + 2, 8).value '@TODO decoupling with a dictionary would be fine
            targetWs.Cells(startRow + 1, 8).value = expenseCategoryRange.Cells(i, 1).value
            targetWs.Cells(startRow, 9).value = amountRange.Cells(i, 1).value
            targetWs.Cells(startRow, 13).value = bankDescRange.Cells(i, 1).value
            '> check for commodity transaction. If it is then you have to use somethings...
            If specialCategoryRange.Cells(i, 1).value <> "" Then 'special transaction
                Select Case specialCategoryRange.Cells(i, 1).value
                Case "Buy/Sell"
                    targetWs.Cells(startRow + 1, 9).value = IIf(amountRange.Cells(i, 1).value < 0, 1, -1) * _
                                                                        CommodityCount(notesRange.Cells(i, 1).value)
                    targetWs.Cells(startRow + 1, 10).value = -1 * (amountRange.Cells(i, 1).value / targetWs.Cells(startRow + 1, 9))
                    targetWs.Cells(startRow + 1, 6).value = IIf(amountRange.Cells(i, 1).value < 0, "Buy", "Sell")
                Case Else
                    targetWs.Cells(startRow + 1, 9).value = amountRange.Cells(i, 1).value * -1
                    targetWs.Cells(startRow + 1, 10).value = 1
                    targetWs.Cells(startRow + 1, 6).value = specialCategoryRange.Cells(i, 1).value
                End Select
            Else
                targetWs.Cells(startRow + 1, 9).value = amountRange.Cells(i, 1).value * -1
                targetWs.Cells(startRow + 1, 10).value = 1
            End If
            '< ...
            targetWs.Cells(startRow, 10).value = 1
            
            ' > adding the UP reconcile note
            reconcileNoteRow = targetWs.Cells(startRow, 11).End(xlUp).Row
            Do Until reconcileNoteRow = 1
                targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
                reconcileNoteRow = targetWs.Cells(reconcileNoteRow, 11).End(xlUp).Row
            Loop
            ' > adding the DOWN reconcile note
            reconcileNoteRow = targetWs.Cells(startRow, 11).End(xlDown).Row
            If targetWs.Cells(reconcileNoteRow, 1).value = datesRange.Cells(i, 1).value Then
                targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
            Else
                ' reconcile that belongs to another date
            End If
            ' > adding the TOP reconcile note
            ' > @TODO what about most top reconcile??
            reconcileNoteRow = targetWs.Cells(2, 11).Row 'weird
            ' Do ?
            
            ' Loop ?
            targetWs.Cells(reconcileNoteRow, 4).value = targetWs.Cells(reconcileNoteRow, 4).value + amountRange.Cells(i, 1).value
    
            
            '
        Else
            'duplicate entry found
            Debug.Print "Duplicate found"
            datesRange.Cells(i, 1).Interior.ColorIndex = 3
            amountRange.Cells(i, 1).Interior.ColorIndex = 3
        End If
    Next i
    ' To not forget. it is a error vector.
    ws.Cells(2, 1).value = ""
' Error Handlers
    Exit Sub
WRONGDATACOUNT:
    MsgBox ("wrong data count.")
    Exit Sub
WRONGPAGENAME:
    MsgBox ("wrong page name you give.")
    Exit Sub
End Sub
Private Function MakeDecimalCalculations(a As Double, b As Double) As String
    MakeDecimalCalculations = CDec(-a) / CDec(b)
End Function


'===================================================
'>> Check Duplicates
'===================================================
Private Sub CheckDuplicate_Test()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("SheKrediKartı")
    ' Should return 0 (no match)
    Debug.Print "Expect 0: " & CheckDuplicate("01.01.2099", CDbl(-100), ws, "TEST")
    ' Should return 0 (date matches, amount doesn't)
    Debug.Print "Expect 0: " & CheckDuplicate("06.01.2023", CDbl(-999.99), ws, "")
    ' Real duplicate: date+amount+desc all match (adjust to a known row in sheet)
    Debug.Print "Expect >0: " & CheckDuplicate("06.01.2023", CDbl(-155), ws, "")
    ' KKDF vs BSMV: same date+amount, different desc -> should NOT be duplicate
    Debug.Print "Expect 0 (diff desc): " & CheckDuplicate("14.02.2024", CDbl(-114.99), ws, "KKDF FARKLI ACIKLAMA")
    ' Same date+amount+desc -> duplicate
    Debug.Print "Expect >0 (same desc): " & CheckDuplicate("14.02.2024", CDbl(-58#), ws, "10314 BORNOVA HÜKÜMET KİZMİR TR")
End Sub

' Returns 0 if not duplicate, row number if duplicate found.
' Matches on col1=date, col9=amount, col13=bankDesc (if provided).
Private Function CheckDuplicate(chkDate As Variant, chkAmount As Double, targetWs As Worksheet, Optional chkDesc As String = "") As Long
    CheckDuplicate = 0
    Dim targetDate As Date
    On Error Resume Next
    targetDate = CDate(chkDate)
    If Err.Number <> 0 Then Exit Function
    On Error GoTo 0

    Dim lastRow As Long
    lastRow = targetWs.Cells(targetWs.Rows.count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If targetWs.Cells(i, 1).value <> "" Then
            Dim cellDate As Date
            On Error Resume Next
            cellDate = CDate(targetWs.Cells(i, 1).value)
            If Err.Number = 0 Then
                If cellDate = targetDate And targetWs.Cells(i, 9).value = chkAmount Then
                    If chkDesc = "" Or StrComp(CStr(targetWs.Cells(i, 13).value), chkDesc, vbTextCompare) = 0 Then
                        CheckDuplicate = i
                        Exit Function
                    End If
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next i
End Function
'===================================================
'<< Check Duplicates
'===================================================

'===================================================
'>> Find Commodity Count
'===================================================
Private Function CommodityCount(str As String) As Double
    Dim result As Variant
    Dim oReg As Object
    Set oReg = CreateObject("VBScript.RegExp")
    With oReg
        .Pattern = "(\d+ Pay)|(x\d+.\d+)"
    End With
    Set result = oReg.Execute(str)
    If result.count = 1 Then
        result = result(0).value
        If left(result, 1) = "x" Then
            result = CDbl(Replace(Replace(result, "x", ""), ".", ","))
        ElseIf right(result, 1) = "y" Then
            result = CDbl(Replace(Replace(result, " Pay", ""), ".", ","))
        Else
            Debug.Print "unknown transaction format"
            result = 666666
        End If
        CommodityCount = result
    Else
        CommodityCount = 666666
    End If
End Function
'===================================================
'<< Find Commodity Count
'===================================================












