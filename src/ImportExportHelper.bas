Attribute VB_Name = "ImportExportHelper"
Option Explicit

Sub ExportAndEraseWorksheet()

    If MsgBox( _
        "Seçili renkteki worksheet'ler CSV olarak export edilecek ve SİLİNECEK." & vbCrLf & _
        "Emin misin?", _
        vbQuestion + vbYesNo + vbDefaultButton2, _
        "Export & Erase") <> vbYes Then
        Exit Sub
    End If
    Application.DisplayAlerts = False
    Dim path As String
    Dim ws As Worksheet
    Dim i As Long: i = 1
    For Each ws In ThisWorkbook.Worksheets
        If ws.Tab.Color = 11854022 Then
            path = ThisWorkbook.path & "\CSVDepot\" & IIf(Len(CStr(i)) < 2, "0" & i, i) & "_" & ws.name & ".csv"
            Call FastCSVExport(ws, path)
            ws.Delete
        Else
        End If
        i = i + 1
    Next ws
    Application.DisplayAlerts = True

End Sub

Private Sub FastCSVExport(ws As Worksheet, path As String)

    Dim arr As Variant
    Dim r As Long, c As Long
    Dim sb As String
    Dim line As String
    
    arr = ws.UsedRange.Value2
    For r = 1 To UBound(arr, 1)
        line = ""
        For c = 1 To UBound(arr, 2)
            line = line & ";" & arr(r, c)
        Next c
        sb = sb & Mid$(line, 2) & vbCrLf
    Next r
    ' UTF-8 (BOM'suz) yaz
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    With stm
    .Type = 2              ' text
    .Charset = "utf-8"
    .Open
    .WriteText sb
    .Position = 3          ' BOM'u atla
    .SaveToFile path, 2
    .Close
    End With

End Sub


Sub ImportCSVsAsNewWorksheet()
    
    Dim path As String
    Dim ws As Worksheet
    Dim baseName As String

    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    Dim fle As Variant
    Dim fleNum As Long
    Dim fleName As String
    Dim flePath As Variant
    Dim files As Variant
    files = GetSortedCSVFiles(ThisWorkbook.path & "\CSVDepot")
    
    For Each flePath In files
        Set fle = fso.GetFile(flePath)
        fleNum = CLng(left(fle.name, InStr(1, fle.name, "_") - 1))
        fleName = Mid(fle.name, InStr(1, fle.name, "_") + 1)
        fleName = left(fleName, Len(fleName) - 4)
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(fleNum - 1))
        ws.name = fleName
        With ws.QueryTables.Add( _
            Connection:="TEXT;" & fle.path, _
            Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileCommaDelimiter = False
            .TextFileConsecutiveDelimiter = False
            .TextFileSpaceDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = True
            .TextFileDecimalSeparator = ","
            .TextFilePlatform = 65001 ' UTF-8
            .TextFileTrailingMinusNumbers = False
            Dim formatArr(1 To 13) As Variant
            Dim i As Long
            For i = LBound(formatArr) To UBound(formatArr)
                formatArr(i) = xlTextFormat
            Next i
            .TextFileColumnDataTypes = formatArr
            .Refresh BackgroundQuery:=False
            .Delete
        End With
        ' > formatting some of the columns
        ws.Columns(9).TextToColumns _
            Destination:=ws.Cells(1, 9), _
            DataType:=xlDelimited, _
            FieldInfo:=Array(1, xlGeneralFormat), _
            DecimalSeparator:=",", _
            ThousandsSeparator:="."
        ws.Columns(9).NumberFormat = "#,##0.00"
        ws.Columns(11).TextToColumns _
            Destination:=ws.Cells(1, 11), _
            DataType:=xlDelimited, _
            FieldInfo:=Array(1, xlGeneralFormat), _
            DecimalSeparator:=",", _
            ThousandsSeparator:="."
        ws.Columns(11).NumberFormat = "#,##0.00"
        ws.Columns(1).TextToColumns _
            Destination:=ws.Cells(1, 1), _
            DataType:=xlDelimited, _
            FieldInfo:=Array(1, xlDMYFormat)
        ws.Columns(1).NumberFormat = "d.mm.yyyy"
        ' < formatting some of the columns
        ws.Tab.Color = 11854022
        ws.activate
        ActiveWindow.Zoom = 100
    Next flePath
End Sub

Function GetSortedCSVFiles(folderPath As String) As Variant

    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim arr() As String
    Dim i As Long: i = 0
    For Each f In fso.GetFolder(folderPath).files
        If LCase(fso.GetExtensionName(f.name)) = "csv" Then
            ReDim Preserve arr(i)
            arr(i) = f.path
            i = i + 1
        End If
    Next
    ' basit bubble sort (dosya sayısı azsa yeter)
    Dim j As Long, tmp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                tmp = arr(i)
                arr(i) = arr(j)
                arr(j) = tmp
            End If
        Next
    Next
    GetSortedCSVFiles = arr

End Function


' ============================================================
' CreateTemplateWorkbook
' Mevcut workbook'un kisisel veri iceRmeyen bir kopyasini olusturur.
' Yesil sekmeli (hesap) sayfalarindaki veri satirlari silinir,
' LOGS ve Bank_Info tamamen temizlenir.
' Sonuc: proje klasorunde TheBerk_Template.xlsm
' ============================================================
Public Sub CreateTemplateWorkbook()
    If MsgBox("Template.xlsm olusturulacak — tum kisisel veri silinip ornek veri eklenecek." & vbCrLf & _
              "Devam?", vbYesNo + vbQuestion, "Create Template") = vbNo Then Exit Sub

    Dim templatePath As String
    templatePath = ThisWorkbook.path & "\Template.xlsm"

    Application.DisplayAlerts = False
    ThisWorkbook.SaveCopyAs templatePath

    Dim twb As Workbook
    Set twb = Workbooks.Open(templatePath)

    ' --- 1. TUM sayfalari temizle ---
    Dim ws As Worksheet
    Dim lastRow As Long
    For Each ws In twb.Worksheets
        If ws.name = "LOGS" Then
            ws.Cells.Delete
        ElseIf ws.name = "Events" Or ws.name = "ExpenseAnalysis" Then
            ws.Delete
        ElseIf ws.name = "IMPORT" Then
            lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If lastRow > 1 Then ws.Rows("4:" & lastRow).value = ""
        Else
            lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
            If lastRow > 1 Then ws.Rows("2:" & lastRow).Delete
        End If
    Next ws

    ' --- 2. ACCOUNTS: ornek hesap hiyerarsisi ---
    Dim wsA As Worksheet
    Set wsA = SheetByCode(twb, "ACCOUNTS")
    If Not wsA Is Nothing Then
        Dim accts As Variant
        accts = Array("Varlıklar", "Varlıklar:Banka", "Varlıklar:Banka:OrnekBanka", _
                      "Varlıklar:Yatırım", "Varlıklar:Yatırım:Hisse", _
                      "Borçlar", "Borçlar:KrediKarti", "Borçlar:KrediKarti:OrnekKart", _
                      "Gider", "Gider:Market", "Gider:Yemek", "Gider:Ulasim", "Gider:Fatura", _
                      "Gelir", "Gelir:Maas", "Gelir:Yatirim")
        Dim i As Long
        For i = 0 To UBound(accts)
            wsA.Cells(i + 2, 1).value = accts(i)
        Next i
    End If

    ' --- 3. COMMODITIES: ornek emtia ---
    Dim wsC As Worksheet
    Set wsC = SheetByCode(twb, "COMMODITIES")
    If Not wsC Is Nothing Then
        wsC.Cells(2, 1).value = "ORNEKSIRKET"
        wsC.Cells(3, 1).value = "ALTIN.S1"
        wsC.Cells(4, 1).value = "USD"
    End If

    ' --- 4. RULES: ornek kategorilendirme kurallari ---
    Dim wsR As Worksheet
    Set wsR = SheetByCode(twb, "RULESS")
    If Not wsR Is Nothing Then
        ' A=Active B=DescRuleType C=Description D=AmountOp E=Amount
        ' F=Account G=ToAccount H=NewDescription I=Special J=Priority
        Dim rData As Variant
        rData = Array( _
            Array(True, "CONTAINS", "MARKET", "", "", "", "Gider:Market", "", "", 1), _
            Array(True, "CONTAINS", "MAAS", ">", 0, "", "Gelir:Maas", "Maas", "", 2), _
            Array(True, "REGEX", "(FATURA|ELEKTRIK|DOGALGAZ)", "", "", "", "Gider:Fatura", "", "", 1), _
            Array(True, "CONTAINS", "HISSE AL", "<", 0, "", "Varlıklar:Yatirim:Hisse", "", "Buy/Sell", 5) _
        )
        Dim r As Long
        For r = 0 To UBound(rData)
            Dim col As Long
            For col = 0 To 9
                wsR.Cells(r + 2, col + 1).value = rData(r)(col)
            Next col
        Next r
    End If

    ' --- 5. Ilk yesil sekmeli sayfaya ornek islemler ---
    Dim wsT As Worksheet
    Dim i As Integer
    For Each ws In twb.Worksheets
        If ws.Tab.Color = 11854022 Then
            i = i + 1
            If i = 1 Then Set wsT = ws
            wsT.name = "ACCOUNT-" & i
        Else
        End If
    Next ws
    If Not wsT Is Nothing Then
        ' Islem 1: Maas geliri (2 satir)
        wsT.Cells(2, 1).value = CDate("15.01.2024")
        wsT.Cells(2, 2).value = "!"
        wsT.Cells(2, 3).value = "Maas"
        wsT.Cells(2, 5).value = "CURRENCY::TRY"
        wsT.Cells(2, 8).value = "Varlıklar:Banka:OrnekBanka"
        wsT.Cells(2, 9).value = 5000
        wsT.Cells(2, 10).value = 1
        wsT.Cells(2, 13).value = "MAAS ODEMESI"
        wsT.Cells(3, 8).value = "Gelir:Maas"
        wsT.Cells(3, 9).value = -5000
        wsT.Cells(3, 10).value = 1
        ' Islem 2: Market gideri (2 satir)
        wsT.Cells(4, 1).value = CDate("16.01.2024")
        wsT.Cells(4, 2).value = "!"
        wsT.Cells(4, 3).value = "Market Alisverisi"
        wsT.Cells(4, 5).value = "CURRENCY::TRY"
        wsT.Cells(4, 8).value = "Varlıklar:Banka:OrnekBanka"
        wsT.Cells(4, 9).value = -150
        wsT.Cells(4, 10).value = 1
        wsT.Cells(4, 13).value = "MIGROS MARKET"
        wsT.Cells(5, 8).value = "Gider:Market"
        wsT.Cells(5, 9).value = 150
        wsT.Cells(5, 10).value = 1
    End If

    ' --- 6. IMPORT hedef adini temizle ---
    On Error Resume Next
    twb.Worksheets("IMPORT").Cells(2, 1).value = ""
    On Error GoTo 0

    twb.Save
    twb.Close
    Application.DisplayAlerts = True

    MsgBox "Template olusturuldu:" & vbCrLf & templatePath, vbInformation, "Create Template"
End Sub

Private Function SheetByCode(wb As Workbook, codeName As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.codeName = codeName Then Set SheetByCode = ws: Exit Function
    Next ws
End Function









