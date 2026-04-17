Attribute VB_Name = "PriceParserMod"
Option Explicit

Public Sub ParsePricesFrom_PortfolioPerformance()
    'portfolio performance yazılımından gelen CSV fiyatlar dosyasını hledger formatına çevirmektedir.
    'Direk dosyayı ramlere alıp basıyor o yüzden baya hızlı vs. ancak sadece
    'bazı entrylerde " ile başlayan ve arasında nokta olan arkadaşlar var onları aradan süzmem lazım
    'bütün dosyaya kolayca regex yapabilirsek kral!
    
    
    Dim inPath As String, outPath As String
    Dim ff As Integer, fout As Integer
    Dim buffer() As Byte, fileSize As Long
    Dim pos As Long, byteVal As Byte, colValue As String
    Dim headers() As String, headersCurrency() As String, headersTicker() As String
    Dim countDict As Object, ignore As Object, currencyMap As Object
    Dim outBuf() As String, outBufSize As Long
    Dim validCols() As Long, validColsCount As Long
    Dim lineParts() As String, linePartsCount As Long
    Dim dateVal As String, val As String, Ticker As Variant
    Dim processedRows As Long, chunkSize As Long
    Dim i As Long
        
    ' ---- FILE PATHS ----
    inPath = Config.COMMODITY_PRICES_FILE
    outPath = Replace(inPath, ".csv", ".hledger")
    
    ' ---- DICTIONARIES ----
    Set countDict = CreateObject("Scripting.Dictionary")
    Set ignore = CreateObject("Scripting.Dictionary")
'    ignore.Add "TR", True
'    ignore.Add "XU100", True
    
    Set currencyMap = CreateObject("Scripting.Dictionary")
    currencyMap.Add "TRY=X", "USD"
    currencyMap.Add "EURTRY=X", "EUR"
    currencyMap.Add "ALTIN", """ALTIN.S1"""
    
    ' ---- BINARY READ ----
    ff = FreeFile
    Open inPath For Binary As #ff
    fileSize = LOF(ff)
    ReDim buffer(1 To fileSize)
    Get #ff, , buffer
    Close #ff
    
    ' ---- HEADER PARSING ----
    ReDim headers(1 To 1024)
    ReDim headersCurrency(1 To 1024)
    ReDim headersTicker(1 To 1024)
    ReDim validCols(1 To 1024)
    validColsCount = 0
    colValue = ""
    pos = 1
    
    Do While pos <= fileSize
        byteVal = buffer(pos)
        If byteVal = 44 Or byteVal = 10 Or byteVal = 13 Then
            ' duplicate fix
            If countDict.Exists(colValue) Then
                countDict(colValue) = countDict(colValue) + 1
                colValue = colValue & "_" & countDict(colValue)
            Else
                countDict.Add colValue, 1
            End If
            
            i = validColsCount + 1
            headers(i) = colValue
            If currencyMap.Exists(colValue) Then
                'headersCurrency(i) = currencyMap(colValue)
                headersTicker(i) = currencyMap(colValue)
            Else
                'headersCurrency(i) = "TRY"
                headersTicker(i) = colValue
            End If
            If ContainsDigit(colValue) Then
                headersTicker(i) = """" & headersTicker(i) & """"
            Else
                headersTicker(i) = headersTicker(i)
            End If
            
'            ' valid columns
'            If Not ignore.Exists(colValue) Then
'                validColsCount = validColsCount + 1
'                validCols(validColsCount) = i
'            End If
            
            ' valid columns only
            If Not ignore.Exists(colValue) And colValue <> "Date" Then
                validColsCount = validColsCount + 1
                validCols(validColsCount) = i
            Else
            End If
            
            colValue = ""
            If byteVal = 10 Then pos = pos + 1: Exit Do
        Else
            colValue = colValue & Chr(byteVal)
        End If
        pos = pos + 1
    Loop
    
    ' ---- OUTPUT SETUP ----
    chunkSize = 1000
    ReDim outBuf(1 To chunkSize)
    outBufSize = 0
    processedRows = 0
    
    ' ---- FILE OPEN FOR OUTPUT ----
    fout = FreeFile
    Open outPath For Output As #fout
    Print #fout, "; ---"
    Print #fout, "; Prices"
    
    ' ---- PROCESS ROWS ----
    ReDim lineParts(1 To 1024)
    Dim inQuotes As Boolean
    Do While pos <= fileSize
        
        ' parse line
        linePartsCount = 0
        colValue = ""
        inQuotes = False
        Do While pos <= fileSize
            byteVal = buffer(pos)
            If byteVal = 34 Then ' 34 = Çift Tırnak (")
                inQuotes = Not inQuotes
            ElseIf byteVal = 44 And Not inQuotes Then ' Tırnak dışındaki virgül (Ayırıcı)
                linePartsCount = linePartsCount + 1
                lineParts(linePartsCount) = colValue
                colValue = ""
            ElseIf (byteVal = 10 Or byteVal = 13) And Not inQuotes Then ' Satır sonu
                linePartsCount = linePartsCount + 1
                lineParts(linePartsCount) = colValue
                colValue = ""
                ' CRLF kontrolü (13'ten sonra 10 gelirse atla)
                If byteVal = 13 And pos < fileSize Then
                    If buffer(pos + 1) = 10 Then pos = pos + 1
                End If
                pos = pos + 1
                Exit Do
            Else
                colValue = colValue & Chr(byteVal)
            End If
            pos = pos + 1
        Loop
        
        If linePartsCount < 1 Then GoTo NextLine
        dateVal = Trim$(lineParts(1)) ' ilk kolon her zaman Date
        If Len(dateVal) < 8 Then GoTo NextLine ' YYYY-MM-DD formatından kısa olamaz
        If Len(dateVal) = 0 Then GoTo NextLine
        
        ' valid columns only
        For i = 1 To validColsCount
            
            Dim c As Long
            c = validCols(i)
            'If headersTicker(c) = "" Then Stop
            If c > linePartsCount Then GoTo NextCol
            
            'If headersTicker(c) = """ALTIN.S1""" And dateVal = "2025-12-26" Then Stop
            val = Trim$(lineParts(c + 1)) '1. sütunu atlıyoruz orası date kolonu oyüzden +1
            
            If Len(val) = 0 Then GoTo NextCol
            
            'If InStr(val, """") > 0 Then Stop
            'If InStr(val, "2025-11-05") > 0 Then Stop
            'val = Replace$(val, """", "")      ' tırnakları temizle
            'val = Replace$(val, ",", "")        ' binlik ayırıcıyı kaldır
            val = Replace$(val, ",", "")      ' ondalığı virgüle çevir
            val = Replace$(val, ".", ",")      ' ondalığı virgüle çevir
            
            outBufSize = outBufSize + 1
            outBuf(outBufSize) = "P    " & dateVal & "    " & headersTicker(c) & "    " & val & " " & headersCurrency(c) & "TRY"
            ' flush if buffer full
            If outBufSize = chunkSize Then
                For Each Ticker In outBuf
                    Print #fout, Ticker
                Next
                outBufSize = 0
            End If
NextCol:
        Next i
        
        processedRows = processedRows + 1
        If processedRows Mod 100 = 0 Then
            Application.StatusBar = "Processing row " & pos / fileSize & "..."
        End If
NextLine:
    Loop
    
    Dim j As Long
    ' flush remaining
    For j = 1 To outBufSize
        Print #fout, outBuf(j)
    Next
    
    Close #fout
    Application.StatusBar = False
    MsgBox "Tamamdır › " & outPath, vbInformation

End Sub

' ---- HELPER ----
Private Function ContainsDigit(s As String) As Boolean
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) Like "#" Then
            ContainsDigit = True
            Exit Function
        End If
    Next i
End Function





