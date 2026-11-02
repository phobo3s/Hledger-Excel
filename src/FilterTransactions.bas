Attribute VB_Name = "FilterTransactions"
'==============================================================================
' FilterTransactions.bas
' hledger-Excel — Çoklu Sütun Filtre Makrosu
' Made with Claude
' KURULUM:
'   1. Bu modülü projeye import et (File > Import File) veya içeriği kopyala
'   2. DATE_COL sabitini kendi yapına göre ayarla
'   3. DATA_START_ROW sabitini ayarla (başlık satırı kaç?)
'   4. Araç çubuğuna veya kısayola FilterByColumns'ı bağla
'
' KULLANIM:
'   - FilterByColumns  › filtre uygula
'   - ClearFilter      › tüm gizlemeyi kaldır
'==============================================================================

Option Explicit

' == Konfigürasyon ============================================================
Private Const DATE_COL       As Long = 1  ' Date sütununun index'i (A=1, B=2, ...)
Private Const DATA_START_ROW As Long = 2  ' İlk veri satırı (başlık atlanır)
Private Const ACCOUNT_COL    As Long = 8  ' Account isimleri kolonu son satırı bulmak için
' =============================================================================

'==============================================================================
' Public: Ana filtre makrosu
'==============================================================================
Public Sub FilterByColumns()

    Dim ws          As Worksheet
    Dim lastRow     As Long
    Dim matchedRows As Object   ' Scripting.Dictionary  {txStart -> True}
    Dim colRange    As Range
    Dim colIndex    As Long
    Dim searchVal   As String
    Dim r           As Long
    Dim isFirstPass As Boolean

    Set ws = ActiveSheet
    lastRow = lastDataRow(ws)

    If lastRow < DATA_START_ROW Then
        MsgBox "Sayfada veri bulunamadı.", vbExclamation
        Exit Sub
    End If

    ' İlk tur için tüm satırları aday olarak başlat
    Set matchedRows = CreateObject("Scripting.Dictionary")
    isFirstPass = True

    '== Sütun seçim döngüsü ====================================================
    Do
        Set colRange = Nothing
        ' Type:=8 › kullanıcı hücre/sütun seçer; Cancel › Nothing
        On Error Resume Next
        Set colRange = Application.InputBox( _
            Prompt:="Filtre uygulanacak sütunu seçin." & vbLf & _
                    "(İptal veya boş bırak › filtreyi uygula)", _
            Title:="Sütun Seç", _
            Type:=8)
        On Error GoTo 0

        ' Kullanıcı iptal etti veya boş geçti › döngüyü bitir
        If colRange Is Nothing Then Exit Do

        colIndex = colRange.Columns(1).Column  ' Seçilen sütunun index'i

        ' Aranacak değeri sor
        searchVal = Trim(Application.InputBox( _
            Prompt:="'" & ws.Cells(1, colIndex).value & "' sütununda aranacak değer:", _
            Title:="Filtre Değeri", _
            Type:=2))   ' Type:=2 › string

        If searchVal = "" Or searchVal = "False" Then
            ' İptal basıldıysa (InputBox False döner) bu turu atla
            GoTo NextIteration
        End If

        '== Eşleşen transaction'ları bul ve AND kesişimi uygula ================
        Dim currentMatch As Object
        Set currentMatch = CreateObject("Scripting.Dictionary")

        Dim thisTxStart As Long
        For r = DATA_START_ROW To lastRow
            Dim cellVal As String
            cellVal = Trim(CStr(ws.Cells(r, colIndex).value))

            If InStr(1, cellVal, searchVal, vbTextCompare) > 0 Then
                thisTxStart = FindTxStart(ws, r)
                If isFirstPass Then
                    currentMatch(thisTxStart) = True
                ElseIf matchedRows.Exists(thisTxStart) Then
                    currentMatch(thisTxStart) = True  ' AND: sadece önceki eşleşenler
                End If
            End If
        Next r

        Set matchedRows = currentMatch
        isFirstPass = False

        If matchedRows.count = 0 Then
            MsgBox "Eşleşen transaction bulunamadı." & vbLf & _
                   "Filtre temizleniyor.", vbInformation
            ClearFilter
            Exit Sub
        End If

NextIteration:
    Loop

    '== Sonuç kalmadıysa çık =====================================================
    If isFirstPass Then
        ' Hiç kriter girilmedi
        Exit Sub
    End If

    If matchedRows.count = 0 Then
        MsgBox "Eşleşen satır yok.", vbInformation
        Exit Sub
    End If

    '== txEnd'leri hesapla: {txStart -> txEnd} ===================================
    Dim txStart As Long, txEnd As Long
    Dim key As Variant

    Dim txWithEnd As Object
    Set txWithEnd = CreateObject("Scripting.Dictionary")
    For Each key In matchedRows.keys
        txStart = CLng(key)
        txWithEnd(txStart) = FindTxEnd(ws, txStart, lastRow)
    Next key

    '== Ardışık aralıkları birleştir ============================================
    Dim txRanges As Collection
    Set txRanges = New Collection

    Dim txKeys As Variant
    txKeys = txWithEnd.keys   ' insertion order = tarama sırası = satır sırası

    Dim mergeStart As Long, mergeEnd As Long
    mergeStart = CLng(txKeys(0))
    mergeEnd = CLng(txWithEnd(txKeys(0)))

    Dim i As Long
    For i = 1 To UBound(txKeys)
        Dim nextStart As Long, nextEnd As Long
        nextStart = CLng(txKeys(i))
        nextEnd = CLng(txWithEnd(txKeys(i)))

        If nextStart = mergeEnd + 1 Then
            mergeEnd = nextEnd              ' ardışık › birleştir
        Else
            txRanges.Add Array(mergeStart, mergeEnd)
            mergeStart = nextStart
            mergeEnd = nextEnd
        End If
    Next i
    txRanges.Add Array(mergeStart, mergeEnd)    ' son aralık

    ApplyVisibility ws, txRanges, lastRow

    MsgBox matchedRows.count & " transaction › " & txRanges.count & _
           " aralığa indirgendi.", vbInformation

End Sub


'==============================================================================
' Public: Tüm gizlemeyi kaldır
'==============================================================================
Public Sub ClearFilter()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    'lastRow = lastDataRow(ws)
    lastRow = ws.UsedRange.Rows(ws.UsedRange.Rows.count).Row

    If lastRow >= DATA_START_ROW Then
        ws.Rows(DATA_START_ROW & ":" & lastRow).Hidden = False
    End If
End Sub

'==============================================================================
' Private: Bir satırdan geriye giderek transaction başını bul
'          (Date sütunu dolu olan ilk satır)
'==============================================================================
Private Function FindTxStart(ws As Worksheet, rowNum As Long) As Long
    Dim r As Long
    r = rowNum

    ' Zaten Date dolu mu?
    If Trim(CStr(ws.Cells(r, DATE_COL).value)) <> "" Then
        FindTxStart = r
        Exit Function
    End If

    ' Yukarı git, Date dolu satırı bul
    Do While r >= DATA_START_ROW
        If Trim(CStr(ws.Cells(r, DATE_COL).value)) <> "" Then
            FindTxStart = r
            Exit Function
        End If
        r = r - 1
    Loop

    ' Bulunamazsa (veri bütünlüğü hatası) orijinal satırı döndür
    FindTxStart = rowNum
    MsgBox "transaction nesnesi başlangıcı bulunamadı"
End Function


'==============================================================================
' Private: Transaction başından aşağı giderek transaction sonunu bul
'          (Bir sonraki Date dolu satırın bir öncesi)
'==============================================================================
Private Function FindTxEnd(ws As Worksheet, rowNum As Long, lastRow As Long) As Long
    Dim r As Long

    ' txStart'ın kendisi Date dolu — bir sonraki Date'i ara
    For r = rowNum + 1 To lastRow
        If Trim(CStr(ws.Cells(r, DATE_COL).value)) <> "" Then
            ' Bir sonraki transaction başladı › bir öncesi bizim sonumuz
            FindTxEnd = r - 1
            Exit Function
        End If
    Next r

    ' Son transaction — lastRow'a kadar
    FindTxEnd = lastRow
End Function


'==============================================================================
' Private: Görünürlük uygula
'          txRanges içindeki satırları göster, geri kalanları gizle
'==============================================================================
Private Sub ApplyVisibility(ws As Worksheet, txRanges As Collection, lastRow As Long)

    ' Önce tüm veri satırlarını gizle
    Application.ScreenUpdating = False
    ws.Rows(DATA_START_ROW & ":" & lastRow).Hidden = True

    ' Sonra eşleşen transaction satırlarını göster
    Dim item As Variant
    Dim txStart As Long, txEnd As Long

    For Each item In txRanges
        txStart = item(0)
        txEnd = item(1)
        ws.Rows(txStart & ":" & txEnd).Hidden = False
    Next item

    Application.ScreenUpdating = True
End Sub


'==============================================================================
' Private: Sayfadaki son dolu satırı bul (Date sütununa göre)
'==============================================================================
Private Function lastDataRow(ws As Worksheet) As Long
    lastDataRow = ws.Cells(ws.Rows.count, ACCOUNT_COL).End(xlUp).Row
End Function


