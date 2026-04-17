Attribute VB_Name = "TestUTF8"
'
' TestUTF8.bas
' UTF-8 Round-Trip Validation Test
' Purpose: Verify Turkish characters survive Excel › Hledger › Excel cycle
'

Option Explicit

Public Sub TestUTF8RoundTrip()
    ' Test Turkish characters in round-trip
    LogManager.LogInfo "=== UTF-8 Round-Trip Test Started ==="

    Dim testCases As Variant
    Dim testString As String
    Dim testFile As String
    Dim fso As Object
    Dim testSheet As Worksheet
    Dim resultRow As Long

    ' Turkish test strings with special characters
    testCases = Array( _
        "Türkçe Test: İ, ş, ç, ğ, ü, ö", _
        "Özel İşlem Açıklaması", _
        "Hesap Dökümü (Tarih: 2024-04-17)", _
        "TEB Yatırım İşlemi - Hisse: TKM", _
        "Kredi Kartı Ödeme: Şube-1" _
    )

    ' Create test worksheet
    On Error Resume Next
    Application.Worksheets("UTF8Test").Delete
    On Error GoTo 0

    Set testSheet = ThisWorkbook.Worksheets.Add
    testSheet.name = "UTF8Test"

    ' Write test cases
    resultRow = 1
    testSheet.Cells(resultRow, 1).value = "Original"
    testSheet.Cells(resultRow, 2).value = "Written to File"
    testSheet.Cells(resultRow, 3).value = "Read Back"
    testSheet.Cells(resultRow, 4).value = "Match?"

    testFile = Config.DATA_FOLDER & "UTF8Test.txt"
    Dim i As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")

    For i = LBound(testCases) To UBound(testCases)
        testString = testCases(i)
        resultRow = resultRow + 1

        ' Write to UTF-8 file
        Call WriteUTF8File(testFile, testString)

        ' Read back
        Dim readBack As String
        readBack = ReadUTF8File(testFile)

        ' Compare
        testSheet.Cells(resultRow, 1).value = testString
        testSheet.Cells(resultRow, 2).value = "? Written"
        testSheet.Cells(resultRow, 3).value = readBack
        testSheet.Cells(resultRow, 4).value = IIf(testString = readBack, "? PASS", "? FAIL")

        If testString = readBack Then
            LogManager.LogInfo "UTF-8 Test PASS: " & testString
        Else
            LogManager.LogWarning "UTF-8 Test FAIL: Expected '" & testString & "', got '" & readBack & "'"
        End If
    Next i

    fso.DeleteFile testFile, True
    LogManager.LogInfo "=== UTF-8 Round-Trip Test Completed ==="
    MsgBox "UTF-8 validation complete. Check 'UTF8Test' sheet.", vbInformation, "UTF-8 Test"
End Sub

Private Sub WriteUTF8File(filePath As String, content As String)
    Dim objFS As Object
    Dim iFile As Integer

    iFile = FreeFile
    Open filePath For Output As #iFile
    Print #iFile, content
    Close #iFile

    ' Convert to UTF-8 without BOM
    Set objFS = CreateObject("ADODB.Stream")
    With objFS
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        .Position = 3 ' Skip BOM
        .Type = 1
    End With

    Dim objFSNoBOM As Object
    Set objFSNoBOM = CreateObject("ADODB.Stream")
    With objFSNoBOM
        .Type = 1
        .Open
        objFS.CopyTo objFSNoBOM
    End With

    objFSNoBOM.SaveToFile filePath, 2
    objFSNoBOM.Close
    objFS.Close
End Sub

Private Function ReadUTF8File(filePath As String) As String
    Dim objFS As Object
    Set objFS = CreateObject("ADODB.Stream")

    With objFS
        .Charset = "UTF-8"
        .Open
        .LoadFromFile filePath
        ReadUTF8File = .ReadText
        .Close
    End With
End Function

' Quick test: run from Immediate Window
' Call TestUTF8RoundTrip()








