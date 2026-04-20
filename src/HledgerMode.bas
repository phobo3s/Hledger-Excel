Attribute VB_Name = "HledgerMode"
Option Explicit

Public Sub InvokeHledgerMode()

    Application.OnKey "%+{X}", "RunHledgerCommand"
    Application.OnKey "%+{C}", "UnInvokeHledgerMode"

End Sub

Private Sub UnInvokeHledgerMode()
    
    Application.OnKey "%+{X}"
    Application.OnKey "%+{C}"

End Sub

Private Sub RunHledgerCommand()

    ActiveSheet.Cells(2, 2).EntireColumn.Resize(, ActiveSheet.Columns.count - 2).value = ""
    Dim cmdText As String
    cmdText = ActiveCell.value
    If cmdText = "" Then Exit Sub
    
    Dim sh As Object
    Set sh = CreateObject("Wscript.Shell")

    Dim shResponse As Object
    Dim shOutput As Object
        
    Dim isOutputCSV As Boolean
    If InStr(UCase(cmdText), "-O CSV") <> 0 Then
        isOutputCSV = True
    Else
        cmdText = cmdText & " -O csv --commodity-column"
        isOutputCSV = True
    End If

    Set shResponse = sh.Exec("cmd.exe /u /c chcp 65001" & "&&" & "hledger " & cmdText & "")
    Set shOutput = shResponse.StdOut

    Dim outputLine As String
    Dim outputLineSplited() As String
    
    Dim i As Integer
    Dim startLagLines As Integer
    If isOutputCSV Then
        startLagLines = 1
    Else
        startLagLines = 1
    End If
    Do While Not shOutput.AtEndOfStream
        If shOutput.line > startLagLines Then
            outputLine = shOutput.ReadLine
            outputLine = ConvertCharsToTurkish(outputLine)
            outputLine = Mid$(outputLine, 2, Len(outputLine) - 2)
            If isOutputCSV Then
                outputLineSplited = Split(outputLine, Chr(34) & "," & Chr(34))
                ActiveSheet.Cells(2 + i, 3).Resize(1, UBound(outputLineSplited) + 1) = outputLineSplited
            Else
                ActiveSheet.Cells(2 + i, 3).value = "'" & outputLine
            End If
            i = i + 1
        Else
            shOutput.ReadLine
        End If
    Loop
    If isOutputCSV Then
        Dim cll As Range
        Dim resultRange As Range
        i = 2
        Do
            Set cll = ActiveSheet.Cells(i, 3)
            cll.offset(0, -1).value = "||"
            i = i + 1
        Loop While (cll <> "")
        Set resultRange = ActiveSheet.Cells(2, 3).Resize(ActiveSheet.UsedRange.Rows.count - 2, ActiveSheet.UsedRange.Columns.count - 2)
        For Each cll In resultRange
            If IsNumeric(cll.value) = True And cll.value <> "" Then cll.value = CDbl(cll.value)
            If IsNumeric(left(cll.value, 4)) = True And _
                Mid(cll.value, 5, 1) = "-" And _
                IsNumeric(Mid(cll.value, 6, 2)) = True And _
                Mid(cll.value, 8, 1) = "-" And _
                IsNumeric(Mid(cll.value, 9, 2)) = True _
                Then cll.value = CDate(cll.value)  '2025-01-31
        Next cll
    Else
    End If

End Sub

Private Function ConvertCharsToTurkish(str As String) As String

    str = Replace(str, "Ä±", "ı")
    str = Replace(str, "Ã¶", "ö")
    str = Replace(str, "Ã§", "ç")
    str = Replace(str, "ÅŸ", "ş")
    str = Replace(str, "ÄŸ", "ğ")
    str = Replace(str, "Ä°", "İ")
    str = Replace(str, "Ã–", "Ö")
    str = Replace(str, "Ãœ", "Ü")
    str = Replace(str, "Ã¼", "ü")
    ConvertCharsToTurkish = str

End Function












