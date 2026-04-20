VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BankGetterHooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Each hook has signature: (chrome As stdChrome, p1 As String, p2 As String, p3 As String)
' Called via: CallByName BankGetterHooks, hookName, VbMethod, chrome, p1, p2, p3

' Iterates 4 TEB deposit accounts using positional Detay link selection.
' p1 = number of accounts (default 4)
Public Sub Hook_TEB_FetchAccounts(chrome As stdChrome, p1 As String, p2 As String, p3 As String)
    Dim accountCount As Integer
    If Len(p1) > 0 And IsNumeric(p1) Then accountCount = CInt(p1) Else accountCount = 4

    Dim detailsLinks As Collection
    Dim detailItem As Variant
    Dim detailNum As Integer
    Dim childi As Variant, itm As Variant
    Dim skipLineCount As Integer
    Dim q As Integer, i As Integer, j As Integer, k As Integer

    With chrome
        For q = 1 To accountCount
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
            i = -1: j = 0
            For Each childi In .accMain.FindFirst(stdLambda.Create("$1.Description like ""Showing * entries"" and $1.Role = ""ROLE_TABLE""")).children
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
                                ActiveCell.offset(i, k - 1).value = "Hesap-" & detailNum
                            End If
                            On Error GoTo 0
                        ElseIf j = 4 Then
                            k = 2: ActiveCell.offset(i, k).value = itm.children.item(1).name
                        ElseIf j = 5 Then
                            k = 3: ActiveCell.offset(i, k).value = CDbl(itm.children.item(1).name)
                        End If
                    Next itm
                    j = 0
                End If
            Next childi
            Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Hesaplar"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
        Next q
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Anasayfa"" and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With
    LogManager.LogInfo "Hook_TEB_FetchAccounts: " & accountCount & " accounts fetched"
End Sub

' Opens date picker and selects a date N days in the past.
' p1 = days back (default 31)
Public Sub Hook_TEB_DatePicker(chrome As stdChrome, p1 As String, p2 As String, p3 As String)
    Dim daysBack As Integer
    If Len(p1) > 0 And IsNumeric(p1) Then daysBack = CInt(p1) Else daysBack = 31

    Dim dateStr As String
    dateStr = Replace(CStr(Date - daysBack), ".", "/")
    If Len(dateStr) = 9 Then dateStr = "0" & dateStr

    With chrome
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""..."" and $1.Role = ""ROLE_PUSHBUTTON""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select year"" and $1.Role = ""ROLE_COMBOBOX""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select year"" and $1.Role = ""ROLE_COMBOBOX""")). _
            AwaitForElement(stdLambda.Create("$1.Name = """ & Mid(dateStr, 7) & """ and $1.Role = ""ROLE_LISTITEM""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select month"" and $1.Role = ""ROLE_COMBOBOX""")).DoDefaultAction
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Select month"" and $1.Role = ""ROLE_COMBOBOX""")). _
            AwaitForElement(stdLambda.Create("$1.Name = """ & Hook_TEB_MonthName(CInt(Mid(dateStr, 4, 2))) & """ and $1.Role = ""ROLE_LISTITEM""")).DoDefaultAction
        Call .accMain.AwaitForElement(stdLambda.Create("$1.Name = """ & CStr(CInt(Mid(dateStr, 1, 2))) & """ and $1.Role = ""ROLE_LINK""")).DoDefaultAction
    End With
    LogManager.LogInfo "Hook_TEB_DatePicker: set to " & dateStr
End Sub

' Selects the "Cari Hesap" checkbox in the account filter panel.
Public Sub Hook_TEB_AccountFilter(chrome As stdChrome, p1 As String, p2 As String, p3 As String)
    With chrome
        Call .AwaitForAccElement(stdLambda.Create("$1.Name = ""Cari Hesap"" and $1.Role = ""ROLE_CELL"""))
        chrome.AwaitForAccElement(stdLambda.Create("$1.Name = ""Cari Hesap"" and $1.Role = ""ROLE_CELL""")) _
            .AwaitForElement(stdLambda.Create("$1.Name = """" and $1.Role = ""ROLE_CHECKBUTTON""")) _
            .parent.children.item(2).DoDefaultAction
    End With
    LogManager.LogInfo "Hook_TEB_AccountFilter: Cari Hesap selected"
End Sub

Private Function Hook_TEB_MonthName(monthNum As Integer) As String
    Hook_TEB_MonthName = Array("Ocak", ChrW(350) & "ubat", "Mart", "Nisan", "May" & ChrW(305) & "s", _
                               "Haziran", "Temmuz", "A" & ChrW(287) & "ustos", _
                               "Eyl" & ChrW(252) & "l", "Ekim", _
                               "Kas" & ChrW(305) & "m", "Aral" & ChrW(305) & "k")(monthNum - 1)
End Function


