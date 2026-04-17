Attribute VB_Name = "Rules"
Option Explicit
Private cachedRules As Object

Public Sub GetRules()
    Dim rowNum As Long
    Dim rulesDict As scripting.Dictionary
    Set rulesDict = New scripting.Dictionary
    Dim aRule As RuleObj
    rowNum = 2
    With RULESS
        Do While .Cells(rowNum, 1) <> ""
            Set aRule = New RuleObj
            aRule.Active = .Cells(rowNum, 1).value
            aRule.DescRuleType = .Cells(rowNum, 2).value
            aRule.Description = .Cells(rowNum, 3).value
            aRule.AmountOp = .Cells(rowNum, 4).value
            aRule.Amount = .Cells(rowNum, 5).value
            aRule.account = .Cells(rowNum, 6).value
            aRule.ToAccount = .Cells(rowNum, 7).value
            aRule.NewDescription = .Cells(rowNum, 8).value
            aRule.Special = .Cells(rowNum, 9).value
            aRule.Priority = .Cells(rowNum, 10).value
            rulesDict.Add rowNum - 1, aRule
            rowNum = rowNum + 1
        Loop
    End With
    Set cachedRules = rulesDict
End Sub

Function CheckRules(ByVal Description As String, ByVal Amount As Double, account As String) As scripting.Dictionary
    
    Call GetRules
    Dim rulesDict As scripting.Dictionary
    Set rulesDict = cachedRules
    
    Dim matches As New Collection
    Dim k As Variant
    Dim aRule As RuleObj
    Dim bestRule As RuleObj
    Dim bestPriority As Long
    bestPriority = -2147483648# ' Long min
    
    For Each k In rulesDict.keys
        Set aRule = rulesDict(k)
        
            ' --- active match
        If Not aRule.Active Then GoTo NextRule
            
            ' --- description match ---
        Select Case UCase(aRule.DescRuleType)
            Case "CONTAINS"
                If InStr(1, Description, aRule.Description, vbTextCompare) = 0 Then GoTo NextRule
            Case "EXACT"
                If StrComp(Description, aRule.Description, vbTextCompare) <> 0 Then GoTo NextRule
            Case "REGEX"
                aRule.NewDescription = RegxReplacer(Description, aRule.Description, aRule.NewDescription)
                If aRule.NewDescription = "FALSE" Then GoTo NextRule
            Case Else
                'nothing? no text match
        End Select
        
           ' --- account match ---
        If aRule.account <> "" And aRule.account <> account Then
            GoTo NextRule
        Else
        End If
               
            ' --- amount match ---
        Select Case aRule.AmountOp
            Case "="
                If Amount <> aRule.Amount Then GoTo NextRule
            Case ">="
                If Amount < aRule.Amount Then GoTo NextRule
            Case ">"
                If Amount <= aRule.Amount Then GoTo NextRule
            Case "<="
                If Amount > aRule.Amount Then GoTo NextRule
            Case "<"
                If Amount >= aRule.Amount Then GoTo NextRule
            Case Else
                'nothing? match
        End Select
        
        ' --- priority check ---
        If aRule.Priority > bestPriority Then
            bestPriority = aRule.Priority
            Set bestRule = aRule
        End If

NextRule:
    Next k
    
    ' --- result ---
    Dim result As scripting.Dictionary
    Set result = New scripting.Dictionary
    
    If Not bestRule Is Nothing Then
        result.Add "toAccount", bestRule.ToAccount
        result.Add "special", bestRule.Special
        result.Add "newDescription", bestRule.NewDescription
    End If
    
    Set CheckRules = result
End Function


Sub DetectEmpty()
 
    Dim empties As Variant
    Set empties = modFindAll64.FindAll(ActiveSheet.Cells(1, 8).Resize(350, 1), "", xlValues, xlWhole)
    Dim wb As Workbook
    Set wb = Workbooks.Add
    wb.ActiveSheet.Range("A1").Select
    Dim cll As Variant
    For Each cll In empties.Cells
        Selection.value = cll.offset(-1, -5).value
        Selection.offset(0, 1).value = cll.offset(-1, 1).value
        Selection.offset(1, 0).Select
    Next cll
    
End Sub

Private Function RegxReplacer(sourcestr As String, regPattern As String, Optional replaceStr As String = "") As String

    RegxReplacer = "FALSE"
    Dim regx As VBA.RegExp
    Set regx = New RegExp
    If regPattern = "" Then Exit Function
    regx.Pattern = regPattern
    regx.ignoreCase = False
    
    If replaceStr <> "" Then
        Dim replaceReg As VBA.RegExp
        Set replaceReg = New RegExp
        replaceReg.Pattern = "$\d+"
        Dim ExpectedSubMatchCount As Long
        If replaceReg.test(replaceStr) = True Then ExpectedSubMatchCount = replaceReg.Execute(replaceStr)(0).SubMatches.Count
        Dim regResult As Variant
        Set regResult = regx.Execute(sourcestr) 'burada hata verirse regex patternde bir sorun var demek
        If regx.test(sourcestr) = True Then
            If regResult(0).SubMatches.Count = ExpectedSubMatchCount Or ExpectedSubMatchCount = 0 Then
                If InStr(1, replaceStr, "$") = 0 Then
                    RegxReplacer = replaceStr
                Else
                    RegxReplacer = regx.Replace(sourcestr, replaceStr)
                End If
            Else
                Debug.Print "Submatch count hatası var regexi kontrol edin"
            End If
        Else
        End If
    Else
        If regx.test(sourcestr) = True Then
            RegxReplacer = regx.Execute(sourcestr)(0).value
        Else
            '??
        End If
    End If
    
End Function





