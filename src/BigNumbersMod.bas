Attribute VB_Name = "BigNumbersMod"
Option Explicit

' -------------------------
' BigDecimal Division for VBA
' -------------------------
' Usage:
'   Debug.Print BigDiv("1,195.70", "3", 10) -> "398,5666666666..."
'   Debug.Print BigDiv("12345678901234567890", "3", 30) -> precise result
' -------------------------

Public Function BigDiv(numA As String, numB As String, Optional PRECISION As Long = 30) As String
    Dim a As String, b As String
    Dim decA As Long, decB As Long
    Dim shiftedA As String, shiftedB As String
    Dim shift As Long
    Dim quotient As String

    ' normalize inputs (remove spaces, quotes)
    a = Trim$(Replace$(numA, """", ""))
    b = Trim$(Replace$(numB, """", ""))
    a = Replace$(a, " ", "")
    b = Replace$(b, " ", "")
    
    If Len(a) = 0 Or Len(b) = 0 Then
        BigDiv = "NaN"
        Exit Function
    End If
    
    ' determine decimal places and remove thousands separators smartly
    a = NormalizeNumberForCalc(a, decA)
    b = NormalizeNumberForCalc(b, decB)
    
    If IsAllZeros(b) Then
        BigDiv = "NaN" ' divide by zero
        Exit Function
    End If
    
    ' We want an integer division:
    ' If shift = decB - decA
    ' if shift >=0 -> scaledA = A * 10^(PRECISION + shift), scaledB = B
    ' else -> scaledA = A * 10^PRECISION, scaledB = B * 10^(-shift)
    shift = decB - decA
    If shift >= 0 Then
        shiftedA = a & String$(PRECISION + shift, "0")
        shiftedB = b
    Else
        shiftedA = a & String$(PRECISION, "0")
        shiftedB = b & String$(-shift, "0")
    End If
    
    ' integer division of shiftedA by shiftedB
    quotient = DivBigInts(shiftedA, shiftedB)
    
    ' ensure quotient has at least PRECISION digits (pad left if necessary)
    If Len(quotient) <= PRECISION Then
        quotient = String$(PRECISION - Len(quotient) + 1, "0") & quotient
    End If
    
    ' insert decimal point before last PRECISION digits
    Dim intPart As String, fracPart As String
    intPart = left$(quotient, Len(quotient) - PRECISION)
    fracPart = right$(quotient, PRECISION)
    
    ' trim leading zeros on integer part
    intPart = TrimLeadingZeros(intPart)
    If Len(intPart) = 0 Then intPart = "0"
    
    ' trim trailing zeros on fraction
    fracPart = TrimTrailingZeros(fracPart)
    
    If Len(fracPart) = 0 Then
        BigDiv = intPart ' no fractional part
    Else
        ' use comma as decimal separator
        BigDiv = intPart & "," & fracPart
    End If
End Function

' -------------------------
' Helpers
' -------------------------

' NormalizeNumberForCalc:
' - removes quotes/spaces
' - decides decimal mark (last occurrence of '.' or ',')
' - removes thousands separators
' - returns pure digits string and sets decPlaces
Private Function NormalizeNumberForCalc(ByVal s As String, ByRef decPlaces As Long) As String
    Dim lastDot As Long, lastComma As Long
    s = Trim$(s)
    s = Replace$(s, """", "")
    s = Replace$(s, " ", "")
    
    lastDot = InStrRev(s, ".")
    lastComma = InStrRev(s, ",")
    
    ' If both present, the rightmost is decimal separator; remove the other as thousands
    If lastDot > 0 And lastComma > 0 Then
        If lastDot > lastComma Then
            ' dot is decimal
            s = Replace$(s, ",", "") ' remove commas (thousands)
            decPlaces = Len(s) - lastDot
            s = ReplaceOnceFromRight(s, ".", "") ' remove that dot but keep decimal info
            ' after removing dot we need pure digits -> will adjust below by splitting
            ' Actually easier: get integer and frac by splitting with Mid
            Dim intPart As String, fracPart As String
            intPart = left$(s, lastDot - 1)
            fracPart = Mid$(s, lastDot + 1)
            s = intPart & fracPart
            decPlaces = Len(fracPart)
            NormalizeNumberForCalc = KeepDigitsOnly(s)
            Exit Function
        Else
            ' comma is decimal
            s = Replace$(s, ".", "") ' remove dots (thousands)
            decPlaces = Len(s) - lastComma
            Dim ip As String, fp As String
            ip = left$(s, lastComma - 1)
            fp = Mid$(s, lastComma + 1)
            s = ip & fp
            decPlaces = Len(fp)
            NormalizeNumberForCalc = KeepDigitsOnly(s)
            Exit Function
        End If
    ElseIf lastDot > 0 Then
        ' only dot present
        ' treat dot as decimal (common)
        Dim idx As Long
        idx = lastDot
        Dim ipt As String, fpt As String
        ipt = left$(s, idx - 1)
        fpt = Mid$(s, idx + 1)
        s = Replace$(ipt, ",", "") & fpt ' remove any commas as thousands
        decPlaces = Len(fpt)
        NormalizeNumberForCalc = KeepDigitsOnly(s)
        Exit Function
    ElseIf lastComma > 0 Then
        ' only comma present
        Dim idx2 As Long
        idx2 = lastComma
        Dim ipt2 As String, fpt2 As String
        ipt2 = left$(s, idx2 - 1)
        fpt2 = Mid$(s, idx2 + 1)
        s = Replace$(ipt2, ".", "") & fpt2 ' remove dots as thousands
        decPlaces = Len(fpt2)
        NormalizeNumberForCalc = KeepDigitsOnly(s)
        Exit Function
    Else
        ' no dot or comma -> integer
        decPlaces = 0
        NormalizeNumberForCalc = KeepDigitsOnly(s)
        Exit Function
    End If
End Function

' Keep only digits 0-9
Private Function KeepDigitsOnly(ByVal s As String) As String
    Dim i As Long, ch As String, outS As String
    outS = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then outS = outS & ch
    Next i
    KeepDigitsOnly = TrimLeadingZeros(outS)
End Function

Private Function TrimLeadingZeros(ByVal s As String) As String
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> "0" Then
            TrimLeadingZeros = Mid$(s, i)
            Exit Function
        End If
    Next i
    TrimLeadingZeros = ""
End Function

Private Function TrimTrailingZeros(ByVal s As String) As String
    Dim i As Long
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) <> "0" Then
            TrimTrailingZeros = left$(s, i)
            Exit Function
        End If
    Next i
    TrimTrailingZeros = ""
End Function

Private Function IsAllZeros(ByVal s As String) As Boolean
    If Len(s) = 0 Then IsAllZeros = True: Exit Function
    Dim i As Long
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> "0" Then IsAllZeros = False: Exit Function
    Next i
    IsAllZeros = True
End Function

' -------------------------
' Big integer operations (strings)
' -------------------------

' Compare big integer strings (no sign, no leading zeros): returns 1 if a>b, 0 if equal, -1 if a<b
Private Function CompareBigInts(ByVal a As String, ByVal b As String) As Long
    a = TrimLeadingZeros(a)
    b = TrimLeadingZeros(b)
    If Len(a) > Len(b) Then CompareBigInts = 1: Exit Function
    If Len(a) < Len(b) Then CompareBigInts = -1: Exit Function
    If a > b Then
        CompareBigInts = 1
    ElseIf a < b Then
        CompareBigInts = -1
    Else
        CompareBigInts = 0
    End If
End Function

' Subtract b from a (a>=b). Returns string.
Private Function SubtractBigInts(ByVal a As String, ByVal b As String) As String
    Dim la As Long, lb As Long, i As Long, carry As Long
    Dim res() As Integer
    a = TrimLeadingZeros(a)
    b = TrimLeadingZeros(b)
    la = Len(a): lb = Len(b)
    ReDim res(1 To la)
    carry = 0
    For i = 0 To la - 1
        Dim da As Long, db As Long, idxA As Long, idxB As Long
        idxA = la - i
        idxB = lb - i
        da = CLng(Mid$(a, idxA, 1))
        If idxB >= 1 Then
            db = CLng(Mid$(b, idxB, 1))
        Else
            db = 0
        End If
        Dim subv As Long
        subv = da - db - carry
        If subv < 0 Then
            subv = subv + 10
            carry = 1
        Else
            carry = 0
        End If
        res(la - i) = subv
    Next i
    Dim outS As String
    outS = ""
    For i = 1 To la
        outS = outS & CStr(res(i))
    Next i
    SubtractBigInts = TrimLeadingZeros(outS)
    If SubtractBigInts = "" Then SubtractBigInts = "0"
End Function

' Multiply big integer by single digit 0..9
Private Function MultiplyBigIntDigit(ByVal a As String, ByVal d As Integer) As String
    If d = 0 Then MultiplyBigIntDigit = "0": Exit Function
    If d = 1 Then MultiplyBigIntDigit = TrimLeadingZeros(a): Exit Function
    Dim carry As Long, i As Long
    Dim la As Long
    la = Len(a)
    ReDim res(1 To la + 1)
    carry = 0
    For i = 0 To la - 1
        Dim da As Long
        da = CLng(Mid$(a, la - i, 1))
        Dim prod As Long
        prod = da * d + carry
        res(la + 1 - i) = prod Mod 10
        carry = prod \ 10
    Next i
    res(1) = carry
    Dim outS As String
    outS = ""
    For i = 1 To UBound(res)
        outS = outS & CStr(res(i))
    Next i
    MultiplyBigIntDigit = TrimLeadingZeros(outS)
End Function

' Multiply big integer by 10^k (i.e., append k zeros)
Private Function MultiplyBy10Power(ByVal a As String, ByVal k As Long) As String
    If a = "0" Then MultiplyBy10Power = "0": Exit Function
    MultiplyBy10Power = a & String$(k, "0")
End Function

' Divides big integer a by big integer b, returns quotient as string (integer division, floor)
' Uses long division with digit-by-digit determination via trial multiplication (0..9)
Private Function DivBigInts(ByVal a As String, ByVal b As String) As String
    a = TrimLeadingZeros(a)
    b = TrimLeadingZeros(b)
    If a = "" Then a = "0"
    If b = "" Then b = "0"
    If b = "0" Then DivBigInts = "NaN": Exit Function
    Dim n As Long, i As Long
    Dim remainder As String, cur As String
    remainder = ""
    Dim q As String
    q = ""
    For i = 1 To Len(a)
        remainder = remainder & Mid$(a, i, 1)
        remainder = TrimLeadingZeros(remainder)
        If remainder = "" Then remainder = "0"
        ' find digit
        Dim d As Integer
        d = 0
        ' binary search 0..9 might be overkill; simple loop 9->0 faster in VBA
        Dim prod As String
        For d = 9 To 0 Step -1
            prod = MultiplyBigIntDigit(b, d)
            If CompareBigInts(prod, remainder) <= 0 Then
                q = q & CStr(d)
                remainder = SubtractBigInts(remainder, prod)
                Exit For
            End If
        Next d
    Next i
    q = TrimLeadingZeros(q)
    If q = "" Then q = "0"
    DivBigInts = q
End Function

' Replace only the first occurrence from right (helper)
Private Function ReplaceOnceFromRight(ByVal s As String, ByVal find As String, ByVal repl As String) As String
    Dim p As Long
    p = InStrRev(s, find)
    If p = 0 Then
        ReplaceOnceFromRight = s
    Else
        ReplaceOnceFromRight = left$(s, p - 1) & repl & Mid$(s, p + Len(find))
    End If
End Function

' -------------------------
' Test routine
' -------------------------
Private Sub TestBigDiv()
    Debug.Print BigDiv("1,195.70", "3", 10)        ' -> 398,5666666666...
    Debug.Print BigDiv("1195.70", "1000", 8)      ' -> 1,19570? Actually -> "1,1957" after comma formatting
    Debug.Print BigDiv("12345678901234567890", "3", 40)
    Debug.Print BigDiv("1000", "7", 30)
End Sub





