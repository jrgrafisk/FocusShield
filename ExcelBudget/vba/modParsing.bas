Attribute VB_Name = "modParsing"
Option Explicit
' Tolerant parsing of the amounts and dates found in bank CSV exports.
' Port of budget_core/parsing.py. Nothing raises - unparsable input
' returns Null so the caller can count and report it.
'
' Requires "Microsoft VBScript Regular Expressions 5.5" behaviour, always
' available late-bound via CreateObject("VBScript.RegExp") on Windows.

Public Const MONTH_NAMES_DA As String = _
    "januar,februar,marts,april,maj,juni,juli,august,september,oktober,november,december"

' -----------------------------------------------------------------------
' Amounts
' -----------------------------------------------------------------------

Private Function NewRegExp(ByVal pattern As String, Optional ByVal ignoreCase As Boolean = True) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = ignoreCase
    re.Global = True
    Set NewRegExp = re
End Function

' Currency markers before or after the number ("1.234,56 DKK", "kr. 199").
Private Function StripCurrency(ByVal s As String) As String
    Dim re As Object
    Set re = NewRegExp("(dkk|sek|nok|isk|eur|usd|gbp|chf|pln|czk|huf|kr\.?|kroner)(?=$|[\s0-9.,+-])")
    s = re.Replace(s, " ")
    s = Replace(s, Chr(8364), "")   ' euro
    s = Replace(s, "$", "")
    s = Replace(s, Chr(163), "")    ' GBP
    s = Replace(s, Chr(165), "")    ' yen
    StripCurrency = s
End Function

' "1.234,56 DR" / "1.234,56 CR" - debit/credit suffix used by some systems.
Private Function StripDcSuffix(ByVal s As String, ByRef negative As Boolean) As String
    Dim re As Object, m As Object
    Set re = NewRegExp("\s*(dr|db|cr|kr)\.?$")
    If re.Test(s) Then
        Set m = re.Execute(s)
        Dim tag As String
        tag = LCase$(m(0).SubMatches(0))
        If tag = "dr" Or tag = "db" Then negative = True
        s = Left$(s, Len(s) - Len(m(0).Value))
    End If
    StripDcSuffix = s
End Function

Private Function IsNumberBody(ByVal s As String) As Boolean
    ' The right single quote is sometimes used as a thousands mark
    ' ("1?234,56"); written as a \u escape (never a literal character) so
    ' the source file's byte encoding can never turn it into mojibake.
    Dim re As Object
    Set re = NewRegExp("^[0-9]+([.,'\u2019][0-9]+)*$")
    IsNumberBody = re.Test(s)
End Function

' Returns the cleaned digits-and-separators body, or "" (check IsEmpty via
' the ok byref) plus whether the number is negative.
Private Function CleanNumberText(ByVal raw As Variant, ByRef ok As Boolean, ByRef negative As Boolean) As String
    ok = False
    negative = False
    If IsNull(raw) Or IsEmpty(raw) Then Exit Function
    Dim s As String
    If IsNumeric(raw) And Not (VarType(raw) = vbString) Then
        Dim v As Double
        v = CDbl(raw)
        negative = (v < 0)
        CleanNumberText = Replace(CStr(Abs(v)), ",", ".")
        ok = True
        Exit Function
    End If
    s = Trim$(CStr(raw))
    If s = "" Then Exit Function

    ' Normalise exotic spaces and dashes.
    s = Replace(s, Chr(160), " ")   ' nbsp
    s = Replace(s, Chr(8201), " ")  ' thin space
    s = Replace(s, Chr(8722), "-")  ' minus sign
    s = Replace(s, Chr(8211), "-")  ' en dash
    s = Replace(s, Chr(8212), "-")  ' em dash

    s = StripCurrency(s)
    s = StripDcSuffix(s, negative)

    s = Trim$(s)
    If Len(s) >= 2 And Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        negative = Not negative
        s = Trim$(Mid$(s, 2, Len(s) - 2))
    End If

    If Left$(s, 1) = "-" Then
        negative = Not negative
        s = Mid$(s, 2)
    ElseIf Left$(s, 1) = "+" Then
        s = Mid$(s, 2)
    End If
    s = Trim$(s)
    If Right$(s, 1) = "-" Then
        negative = Not negative
        s = Left$(s, Len(s) - 1)
    ElseIf Right$(s, 1) = "+" Then
        s = Left$(s, Len(s) - 1)
    End If

    s = Replace(s, " ", "")
    s = Replace(s, "'", "")
    s = Replace(s, Chr(8217), "")   ' right single quote

    If s = "" Or Not IsNumberBody(s) Then Exit Function
    CleanNumberText = s
    ok = True
End Function

Private Function IsGrouped(ByVal body As String, ByVal sep As String) As Boolean
    Dim re As Object
    Set re = NewRegExp("^\d{1,3}(?:" & EscapeRegex(sep) & "\d{3})+$")
    IsGrouped = re.Test(body)
End Function

Private Function EscapeRegex(ByVal s As String) As String
    Dim chars As Variant, i As Long
    chars = Array(".", "\", "+", "*", "?", "[", "^", "]", "$", "(", ")", "{", "}", "=", "!", "<", ">", "|", ":", "-", "#")
    For i = LBound(chars) To UBound(chars)
        s = Replace(s, chars(i), "\" & chars(i))
    Next i
    EscapeRegex = s
End Function

' Guess the decimal separator of a single cleaned number body ("," or "."
' or "" when there is nothing to decide).
Private Function ValueDecimalSeparator(ByVal body As String) As String
    Dim hasDot As Boolean, hasComma As Boolean
    hasDot = (InStr(body, ".") > 0)
    hasComma = (InStr(body, ",") > 0)
    If hasDot And hasComma Then
        If InStrRev(body, ".") > InStrRev(body, ",") Then
            ValueDecimalSeparator = "."
        Else
            ValueDecimalSeparator = ","
        End If
        Exit Function
    End If
    Dim sep As String, other As String
    Dim pairs(1, 1) As String
    pairs(0, 0) = ",": pairs(0, 1) = "."
    pairs(1, 0) = ".": pairs(1, 1) = ","
    Dim i As Long
    For i = 0 To 1
        sep = pairs(i, 0): other = pairs(i, 1)
        If InStr(body, sep) = 0 Then GoTo ContinueLoop
        If CountChar(body, sep) > 1 Then
            ValueDecimalSeparator = other
            Exit Function
        End If
        Dim tail As String
        tail = Mid$(body, InStrRev(body, sep) + 1)
        If Len(tail) = 3 And IsGrouped(body, sep) Then
            ValueDecimalSeparator = other
        Else
            ValueDecimalSeparator = sep
        End If
        Exit Function
ContinueLoop:
    Next i
    ValueDecimalSeparator = ""
End Function

Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    CountChar = Len(s) - Len(Replace(s, ch, ""))
End Function

Private Function AssembleNumber(ByVal body As String, ByVal decimalSep As String) As String
    Dim grouping As String
    grouping = IIf(decimalSep = ".", ",", ".")
    body = Replace(body, grouping, "")
    Dim parts() As String
    parts = Split(body, decimalSep)
    If UBound(parts) = 0 Then
        AssembleNumber = parts(0)
        Exit Function
    End If
    If UBound(parts) > 1 Then
        Dim last As String
        last = parts(UBound(parts))
        If Len(last) = 1 Or Len(last) = 2 Then
            Dim joined As String, i As Long
            joined = ""
            For i = 0 To UBound(parts) - 1
                joined = joined & parts(i)
            Next i
            AssembleNumber = joined & "." & last
        Else
            AssembleNumber = Join(parts, "")
        End If
        Exit Function
    End If
    AssembleNumber = parts(0) & "." & parts(1)
End Function

' Parse raw into a Double, or return Null.  decimalSep forces "," or "."; a
' blank string lets the value decide for itself.
Public Function ParseAmount(ByVal raw As Variant, Optional ByVal decimalSep As String = "") As Variant
    Dim ok As Boolean, negative As Boolean, body As String
    body = CleanNumberText(raw, ok, negative)
    If Not ok Then
        ParseAmount = Null
        Exit Function
    End If
    Dim sep As String, txt As String
    sep = decimalSep
    If sep = "" Then sep = ValueDecimalSeparator(body)
    If sep = "" Then
        txt = body
    Else
        txt = AssembleNumber(body, sep)
    End If
    If Not IsNumeric(txt) Then
        ParseAmount = Null
        Exit Function
    End If
    Dim value As Double
    value = Val(txt)   ' Val() always uses "." regardless of locale.
    If negative Then value = -value
    ParseAmount = value
End Function

Public Function LooksLikeAmount(ByVal raw As Variant) As Boolean
    Dim ok As Boolean, negative As Boolean
    CleanNumberText raw, ok, negative
    LooksLikeAmount = ok
End Function

' Decide whether a column uses "," or "." as decimal separator, or ""
' when undecided (caller should default to ",").
Public Function DetectDecimalSeparator(ByVal samples As Collection) As String
    Dim strongComma As Long, strongDot As Long, weakComma As Long, weakDot As Long
    Dim raw As Variant, ok As Boolean, negative As Boolean, body As String
    For Each raw In samples
        body = CleanNumberText(raw, ok, negative)
        If Not ok Or body = "" Then GoTo ContinueLoop
        Dim hasDot As Boolean, hasComma As Boolean
        hasDot = (InStr(body, ".") > 0)
        hasComma = (InStr(body, ",") > 0)
        If hasDot And hasComma Then
            If InStrRev(body, ".") > InStrRev(body, ",") Then
                strongDot = strongDot + 2
            Else
                strongComma = strongComma + 2
            End If
            GoTo ContinueLoop
        End If
        Dim pairs(1, 1) As String
        pairs(0, 0) = ",": pairs(0, 1) = "."
        pairs(1, 0) = ".": pairs(1, 1) = ","
        Dim i As Long, sep As String, other As String
        For i = 0 To 1
            sep = pairs(i, 0): other = pairs(i, 1)
            If InStr(body, sep) = 0 Then GoTo NextPair
            If CountChar(body, sep) > 1 Then
                If other = "," Then strongComma = strongComma + 1 Else strongDot = strongDot + 1
            Else
                Dim tail As String
                tail = Mid$(body, InStrRev(body, sep) + 1)
                If Len(tail) = 1 Or Len(tail) = 2 Or Len(tail) > 3 Then
                    If sep = "," Then strongComma = strongComma + 1 Else strongDot = strongDot + 1
                ElseIf IsGrouped(body, sep) Then
                    If other = "," Then weakComma = weakComma + 1 Else weakDot = weakDot + 1
                Else
                    If sep = "," Then strongComma = strongComma + 1 Else strongDot = strongDot + 1
                End If
            End If
            Exit For
NextPair:
        Next i
ContinueLoop:
    Next raw
    If strongComma <> strongDot Then
        DetectDecimalSeparator = IIf(strongComma > strongDot, ",", ".")
        Exit Function
    End If
    If weakComma <> weakDot Then
        DetectDecimalSeparator = IIf(weakComma > weakDot, ",", ".")
        Exit Function
    End If
    If strongComma > 0 Or weakComma > 0 Then
        DetectDecimalSeparator = ","
        Exit Function
    End If
    DetectDecimalSeparator = ""
End Function

' -----------------------------------------------------------------------
' Dates
' -----------------------------------------------------------------------

Private Function MonthNumberFromName(ByVal name As String) As Long
    name = FoldText(name)
    Dim names As Variant, i As Long
    names = Split(MONTH_NAMES_DA, ",")
    For i = 0 To UBound(names)
        If name = names(i) Or name = Left$(names(i), 3) Then
            MonthNumberFromName = i + 1
            Exit Function
        End If
    Next i
    Dim enNames As Variant
    enNames = Array("january", "february", "march", "april", "may", "june", _
                    "july", "august", "september", "october", "november", "december")
    For i = 0 To UBound(enNames)
        If name = enNames(i) Or name = Left$(enNames(i), 3) Then
            MonthNumberFromName = i + 1
            Exit Function
        End If
    Next i
    Select Case name
        Case "des": MonthNumberFromName = 12
        Case "sept": MonthNumberFromName = 9
        Case "okt", "oct": MonthNumberFromName = 10
        Case Else: MonthNumberFromName = 0
    End Select
End Function

Private Function MakeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long) As Variant
    If y < 100 Then y = y + IIf(y < 70, 2000, 1900)
    If y < 1900 Or y > 2200 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then
        MakeDate = Null
        Exit Function
    End If
    On Error GoTo Fail
    MakeDate = DateSerial(y, m, d)
    Exit Function
Fail:
    MakeDate = Null
End Function

' Parse a date written in (almost) any common layout: "2024-05-31",
' "31-05-2024", "31/05/24", "31.05.2024", "20240531", "31052024",
' "31. maj 2024", "May 31, 2024", any of those with a trailing clock time.
Public Function ParseDate(ByVal raw As Variant, Optional ByVal dayFirst As Boolean = True) As Variant
    If IsNull(raw) Or IsEmpty(raw) Then
        ParseDate = Null
        Exit Function
    End If
    If VarType(raw) = vbDate Then
        ParseDate = CDate(raw)
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(raw))
    If s = "" Then
        ParseDate = Null
        Exit Function
    End If

    Dim compact As String
    compact = Replace(s, " ", "")
    If Len(compact) = 8 And IsNumeric(compact) Then
        Dim r As Variant
        r = MakeDate(CLng(Left$(compact, 4)), CLng(Mid$(compact, 5, 2)), CLng(Mid$(compact, 7, 2)))
        If Not IsNull(r) Then
            ParseDate = r
            Exit Function
        End If
        r = MakeDate(CLng(Right$(compact, 4)), CLng(Mid$(compact, 3, 2)), CLng(Left$(compact, 2)))
        If Not IsNull(r) Then
            ParseDate = r
            Exit Function
        End If
        r = MakeDate(CLng(Right$(compact, 4)), CLng(Left$(compact, 2)), CLng(Mid$(compact, 3, 2)))
        If Not IsNull(r) Then
            ParseDate = r
            Exit Function
        End If
    End If

    Dim re As Object, m As Object, tok As Object
    Set re = NewRegExp("[0-9]+|[a-zA-Z\u00E6\u00F8\u00E5]+")
    Set m = re.Execute(s)
    If m.Count < 3 Then
        ParseDate = Null
        Exit Function
    End If

    Dim monthFromName As Long
    monthFromName = 0
    Dim numbers As New Collection
    Dim t As String, mn As Long
    Dim idx As Long
    For idx = 0 To m.Count - 1
        t = m(idx).Value
        If IsNumeric(t) Then
            If Len(t) > 4 Then
                ParseDate = Null
                Exit Function
            End If
            numbers.Add CLng(t)
        Else
            mn = MonthNumberFromName(t)
            If mn > 0 And monthFromName = 0 Then
                monthFromName = mn
            ElseIf numbers.Count > 0 Then
                Exit For
            End If
        End If
        If monthFromName > 0 And numbers.Count >= 2 Then Exit For
        If monthFromName = 0 And numbers.Count >= 3 Then Exit For
    Next idx

    If monthFromName > 0 Then
        If numbers.Count < 2 Then
            ParseDate = Null
            Exit Function
        End If
        Dim a As Long, b As Long, yr As Long, dy As Long
        a = numbers(1): b = numbers(2)
        If a > 31 Then
            yr = a: dy = b
        Else
            yr = b: dy = a
        End If
        ParseDate = MakeDate(yr, monthFromName, dy)
        Exit Function
    End If

    If numbers.Count < 3 Then
        ParseDate = Null
        Exit Function
    End If
    Dim n1 As Long, n2 As Long, n3 As Long
    n1 = numbers(1): n2 = numbers(2): n3 = numbers(3)
    If n1 > 31 Then
        ParseDate = MakeDate(n1, n2, n3)
        Exit Function
    End If
    If n1 > 12 Then
        ParseDate = MakeDate(n3, n2, n1)
        Exit Function
    End If
    If n2 > 12 Then
        ParseDate = MakeDate(n3, n1, n2)
        Exit Function
    End If
    Dim result As Variant
    If dayFirst Then
        result = MakeDate(n3, n2, n1)
        If IsNull(result) Then result = MakeDate(n3, n1, n2)
    Else
        result = MakeDate(n3, n1, n2)
        If IsNull(result) Then result = MakeDate(n3, n2, n1)
    End If
    ParseDate = result
End Function

Public Function LooksLikeDate(ByVal raw As Variant) As Boolean
    LooksLikeDate = Not IsNull(ParseDate(raw))
End Function

' Decide between 31/05/2024 (day first) and 05/31/2024 (month first).
Public Function DetectDayfirst(ByVal samples As Collection) As Boolean
    Dim dayVotes As Long, monthVotes As Long
    Dim raw As Variant, s As String
    Dim re As Object, m As Object
    Set re = NewRegExp("[0-9]+")
    For Each raw In samples
        If IsNull(raw) Then GoTo ContinueLoop
        s = Trim$(CStr(raw))
        If s = "" Then GoTo ContinueLoop
        If Len(Replace(s, " ", "")) = 8 And IsNumeric(Replace(s, " ", "")) Then GoTo ContinueLoop
        Set m = re.Execute(s)
        Dim numbers As New Collection
        Dim i As Long
        For i = 0 To m.Count - 1
            If Len(m(i).Value) <= 4 Then numbers.Add CLng(m(i).Value)
        Next i
        If numbers.Count < 3 Then GoTo ContinueLoop
        Dim a As Long, b As Long
        a = numbers(1): b = numbers(2)
        If a > 31 Then GoTo ContinueLoop
        If a > 12 Then
            dayVotes = dayVotes + 1
        ElseIf b > 12 Then
            monthVotes = monthVotes + 1
        End If
ContinueLoop:
    Next raw
    DetectDayfirst = Not (monthVotes > dayVotes)
End Function

Public Function MonthKeyOf(ByVal d As Date) As String
    MonthKeyOf = Format(Year(d), "0000") & "-" & Format(Month(d), "00")
End Function

Public Function MonthLabelOf(ByVal key As String) As String
    On Error GoTo Fail
    Dim parts() As String
    parts = Split(key, "-")
    Dim names As Variant
    names = Split(MONTH_NAMES_DA, ",")
    MonthLabelOf = names(CLng(parts(1)) - 1) & " " & parts(0)
    Exit Function
Fail:
    MonthLabelOf = key
End Function
