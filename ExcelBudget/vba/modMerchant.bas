Attribute VB_Name = "modMerchant"
Option Explicit
' Find the recognisable shop or company inside a transaction text.
' Port of budget_core/merchant.py.

Public Const RULE_WORDS As Long = 3
Public Const RULE_CHARS As Long = 32
Public Const MIN_NAME_LENGTH As Long = 3

Private mNoiseWords As Object   ' Scripting.Dictionary, folded word -> True

Private Sub EnsureNoiseWords()
    If Not mNoiseWords Is Nothing Then Exit Sub
    Set mNoiseWords = CreateObject("Scripting.Dictionary")
    Dim words As Variant, i As Long
    words = Array( _
        "dankort", "nota", "dankortnota", "kortkob", "kortkoeb", "kort", "kortnr", _
        "visa", "mastercard", "maestro", "eurocard", "kreditkort", "betalingskort", _
        "debitkort", "chipkort", "haevekort", "kontokort", "korttype", _
        "kortbetaling", "korttransaktion", "kortnota", "kortholder", _
        "kortoplysninger", "kortkobsnota", _
        "kreditcard", "debitcard", "betalingscard", "chipcard", "bankcard", _
        "kontocard", "cardtype", "cardbetaling", "cardtransaktion", "cardnota", _
        "cardholder", "cardoplysninger", "cardkobsnota", "cardkob", "cardkoeb", _
        "cardnr", "card", _
        "americanexpress", "amex", "dinersclub", "diners", "jcb", "unionpay", _
        "bs", "pbs", "betaling", "betalingsservice", "indbetalingskort", _
        "ref", "refnr", "reference", "referencenr", "id", "idnr", "nr", "no", _
        "kl", "den", "d", "dato", "tid", "kvittering", "bilag", "faktura", "fakt", _
        "kob", "koeb", "purchase", "payment", "pos", "atm", "automat", "netbank", _
        "mobilbank", "onlinebank", "udland", "udl", "valutakurs", "kurs", _
        "dk", "dkk", "sek", "nok", "eur", "usd", "gbp", "kr", "kroner", _
        "fra", "til", "med", "m", "v", "og", "af", "pa", "konto", "kontonr", _
        "konto-nr", _
        "transaktion", "postering", "posteringstekst", "tekst", "note", _
        "aut", "auto", "automatisk", "straks", "straksoverforsel")
    For i = LBound(words) To UBound(words)
        mNoiseWords(words(i)) = True
    Next i
End Sub

Private Function IsDateOrTimeToken(ByVal token As String) As Boolean
    Static reDate As Object, reTime As Object
    If reDate Is Nothing Then
        Set reDate = CreateObject("VBScript.RegExp")
        reDate.pattern = "^\d{1,4}[./-]\d{1,2}([./-]\d{2,4})?\.?$"
        Set reTime = CreateObject("VBScript.RegExp")
        reTime.pattern = "^\d{1,2}[:.]\d{2}$"
    End If
    IsDateOrTimeToken = reDate.Test(token) Or reTime.Test(token)
End Function

Private Function HasLetter(ByVal token As String) As Boolean
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.pattern = "[^\W\d_]"
    End If
    HasLetter = re.Test(token)
End Function

Private Function TrimToken(ByVal token As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.pattern = "^[^\w&]+|[^\w&%]+$"
    End If
    Dim result As String
    result = re.Replace(token, "")
    If result = "" Then result = token
    TrimToken = result
End Function

Private Function DigitRatioAtLeastHalf(ByVal token As String) As Boolean
    Dim i As Long, digits As Long
    For i = 1 To Len(token)
        If Mid$(token, i, 1) Like "#" Then digits = digits + 1
    Next i
    DigitRatioAtLeastHalf = (digits > 0 And digits * 2 >= Len(token))
End Function

Private Function IsNoise(ByVal token As String) As Boolean
    EnsureNoiseWords
    If Not HasLetter(token) Then
        IsNoise = True             ' pure numbers, "//", "-", "***"
        Exit Function
    End If
    If IsDateOrTimeToken(token) Then
        IsNoise = True
        Exit Function
    End If
    If LooksLikeAmount(token) Then
        IsNoise = True
        Exit Function
    End If
    Dim folded As String
    folded = FoldText(token)
    If mNoiseWords.Exists(folded) Then
        IsNoise = True
        Exit Function
    End If
    ' "Dankort-nota", "kortkob/visa": noise when every part is noise.
    Dim parts() As String, i As Long, allNoise As Boolean
    parts = SplitAny(folded, "-/.")
    If UBound(parts) >= 1 Then
        allNoise = True
        For i = LBound(parts) To UBound(parts)
            If parts(i) <> "" And Not mNoiseWords.Exists(parts(i)) Then
                allNoise = False
                Exit For
            End If
        Next i
        If allNoise Then
            IsNoise = True
            Exit Function
        End If
    End If
    If DigitRatioAtLeastHalf(token) Then
        IsNoise = True
        Exit Function
    End If
    IsNoise = False
End Function

Private Function SplitAny(ByVal s As String, ByVal seps As String) As String()
    Dim i As Long, ch As String
    For i = 1 To Len(seps)
        ch = Mid$(seps, i, 1)
        s = Replace(s, ch, Chr(1))
    Next i
    SplitAny = Split(s, Chr(1))
End Function

Private Function IsConnector(ByVal token As String) As Boolean
    ' "&", "-", "/" and friends: part of a name, not a word of their own.
    If Len(token) > 2 Then
        IsConnector = False
        Exit Function
    End If
    If HasLetter(token) Then
        IsConnector = False
        Exit Function
    End If
    Dim i As Long
    For i = 1 To Len(token)
        If Mid$(token, i, 1) Like "#" Then
            IsConnector = False
            Exit Function
        End If
    Next i
    IsConnector = True
End Function

' Tokens, a parallel "keep" flag array, and a parallel "connector" flag
' array, all 0-based, matching the words of text split on whitespace.
Private Sub Analyse(ByVal text As String, ByRef tokens() As String, _
                    ByRef keep() As Boolean, ByRef connectors() As Boolean)
    Dim raw() As String
    raw = Split(SqueezeText(text), " ")
    Dim n As Long, i As Long, j As Long
    n = 0
    For i = LBound(raw) To UBound(raw)
        If raw(i) <> "" Then n = n + 1
    Next i
    ReDim tokens(0 To IIf(n > 0, n - 1, 0))
    ReDim keep(0 To IIf(n > 0, n - 1, 0))
    ReDim connectors(0 To IIf(n > 0, n - 1, 0))
    j = 0
    For i = LBound(raw) To UBound(raw)
        If raw(i) <> "" Then
            tokens(j) = raw(i)
            Dim trimmed As String, connector As Boolean
            trimmed = TrimToken(raw(i))
            connector = IsConnector(raw(i))
            connectors(j) = connector
            keep(j) = (Not connector) And (Not IsNoise(trimmed))
            j = j + 1
        End If
    Next i
    If n = 0 Then
        ReDim tokens(0 To -1)
        ReDim keep(0 To -1)
        ReDim connectors(0 To -1)
    End If
End Sub

Private Function CleanTokensJoined(ByVal text As String) As String
    Dim tokens() As String, keep() As Boolean, connectors() As Boolean
    Analyse text, tokens, keep, connectors
    Dim out As String, i As Long
    out = ""
    If UBound(tokens) < LBound(tokens) Then
        CleanTokensJoined = ""
        Exit Function
    End If
    For i = LBound(tokens) To UBound(tokens)
        If keep(i) Then
            Dim trimmed As String
            trimmed = TrimToken(tokens(i))
            If trimmed = "" Then trimmed = tokens(i)
            If out <> "" Then out = out & " "
            out = out & trimmed
        End If
    Next i
    CleanTokensJoined = out
End Function

' The shop/company part of a transaction text (original spelling).
Public Function MerchantName(ByVal text As Variant) As String
    Dim s As String
    s = CStr(text)
    Dim name As String
    name = Trim$(CleanTokensJoined(s))
    If Len(name) < MIN_NAME_LENGTH Then
        MerchantName = SqueezeText(s)
    Else
        MerchantName = name
    End If
End Function

' A grouping key: two texts from the same shop give the same key.
Public Function MerchantKey(ByVal text As Variant) As String
    Dim k As String
    k = FoldText(MerchantName(text))
    If k = "" Then k = FoldText(CStr(text))
    MerchantKey = k
End Function

' A keyword to save as a rule - short enough to match next month too, and
' guaranteed to be an unbroken substring of the original text.
Public Function RuleKeyword(ByVal text As Variant) As String
    Dim s As String
    s = CStr(text)
    Dim tokens() As String, keep() As Boolean, connectors() As Boolean
    Analyse s, tokens, keep, connectors
    If UBound(tokens) < LBound(tokens) Then
        RuleKeyword = Left$(SqueezeText(s), RULE_CHARS)
        Exit Function
    End If
    Dim anyKeep As Boolean, i As Long
    anyKeep = False
    For i = LBound(keep) To UBound(keep)
        If keep(i) Then anyKeep = True: Exit For
    Next i
    If Not anyKeep Then
        RuleKeyword = Left$(SqueezeText(s), RULE_CHARS)
        Exit Function
    End If

    Dim first As Long
    first = -1
    For i = LBound(keep) To UBound(keep)
        If keep(i) Then first = i: Exit For
    Next i
    Dim last As Long, words As Long
    last = first
    words = 0
    For i = first To UBound(tokens)
        If connectors(i) Then GoTo ContinueLoop     ' "&" does not end the name
        If Not keep(i) Then Exit For                ' a shop number ends it
        words = words + 1
        If words > RULE_WORDS Then Exit For
        last = i
ContinueLoop:
    Next i

    Dim joined As String
    joined = ""
    For i = first To last
        If joined <> "" Then joined = joined & " "
        joined = joined & tokens(i)
    Next i
    Dim keyword As String
    keyword = TrimToken(joined)
    If keyword = joined And joined = "" Then keyword = ""
    ' TrimToken() falls back to the input when the result is empty, which
    ' is wrong here (an all-punctuation joined string should stay empty).
    Dim reCheck As Object
    Set reCheck = CreateObject("VBScript.RegExp")
    reCheck.Global = True
    reCheck.pattern = "^[^\w&]+|[^\w&%]+$"
    keyword = reCheck.Replace(joined, "")

    If Len(keyword) > RULE_CHARS Then
        keyword = RTrim$(Left$(keyword, RULE_CHARS))
    End If
    If keyword = "" Then keyword = Left$(SqueezeText(s), RULE_CHARS)
    RuleKeyword = keyword
End Function
