Attribute VB_Name = "modColumns"
Option Explicit
' Figure out which column holds the date, the text and the amount.
' Port of budget_core/columns.py.

Private Function WA(ParamArray items() As Variant) As Variant
    WA = items
End Function

Private mDateWords As Variant, mDateWordsPrimary As Variant
Private mTextWords As Variant, mAmountWords As Variant
Private mInWords As Variant, mOutWords As Variant
Private mBalanceWords As Variant, mCurrencyWords As Variant, mIgnoreWords As Variant

Private Sub EnsureWordLists()
    If Not IsEmpty(mDateWords) Then Exit Sub
    mDateWords = WA("dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato", _
                   "transaktionsdato", "valor", "valordato", "posting", "booking", "fecha")
    mDateWordsPrimary = WA("dato", "date", "bogfort", "posteringsdato", "transaktionsdato", _
                          "posting", "booking")
    mTextWords = WA("tekst", "text", "beskrivelse", "description", "posteringstekst", _
                   "narrative", "details", "detaljer", "modtager", "afsender", "payee", _
                   "merchant", "reference", "memo", "note", "navn", "name", "titel", _
                   "forklaring", "kommentar", "bemaerkning")
    mAmountWords = WA("belob", "beloeb", "amount", "sum", "betrag", "montant", "value", _
                     "transaktionsbelob", "bevaegelse", "posteringsbelob")
    mInWords = WA("indsat", "indbetaling", "indbetalt", "ind", "kredit", "credit", _
                 "deposit", "modtaget", "tilgang", "indgaende", "income", "haevet ind")
    mOutWords = WA("haevet", "udbetaling", "udbetalt", "ud", "debet", "debit", _
                  "withdrawal", "betalt", "afgang", "udgaende", "expense", "traek")
    mBalanceWords = WA("saldo", "balance", "beholdning", "kontosaldo", "running balance", _
                      "ny saldo")
    mCurrencyWords = WA("valuta", "currency", "mont", "vaeluta")
    mIgnoreWords = WA("kontonummer", "account number", "kortnummer", "id", "reg nr", _
                     "registreringsnummer", "afstemt", "status")
End Sub

Private Function HasWord(ByVal header As String, ByVal words As Variant) As Boolean
    Dim folded As String
    folded = FoldText(header)
    If folded = "" Then
        HasWord = False
        Exit Function
    End If
    Dim i As Long
    For i = LBound(words) To UBound(words)
        If folded = words(i) Then
            HasWord = True
            Exit Function
        End If
    Next i
    Dim cleaned As String
    cleaned = Replace(Replace(Replace(folded, "/", " "), "-", " "), ".", " ")
    Dim tokens() As String, t As Long
    tokens = Split(cleaned, " ")
    For t = LBound(tokens) To UBound(tokens)
        If tokens(t) <> "" Then
            For i = LBound(words) To UBound(words)
                If tokens(t) = words(i) Then
                    HasWord = True
                    Exit Function
                End If
            Next i
        End If
    Next t
    For i = LBound(words) To UBound(words)
        If Len(CStr(words(i))) > 4 And InStr(folded, words(i)) > 0 Then
            HasWord = True
            Exit Function
        End If
    Next i
    HasWord = False
End Function

Public Function AnalyseColumns(ByVal t As clsTable, Optional ByVal limit As Long = 300) As Collection
    Dim out As New Collection, i As Long
    For i = 0 To t.nColumns - 1
        Dim st As New clsColumnStats
        st.Init i, t.HeaderName(i), t.ColumnValues(i, limit)
        out.Add st
    Next i
    Set AnalyseColumns = out
End Function

Private Function PickDate(ByVal stats As Collection) As Long
    EnsureWordLists
    Dim candidates As New Collection, s As clsColumnStats
    For Each s In stats
        If s.IsDateish() Then candidates.Add s
    Next s
    If candidates.Count = 0 Then
        For Each s In stats
            If s.DateRatio >= 0.3 Then candidates.Add s
        Next s
    End If
    If candidates.Count = 0 Then
        PickDate = -1
        Exit Function
    End If

    For Each s In candidates
        If HasWord(s.header, mDateWordsPrimary) Then
            PickDate = s.Index
            Exit Function
        End If
    Next s
    For Each s In candidates
        If HasWord(s.header, mDateWords) Then
            PickDate = s.Index
            Exit Function
        End If
    Next s

    Dim best As clsColumnStats
    For Each s In candidates
        If best Is Nothing Then
            Set best = s
        ElseIf s.DateRatio > best.DateRatio Then
            Set best = s
        End If
    Next s
    PickDate = best.Index
End Function

Private Function LooksLikeBalance(ByVal amountCol As clsColumnStats, ByVal balanceCol As clsColumnStats, _
                                  ByVal decimalSep As String) As Boolean
    Dim hits As Long, tested As Long
    Dim prevSet As Boolean, prev As Double
    Dim av As Variant, bv As Variant
    Dim n As Long, i As Long
    n = WorksheetFunction.Min(amountCol.Values.Count, balanceCol.Values.Count)
    For i = 1 To n
        Dim balanceParsed As Variant, amountParsed As Variant
        balanceParsed = ParseAmount(balanceCol.Values(i), decimalSep)
        amountParsed = ParseAmount(amountCol.Values(i), decimalSep)
        If IsNull(balanceParsed) Or IsNull(amountParsed) Then
            If Not IsNull(balanceParsed) Then
                prev = balanceParsed
                prevSet = True
            End If
        Else
            If prevSet Then
                tested = tested + 1
                If Abs((balanceParsed - prev) - amountParsed) < 0.02 Or _
                   Abs((prev - balanceParsed) - amountParsed) < 0.02 Then
                    hits = hits + 1
                End If
            End If
            prev = balanceParsed
            prevSet = True
        End If
    Next i
    LooksLikeBalance = (tested >= 3 And hits >= tested * 0.7)
End Function

Private Sub PickAmount(ByVal stats As Collection, ByVal dateIndex As Long, ByVal mapping As clsColumnMapping)
    EnsureWordLists
    Dim numeric As New Collection, s As clsColumnStats
    For Each s In stats
        If s.IsNumericCol() And s.Index <> dateIndex And s.LongIntRatio < 0.8 Then numeric.Add s
    Next s
    If numeric.Count = 0 Then
        mapping.Notes.Add DK("Fandt ingen talkolonne - v~ae~lg bel~oe~bskolonnen manuelt.")
        Exit Sub
    End If

    ' 1) An explicit balance column is never the amount.
    Dim balances As New Collection
    For Each s In numeric
        If HasWord(s.header, mBalanceWords) Then balances.Add s
    Next s
    If balances.Count > 0 Then
        mapping.BalanceCol = balances(1).Index
        Dim rest0 As New Collection
        For Each s In numeric
            If s.Index <> mapping.BalanceCol Then rest0.Add s
        Next s
        Set numeric = rest0
    End If

    ' 2) Separate in/out columns?
    Dim ins As New Collection, outs As New Collection
    For Each s In numeric
        If HasWord(s.header, mInWords) Then ins.Add s
    Next s
    For Each s In numeric
        If HasWord(s.header, mOutWords) Then outs.Add s
    Next s
    If ins.Count > 0 And outs.Count > 0 And ins(1).Index <> outs(1).Index Then
        mapping.AmountInCol = ins(1).Index
        mapping.AmountOutCol = outs(1).Index
        Exit Sub
    End If
    If numeric.Count = 2 Then
        Dim anyNamed As Boolean
        anyNamed = False
        For Each s In numeric
            If HasWord(s.header, mAmountWords) Then anyNamed = True
        Next s
        If Not anyNamed Then
            Dim first As clsColumnStats, second As clsColumnStats
            Dim idx2 As Long
            idx2 = 0
            For Each s In numeric
                idx2 = idx2 + 1
                If idx2 = 1 Then Set first = s Else Set second = s
            Next s
            Dim exclusive As Long, tested As Long, i As Long
            Dim n As Long
            n = WorksheetFunction.Min(first.Values.Count, second.Values.Count)
            For i = 1 To n
                Dim aFilled As Boolean, bFilled As Boolean
                aFilled = (Trim$(CStr(first.Values(i))) <> "")
                bFilled = (Trim$(CStr(second.Values(i))) <> "")
                If aFilled Or bFilled Then
                    tested = tested + 1
                    If aFilled <> bFilled Then exclusive = exclusive + 1
                End If
            Next i
            If tested >= 3 And exclusive >= tested * 0.8 Then
                mapping.AmountInCol = first.Index
                mapping.AmountOutCol = second.Index
                mapping.Notes.Add DK("To kolonner udfyldes skiftevis - tolkes som ind- og udbetaling. " & _
                                     "Byt om i dialogen hvis det er omvendt.")
                Exit Sub
            End If
        End If
    End If

    If numeric.Count = 1 Then
        mapping.AmountCol = numeric(1).Index
        Exit Sub
    End If

    ' 3) Named amount column wins.
    Dim named As New Collection
    For Each s In numeric
        If HasWord(s.header, mAmountWords) Then named.Add s
    Next s
    If named.Count > 0 Then
        mapping.AmountCol = named(1).Index
        If mapping.BalanceCol = -1 Then
            For Each s In numeric
                If s.Index <> mapping.AmountCol Then
                    If LooksLikeBalance(named(1), s, mapping.DecimalSep) Then
                        mapping.BalanceCol = s.Index
                        Exit For
                    End If
                End If
            Next s
        End If
        Exit Sub
    End If

    ' 4) Otherwise look for the running balance pattern.
    Dim cand As clsColumnStats, other As clsColumnStats
    For Each cand In numeric
        For Each other In numeric
            If other.Index <> cand.Index Then
                If LooksLikeBalance(cand, other, mapping.DecimalSep) Then
                    mapping.AmountCol = cand.Index
                    mapping.BalanceCol = other.Index
                    mapping.Notes.Add DK("Kolonnen """) & other.header & _
                        DK(""" ser ud til at v~ae~re en l~oe~bende saldo og springes over.")
                    Exit Sub
                End If
            End If
        Next other
    Next cand

    ' 5) Give up gracefully: prefer decimals, signs and a late position.
    Dim bestCol As clsColumnStats, bestMixed As Long, bestDecimal As Double, bestIdx As Long
    Dim first2 As Boolean
    first2 = True
    For Each s In numeric
        Dim hasNeg As Boolean, hasPos As Boolean, v As Variant
        hasNeg = False: hasPos = False
        For Each v In s.Filled
            Dim pv As Variant
            pv = ParseAmount(v, mapping.DecimalSep)
            If Not IsNull(pv) Then
                If pv < 0 Then hasNeg = True
                If pv > 0 Then hasPos = True
            End If
        Next v
        Dim mixed As Long
        mixed = IIf(hasNeg And hasPos, 1, 0)
        If first2 Then
            Set bestCol = s: bestMixed = mixed: bestDecimal = s.DecimalRatio: bestIdx = s.Index
            first2 = False
        ElseIf (mixed > bestMixed) Or _
              (mixed = bestMixed And s.DecimalRatio > bestDecimal) Or _
              (mixed = bestMixed And s.DecimalRatio = bestDecimal And s.Index < bestIdx) Then
            Set bestCol = s: bestMixed = mixed: bestDecimal = s.DecimalRatio: bestIdx = s.Index
        End If
    Next s
    mapping.AmountCol = bestCol.Index
    mapping.Notes.Add DK("Flere talkolonner - valgte """) & bestCol.header & _
        DK(""" som bel~oe~b. Ret det i dialogen hvis det er forkert.")
End Sub

Private Function PickText(ByVal stats As Collection, ByVal usedIndexes As Collection) As Collection
    EnsureWordLists
    Dim used As Object
    Set used = CreateObject("Scripting.Dictionary")
    Dim u As Variant
    For Each u In usedIndexes
        used(CLng(u)) = True
    Next u

    Dim candidates As New Collection, s As clsColumnStats
    For Each s In stats
        If Not used.Exists(s.Index) And s.NumericRatio < 0.5 And s.DateRatio < 0.5 And _
           s.FillRatio > 0.2 And Not HasWord(s.header, mIgnoreWords) Then
            candidates.Add s
        End If
    Next s
    Set PickText = New Collection
    If candidates.Count = 0 Then Exit Function

    ' Score-sort candidates descending: (named, distinctRatio, min(avgLen,40), -index)
    Dim arr() As Object, n As Long, i As Long, j As Long
    n = candidates.Count
    ReDim arr(1 To n)
    i = 0
    For Each s In candidates
        i = i + 1
        Set arr(i) = s
    Next s
    For i = 2 To n
        Dim keyObj As clsColumnStats
        Set keyObj = arr(i)
        j = i - 1
        Do While j >= 1
            If ScoreLess(arr(j), keyObj) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = keyObj
    Next i

    Dim best As clsColumnStats
    Set best = arr(1)
    Dim chosen As New Collection
    chosen.Add best.Index

    If HasWord(best.header, mTextWords) Then
        For i = 2 To n
            If chosen.Count >= 2 Then Exit For
            Dim otherS As clsColumnStats
            Set otherS = arr(i)
            If HasWord(otherS.header, mTextWords) And otherS.FillRatio > 0.5 And otherS.AvgLength >= 4 Then
                chosen.Add otherS.Index
            End If
        Next i
    ElseIf best.DistinctRatio < 0.25 And n > 1 Then
        chosen.Add arr(2).Index
    End If

    ' sorted ascending
    Dim vals() As Long, m As Long
    m = chosen.Count
    ReDim vals(1 To m)
    Dim c As Variant
    i = 0
    For Each c In chosen
        i = i + 1
        vals(i) = c
    Next c
    For i = 2 To m
        Dim keyV As Long
        keyV = vals(i)
        j = i - 1
        Do While j >= 1
            If vals(j) > keyV Then
                vals(j + 1) = vals(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        vals(j + 1) = keyV
    Next i
    Dim result As New Collection
    For i = 1 To m
        result.Add vals(i)
    Next i
    Set PickText = result
End Function

' True when a should sort AFTER b in a descending-score list (i.e. a's
' score is lower than b's) - used to keep the insertion sort stable.
Private Function ScoreLess(ByVal a As clsColumnStats, ByVal b As clsColumnStats) As Boolean
    Dim aNamed As Long, bNamed As Long
    aNamed = IIf(HasWord(a.header, mTextWords), 1, 0)
    bNamed = IIf(HasWord(b.header, mTextWords), 1, 0)
    If aNamed <> bNamed Then
        ScoreLess = (aNamed < bNamed)
        Exit Function
    End If
    If a.DistinctRatio <> b.DistinctRatio Then
        ScoreLess = (a.DistinctRatio < b.DistinctRatio)
        Exit Function
    End If
    Dim aLen As Double, bLen As Double
    aLen = WorksheetFunction.Min(a.AvgLength, 40)
    bLen = WorksheetFunction.Min(b.AvgLength, 40)
    If aLen <> bLen Then
        ScoreLess = (aLen < bLen)
        Exit Function
    End If
    ScoreLess = (-a.Index < -b.Index)
End Function

Public Function DetectMapping(ByVal t As clsTable, Optional ByVal limit As Long = 300) As clsColumnMapping
    Dim mapping As New clsColumnMapping
    Dim stats As Collection
    Set stats = AnalyseColumns(t, limit)
    If stats.Count = 0 Then
        mapping.Notes.Add "Tabellen er tom."
        Set DetectMapping = mapping
        Exit Function
    End If

    mapping.DateCol = PickDate(stats)
    If mapping.DateCol >= 0 Then
        Dim s As clsColumnStats
        For Each s In stats
            If s.Index = mapping.DateCol Then
                mapping.dayFirst = DetectDayfirst(s.Filled)
                Exit For
            End If
        Next s
    End If

    Dim numericValues As New Collection
    For Each s In stats
        If s.IsNumericCol() And s.Index <> mapping.DateCol Then
            Dim v As Variant
            For Each v In s.Filled
                numericValues.Add v
            Next v
        End If
    Next s
    mapping.DecimalSep = DetectDecimalSeparator(numericValues)
    If mapping.DecimalSep = "" Then mapping.DecimalSep = ","

    PickAmount stats, mapping.DateCol, mapping

    Dim used As New Collection
    If mapping.DateCol >= 0 Then used.Add mapping.DateCol
    If mapping.AmountCol >= 0 Then used.Add mapping.AmountCol
    If mapping.AmountInCol >= 0 Then used.Add mapping.AmountInCol
    If mapping.AmountOutCol >= 0 Then used.Add mapping.AmountOutCol
    If mapping.BalanceCol >= 0 Then used.Add mapping.BalanceCol

    For Each s In stats
        If Not InCollection(used, s.Index) And HasWord(s.header, mCurrencyWords) And s.AvgLength <= 6 Then
            mapping.CurrencyCol = s.Index
            used.Add s.Index
            Exit For
        End If
    Next s

    Dim textCols As Collection
    Set textCols = PickText(stats, used)
    Set mapping.TextCols = textCols
    If textCols.Count = 0 Then
        mapping.Notes.Add DK("Fandt ingen tekstkolonne - alle poster f~aa~r samme beskrivelse.")
    End If
    If mapping.DateCol = -1 Then
        mapping.Notes.Add DK("Fandt ingen datokolonne - v~ae~lg den manuelt.")
    End If
    Set DetectMapping = mapping
End Function

Private Function InCollection(ByVal col As Collection, ByVal value As Long) As Boolean
    Dim v As Variant
    For Each v In col
        If CLng(v) = value Then
            InCollection = True
            Exit Function
        End If
    Next v
    InCollection = False
End Function
