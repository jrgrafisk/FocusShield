Attribute VB_Name = "modBudgetCore"
Option Explicit
' Turn a detected table into categorised transactions and a budget.
' Port of budget_core/budget.py.

Public Const KIND_EXPENSE As String = "Udgift"

Public Function KIND_INCOME() As String
    KIND_INCOME = DK("Indt~ae~gt")
End Function

Public Const SIGN_AUTO As String = "auto"
Public Const SIGN_NEGATIVE_IS_EXPENSE As String = "negativ"
Public Const SIGN_POSITIVE_IS_EXPENSE As String = "positiv"

Private Function CellAt(ByRef row() As String, ByVal index As Long) As String
    If index < 0 Or index > UBound(row) Or index < LBound(row) Then
        CellAt = ""
    Else
        CellAt = row(index)
    End If
End Function

Private Function RowAmount(ByRef row() As String, ByVal mapping As clsColumnMapping, _
                           ByVal decimalSep As String) As Variant
    If mapping.AmountInCol >= 0 Or mapping.AmountOutCol >= 0 Then
        Dim valueIn As Variant, valueOut As Variant
        valueIn = Null: valueOut = Null
        If mapping.AmountInCol >= 0 Then valueIn = ParseAmount(CellAt(row, mapping.AmountInCol), decimalSep)
        If mapping.AmountOutCol >= 0 Then valueOut = ParseAmount(CellAt(row, mapping.AmountOutCol), decimalSep)
        If IsNull(valueIn) And IsNull(valueOut) Then
            RowAmount = Null
            Exit Function
        End If
        Dim inAbs As Double, outAbs As Double
        inAbs = 0#: outAbs = 0#
        If Not IsNull(valueIn) Then inAbs = Abs(CDbl(valueIn))
        If Not IsNull(valueOut) Then outAbs = Abs(CDbl(valueOut))
        RowAmount = inAbs - outAbs
        Exit Function
    End If
    RowAmount = ParseAmount(CellAt(row, mapping.AmountCol), decimalSep)
End Function

Private Function DecideSign(ByVal values As Collection, ByVal mode As String) As Boolean
    If mode = SIGN_POSITIVE_IS_EXPENSE Then
        DecideSign = True
        Exit Function
    End If
    If mode = SIGN_NEGATIVE_IS_EXPENSE Then
        DecideSign = False
        Exit Function
    End If
    Dim negatives As Long, v As Variant
    For Each v In values
        If v < 0 Then negatives = negatives + 1
    Next v
    DecideSign = (negatives = 0 And values.Count > 0)
End Function

Public Function BuildTransactions(ByVal t As clsTable, ByVal mapping As clsColumnMapping, _
                                  ByVal ruleset As clsRuleSet, _
                                  Optional ByVal decimalSep As String = "", _
                                  Optional ByVal dayFirstOverride As Variant, _
                                  Optional ByVal signMode As String = SIGN_AUTO, _
                                  Optional ByVal skipZero As Boolean = True, _
                                  Optional ByVal dateFrom As Variant, _
                                  Optional ByVal dateTo As Variant) As clsBuildResult
    Dim result As New clsBuildResult
    Dim w As Variant
    For Each w In mapping.Notes
        result.Warnings.Add w
    Next w

    Dim decimal2 As String
    decimal2 = decimalSep
    If decimal2 = "" Then decimal2 = mapping.DecimalSep

    Dim dayFirst2 As Boolean
    If IsMissing(dayFirstOverride) Then
        dayFirst2 = mapping.dayFirst
    Else
        dayFirst2 = CBool(dayFirstOverride)
    End If

    If mapping.DateCol < 0 Then
        result.Warnings.Add DK("Ingen datokolonne valgt.")
        Set BuildTransactions = result
        Exit Function
    End If
    If Not mapping.HasAmount() Then
        result.Warnings.Add DK("Ingen bel~oe~bskolonne valgt.")
        Set BuildTransactions = result
        Exit Function
    End If

    Dim n As Long
    n = t.Rows.Count
    Dim parsedDates() As Date, parsedTexts() As String, parsedAmounts() As Double
    Dim parsedCurrencies() As String, parsedIndexes() As Long
    ReDim parsedDates(1 To WorksheetFunction.Max(n, 1))
    ReDim parsedTexts(1 To WorksheetFunction.Max(n, 1))
    ReDim parsedAmounts(1 To WorksheetFunction.Max(n, 1))
    ReDim parsedCurrencies(1 To WorksheetFunction.Max(n, 1))
    ReDim parsedIndexes(1 To WorksheetFunction.Max(n, 1))
    Dim parsedCount As Long
    parsedCount = 0

    Dim skippedNoDate As Long, skippedNoAmount As Long
    Dim rowIdx As Long, r As Variant
    rowIdx = -1
    For Each r In t.Rows
        rowIdx = rowIdx + 1
        Dim row() As String
        row = r
        Dim parsedDate As Variant
        parsedDate = ParseDate(CellAt(row, mapping.DateCol), dayFirst2)
        Dim amountVal As Variant
        amountVal = RowAmount(row, mapping, decimal2)

        If IsNull(parsedDate) Then
            Dim anyText As Boolean, cellV As Variant
            anyText = Not IsNull(amountVal)
            If Not anyText Then
                For Each cellV In row
                    If Trim$(CStr(cellV)) <> "" Then anyText = True: Exit For
                Next cellV
            End If
            If anyText Then skippedNoDate = skippedNoDate + 1
            GoTo ContinueLoop
        End If
        If IsNull(amountVal) Then
            skippedNoAmount = skippedNoAmount + 1
            GoTo ContinueLoop
        End If

        Dim amountDbl As Double
        amountDbl = CDbl(amountVal)
        If skipZero And Abs(amountDbl) < 0.0000001 Then GoTo ContinueLoop
        If Not IsMissing(dateFrom) Then
            If CDate(parsedDate) < CDate(dateFrom) Then GoTo ContinueLoop
        End If
        If Not IsMissing(dateTo) Then
            If CDate(parsedDate) > CDate(dateTo) Then GoTo ContinueLoop
        End If

        Dim txText As String
        txText = ""
        Dim ci As Variant
        For Each ci In mapping.TextCols
            Dim part As String
            part = SqueezeText(CellAt(row, CLng(ci)))
            If part <> "" Then
                If txText <> "" Then txText = txText & " " & ChrW$(8211) & " "
                txText = txText & part
            End If
        Next ci
        If txText = "" Then txText = "(ingen tekst)"

        Dim currency2 As String
        currency2 = ""
        If mapping.CurrencyCol >= 0 Then currency2 = SqueezeText(CellAt(row, mapping.CurrencyCol))

        parsedCount = parsedCount + 1
        parsedDates(parsedCount) = CDate(parsedDate)
        parsedTexts(parsedCount) = txText
        parsedAmounts(parsedCount) = amountDbl
        parsedCurrencies(parsedCount) = currency2
        parsedIndexes(parsedCount) = rowIdx
ContinueLoop:
    Next r

    result.SkippedNoDate = skippedNoDate
    result.SkippedNoAmount = skippedNoAmount

    If parsedCount = 0 Then
        result.Warnings.Add DK("Ingen brugbare r~ae~kker. Kontroll~ee~r kolonnevalget i dialogen.")
        Set BuildTransactions = result
        Exit Function
    End If

    Dim flip As Boolean
    flip = False
    If mapping.AmountCol >= 0 And mapping.AmountInCol < 0 Then
        Dim amountsCol As New Collection
        Dim i As Long
        For i = 1 To parsedCount
            amountsCol.Add parsedAmounts(i)
        Next i
        flip = DecideSign(amountsCol, signMode)
        If flip And signMode = SIGN_AUTO Then
            result.Warnings.Add DK("Bel~oe~bskolonnen indeholder ingen negative tal - alle poster " & _
                "behandles som udgifter. Skift fortegnsregel i dialogen hvis " & _
                "filen ogs~aa~ indeholder indt~ae~gter.")
        End If
    End If

    Dim transactions As New Collection
    For i = 1 To parsedCount
        Dim value As Double
        value = parsedAmounts(i)
        If flip Then value = -value
        Dim tx As New clsTransaction
        tx.Init parsedDates(i), parsedTexts(i), value, ruleset.Categorise(parsedTexts(i)), _
               parsedCurrencies(i), parsedIndexes(i)
        transactions.Add tx
    Next i

    Set transactions = SortTransactions(transactions)
    Set result.Transactions = transactions

    ' Group uncategorised entries by shop for the "assign a category" flow.
    Dim groups As Object
    Set groups = CreateObject("Scripting.Dictionary")
    Dim order As New Collection
    Dim one As clsTransaction
    For Each one In transactions
        If one.Category = DEFAULT_CATEGORY Then
            Dim key As String
            key = MerchantKey(one.Text)
            If Not groups.Exists(key) Then
                groups(key) = Array(MerchantName(one.Text), 0, 0#, RuleKeyword(one.Text))
                order.Add key
            End If
            Dim entry As Variant
            entry = groups(key)
            entry(1) = entry(1) + 1
            entry(2) = entry(2) + one.Amount
            groups(key) = entry
        End If
    Next one

    Dim rawGroups As New Collection
    Dim k As Variant
    For Each k In order
        rawGroups.Add groups(k)
    Next k
    Set result.Uncategorised = SortGroupsByAbsTotalDesc(rawGroups)

    Set BuildTransactions = result
End Function

Private Function SortTransactions(ByVal col As Collection) As Collection
    Dim n As Long, arr() As Object, i As Long, j As Long
    n = col.Count
    Set SortTransactions = New Collection
    If n = 0 Then Exit Function
    ReDim arr(1 To n)
    i = 0
    Dim t As clsTransaction
    For Each t In col
        i = i + 1
        Set arr(i) = t
    Next t
    For i = 2 To n
        Dim keyT As clsTransaction
        Set keyT = arr(i)
        j = i - 1
        Do While j >= 1
            If TxLess(keyT, arr(j)) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = keyT
    Next i
    Dim out As New Collection
    For i = 1 To n
        out.Add arr(i)
    Next i
    Set SortTransactions = out
End Function

' True when "a" sorts before "b": earlier date, then earlier source row.
Private Function TxLess(ByVal a As clsTransaction, ByVal b As clsTransaction) As Boolean
    If a.TxDate <> b.TxDate Then
        TxLess = (a.TxDate < b.TxDate)
    Else
        TxLess = (a.SourceRow < b.SourceRow)
    End If
End Function

Private Function SortGroupsByAbsTotalDesc(ByVal col As Collection) As Collection
    Dim n As Long, arr() As Variant, i As Long, j As Long
    n = col.Count
    Set SortGroupsByAbsTotalDesc = New Collection
    If n = 0 Then Exit Function
    ReDim arr(1 To n)
    i = 0
    Dim g As Variant
    For Each g In col
        i = i + 1
        arr(i) = g
    Next g
    For i = 2 To n
        Dim keyG As Variant
        keyG = arr(i)
        j = i - 1
        Do While j >= 1 And Abs(arr(j)(2)) < Abs(keyG(2))
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = keyG
    Next i
    Dim out As New Collection
    For i = 1 To n
        out.Add arr(i)
    Next i
    Set SortGroupsByAbsTotalDesc = out
End Function
