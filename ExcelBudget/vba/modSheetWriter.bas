Attribute VB_Name = "modSheetWriter"
Option Explicit
' Writes the budget into an Excel workbook. Port of office/budget_office.py.
' Danish sheet names with special letters are Functions (DK-decoded), not
' Const, since VBA Const requires a compile-time literal - see modI18n.

Public Const SHEET_TX As String = "Transaktioner"
Public Const SHEET_RULES As String = "Kategorier"
Public Const SHEET_PLAN As String = "Budgetforslag"
Public Const SHEET_FORECAST As String = "Prognose"
Public Const SHEET_ACCOUNTS As String = "Konti"
Public Const SHEET_BANK As String = "Bankbudget"

Public Function SHEET_MONTHS() As String
    SHEET_MONTHS = DK("Alle m~aa~neder")
End Function

Public Const TX_HEADER_ROW As Long = 4
Public Const TX_FIRST_ROW As Long = 5          ' first data row, 1-based
Public Const TX_MAX_ROW As Long = 5000
Public Const COL_EXP_DATE As Long = 2
Public Const COL_EXP_AMOUNT As Long = 3
Public Const COL_EXP_TEXT As Long = 4
Public Const COL_EXP_CAT As Long = 5
Public Const COL_INC_DATE As Long = 7
Public Const COL_INC_AMOUNT As Long = 8
Public Const COL_INC_TEXT As Long = 9
Public Const COL_INC_CAT As Long = 10

Public Const ACC_HEADER_ROW As Long = 4
Public Const ACC_FIRST_ROW As Long = 5
Public Const ACC_COL_NAME As Long = 2
Public Const ACC_COL_TYPE As Long = 3
Public Const ACC_COL_BALANCE As Long = 4
Public Const ACC_COL_NUMBER As Long = 5

Public Const PLAN_HEADER_ROW As Long = 5

' -----------------------------------------------------------------------
' Small helpers
' -----------------------------------------------------------------------

Public Function ColLetter(ByVal col As Long) As String
    Dim n As Long, s As String
    n = col
    Do While n > 0
        Dim rem2 As Long
        rem2 = (n - 1) Mod 26
        s = Chr(65 + rem2) & s
        n = (n - rem2 - 1) \ 26
    Loop
    ColLetter = s
End Function

Public Function CellRef(ByVal col1 As Long, ByVal row1 As Long) As String
    CellRef = ColLetter(col1) & row1
End Function

Public Function RangeRef(ByVal col1a As Long, ByVal row1a As Long, _
                        ByVal col1b As Long, ByVal row1b As Long) As String
    RangeRef = ColLetter(col1a) & row1a & ":" & ColLetter(col1b) & row1b
End Function

Public Function EnsureSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = sheetName
    EnsureSheet = ws
End Function

Public Function UsedRowCount(ByVal ws As Worksheet) As Long
    On Error Resume Next
    UsedRowCount = ws.Cells.Find("*", ws.Cells(1, 1), xlFormulas, , xlByRows, xlPrevious).row
    On Error GoTo 0
    If UsedRowCount = 0 Then UsedRowCount = 1
End Function

Public Function NewPen(ByVal ws As Worksheet) As clsPen
    Dim p As New clsPen
    p.Init ws
    Set NewPen = p
End Function

Public Function Serial(ByVal d As Date) As Double
    Serial = CDbl(d)
End Function

' -----------------------------------------------------------------------
' Transaktioner
' -----------------------------------------------------------------------

Public Sub WriteTransactions(ByVal ws As Worksheet, ByVal transactions As Collection)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 11
    pen.ColumnWidth 2, 15
    pen.ColumnWidth 3, 30
    pen.ColumnWidth 4, 16
    pen.ColumnWidth 5, 4
    pen.ColumnWidth 6, 11
    pen.ColumnWidth 7, 15
    pen.ColumnWidth 8, 30
    pen.ColumnWidth 9, 16

    pen.Merge "B1:J1"
    pen.Text "B1", DK("Skift eller tilf~oe~j kategorier i kolonnerne nedenfor - og i " & _
                     "arket ""Kategorier"", hvis de skal s~ae~ttes automatisk."), _
             size:=10, color:=LIGHT_TEXT, bg:=NAVY, italic:=True, align:="left", valign:="center"
    pen.RowHeight 1, 22

    pen.Text "B2", "Udgifter", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE
    pen.Text "G2", DK("Indt~ae~gter"), fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE
    pen.RowHeight 2, 26

    Dim headers As Variant
    headers = Array("Dato", DK("Bel~oe~b"), "Beskrivelse", "Kategori")
    Dim cols1 As Variant, cols2 As Variant, i As Long
    cols1 = Array("B", "C", "D", "E")
    cols2 = Array("G", "H", "I", "J")
    For i = 0 To 3
        pen.Text cols1(i) & TX_HEADER_ROW, headers(i), bold:=True, size:=11, color:=NAVY
        pen.Text cols2(i) & TX_HEADER_ROW, headers(i), bold:=True, size:=11, color:=NAVY
    Next i

    Dim expenses As New Collection, income As New Collection
    Dim t As clsTransaction
    For Each t In transactions
        If t.Amount < 0 Then expenses.Add t Else income.Add t
    Next t
    WriteBlock ws, expenses, COL_EXP_DATE, TX_FIRST_ROW
    WriteBlock ws, income, COL_INC_DATE, TX_FIRST_ROW

    Dim rows2 As Long
    rows2 = WorksheetFunction.Max(expenses.Count, income.Count, 1)
    Dim last As Long
    last = TX_FIRST_ROW + rows2 - 1
    StyleTxRange pen, last

    DimIgnoredRows pen, expenses, COL_EXP_DATE
    DimIgnoredRows pen, income, COL_INC_DATE

    AddCategoryDropdown ws, COL_EXP_CAT, last
    AddCategoryDropdown ws, COL_INC_CAT, last
    FreezeAt ws, 1, TX_HEADER_ROW
End Sub

Public Sub StyleTxRange(ByVal pen As clsPen, ByVal last As Long)
    Dim firstCol As Variant
    For Each firstCol In Array(COL_EXP_DATE, COL_INC_DATE)
        pen.Style RangeRef(firstCol, TX_FIRST_ROW, firstCol, last), fmt:=FMT_DATE, _
                 color:=MUTED, align:="left"
        pen.Style RangeRef(firstCol + 1, TX_FIRST_ROW, firstCol + 1, last), fmt:=FMT_CURRENCY2, _
                 color:=TEXT_GREY, bold:=True, align:="left"
        pen.Style RangeRef(firstCol + 2, TX_FIRST_ROW, firstCol + 3, last), color:=TEXT_GREY, _
                 align:="left"
    Next firstCol
End Sub

Public Sub WriteBlock(ByVal ws As Worksheet, ByVal transactions As Collection, _
                       ByVal firstCol As Long, ByVal startRow As Long)
    If transactions.Count = 0 Then Exit Sub
    Dim n As Long
    n = transactions.Count
    Dim data() As Variant
    ReDim data(1 To n, 1 To 4)
    Dim i As Long, t As clsTransaction
    i = 0
    For Each t In transactions
        i = i + 1
        data(i, 1) = t.TxDate
        data(i, 2) = Abs(t.Amount)
        data(i, 3) = t.Text
        data(i, 4) = t.Category
    Next t
    ws.Range(ws.Cells(startRow, firstCol), ws.Cells(startRow + n - 1, firstCol + 3)).Value = data
End Sub

Private Sub DimIgnoredRows(ByVal pen As clsPen, ByVal transactions As Collection, ByVal firstCol As Long)
    Dim offset As Long, t As clsTransaction
    offset = 0
    For Each t In transactions
        If t.Category = IGNORED_CATEGORY Then
            Dim row2 As Long
            row2 = TX_FIRST_ROW + offset
            pen.Style RangeRef(firstCol, row2, firstCol + 3, row2), color:=MUTED, italic:=True
        End If
        offset = offset + 1
    Next t
End Sub

Public Sub AddCategoryDropdown(ByVal ws As Worksheet, ByVal col As Long, ByVal lastRow As Long)
    If lastRow < TX_FIRST_ROW Then Exit Sub
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.Dropdown RangeRef(col, TX_FIRST_ROW, col, lastRow), "'" & SHEET_RULES & "'!$E$5:$E$80"
End Sub

Public Sub FreezeAt(ByVal ws As Worksheet, ByVal col0 As Long, ByVal row0 As Long)
    ws.Activate
    ws.Cells(row0 + 1, col0 + 1).Select
    ActiveWindow.FreezePanes = True
End Sub

' Read every transaction back from the sheet. Mirrors budget_office's
' read_transactions(): the two blocks share the same starting row, so a
' given physical row can hold one expense AND one unrelated income entry.
Public Function ReadTransactions(ByVal wb As Workbook) As Collection
    Dim out As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_TX)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ReadTransactions = out
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    If lastRow < TX_FIRST_ROW Then
        Set ReadTransactions = out
        Exit Function
    End If

    Dim blocks As Variant
    blocks = Array(Array(COL_EXP_DATE, COL_EXP_AMOUNT, COL_EXP_TEXT, COL_EXP_CAT, -1#), _
                   Array(COL_INC_DATE, COL_INC_AMOUNT, COL_INC_TEXT, COL_INC_CAT, 1#))
    Dim b As Variant
    For Each b In blocks
        Dim dateCol As Long, amtCol As Long, textCol As Long, catCol As Long, sign As Double
        dateCol = b(0): amtCol = b(1): textCol = b(2): catCol = b(3): sign = b(4)
        Dim row2 As Long
        For row2 = TX_FIRST_ROW To lastRow
            Dim dateVal As Variant
            dateVal = ws.Cells(row2, dateCol).Value
            If Not IsDate(dateVal) Then GoTo ContinueRow
            Dim amtVal As Variant
            amtVal = ws.Cells(row2, amtCol).Value
            If Not IsNumeric(amtVal) Then GoTo ContinueRow
            If CDbl(amtVal) = 0 Then GoTo ContinueRow
            Dim tx As New clsTransaction
            tx.Init CDate(dateVal), CStr(ws.Cells(row2, textCol).Value), _
                   sign * Abs(CDbl(amtVal)), Trim$(CStr(ws.Cells(row2, catCol).Value)), _
                   "", row2
            out.Add tx
ContinueRow:
        Next row2
    Next b

    Set ReadTransactions = SortTransactionsByDateThenRow(out)
End Function

Private Function SortTransactionsByDateThenRow(ByVal col As Collection) As Collection
    Dim n As Long, arr() As Object, i As Long, j As Long
    n = col.Count
    Set SortTransactionsByDateThenRow = New Collection
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
            Dim swap As Boolean
            If arr(j).TxDate <> keyT.TxDate Then
                swap = (arr(j).TxDate > keyT.TxDate)
            Else
                swap = (arr(j).SourceRow > keyT.SourceRow)
            End If
            If swap Then
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
    Set SortTransactionsByDateThenRow = out
End Function

Public Function HasBudget(ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_TX)
    On Error GoTo 0
    HasBudget = Not (ws Is Nothing)
End Function

' -----------------------------------------------------------------------
' Kategorier: the rules table (B:C) and the category/account list (E:F)
' used as the dropdown source everywhere else in the workbook.
' -----------------------------------------------------------------------

Public Sub WriteRules(ByVal ws As Worksheet, ByVal ruleset As clsRuleSet, Optional ByVal accounts As Object)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 26
    pen.ColumnWidth 2, 22
    pen.ColumnWidth 3, 4
    pen.ColumnWidth 4, 22
    pen.ColumnWidth 5, 20

    pen.Merge "B2:F2"
    pen.Text "B2", "Kategorier", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"
    pen.Merge "B3:F3"
    pen.Text "B3", DK("Tilf~oe~j, ret eller slet regler i venstre tabel (n~oe~gleord -> " & _
                     "kategori). H~oe~jre tabel viser alle kategorier og hvilken konto, " & _
                     "de tr~ae~kkes fra."), _
             size:=9, italic:=True, color:=MUTED, align:="left", wrap:=True, valign:="center"
    pen.RowHeight 3, 34

    pen.Text "B4", DK("N~oe~gleord"), bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "C4", "Kategori", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "E4", "Kategori", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "F4", "Konto", bold:=True, size:=11, color:=NAVY, align:="left"

    Dim rows As Collection
    Set rows = ruleset.ToRows()
    Dim r As Long
    r = 0
    Dim row2 As Variant
    For Each row2 In rows
        r = r + 1
        pen.Text CellRef(2, 4 + r), CStr(row2(0)), size:=10, color:=DARK, align:="left"
        pen.Text CellRef(3, 4 + r), CStr(row2(1)), size:=10, color:=DARK, align:="left"
    Next row2

    Dim categories As Collection
    Set categories = ruleset.Categories()
    Dim c As Long
    c = 0
    Dim cat As Variant
    For Each cat In categories
        c = c + 1
        pen.Text CellRef(5, 4 + c), CStr(cat), size:=10, color:=DARK, align:="left"
        Dim acct As String
        acct = ""
        If Not accounts Is Nothing Then
            If accounts.Exists(CStr(cat)) Then acct = accounts(CStr(cat))
        End If
        pen.Text CellRef(6, 4 + c), acct, size:=10, color:=DARK, align:="left"
    Next cat
    Dim lastCatRow As Long
    lastCatRow = 4 + WorksheetFunction.Max(c, 1)
    pen.Dropdown RangeRef(6, 5, 6, lastCatRow), "'" & SHEET_ACCOUNTS & "'!$B$5:$B$50"

    FreezeAt ws, 0, 4
End Sub

Public Function ReadRules(ByVal wb As Workbook) As Collection
    Dim out As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_RULES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ReadRules = out
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    Dim r As Long
    For r = 5 To lastRow
        Dim kw As String, cat As String
        kw = Trim$(CStr(ws.Cells(r, 2).Value))
        cat = Trim$(CStr(ws.Cells(r, 3).Value))
        If kw <> "" And cat <> "" Then out.Add Array(kw, cat)
    Next r
    Set ReadRules = out
End Function

Public Function ReadCategoryAccounts(ByVal wb As Workbook) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_RULES)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ReadCategoryAccounts = out
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    Dim r As Long
    For r = 5 To lastRow
        Dim cat As String, acct As String
        cat = Trim$(CStr(ws.Cells(r, 5).Value))
        acct = Trim$(CStr(ws.Cells(r, 6).Value))
        If cat <> "" And acct <> "" Then out(cat) = acct
    Next r
    Set ReadCategoryAccounts = out
End Function

' -----------------------------------------------------------------------
' Konti
' -----------------------------------------------------------------------

' rows: Collection of Array(navn, type, saldo, kontonummer)
Public Sub WriteAccounts(ByVal ws As Worksheet, ByVal rows As Collection)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 22
    pen.ColumnWidth 2, 18
    pen.ColumnWidth 3, 16
    pen.ColumnWidth 4, 24

    pen.Merge "B2:F2"
    pen.Text "B2", "Konti", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"
    pen.Merge "B3:F3"
    pen.Text "B3", DK("Nuv~ae~rende saldo - Prognosen bruger summen som startsaldo. S~ae~t " & _
                     "dit eget kontonummer under ""Kontonummer(e)"" og tryk Opdat~ee~r for " & _
                     "at f~aa~ overf~oe~rsler mellem dine egne konti sat til """ & _
                     IGNORED_CATEGORY & """ i stedet for at t~ae~lle som indt~ae~gt/udgift " & _
                     "- adskil flere numre med komma."), _
             size:=9, italic:=True, color:=MUTED, align:="left", wrap:=True, valign:="center"
    pen.RowHeight 3, 60

    pen.Text "B4", "Konto", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "C4", "Type", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "D4", "Saldo", bold:=True, size:=11, color:=NAVY, align:="right"
    pen.Text "E4", DK("Kontonummer(e)"), bold:=True, size:=11, color:=NAVY, align:="left"

    Dim r As Long
    r = 0
    Dim row2 As Variant
    For Each row2 In rows
        r = r + 1
        Dim rr As Long
        rr = ACC_HEADER_ROW + r
        pen.Text CellRef(ACC_COL_NAME, rr), CStr(row2(0)), bold:=True, size:=10, color:=DARK, align:="left"
        pen.Text CellRef(ACC_COL_TYPE, rr), CStr(row2(1)), size:=10, color:=TEXT_GREY, align:="left"
        pen.Number CellRef(ACC_COL_BALANCE, rr), CDbl(row2(2)), fmt:=FMT_CURRENCY, size:=10, _
                  color:=DARK, align:="right", bg:=PEACH
        pen.Text CellRef(ACC_COL_NUMBER, rr), CStr(row2(3)), size:=10, color:=DARK, align:="left", bg:=PEACH
    Next row2

    Dim n As Long
    n = WorksheetFunction.Max(rows.Count, 1)
    Dim lastRow As Long
    lastRow = ACC_HEADER_ROW + n
    Dim totalRow As Long
    totalRow = lastRow + 2
    pen.Text CellRef(ACC_COL_NAME, totalRow), "I alt", bold:=True, size:=10, color:=NAVY, align:="left"
    pen.Formula CellRef(ACC_COL_BALANCE, totalRow), _
               "=SUM(" & RangeRef(ACC_COL_BALANCE, ACC_FIRST_ROW, ACC_COL_BALANCE, lastRow) & ")", _
               fmt:=FMT_CURRENCY, bold:=True, color:=NAVY, align:="right"

    FreezeAt ws, 0, ACC_HEADER_ROW
End Sub

Public Sub EnsureAccountsSheet(ByVal wb As Workbook)
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = wb.Worksheets(SHEET_ACCOUNTS)
    On Error GoTo 0
    If Not existing Is Nothing Then Exit Sub
    Dim ws As Worksheet
    Set ws = EnsureSheet(wb, SHEET_ACCOUNTS)
    Dim seed As New Collection
    seed.Add Array(DK("L~oe~nkonto"), DK("L~oe~nkonto"), 0#, "")
    seed.Add Array("Budgetkonto", "Budgetkonto", 0#, "")
    seed.Add Array("Opsparingskonto", "Opsparingskonto", 0#, "")
    WriteAccounts ws, seed
End Sub

' [(navn, type, saldo, kontonummer, ui_row)]
Public Function ScanAccounts(ByVal wb As Workbook) As Collection
    Dim out As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_ACCOUNTS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ScanAccounts = out
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    Dim r As Long
    For r = ACC_FIRST_ROW To lastRow
        Dim name2 As String
        name2 = Trim$(CStr(ws.Cells(r, ACC_COL_NAME).Value))
        If name2 = "" Or FoldText(name2) = "i alt" Then GoTo ContinueLoop
        Dim kind As String, balance As Double, number As String
        kind = Trim$(CStr(ws.Cells(r, ACC_COL_TYPE).Value))
        If IsNumeric(ws.Cells(r, ACC_COL_BALANCE).Value) Then
            balance = CDbl(ws.Cells(r, ACC_COL_BALANCE).Value)
        End If
        number = Trim$(CStr(ws.Cells(r, ACC_COL_NUMBER).Value))
        out.Add Array(name2, kind, balance, number, r)
ContinueLoop:
    Next r
    Set ScanAccounts = out
End Function

Public Function ReadAccountNumbers(ByVal wb As Workbook) As Collection
    Dim out As New Collection
    Dim a As Variant
    For Each a In ScanAccounts(wb)
        Dim n As Variant
        For Each n In ParseAccountNumbers(CStr(a(3)))
            out.Add n
        Next n
    Next a
    Set ReadAccountNumbers = out
End Function

Public Function AccountsTotal(ByVal wb As Workbook) As Double
    Dim a As Variant, s As Double
    For Each a In ScanAccounts(wb)
        s = s + CDbl(a(2))
    Next a
    AccountsTotal = s
End Function

' -----------------------------------------------------------------------
' Budgetforslag
' -----------------------------------------------------------------------

' Writes the plan and returns a Dictionary of row markers ("income",
' "expense", "savings" -> row number) that Prognose's formulas point at,
' plus "accounts" -> Dictionary(account name -> Collection of Array(row, sign)).
Public Function WritePlan(ByVal ws As Worksheet, ByVal plan As clsBudgetPlan, _
                          Optional ByVal targets As Object, Optional ByVal accounts As Object, _
                          Optional ByVal planAccounts As Object) As Object
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 26
    pen.ColumnWidth 2, 16
    pen.ColumnWidth 3, 16
    pen.ColumnWidth 4, 14
    pen.ColumnWidth 5, 18

    pen.Merge "B2:F2"
    pen.Text "B2", "Budgetforslag", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"
    pen.Merge "B3:F3"
    Dim introText As String
    If plan.IsUncertain() Then
        introText = DK("Udkast til et fast m~aa~nedsbudget. Bem~ae~rk: bygger kun p~aa~ " & _
                       CStr(plan.Months.Count) & DK(" m~aa~ned(er) - tallene er usikre. " & _
                       "Ret M~aa~l-kolonnen som du vil."))
    Else
        introText = DK("Udkast til et fast m~aa~nedsbudget ud fra hele perioden. Ret " & _
                       "M~aa~l-kolonnen som du vil - resten genberegnes automatisk.")
    End If
    pen.Text "B3", introText, size:=9, italic:=True, color:=MUTED, align:="left", _
             wrap:=True, valign:="center"
    pen.RowHeight 3, 34

    Dim header As Long
    header = PLAN_HEADER_ROW
    pen.Text CellRef(2, header), "Kategori", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text CellRef(3, header), DK("Gennemsnit/md"), bold:=True, size:=10, color:=NAVY, align:="right"
    pen.Text CellRef(4, header), DK("M~aa~l"), bold:=True, size:=11, color:=NAVY, align:="right"
    pen.Text CellRef(5, header), "Forskel", bold:=True, size:=10, color:=NAVY, align:="right"
    pen.Text CellRef(6, header), "Konto", bold:=True, size:=10, color:=NAVY, align:="left"

    Dim row2 As Long
    row2 = header + 1
    Dim accountRows As Object
    Set accountRows = CreateObject("Scripting.Dictionary")

    ' -- income --------------------------------------------------------
    pen.Text CellRef(2, row2), DK("INDT~AE~GTER"), bold:=True, size:=10, color:=ORANGE, align:="left"
    row2 = row2 + 1
    Dim incomeFirst As Long
    incomeFirst = row2
    Dim c As clsCategoryPlan
    For Each c In plan.Income
        WritePlanRow pen, ws, row2, c, targets, accounts, planAccounts, accountRows, 1
        row2 = row2 + 1
    Next c
    Dim incomeLast As Long
    incomeLast = row2 - 1
    Dim incomeTotalRow As Long
    row2 = row2 + 1
    incomeTotalRow = row2
    pen.Text CellRef(2, row2), DK("INDT~AE~GTER I ALT"), bold:=True, size:=10, color:=NAVY, align:="left"
    If incomeLast >= incomeFirst Then
        pen.Formula CellRef(4, row2), "=SUM(" & RangeRef(4, incomeFirst, 4, incomeLast) & ")", _
                   fmt:=FMT_CURRENCY, bold:=True, color:=NAVY, align:="right"
    Else
        pen.Number CellRef(4, row2), 0, fmt:=FMT_CURRENCY, bold:=True, color:=NAVY, align:="right"
    End If
    row2 = row2 + 2

    ' -- expenses, grouped -----------------------------------------------
    Dim groupNames As Variant
    groupNames = Array(GROUP_FIXED, GROUP_VARIABLE, GROUP_PERIODIC)
    Dim expenseFirst As Long, expenseLast As Long
    expenseFirst = row2
    Dim g As Variant
    For Each g In groupNames
        Dim members As Collection
        Set members = plan.GroupCategories(CStr(g))
        If members.Count > 0 Then
            pen.Text CellRef(2, row2), GroupLabel(CStr(g)), bold:=True, size:=10, color:=ORANGE, align:="left"
            row2 = row2 + 1
            For Each c In members
                WritePlanRow pen, ws, row2, c, targets, accounts, planAccounts, accountRows, -1
                row2 = row2 + 1
            Next c
            row2 = row2 + 1
        End If
    Next g
    expenseLast = row2 - 1

    Dim expenseTotalRow As Long
    expenseTotalRow = row2
    pen.Text CellRef(2, row2), "UDGIFTER I ALT", bold:=True, size:=10, color:=NAVY, align:="left"
    ' SUM() over the whole block ignores the group-label text rows
    ' automatically, so no need to skip them explicitly.
    pen.Formula CellRef(4, row2), "=SUM(" & RangeRef(4, expenseFirst, 4, expenseLast) & ")", _
               fmt:=FMT_CURRENCY, bold:=True, color:=NAVY, align:="right"
    row2 = row2 + 2

    Dim savingsRow As Long
    savingsRow = row2
    pen.Text CellRef(2, row2), "TIL OPSPARING", bold:=True, size:=11, color:=ORANGE, align:="left"
    pen.Formula CellRef(4, row2), "=" & CellRef(4, incomeTotalRow) & "-" & CellRef(4, expenseTotalRow), _
               fmt:=FMT_SIGNED, bold:=True, color:=ORANGE, align:="right"

    FreezeAt ws, 0, header

    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    out("income") = incomeTotalRow
    out("expense") = expenseTotalRow
    out("savings") = savingsRow
    Set out("accounts") = accountRows
    Set WritePlan = out
End Function

Private Sub WritePlanRow(ByVal pen As clsPen, ByVal ws As Worksheet, ByVal row2 As Long, _
                         ByVal c As clsCategoryPlan, ByVal targets As Object, _
                         ByVal accounts As Object, ByVal planAccounts As Object, _
                         ByVal accountRows As Object, ByVal sign As Long)
    pen.Text CellRef(2, row2), c.Category, size:=10, color:=DARK, align:="left"
    pen.Number CellRef(3, row2), c.MeanAll, fmt:=FMT_CURRENCY, size:=10, color:=MUTED, align:="right"

    Dim target As Double
    target = c.Suggestion
    Dim key As String
    key = c.Kind & "|" & c.Category
    If Not targets Is Nothing Then
        If targets.Exists(key) Then target = CDbl(targets(key))
    End If
    pen.Number CellRef(4, row2), target, fmt:=FMT_CURRENCY, size:=10, bold:=True, _
              color:=DARK, align:="right", bg:=PEACH

    pen.Formula CellRef(5, row2), "=" & CellRef(4, row2) & "-" & CellRef(3, row2), _
               fmt:=FMT_SIGNED, size:=10, color:=TEXT_GREY, align:="right"

    Dim acct As String
    acct = ""
    If Not planAccounts Is Nothing Then
        If planAccounts.Exists(key) Then acct = planAccounts(key)
    End If
    If acct = "" And Not accounts Is Nothing Then
        If accounts.Exists(c.Category) Then acct = accounts(c.Category)
    End If
    pen.Text CellRef(6, row2), acct, size:=10, color:=DARK, align:="left"
    If acct <> "" Then
        If Not accountRows.Exists(acct) Then Set accountRows(acct) = New Collection
        accountRows(acct).Add Array(row2, sign)
    End If

    Dim noteText As String
    noteText = DK("Gennemsnit: ") & Format(c.MeanAll, "#,##0") & " kr./md" & vbLf & _
              CStr(c.TransactionCount) & " postering(er) i perioden."
    pen.Note CellRef(2, row2), noteText
End Sub

' True for the section/total label rows on Budgetforslag ("INDT?GTER",
' "FASTE UDGIFTER", "UDGIFTER I ALT", ...) - never for a real category, even
' one literally named "Indt?gt" (a real income category name).
Private Function IsPlanLabelRow(ByVal cat As String) As Boolean
    Dim folded As String
    folded = FoldText(cat)
    Dim labels As Variant
    labels = Array(FoldText(DK("INDT~AE~GTER")), FoldText(DK("INDT~AE~GTER I ALT")), _
                   "UDGIFTER I ALT", "TIL OPSPARING", _
                   FoldText(GroupLabel(GROUP_FIXED)), FoldText(GroupLabel(GROUP_VARIABLE)), _
                   FoldText(GroupLabel(GROUP_PERIODIC)))
    Dim l As Variant
    For Each l In labels
        If folded = FoldText(CStr(l)) Then
            IsPlanLabelRow = True
            Exit Function
        End If
    Next l
    IsPlanLabelRow = False
End Function

Public Function ReadTargets(ByVal wb As Workbook) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_PLAN)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ReadTargets = out
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    Dim r As Long
    For r = PLAN_HEADER_ROW + 1 To lastRow
        Dim cat As String, valCell As Variant
        cat = Trim$(CStr(ws.Cells(r, 2).Value))
        If cat = "" Then GoTo ContinueLoop
        If IsPlanLabelRow(cat) Then GoTo ContinueLoop
        valCell = ws.Cells(r, 4).Value
        If Not IsNumeric(valCell) Then GoTo ContinueLoop
        ' Figure out kind from which side of the sheet the category total
        ' fell under: rows above the expense total are still ambiguous by
        ' name alone, so store keyed only by category name - matched to
        ' both kinds when reapplied (harmless since names rarely collide).
        out(KIND_EXPENSE & "|" & cat) = CDbl(valCell)
        out(KIND_INCOME() & "|" & cat) = CDbl(valCell)
ContinueLoop:
    Next r
    Set ReadTargets = out
End Function

' -----------------------------------------------------------------------
' Prognose
' -----------------------------------------------------------------------

Public Const FORECAST_MONTHS As Long = 24
Public Const FORECAST_HEADER_ROW As Long = 9

Public Sub WriteForecast(ByVal ws As Worksheet, ByVal wb As Workbook, _
                         ByVal planRows As Object, ByVal lastMonthKey As String, _
                         ByVal fallbackStart As Double)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 16
    pen.ColumnWidth 2, 15
    pen.ColumnWidth 3, 15
    pen.ColumnWidth 4, 15

    pen.Merge "B2:E2"
    pen.Text "B2", "Prognose", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"
    pen.Merge "B3:E3"
    pen.Text "B3", DK("S~aa~dan udvikler saldoen sig de n~ae~ste " & FORECAST_MONTHS & _
                     " m~aa~neder, hvis du rammer m~aa~lene i ""Budgetforslag""."), _
             size:=9, italic:=True, color:=MUTED, align:="left", wrap:=True, valign:="center"
    pen.RowHeight 3, 30

    Dim accountsTotal As Double
    accountsTotal = modSheetWriter.AccountsTotal(wb)
    pen.Text "B5", "Startsaldo", bold:=True, size:=10, color:=NAVY, align:="left"
    If accountsTotal <> 0 Then
        pen.Formula "D5", "=SUM('" & SHEET_ACCOUNTS & "'!" & _
                   RangeRef(ACC_COL_BALANCE, ACC_FIRST_ROW, ACC_COL_BALANCE, ACC_FIRST_ROW + 20) & ")", _
                   fmt:=FMT_CURRENCY, size:=10, color:=DARK, align:="right"
    Else
        pen.Number "D5", fallbackStart, fmt:=FMT_CURRENCY, size:=10, color:=DARK, align:="right", bg:=PEACH
    End If
    pen.Text "B6", DK("Netto pr. m~aa~ned med dine m~aa~l"), size:=10, color:=TEXT_GREY, align:="left"
    pen.Formula "D6", "='" & SHEET_PLAN & "'!" & CellRef(4, CLng(planRows("savings"))), _
               fmt:=FMT_SIGNED, size:=10, bold:=True, color:=ORANGE, align:="right"

    Dim head As Long
    head = FORECAST_HEADER_ROW
    pen.Text CellRef(2, head), DK("M~aa~ned"), bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text CellRef(3, head), "Netto/md", bold:=True, size:=10, color:=NAVY, align:="right"
    pen.Text CellRef(4, head), "Saldo", bold:=True, size:=11, color:=NAVY, align:="right"

    Dim startMonth As Date
    startMonth = NextMonth(lastMonthKey)
    Dim i As Long, r As Long
    For i = 0 To FORECAST_MONTHS - 1
        r = head + 1 + i
        pen.DateCell CellRef(2, r), DateAdd("m", i, startMonth), fmt:=FMT_MONTH, size:=10, _
                    color:=DARK, align:="left"
        pen.Formula CellRef(3, r), "=$D$6", fmt:=FMT_SIGNED, size:=10, color:=ORANGE, align:="right"
        Dim prevRef As String
        If i = 0 Then prevRef = "$D$5" Else prevRef = CellRef(4, r - 1)
        pen.Formula CellRef(4, r), "=" & prevRef & "+" & CellRef(3, r), fmt:=FMT_CURRENCY, _
                   size:=10, bold:=True, color:=ORANGE, align:="right"
    Next i

    Dim lastRow2 As Long
    lastRow2 = head + FORECAST_MONTHS
    AddForecastChart ws, head, lastRow2
    FreezeAt ws, 0, head
End Sub

Private Function NextMonth(ByVal monthKey As String) As Date
    If monthKey = "" Then
        NextMonth = DateSerial(Year(Date), Month(Date), 1)
        Exit Function
    End If
    Dim parts() As String
    parts = Split(monthKey, "-")
    NextMonth = DateAdd("m", 1, DateSerial(CLng(parts(0)), CLng(parts(1)), 1))
End Function

Private Sub AddForecastChart(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal lastRow As Long)
    Dim chartName As String
    chartName = DK("Saldoudvikling")
    Dim co As ChartObject
    Dim existing As ChartObject
    On Error Resume Next
    Set existing = ws.ChartObjects(chartName)
    On Error GoTo 0
    If Not existing Is Nothing Then existing.Delete

    Set co = ws.ChartObjects.Add(Left:=ws.Columns("F").Left, Top:=ws.Rows(2).Top, _
                                 Width:=420, Height:=260)
    co.Name = chartName
    Dim src As Range
    Set src = Application.Union(ws.Range(RangeRef(2, headerRow, 2, lastRow)), _
                                ws.Range(RangeRef(4, headerRow, 4, lastRow)))
    With co.Chart
        .SetSourceData Source:=src
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = chartName
    End With
End Sub

' -----------------------------------------------------------------------
' Alle maaneder
' -----------------------------------------------------------------------

Public Sub WriteMonths(ByVal ws As Worksheet, ByVal summary As clsSummary)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 26

    pen.Merge "B2:D2"
    pen.Text "B2", SHEET_MONTHS(), fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"

    Dim head As Long
    head = 4
    pen.Text CellRef(2, head), "Kategori", bold:=True, size:=11, color:=NAVY, align:="left"
    Dim m As Variant, col As Long
    col = 3
    For Each m In summary.Months
        pen.DateCell CellRef(col, head), MonthKeyToDate(CStr(m)), fmt:="mmm yyyy", bold:=True, _
                    size:=10, color:=NAVY, align:="right"
        col = col + 1
    Next m

    Dim row2 As Long
    row2 = head + 1
    Dim cat As Variant
    For Each cat In summary.ExpenseCategories
        pen.Text CellRef(2, row2), CStr(cat), size:=10, color:=DARK, align:="left"
        col = 3
        For Each m In summary.Months
            pen.Number CellRef(col, row2), summary.Value(KIND_EXPENSE, CStr(cat), CStr(m)), _
                      fmt:=FMT_CURRENCY, size:=10, color:=TEXT_GREY, align:="right"
            col = col + 1
        Next m
        row2 = row2 + 1
    Next cat
    row2 = row2 + 1
    For Each cat In summary.IncomeCategories
        pen.Text CellRef(2, row2), CStr(cat), size:=10, color:=DARK, align:="left"
        col = 3
        For Each m In summary.Months
            pen.Number CellRef(col, row2), summary.Value(KIND_INCOME(), CStr(cat), CStr(m)), _
                      fmt:=FMT_CURRENCY, size:=10, color:=NAVY, align:="right"
            col = col + 1
        Next m
        row2 = row2 + 1
    Next cat

    FreezeAt ws, 1, head
End Sub

Private Function MonthKeyToDate(ByVal monthKey As String) As Date
    Dim parts() As String
    parts = Split(monthKey, "-")
    MonthKeyToDate = DateSerial(CLng(parts(0)), CLng(parts(1)), 1)
End Function

' -----------------------------------------------------------------------
' Bankbudget: a plain-value snapshot of the target column, for handing to
' the bank. Unlike every other sheet this one is NOT rebuilt by Refresh -
' only by explicitly running "Opret bankbudget" again.
' -----------------------------------------------------------------------

Public Sub WriteBankBudget(ByVal ws As Worksheet, ByVal targets As Object)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    pen.ColumnWidth 0, 4
    pen.ColumnWidth 1, 30
    pen.ColumnWidth 2, 16

    pen.Merge "B2:C2"
    pen.Text "B2", "Bankbudget", fontName:=FONT_TITLE, size:=18, bold:=True, color:=ORANGE, align:="left"
    pen.Merge "B3:C3"
    pen.Text "B3", DK("~OE~jebliksbillede af m~aa~lene i Budgetforslag, som faste tal - " & _
                     "opdateres ikke automatisk. K~oe~r ""Opret bankbudget"" igen for en ny version."), _
             size:=9, italic:=True, color:=MUTED, align:="left", wrap:=True, valign:="center"
    pen.RowHeight 3, 34

    pen.Text "B5", "Kategori", bold:=True, size:=11, color:=NAVY, align:="left"
    pen.Text "C5", DK("M~aa~ned"), bold:=True, size:=11, color:=NAVY, align:="right"

    Dim row2 As Long
    row2 = 6
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim key As Variant
    For Each key In targets.Keys
        Dim parts() As String
        parts = Split(CStr(key), "|")
        Dim cat As String
        cat = parts(1)
        If Not seen.Exists(cat) Then
            seen(cat) = True
            pen.Text CellRef(2, row2), cat, size:=10, color:=DARK, align:="left"
            pen.Number CellRef(3, row2), CDbl(targets(key)), fmt:=FMT_CURRENCY, size:=10, _
                      color:=DARK, align:="right"
            row2 = row2 + 1
        End If
    Next key
    FreezeAt ws, 0, 5
End Sub

Public Function ReadForecastStart(ByVal wb As Workbook) As Variant
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_FORECAST)
    On Error GoTo 0
    If ws Is Nothing Then
        ReadForecastStart = Null
        Exit Function
    End If
    On Error Resume Next
    Dim v As Variant
    v = ws.Range("D5").Value
    On Error GoTo 0
    If IsNumeric(v) Then ReadForecastStart = CDbl(v) Else ReadForecastStart = Null
End Function

Public Function ReadPlanAccounts(ByVal wb As Workbook) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_PLAN)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ReadPlanAccounts = out
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = UsedRowCount(ws)
    Dim r As Long
    For r = PLAN_HEADER_ROW + 1 To lastRow
        Dim cat As String, acct As String
        cat = Trim$(CStr(ws.Cells(r, 2).Value))
        If cat = "" Or IsPlanLabelRow(cat) Then GoTo ContinueLoop
        acct = Trim$(CStr(ws.Cells(r, 6).Value))
        If acct <> "" Then
            out(KIND_EXPENSE & "|" & cat) = acct
            out(KIND_INCOME() & "|" & cat) = acct
        End If
ContinueLoop:
    Next r
    Set ReadPlanAccounts = out
End Function
