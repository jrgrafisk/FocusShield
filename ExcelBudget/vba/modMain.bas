Attribute VB_Name = "modMain"
Option Explicit
' Top-level orchestration and the macros the "Budget" menu calls.
' Port of oxt/budget_extension.py, adapted for Excel: no modal dialogs (see
' the setup guide for why) - CSV import uses the auto-detected column
' mapping directly, and "assigning a category" happens by editing the
' Transaktioner/Kategorier sheets directly instead of through a popup.

Private Const APP_TITLE As String = "Budget fra CSV"

Private Function Msg(ByVal text2 As String)
    MsgBox text2, vbOKOnly + vbInformation, APP_TITLE
End Function

Private Function ErrMsg(ByVal text2 As String)
    MsgBox text2, vbOKOnly + vbCritical, APP_TITLE
End Function

Private Function JoinLines(ByVal lines As Collection) As String
    Dim out As String, first2 As Boolean
    first2 = True
    Dim l As Variant
    For Each l In lines
        If Not first2 Then out = out & vbLf
        out = out & l
        first2 = False
    Next l
    JoinLines = out
End Function

' -- rules persistence (a simple text file next to the workbook) ----------

Private Function RulesFilePath() As String
    RulesFilePath = Environ$("APPDATA") & "\BudgetFraCSV\kategoriregler.csv"
End Function

Public Function LoadRules() As clsRuleSet
    Dim rs As New clsRuleSet
    Dim path As String
    path = RulesFilePath()
    If Len(Dir$(path)) > 0 Then
        rs.LoadFromFile path
    Else
        rs.InitDefault
    End If
    Set LoadRules = rs
End Function

Private Sub SaveRulesQuietly(ByVal rs As clsRuleSet)
    On Error Resume Next
    rs.SaveToFile RulesFilePath()
    On Error GoTo 0
End Sub

' -- file picking -----------------------------------------------------

Private Function PickCsvFile() As String
    Dim result As Variant
    result = Application.GetOpenFilename( _
        FileFilter:="CSV-filer (*.csv;*.txt),*.csv;*.txt,Alle filer (*.*),*.*", _
        Title:=DK("V~ae~lg CSV-fil"))
    If result = False Then
        PickCsvFile = ""
    Else
        PickCsvFile = CStr(result)
    End If
End Function

' -- import -------------------------------------------------------------

Public Sub ImportCsv()
    Dim path As String
    path = PickCsvFile()
    If path = "" Then Exit Sub

    Dim t As clsTable
    On Error GoTo ReadFail
    Set t = ReadTable(path)
    On Error GoTo 0
    ImportTable t, Mid$(path, InStrRev(path, "\") + 1)
    Exit Sub
ReadFail:
    ErrMsg DK("Kunne ikke l~ae~se filen:") & vbLf & Err.Description
End Sub

Public Sub ImportActiveSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then
        Msg DK("~AA~bn f~oe~rst et regneark med posteringerne.")
        Exit Sub
    End If
    Dim t As clsTable
    Set t = TableFromSheet(ws)
    If t.Rows.Count = 0 Then
        Msg "Arket er tomt."
        Exit Sub
    End If
    ImportTable t, "arket """ & ws.Name & """"
End Sub

Private Function TableFromSheet(ByVal ws As Worksheet) As clsTable
    ' Reads .Text (the displayed string), not .Value: a real Excel date
    ' cell's .Value is a numeric serial, not text, and would fail to parse
    ' as a date later - .Text gives back what the user actually sees, the
    ' same kind of string a CSV file would have had in the first place.
    Dim used As Range
    Set used = ws.UsedRange
    Dim rows As New Collection
    Dim r As Long, c As Long, nRows As Long, nCols As Long
    nRows = used.Rows.Count
    nCols = used.Columns.Count
    For r = 1 To nRows
        Dim row2() As String
        ReDim row2(0 To nCols - 1)
        For c = 1 To nCols
            row2(c - 1) = used.Cells(r, c).Text
        Next c
        rows.Add row2
    Next r
    Set TableFromSheet = TableFromRows(rows, Empty, "utf-8", ";", ws.Name)
End Function

Private Sub ImportTable(ByVal t As clsTable, ByVal sourceName As String)
    If t.Rows.Count = 0 Then
        Msg "Fandt ingen datar" & DK("~ae~kker i ") & sourceName & "."
        Exit Sub
    End If
    Dim mapping As clsColumnMapping
    Set mapping = DetectMapping(t)

    Dim ruleset As clsRuleSet
    Set ruleset = LoadRules()
    Dim result As clsBuildResult
    Set result = BuildTransactions(t, mapping, ruleset)
    If Not result.Ok() Then
        ErrMsg DK("Der kunne ikke laves et budget.") & vbLf & vbLf & JoinLines(result.SummaryLines())
        Exit Sub
    End If

    Dim wb As Workbook
    Set wb = Application.Workbooks.Add
    Dim startBalance As Double
    startBalance = 0
    Dim summary As clsSummary
    Set summary = BuildWorkbook(wb, result, ruleset, startBalance)
    SaveRulesQuietly ruleset

    Msg DoneText(result, summary, mapping, t)
End Sub

Public Sub ImportCsvAppend()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If
    Dim path As String
    path = PickCsvFile()
    If path = "" Then Exit Sub

    Dim t As clsTable
    On Error GoTo ReadFail
    Set t = ReadTable(path)
    On Error GoTo 0
    If t.Rows.Count = 0 Then
        Msg "Fandt ingen datar" & DK("~ae~kker i ") & path & "."
        Exit Sub
    End If

    Dim mapping As clsColumnMapping
    Set mapping = DetectMapping(t)
    Dim rows As Collection
    Set rows = ReadRules(wb)
    Dim ruleset As New clsRuleSet
    If rows.Count > 0 Then ruleset.InitFromRows rows Else ruleset.InitDefault

    Dim result As clsBuildResult
    Set result = BuildTransactions(t, mapping, ruleset)
    If Not result.Ok() Then
        ErrMsg DK("Der kunne ikke tilf~oe~jes posteringer.") & vbLf & vbLf & JoinLines(result.SummaryLines())
        Exit Sub
    End If

    Dim added As Long, skipped As Long, truncated As Boolean
    Dim summary As clsSummary
    AppendTransactions wb, result.Transactions, added, skipped, truncated, summary
    If summary Is Nothing Then
        Msg "Ingen nye posteringer - alle " & skipped & " postering(er) fandtes i forvejen."
        Exit Sub
    End If

    Dim lines As String
    lines = added & " ny(e) postering(er) tilf" & DK("~oe~jet.")
    If skipped > 0 Then lines = lines & vbLf & skipped & " postering(er) fandtes allerede og blev sprunget over."
    If truncated Then lines = lines & vbLf & DK("Bem~ae~rk: nogle af de nyeste blev ikke tilf~oe~jet (pladsen i Transaktioner er fuld).")
    lines = lines & vbLf & summary.Months.Count & " m" & DK("~aa~ned(er) i budgettet nu.")
    Msg lines
    Exit Sub
ReadFail:
    ErrMsg DK("Kunne ikke l~ae~se filen:") & vbLf & Err.Description
End Sub

Private Function ActiveWorkbookOrNothing() As Workbook
    On Error Resume Next
    Set ActiveWorkbookOrNothing = Application.ActiveWorkbook
    On Error GoTo 0
End Function

Private Function DoneText(ByVal result As clsBuildResult, ByVal summary As clsSummary, _
                          ByVal mapping As clsColumnMapping, ByVal t As clsTable) As String
    Dim lines As New Collection
    Dim l As Variant
    For Each l In result.SummaryLines()
        lines.Add l
    Next l
    lines.Add ""
    lines.Add t.Describe()
    DoneText = JoinLines(lines)
End Function

' -- refresh (Opdater) --------------------------------------------------

Public Sub RefreshBudget()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If

    Dim rows As Collection
    Set rows = ReadRules(wb)
    Dim ruleset As New clsRuleSet
    If rows.Count > 0 Then ruleset.InitFromRows rows Else Set ruleset = LoadRules()

    Dim transactions As Collection
    Set transactions = ReadTransactions(wb)
    If transactions.Count = 0 Then
        Msg DK("Fandt ingen posteringer at opdatere.")
        Exit Sub
    End If

    Dim accountNumbers As Collection
    Set accountNumbers = ReadAccountNumbers(wb)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(SHEET_TX)
    Dim changed As Long
    Dim t As clsTransaction
    For Each t In transactions
        Dim category As String
        If accountNumbers.Count > 0 And TextMentionsAccount(t.Text, accountNumbers) Then
            category = IGNORED_CATEGORY
        Else
            category = ruleset.Categorise(t.Text)
            If category = DEFAULT_CATEGORY And t.Category <> "" Then category = t.Category
        End If
        If category <> t.Category Then
            Dim catCol As Long
            catCol = IIf(t.Amount >= 0, COL_INC_CAT, COL_EXP_CAT)
            ws.Cells(t.SourceRow, catCol).Value = category
            t.Category = category
            changed = changed + 1
        End If
    Next t
    ReapplyIgnoredStyling ws, transactions

    Dim summary As clsSummary
    Set summary = RebuildSummaries(wb, transactions)
    SaveRulesQuietly ruleset

    Msg DK("Budgettet er opdateret.") & vbLf & vbLf & changed & " postering(er) fik ny kategori." _
       & vbLf & summary.Months.Count & " m" & DK("~aa~ned(er) i budgettet.")
End Sub

Private Sub ReapplyIgnoredStyling(ByVal ws As Worksheet, ByVal transactions As Collection)
    Dim pen As clsPen
    Set pen = NewPen(ws)
    Dim t As clsTransaction
    For Each t In transactions
        Dim firstCol As Long
        firstCol = IIf(t.Amount >= 0, COL_INC_DATE, COL_EXP_DATE)
        If t.Category = IGNORED_CATEGORY Then
            pen.Style RangeRef(firstCol, t.SourceRow, firstCol + 3, t.SourceRow), color:=MUTED, italic:=True
        Else
            pen.Style RangeRef(firstCol, t.SourceRow, firstCol, t.SourceRow), color:=MUTED, italic:=False
            pen.Style RangeRef(firstCol + 1, t.SourceRow, firstCol + 3, t.SourceRow), color:=TEXT_GREY, italic:=False
        End If
    Next t
End Sub

' -- category rules -------------------------------------------------------

Public Sub ShowRules()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If
    wb.Worksheets(SHEET_RULES).Activate
End Sub

Public Sub ResetRules()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If
    If MsgBox(DK("Nulstil alle kategoriregler til standarden? Dine egne regler g~aa~r tabt."), _
             vbYesNo + vbQuestion, APP_TITLE) <> vbYes Then Exit Sub

    Dim accounts As Object
    Set accounts = ReadCategoryAccounts(wb)
    Dim ruleset As New clsRuleSet
    ruleset.InitDefault
    Dim ws As Worksheet
    Set ws = EnsureSheet(wb, SHEET_RULES)
    WriteRules ws, ruleset, accounts
    SaveRulesQuietly ruleset
    Msg DK("Kategoriregler er nulstillet til standarden.")
End Sub

' -- accounts -------------------------------------------------------------

Public Sub ShowAccounts()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If
    EnsureAccountsSheet wb
    wb.Worksheets(SHEET_ACCOUNTS).Activate
End Sub

' -- bank budget snapshot ---------------------------------------------------

Public Sub CreateBankBudget()
    Dim wb As Workbook
    Set wb = ActiveWorkbookOrNothing()
    If wb Is Nothing Or Not HasBudget(wb) Then
        Msg DK("~AA~bn budgettet f~oe~rst.")
        Exit Sub
    End If
    Dim existing As Worksheet
    On Error Resume Next
    Set existing = wb.Worksheets(SHEET_PLAN)
    On Error GoTo 0
    If existing Is Nothing Then
        Msg DK("Der er intet budgetforslag endnu - import~ee~r en fil med mere end " & _
              DK("~ee~n m~aa~ned f~oe~rst."))
        Exit Sub
    End If
    Dim targets As Object
    Set targets = ReadTargets(wb)
    If targets.Count = 0 Then
        Msg "Fandt ingen m" & DK("~aa~l i """) & SHEET_PLAN & """."
        Exit Sub
    End If
    Dim ws As Worksheet
    Set ws = EnsureSheet(wb, SHEET_BANK)
    WriteBankBudget ws, targets
    ws.Activate
    Msg DK("Bankbudget er oprettet.")
End Sub

Public Sub ShowAbout()
    Msg "Budget fra CSV" & vbLf & DK("En Excel-udgave af LibreOffice-udvidelsen med samme navn.") & _
       vbLf & vbLf & DK("Import~ee~r en bank-CSV, kategoris~ee~r automatisk, og f~aa~ et " & _
       "udkast til et fast m~aa~nedsbudget.")
End Sub

' -----------------------------------------------------------------------
' Orchestration: building and rebuilding the whole workbook.
' -----------------------------------------------------------------------

Public Function BuildWorkbook(ByVal wb As Workbook, ByVal result As clsBuildResult, _
                              ByVal ruleset As clsRuleSet, Optional ByVal startBalance As Double = 0) As clsSummary
    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    Dim placeholderName As String
    placeholderName = "__tmp__"
    Dim i As Long
    For i = wb.Worksheets.Count To 2 Step -1
        Application.DisplayAlerts = False
        wb.Worksheets(i).Delete
        Application.DisplayAlerts = True
    Next i
    wb.Worksheets(1).Name = placeholderName

    Dim txWs As Worksheet
    Set txWs = EnsureSheet(wb, SHEET_TX)
    WriteTransactions txWs, result.Transactions

    Dim rulesWs As Worksheet
    Set rulesWs = EnsureSheet(wb, SHEET_RULES)
    WriteRules rulesWs, ruleset

    EnsureAccountsSheet wb

    Dim summary As New clsSummary
    summary.Init result.Transactions

    If summary.Months.Count > 1 Then
        Dim plan As clsBudgetPlan
        Set plan = BuildPlan(summary, result.Transactions)
        Dim planWs As Worksheet
        Set planWs = EnsureSheet(wb, SHEET_PLAN)
        Dim planRows As Object
        Set planRows = WritePlan(planWs, plan)

        Dim accountsTotal As Double
        accountsTotal = modSheetWriter.AccountsTotal(wb)
        Dim fallback As Double
        If accountsTotal <> 0 Then
            fallback = accountsTotal
        Else
            fallback = startBalance + summary.NetTotal()
        End If
        Dim forecastWs As Worksheet
        Set forecastWs = EnsureSheet(wb, SHEET_FORECAST)
        WriteForecast forecastWs, wb, planRows, LastMonthKey(summary), fallback

        Dim monthsWs As Worksheet
        Set monthsWs = EnsureSheet(wb, SHEET_MONTHS())
        WriteMonths monthsWs, summary
    End If

    OrderSheets wb
    Application.DisplayAlerts = False
    wb.Worksheets(placeholderName).Delete
    Application.DisplayAlerts = True

    ActivateFirstSheet wb
    Application.Calculate
    Set BuildWorkbook = summary
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

Private Function LastMonthKey(ByVal summary As clsSummary) As String
    Dim m As Variant, last2 As String
    For Each m In summary.Months
        last2 = CStr(m)
    Next m
    LastMonthKey = last2
End Function

Private Sub OrderSheets(ByVal wb As Workbook)
    Dim order As Variant
    order = Array(SHEET_PLAN, SHEET_FORECAST, SHEET_TX, SHEET_MONTHS(), SHEET_RULES, _
                  SHEET_ACCOUNTS, SHEET_BANK)
    Dim pos As Long
    pos = 1
    Dim n As Variant
    For Each n In order
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = Nothing
        Set ws = wb.Worksheets(CStr(n))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Move Before:=wb.Worksheets(pos)
            pos = pos + 1
        End If
    Next n
End Sub

Private Sub ActivateFirstSheet(ByVal wb As Workbook)
    Dim planWs As Worksheet
    On Error Resume Next
    Set planWs = wb.Worksheets(SHEET_PLAN)
    On Error GoTo 0
    If Not planWs Is Nothing Then
        planWs.Activate
    Else
        wb.Worksheets(SHEET_TX).Activate
    End If
End Sub

' Recreate Budgetforslag/Prognose/Alle maaneder from a (possibly changed)
' transaction list, preserving whatever the user has typed into the Maal/
' Konto columns. Used by both "Opdater" and "Tilfoej flere posteringer".
Public Function RebuildSummaries(ByVal wb As Workbook, ByVal transactions As Collection) As clsSummary
    Dim summary As New clsSummary
    summary.Init transactions

    Dim targets As Object, planAccounts As Object, categoryAccounts As Object
    Set targets = ReadTargets(wb)
    Set planAccounts = ReadPlanAccounts(wb)
    Set categoryAccounts = ReadCategoryAccounts(wb)
    Dim forecastStart As Variant
    forecastStart = ReadForecastStart(wb)

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    Dim names As Variant
    names = Array(SHEET_PLAN, SHEET_FORECAST, SHEET_MONTHS())
    Dim n As Variant
    For Each n In names
        Dim existing As Worksheet
        On Error Resume Next
        Set existing = Nothing
        Set existing = wb.Worksheets(CStr(n))
        On Error GoTo 0
        If Not existing Is Nothing Then
            Application.DisplayAlerts = False
            existing.Delete
            Application.DisplayAlerts = True
        End If
    Next n

    EnsureAccountsSheet wb

    If summary.Months.Count > 1 Then
        Dim plan As clsBudgetPlan
        Set plan = BuildPlan(summary, transactions)
        Dim planWs As Worksheet
        Set planWs = EnsureSheet(wb, SHEET_PLAN)
        Dim planRows As Object
        Set planRows = WritePlan(planWs, plan, targets, categoryAccounts, planAccounts)

        Dim fallback As Double
        If Not IsNull(forecastStart) Then
            fallback = CDbl(forecastStart)
        Else
            Dim accountsTotal As Double
            accountsTotal = modSheetWriter.AccountsTotal(wb)
            If accountsTotal <> 0 Then fallback = accountsTotal Else fallback = summary.NetTotal()
        End If
        Dim forecastWs As Worksheet
        Set forecastWs = EnsureSheet(wb, SHEET_FORECAST)
        WriteForecast forecastWs, wb, planRows, LastMonthKey(summary), fallback

        Dim monthsWs As Worksheet
        Set monthsWs = EnsureSheet(wb, SHEET_MONTHS())
        WriteMonths monthsWs, summary
    End If

    OrderSheets wb
    Application.Calculate
    Set RebuildSummaries = summary
CleanUp:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Function

' -- append more transactions to an existing budget ------------------------

Public Sub AppendTransactions(ByVal wb As Workbook, ByVal newTransactions As Collection, _
                              ByRef added As Long, ByRef skipped As Long, ByRef truncated As Boolean, _
                              ByRef summaryOut As clsSummary)
    added = 0: skipped = 0: truncated = False
    Set summaryOut = Nothing

    Dim ws As Worksheet
    Set ws = wb.Worksheets(SHEET_TX)
    Dim existing As Collection
    Set existing = ReadTransactions(wb)

    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim t As clsTransaction
    For Each t In existing
        Dim k As String
        k = DedupKey(t)
        If seen.Exists(k) Then seen(k) = seen(k) + 1 Else seen(k) = 1
    Next t

    Dim toAdd As New Collection
    For Each t In newTransactions
        Dim k2 As String
        k2 = DedupKey(t)
        If seen.Exists(k2) Then
            If seen(k2) > 0 Then
                seen(k2) = seen(k2) - 1
                skipped = skipped + 1
                GoTo ContinueLoop
            End If
        End If
        toAdd.Add t
ContinueLoop:
    Next t

    If toAdd.Count = 0 Then Exit Sub

    Dim existingExpenses As New Collection, existingIncome As New Collection
    For Each t In existing
        If t.Amount < 0 Then existingExpenses.Add t Else existingIncome.Add t
    Next t
    Dim newExpenses As New Collection, newIncome As New Collection
    For Each t In toAdd
        If t.Amount < 0 Then newExpenses.Add t Else newIncome.Add t
    Next t

    Dim roomExp As Long, roomInc As Long
    roomExp = WorksheetFunction.Max(TX_MAX_ROW - TX_FIRST_ROW + 1 - existingExpenses.Count, 0)
    roomInc = WorksheetFunction.Max(TX_MAX_ROW - TX_FIRST_ROW + 1 - existingIncome.Count, 0)
    If newExpenses.Count > roomExp Then
        Set newExpenses = TakeFirst(newExpenses, roomExp)
        truncated = True
    End If
    If newIncome.Count > roomInc Then
        Set newIncome = TakeFirst(newIncome, roomInc)
        truncated = True
    End If

    Application.ScreenUpdating = False
    On Error GoTo CleanUp

    WriteBlock ws, newExpenses, COL_EXP_DATE, TX_FIRST_ROW + existingExpenses.Count
    WriteBlock ws, newIncome, COL_INC_DATE, TX_FIRST_ROW + existingIncome.Count

    Dim newExpCount As Long, newIncCount As Long
    newExpCount = existingExpenses.Count + newExpenses.Count
    newIncCount = existingIncome.Count + newIncome.Count
    Dim last2 As Long
    last2 = TX_FIRST_ROW + WorksheetFunction.Max(newExpCount, newIncCount, 1) - 1
    Dim pen As clsPen
    Set pen = NewPen(ws)
    StyleTxRange pen, last2

    Dim allTransactions As Collection
    Set allTransactions = ReadTransactions(wb)
    ReapplyIgnoredStyling ws, allTransactions

    AddCategoryDropdown ws, COL_EXP_CAT, last2
    AddCategoryDropdown ws, COL_INC_CAT, last2

    added = toAdd.Count
    Set summaryOut = RebuildSummaries(wb, allTransactions)
CleanUp:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then Err.Raise Err.Number, , Err.Description
End Sub

Private Function DedupKey(ByVal t As clsTransaction) As String
    DedupKey = Format(t.TxDate, "yyyy-mm-dd") & "|" & t.Text & "|" & Format(t.Amount, "0.00")
End Function

Private Function TakeFirst(ByVal col As Collection, ByVal n As Long) As Collection
    Dim out As New Collection, i As Long
    i = 0
    Dim v As Variant
    For Each v In col
        If i >= n Then Exit For
        out.Add v
        i = i + 1
    Next v
    Set TakeFirst = out
End Function
