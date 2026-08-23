Attribute VB_Name = "modPlan"
Option Explicit
' Turn a year of transactions into a draft budget. Port of budget_core/plan.py.

Public Const SAMPLE_SIZE As Long = 8

Public Const GROUP_FIXED As String = "fast"
Public Const GROUP_VARIABLE As String = "variabel"
Public Const GROUP_PERIODIC As String = "periodisk"

Public Const MONTHLY_COVERAGE As Double = 0.9
Public Const FIXED_COVERAGE As Double = 0.7
Public Const FIXED_VARIATION As Double = 0.3
Public Const STEADY_VARIATION As Double = 0.1
Public Const PERIODIC_COVERAGE As Double = 0.4
Public Const UNCERTAIN_MONTHS As Long = 6

Public Function GroupLabel(ByVal groupName As String) As String
    Select Case groupName
        Case GROUP_FIXED: GroupLabel = "FASTE UDGIFTER"
        Case GROUP_VARIABLE: GroupLabel = "VARIABLE UDGIFTER"
        Case GROUP_PERIODIC: GroupLabel = "PERIODISKE UDGIFTER"
        Case Else: GroupLabel = ""
    End Select
End Function

Public Function GroupHelp(ByVal groupName As String) As String
    Select Case groupName
        Case GROUP_FIXED
            GroupHelp = DK("Samme bel~oe~b hver m~aa~ned - budgettet er gennemsnittet.")
        Case GROUP_VARIABLE
            GroupHelp = DK("Svinger fra m~aa~ned til m~aa~ned - forslaget er en typisk " & _
                          "m~aa~ned (median), ikke gennemsnittet.")
        Case GROUP_PERIODIC
            GroupHelp = DK("Kommer f~aa~ gange om ~aa~ret - her henlagt pr. m~aa~ned, s~aa~ " & _
                          "~aa~rets samlede bel~oe~b er delt ligeligt ud.")
        Case Else
            GroupHelp = ""
    End Select
End Function

Private Function FixedByNature() As Object
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        Dim words As Variant, i As Long
        words = Array("bolig", DK("el, vand og varme"), DK("forsikring og pension"), _
                      DK("telefon og internet"), "abonnementer", DK("l~ae~n og afdrag"), _
                      DK("b~oe~rn og uddannelse"), "gebyrer og renter")
        For i = LBound(words) To UBound(words)
            dict(words(i)) = True
        Next i
    End If
    Set FixedByNature = dict
End Function

Private Function VariableByNature() As Object
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        Dim words As Variant, i As Long
        words = Array("dagligvarer", "restaurant", "shopping", "transport", _
                      DK("fritid og sport"), "sundhed", "rejser", "mobilepay", _
                      DK("overf~oe~rsel"), "ukategoriseret")
        For i = LBound(words) To UBound(words)
            dict(words(i)) = True
        Next i
    End If
    Set VariableByNature = dict
End Function

Private Function RoundTo(ByVal value As Double, ByVal step As Long, Optional ByVal mode As String = "nearest") As Double
    If step <= 0 Or value = 0 Then
        RoundTo = Int(value * 100 + IIf(value >= 0, 0.5, -0.5)) / 100#
        Exit Function
    End If
    Dim quotient As Double
    quotient = value / CDbl(step)
    Select Case mode
        Case "up"
            quotient = -Int(-quotient)          ' ceiling
        Case "down"
            quotient = Int(quotient)            ' floor
        Case Else
            quotient = Int(quotient + 0.5)       ' floor(quotient + 0.5)
    End Select
    RoundTo = quotient * step
End Function

Private Function Classify(ByVal plan As clsCategoryPlan) As String
    If plan.Kind = KIND_INCOME() Then
        Classify = GROUP_FIXED
        Exit Function
    End If
    Dim name As String
    name = FoldText(plan.Category)
    If FixedByNature().Exists(name) Then
        If plan.Coverage() >= MONTHLY_COVERAGE Then
            Classify = GROUP_FIXED
        Else
            Classify = GROUP_PERIODIC
        End If
        Exit Function
    End If
    If plan.Coverage() <= PERIODIC_COVERAGE Then
        Classify = GROUP_PERIODIC
        Exit Function
    End If
    Dim steady As Boolean
    steady = (plan.Coverage() >= FIXED_COVERAGE And plan.Variation <= FIXED_VARIATION)
    If steady And Not (VariableByNature().Exists(name) And plan.Variation > STEADY_VARIATION) Then
        Classify = GROUP_FIXED
        Exit Function
    End If
    Classify = GROUP_VARIABLE
End Function

Private Function Suggest(ByVal plan As clsCategoryPlan) As Double
    If plan.Kind = KIND_INCOME() Then
        Dim base As Double
        If plan.Coverage() >= FIXED_COVERAGE Then base = plan.Median Else base = plan.MeanAll
        Suggest = RoundTo(base, 100, "down")
        Exit Function
    End If
    If plan.GroupName = GROUP_FIXED Then
        Suggest = RoundTo(plan.MeanActive, 10, "up")
        Exit Function
    End If
    If plan.GroupName = GROUP_PERIODIC Then
        Suggest = RoundTo(plan.MeanAll, 50, "up")
        Exit Function
    End If
    Dim baseVariable As Double
    baseVariable = plan.Median
    If baseVariable = 0 Then baseVariable = plan.MeanActive
    Suggest = RoundTo(baseVariable, 50)
End Function

Public Function BuildPlan(ByVal summary As clsSummary, Optional ByVal transactions As Collection, _
                          Optional ByVal months As Collection) As clsBudgetPlan
    Dim useMonths As Collection
    If months Is Nothing Then
        Set useMonths = summary.Months
    Else
        Set useMonths = months
    End If

    ' Group transactions by "kind|category" for the sample.
    Dim byCategory As Object
    Set byCategory = CreateObject("Scripting.Dictionary")
    If Not (transactions Is Nothing) Then
        Dim monthSet As Object
        Set monthSet = CreateObject("Scripting.Dictionary")
        Dim m As Variant
        For Each m In useMonths
            monthSet(CStr(m)) = True
        Next m
        Dim t As clsTransaction
        For Each t In transactions
            If monthSet.Exists(t.MonthKeyStr) Then
                Dim key As String
                key = t.Kind & "|" & t.Category
                If Not byCategory.Exists(key) Then Set byCategory(key) = New Collection
                byCategory(key).Add t
            End If
        Next t
    End If

    Dim categories As New Collection
    Dim kindsAndNames As Collection
    Set kindsAndNames = New Collection
    kindsAndNames.Add Array(KIND_INCOME(), summary.IncomeCategories)
    kindsAndNames.Add Array(KIND_EXPENSE, summary.ExpenseCategories)

    Dim pair As Variant
    For Each pair In kindsAndNames
        Dim kind As String, names As Collection
        kind = pair(0)
        Set names = pair(1)
        Dim nm As Variant
        For Each nm In names
            Dim values As New Collection
            For Each m In useMonths
                values.Add summary.Value(kind, CStr(nm), CStr(m))
            Next m
            Dim key2 As String
            key2 = kind & "|" & nm
            Dim plan As New clsCategoryPlan
            If byCategory.Exists(key2) Then
                plan.Init kind, CStr(nm), values, byCategory(key2)
            Else
                plan.Init kind, CStr(nm), values
            End If
            plan.GroupName = Classify(plan)
            plan.Suggestion = Suggest(plan)
            categories.Add plan
        Next nm
    Next pair

    Set categories = SortByMeanDesc(categories)

    Dim result As New clsBudgetPlan
    result.Init useMonths, categories
    Set BuildPlan = result
End Function

Private Function SortByMeanDesc(ByVal col As Collection) As Collection
    Dim n As Long, arr() As Object, i As Long, j As Long
    n = col.Count
    Set SortByMeanDesc = New Collection
    If n = 0 Then Exit Function
    ReDim arr(1 To n)
    i = 0
    Dim c As clsCategoryPlan
    For Each c In col
        i = i + 1
        Set arr(i) = c
    Next c
    For i = 2 To n
        Dim keyC As clsCategoryPlan
        Set keyC = arr(i)
        j = i - 1
        Do While j >= 1
            If PlanLess(keyC, arr(j)) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = keyC
    Next i
    Dim out As New Collection
    For i = 1 To n
        out.Add arr(i)
    Next i
    Set SortByMeanDesc = out
End Function

' True when "a" sorts before "b": bigger MeanAll first, then category name.
Private Function PlanLess(ByVal a As clsCategoryPlan, ByVal b As clsCategoryPlan) As Boolean
    If a.MeanAll <> b.MeanAll Then
        PlanLess = (a.MeanAll > b.MeanAll)
    Else
        PlanLess = (a.Category < b.Category)
    End If
End Function
