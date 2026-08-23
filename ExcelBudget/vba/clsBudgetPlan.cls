VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBudgetPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A draft budget for a normal month. Port of budget_core.plan.BudgetPlan.

Public Months As Collection            ' String, sorted month keys
Public Categories As Collection        ' clsCategoryPlan
Public Income As Collection            ' clsCategoryPlan, kind = income
Public Expenses As Collection          ' clsCategoryPlan, kind = expense

Private mGroups As Object              ' groupName -> Collection(clsCategoryPlan)

Public Sub Init(ByVal months As Collection, ByVal categories As Collection)
    Set Me.Months = months
    Set Me.Categories = categories
    Set Income = New Collection
    Set Expenses = New Collection
    Set mGroups = CreateObject("Scripting.Dictionary")
    mGroups(GROUP_FIXED) = New Collection
    mGroups(GROUP_VARIABLE) = New Collection
    mGroups(GROUP_PERIODIC) = New Collection

    Dim c As clsCategoryPlan
    For Each c In categories
        If c.Kind = KIND_INCOME() Then
            Income.Add c
        Else
            Expenses.Add c
            If mGroups.Exists(c.GroupName) Then mGroups(c.GroupName).Add c
        End If
    Next c
End Sub

Public Function GroupCategories(ByVal groupName As String) As Collection
    If mGroups.Exists(groupName) Then
        Set GroupCategories = mGroups(groupName)
    Else
        Set GroupCategories = New Collection
    End If
End Function

Public Function MonthCount() As Long
    MonthCount = WorksheetFunction.Max(1, Months.Count)
End Function

Public Function IncomeMean() As Double
    Dim c As clsCategoryPlan, s As Double
    For Each c In Income
        s = s + c.MeanAll
    Next c
    IncomeMean = s
End Function

Public Function ExpenseMean() As Double
    Dim c As clsCategoryPlan, s As Double
    For Each c In Expenses
        s = s + c.MeanAll
    Next c
    ExpenseMean = s
End Function

Public Function SavingsMean() As Double
    SavingsMean = IncomeMean() - ExpenseMean()
End Function

Public Function IncomeSuggested() As Double
    Dim c As clsCategoryPlan, s As Double
    For Each c In Income
        s = s + c.Suggestion
    Next c
    IncomeSuggested = s
End Function

Public Function ExpenseSuggested() As Double
    Dim c As clsCategoryPlan, s As Double
    For Each c In Expenses
        s = s + c.Suggestion
    Next c
    ExpenseSuggested = s
End Function

Public Function SavingsSuggested() As Double
    SavingsSuggested = IncomeSuggested() - ExpenseSuggested()
End Function

Public Function IsUncertain() As Boolean
    IsUncertain = (Months.Count < UNCERTAIN_MONTHS)
End Function

Public Function GroupMean(ByVal groupName As String) As Double
    Dim c As clsCategoryPlan, s As Double
    For Each c In GroupCategories(groupName)
        s = s + c.MeanAll
    Next c
    GroupMean = s
End Function

Public Function SummaryLines() As Collection
    Dim out As New Collection
    out.Add DK("Gennemsnitlig indt~ae~gt: ") & Format(IncomeMean(), "0") & " kr./md"
    out.Add DK("Faste udgifter: ") & Format(GroupMean(GROUP_FIXED), "0") & " kr./md"
    out.Add DK("Variable udgifter: ") & Format(GroupMean(GROUP_VARIABLE), "0") & " kr./md"
    out.Add DK("Periodiske udgifter: ") & Format(GroupMean(GROUP_PERIODIC), "0") & " kr./md"
    out.Add DK("Til opsparing i gennemsnit: ") & Format(SavingsMean(), "0") & " kr./md"
    out.Add DK("Forslaget giver plads til ") & Format(SavingsSuggested(), "0") & DK(" kr./md i opsparing")
    If IsUncertain() Then
        out.Add DK("Bem~ae~rk: forslaget bygger kun p~aa~ ") & Months.Count & _
               DK(" m~aa~ned(er) - tallene er usikre.")
    End If
    Set SummaryLines = out
End Function
