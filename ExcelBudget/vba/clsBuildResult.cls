VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBuildResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Port of budget_core.budget.BuildResult.

Public Transactions As Collection      ' clsTransaction
Public Warnings As Collection          ' String
Public SkippedNoDate As Long
Public SkippedNoAmount As Long
Public Uncategorised As Collection     ' Array(name, count, total, keyword)

Private Sub Class_Initialize()
    Set Transactions = New Collection
    Set Warnings = New Collection
    Set Uncategorised = New Collection
End Sub

Public Function Ok() As Boolean
    Ok = (Transactions.Count > 0)
End Function

Public Function SummaryLines() As Collection
    Dim out As New Collection
    out.Add Transactions.Count & " posteringer indl" & DK("~ae~st.")
    If Transactions.Count > 0 Then
        Dim firstDate As Date, lastDate As Date, t As clsTransaction, first2 As Boolean
        first2 = True
        For Each t In Transactions
            If first2 Then
                firstDate = t.TxDate: lastDate = t.TxDate
                first2 = False
            Else
                If t.TxDate < firstDate Then firstDate = t.TxDate
                If t.TxDate > lastDate Then lastDate = t.TxDate
            End If
        Next t
        out.Add "Periode: " & Format(firstDate, "yyyy-mm-dd") & " - " & Format(lastDate, "yyyy-mm-dd")
    End If
    If SkippedNoDate > 0 Then
        out.Add SkippedNoDate & DK(" r~ae~kke(r) uden gyldig dato blev sprunget over.")
    End If
    If SkippedNoAmount > 0 Then
        out.Add SkippedNoAmount & DK(" r~ae~kke(r) uden gyldigt bel~oe~b blev sprunget over.")
    End If
    Dim uncategorisedCount As Long, ignoredCount As Long
    For Each t In Transactions
        If t.Category = DEFAULT_CATEGORY Then uncategorisedCount = uncategorisedCount + 1
        If t.Category = IGNORED_CATEGORY Then ignoredCount = ignoredCount + 1
    Next t
    If uncategorisedCount > 0 Then
        out.Add uncategorisedCount & " postering(er) er ikke kategoriseret."
    End If
    If ignoredCount > 0 Then
        out.Add ignoredCount & DK(" postering(er) er markeret som ignoreret og t~ae~ller ikke med i budgettet.")
    End If
    Dim w As Variant
    For Each w In Warnings
        out.Add w
    Next w
    Set SummaryLines = out
End Function
