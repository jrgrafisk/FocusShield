VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Monthly totals per category - the actual budget.
' Port of budget_core.budget.Summary.
'
' Cells are keyed by "kind|category|month" strings in a Dictionary since
' VBA has no tuple keys.

Public Months As Collection            ' sorted String, "2024-05"
Public ExpenseCategories As Collection
Public IncomeCategories As Collection
Public TransactionCount As Long
Public IgnoredCount As Long

Private mCells As Object               ' key -> Double
Private mCounts As Object              ' "kind|category" -> Long

Private Sub Class_Initialize()
    Set mCells = CreateObject("Scripting.Dictionary")
    Set mCounts = CreateObject("Scripting.Dictionary")
    Set Months = New Collection
    Set ExpenseCategories = New Collection
    Set IncomeCategories = New Collection
End Sub

Private Function CellKey(ByVal kind As String, ByVal category As String, ByVal monthKey As String) As String
    CellKey = kind & "|" & category & "|" & monthKey
End Function

Public Sub Init(ByVal transactions As Collection)
    Dim monthsSeen As Object
    Set monthsSeen = CreateObject("Scripting.Dictionary")
    Dim t As clsTransaction, counted As New Collection
    For Each t In transactions
        If t.Category <> IGNORED_CATEGORY Then
            counted.Add t
            If Not monthsSeen.Exists(t.MonthKeyStr) Then monthsSeen(t.MonthKeyStr) = True
        End If
    Next t

    Set Months = SortedKeys(monthsSeen)
    TransactionCount = counted.Count
    IgnoredCount = transactions.Count - counted.Count

    For Each t In counted
        Dim key As String, countKey As String
        key = CellKey(t.Kind, t.Category, t.MonthKeyStr)
        If mCells.Exists(key) Then
            mCells(key) = mCells(key) + t.Amount
        Else
            mCells(key) = t.Amount
        End If
        countKey = t.Kind & "|" & t.Category
        If mCounts.Exists(countKey) Then
            mCounts(countKey) = mCounts(countKey) + 1
        Else
            mCounts(countKey) = 1
        End If
    Next t

    Set ExpenseCategories = CategoriesFor(KIND_EXPENSE)
    Set IncomeCategories = CategoriesFor(KIND_INCOME)
End Sub

Private Function SortedKeys(ByVal dict As Object) As Collection
    Dim n As Long, arr() As String, i As Long, j As Long, tmp As String
    n = dict.Count
    Set SortedKeys = New Collection
    If n = 0 Then Exit Function
    ReDim arr(1 To n)
    Dim k As Variant
    i = 0
    For Each k In dict.Keys
        i = i + 1
        arr(i) = k
    Next k
    For i = 2 To n
        tmp = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) > tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i
    For i = 1 To n
        SortedKeys.Add arr(i)
    Next i
End Function

Private Function CategoriesFor(ByVal kind As String) As Collection
    Dim totals As Object
    Set totals = CreateObject("Scripting.Dictionary")
    Dim key As Variant, parts() As String
    For Each key In mCells.Keys
        parts = Split(CStr(key), "|")
        If parts(0) = kind Then
            If totals.Exists(parts(1)) Then
                totals(parts(1)) = totals(parts(1)) + mCells(key)
            Else
                totals(parts(1)) = mCells(key)
            End If
        End If
    Next key

    ' Sort by descending |total|, category name as tiebreak (ascending).
    Dim n As Long, names() As String, i As Long, j As Long
    n = totals.Count
    If n = 0 Then
        Set CategoriesFor = New Collection
        Exit Function
    End If
    ReDim names(1 To n)
    i = 0
    Dim kk As Variant
    For Each kk In totals.Keys
        i = i + 1
        names(i) = kk
    Next kk
    For i = 2 To n
        Dim keyName As String
        keyName = names(i)
        j = i - 1
        Do While j >= 1 And IsHigherPriority(keyName, names(j), totals)
            names(j + 1) = names(j)
            j = j - 1
        Loop
        names(j + 1) = keyName
    Next i
    Dim out As New Collection
    For i = 1 To n
        out.Add names(i)
    Next i
    Set CategoriesFor = out
End Function

' True when "a" should sort BEFORE "b" (a has priority): bigger |total|
' first, then alphabetical.
Private Function IsHigherPriority(ByVal a As String, ByVal b As String, ByVal totals As Object) As Boolean
    Dim absA As Double, absB As Double
    absA = Abs(totals(a))
    absB = Abs(totals(b))
    If absA <> absB Then
        IsHigherPriority = (absA > absB)
    Else
        IsHigherPriority = (a < b)
    End If
End Function

Public Function Value(ByVal kind As String, ByVal category As String, ByVal monthKey As String) As Double
    Dim key As String
    key = CellKey(kind, category, monthKey)
    If mCells.Exists(key) Then
        Value = mCells(key)
    Else
        Value = 0
    End If
End Function

Public Function CategoryTotal(ByVal kind As String, ByVal category As String) As Double
    Dim total As Double, m As Variant
    For Each m In Months
        total = total + Value(kind, category, CStr(m))
    Next m
    CategoryTotal = total
End Function

Public Function MonthTotal(ByVal kind As String, ByVal monthKey As String) As Double
    Dim cats As Collection, c As Variant, total As Double
    If kind = KIND_EXPENSE Then
        Set cats = ExpenseCategories
    Else
        Set cats = IncomeCategories
    End If
    For Each c In cats
        total = total + Value(kind, CStr(c), monthKey)
    Next c
    MonthTotal = total
End Function

Public Function total(ByVal kind As String) As Double
    Dim key As Variant, sum As Double, parts() As String
    For Each key In mCells.Keys
        parts = Split(CStr(key), "|")
        If parts(0) = kind Then sum = sum + mCells(key)
    Next key
    total = sum
End Function

Public Function NetTotal() As Double
    NetTotal = total(KIND_INCOME) + total(KIND_EXPENSE)
End Function

Public Function MonthCount() As Long
    MonthCount = WorksheetFunction.Max(1, Months.Count)
End Function
