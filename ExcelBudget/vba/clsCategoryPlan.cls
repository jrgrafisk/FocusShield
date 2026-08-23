VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCategoryPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' What one category costs per month, and what is suggested to budget.
' Port of budget_core.plan.CategoryPlan.

Public Kind As String
Public Category As String
Public GroupName As String          ' GROUP_FIXED / GROUP_VARIABLE / GROUP_PERIODIC
Public Values As Collection         ' Double, abs(monthly total), one per month
Public MonthsPresent As Long
Public total As Double
Public MeanAll As Double
Public MeanActive As Double
Public Median As Double
Public Low As Double
Public High As Double
Public Variation As Double
Public Suggestion As Double
Public Sample As Collection         ' clsTransaction, biggest first, up to SAMPLE_SIZE
Public TransactionCount As Long
Public Account As String

Public Sub Init(ByVal kind As String, ByVal category As String, ByVal rawValues As Collection, _
                Optional ByVal transactions As Collection)
    Kind = kind
    Category = category
    Account = ""

    Set Values = New Collection
    Dim v As Variant
    For Each v In rawValues
        Values.Add Abs(CDbl(v))
    Next v

    Dim active As New Collection
    For Each v In Values
        If v > 0.005 Then active.Add v
    Next v

    Dim monthsCount As Long
    monthsCount = Values.Count
    If monthsCount = 0 Then monthsCount = 1

    MonthsPresent = active.Count
    total = SumCollection(Values)
    MeanAll = total / CDbl(monthsCount)
    If active.Count > 0 Then
        MeanActive = SumCollection(active) / CDbl(active.Count)
    Else
        MeanActive = 0
    End If
    Median = MedianOf(active)
    Low = MinOf(active)
    High = MaxOf(active)
    If MeanActive <> 0 Then
        Variation = StdevOf(active) / MeanActive
    Else
        Variation = 0
    End If
    GroupName = GROUP_VARIABLE
    Suggestion = 0

    Set Sample = New Collection
    TransactionCount = 0
    If Not (transactions Is Nothing) Then
        TransactionCount = transactions.Count
        Set Sample = TopByAbsAmount(transactions, SAMPLE_SIZE)
    End If
End Sub

Public Function Coverage() As Double
    Dim d As Long
    d = Values.Count
    If d = 0 Then d = 1
    Coverage = MonthsPresent / CDbl(d)
End Function

Private Function SumCollection(ByVal col As Collection) As Double
    Dim v As Variant, s As Double
    For Each v In col
        s = s + v
    Next v
    SumCollection = s
End Function

Private Function MinOf(ByVal col As Collection) As Double
    If col.Count = 0 Then
        MinOf = 0
        Exit Function
    End If
    Dim v As Variant, m As Double, first2 As Boolean
    first2 = True
    For Each v In col
        If first2 Then m = v: first2 = False
        If v < m Then m = v
    Next v
    MinOf = m
End Function

Private Function MaxOf(ByVal col As Collection) As Double
    If col.Count = 0 Then
        MaxOf = 0
        Exit Function
    End If
    Dim v As Variant, m As Double, first2 As Boolean
    first2 = True
    For Each v In col
        If first2 Then m = v: first2 = False
        If v > m Then m = v
    Next v
    MaxOf = m
End Function

Private Function MedianOf(ByVal col As Collection) As Double
    Dim n As Long
    n = col.Count
    If n = 0 Then
        MedianOf = 0
        Exit Function
    End If
    Dim arr() As Double, i As Long, j As Long
    ReDim arr(1 To n)
    i = 0
    Dim v As Variant
    For Each v In col
        i = i + 1
        arr(i) = v
    Next v
    For i = 2 To n
        Dim keyV As Double
        keyV = arr(i)
        j = i - 1
        Do While j >= 1
            If arr(j) > keyV Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = keyV
    Next i
    Dim mid As Long
    mid = n \ 2
    If n Mod 2 = 1 Then
        MedianOf = arr(mid + 1)
    Else
        MedianOf = (arr(mid) + arr(mid + 1)) / 2#
    End If
End Function

Private Function StdevOf(ByVal col As Collection) As Double
    Dim n As Long
    n = col.Count
    If n < 2 Then
        StdevOf = 0
        Exit Function
    End If
    Dim mean As Double
    mean = SumCollection(col) / CDbl(n)
    Dim v As Variant, variance As Double
    For Each v In col
        variance = variance + (v - mean) * (v - mean)
    Next v
    variance = variance / CDbl(n - 1)
    StdevOf = Sqr(variance)
End Function

Private Function TopByAbsAmount(ByVal transactions As Collection, ByVal count As Long) As Collection
    Dim n As Long, arr() As Object, i As Long, j As Long
    n = transactions.Count
    Set TopByAbsAmount = New Collection
    If n = 0 Then Exit Function
    ReDim arr(1 To n)
    i = 0
    Dim t As clsTransaction
    For Each t In transactions
        i = i + 1
        Set arr(i) = t
    Next t
    For i = 2 To n
        Dim keyT As clsTransaction
        Set keyT = arr(i)
        j = i - 1
        Do While j >= 1
            If Abs(arr(j).Amount) < Abs(keyT.Amount) Then
                Set arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        Set arr(j + 1) = keyT
    Next i
    Dim out As New Collection, limit As Long
    limit = WorksheetFunction.Min(count, n)
    For i = 1 To limit
        out.Add arr(i)
    Next i
    Set TopByAbsAmount = out
End Function
